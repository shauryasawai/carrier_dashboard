from __future__ import annotations

import functools
import json
import os
import re
from datetime import date, timedelta

from . import kpi

_TRUE_VALUES = ("1", "true", "yes", "on")


def _env_flag(name: str) -> bool:
    return os.environ.get(name, "0").strip().lower() in _TRUE_VALUES


# Hard upper bound on the lookback window. Configurable so memory/timeout-
# constrained hosts (e.g. Vercel, where each request re-queries because the
# in-memory cache doesn't persist) can clamp it small, e.g. BQ_MAX_LOOKBACK_DAYS=7.
def _max_lookback_days() -> int:
    try:
        return max(1, int(os.environ.get("BQ_MAX_LOOKBACK_DAYS")))
    except (TypeError, ValueError):
        return 366


MAX_LOOKBACK_DAYS = _max_lookback_days()
# Fallback window (days) used when BQ_LOOKBACK_DAYS is unset or invalid.
DEFAULT_LOOKBACK_DAYS = 30

DEFAULT_COLUMN_MAP = {
    "carrier": "courier_partner",
    "account": "account_code",
    "awb": "awb",                    # waybill
    "order_id": "order_id",          # order reference
    "weight": "shipment_weight",
    "payment": "payment_mode",
    "pickup_pin": "pickup_pincode",
    "drop_pin": "drop_pincode",
    "drop_city": "drop_city",
    "pickup_ts": "pickup_date",
    "order_ts": "order_date",          # (O2S processing time)
    "window_ts": "partition_date",     # load-window/partition date (= the day the
                                       # KPI cards count on); drives the per-day chart
    "delivery_ts": "delivery_date",
    "ofd1_ts": "out_for_delivery_1st_attempt",
    "edd_ts": "expected_delivery_date_by_courier_partner",
    "zone": "zone",
    "channel": "channel_name",
    "delivery_type": "shipment_type",
    "status": "clickpost_unified_status",
    "attempts": "out_for_delivery_attempts",
    "item_names": "items",
    "sku": "product_sku_code",         # product SKU code(s) for the shipment
    "order_value": "invoice_value",    # declared order/invoice value (revenue, ₹)
    "cod_value": "cod_value",           # cash-on-delivery collectable amount (₹)
}

_IDENT = re.compile(r"^[A-Za-z0-9_-]+$")


def _check_identifier(name: str, dotted: bool = False) -> None:
    parts = name.split(".") if dotted else [name]
    if not name or not all(_IDENT.match(p) for p in parts):
        raise ValueError(f"Unsafe BigQuery identifier: {name!r}")


def _column_map() -> dict:
    raw = os.environ.get("BQ_COLUMN_MAP")
    if raw:
        try:
            m = json.loads(raw)
            if isinstance(m, dict) and m:
                return m
        except ValueError:
            pass
    return DEFAULT_COLUMN_MAP


def _table_ref():
    return (os.environ.get("BQ_PROJECT"), os.environ.get("BQ_DATASET"),
            os.environ.get("BQ_TABLE"))


def _date_column() -> str:
    return os.environ.get("BQ_DATE_COLUMN")


def _date_is_string() -> bool:
    """Whether the lookback column is stored as a STRING rather than a DATE/
    TIMESTAMP. Set BQ_DATE_COLUMN_IS_STRING=1 when the column holds text like
    '2024-05-11' or '2024-05-11 10:30:00' (BigQuery can't compare STRING >= DATE
    directly, so we cast it; this disables partition pruning, so prefer a real
    DATE partition column when one is available)."""
    return _env_flag("BQ_DATE_COLUMN_IS_STRING")


def _probe_disabled() -> bool:
    """Whether to skip the get_table metadata probe in _partition_column.

    The probe costs one extra BigQuery API round-trip. On serverless hosts
    (Vercel) every request is a cold process, so the lru_cache below never
    survives between loads and that round-trip is paid on EVERY load. When the
    target table isn't partitioned (or you've set BQ_PARTITION_COLUMN explicitly),
    the probe can never help, so set BQ_DISABLE_PARTITION_PROBE=1 to skip it and
    fall straight through to the configured BQ_DATE_COLUMN predicate.
    """
    return _env_flag("BQ_DISABLE_PARTITION_PROBE")


@functools.lru_cache(maxsize=8)
def _partition_column(project: str, dataset: str, table: str) -> tuple:
    """Return (column_name, field_type) of the table's partition column, or
    (None, None). Filtering on this column lets BigQuery prune partitions —
    the single biggest speed/cost win for the lookback query, since it avoids
    scanning the whole ~1-year table.

    Resolution order:
      1. BQ_PARTITION_COLUMN env (skips the metadata lookup entirely). Its type
         can be hinted via BQ_PARTITION_COLUMN_TYPE (default DATE) so even the
         type lookup avoids a get_table call.
      2. The table's declared time/range partitioning, read once via get_table
         (cached for the process). Needs only bigquery.tables.get (included in
         the Data Viewer role); any failure falls back to (None, None).

    When BQ_DISABLE_PARTITION_PROBE is set, step 2 is skipped entirely (no API
    call): either the explicit column is returned, or (None, None) so the caller
    uses the configured BQ_DATE_COLUMN.
    """
    explicit = os.environ.get("BQ_PARTITION_COLUMN")

    # No-API fast paths: avoid the get_table round-trip when we can.
    if explicit:
        return (explicit, os.environ.get("BQ_PARTITION_COLUMN_TYPE", "DATE"))
    if _probe_disabled():
        return (None, None)

    try:
        tbl = _client().get_table(f"{project}.{dataset}.{table}")
    except Exception:  # noqa: BLE001 - metadata is best-effort; degrade gracefully
        return (None, None)

    tp = getattr(tbl, "time_partitioning", None)
    rp = getattr(tbl, "range_partitioning", None)
    field = (tp.field if tp else None) or (rp.field if rp else None)
    if field is None:
        # Ingestion-time partitioning exposes a TIMESTAMP pseudo-column.
        return ("_PARTITIONTIME", "TIMESTAMP") if tp is not None else (None, None)
    return (field, _column_type_lookup(tbl, field))


def _column_type_lookup(tbl, field):
    """Best-effort BigQuery field type (e.g. DATE/TIMESTAMP/DATETIME) for a
    column, or None if it can't be determined."""
    if tbl is None or not field:
        return None
    for col in tbl.schema:
        if col.name == field:
            return col.field_type
    return None


_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _date_literal(value: str) -> str:
    """Validate a YYYY-MM-DD string (from the UI date pickers) and return it as
    a BigQuery DATE literal. Rejects anything else so user input can never be
    interpolated as raw SQL."""
    s = (value or "").strip()
    if not _DATE_RE.match(s):
        raise ValueError(f"Invalid date (expected YYYY-MM-DD): {value!r}")
    return f"DATE '{s}'"


def _filter_column_sql(project, dataset, table) -> tuple:
    """Resolve the column to filter the lookback/range on.

    Returns (col_sql, kind) where col_sql is the SQL expression to compare
    (a backticked column, or a SAFE_CAST wrapper for STRING date columns ready
    to compare against a DATE) and kind is 'TIMESTAMP' | 'DATETIME' | 'DATE'.

    Prefers the partition column (prunes partitions); falls back to the
    configured BQ_DATE_COLUMN, casting it when it's a STRING. The cast path
    scans the whole table, so a real partition/date column is what keeps the
    query fast.
    """
    part_col, part_type = _partition_column(project, dataset, table)
    if part_col:
        _check_identifier(part_col)
        ftype = (part_type or "").upper()
        if ftype == "TIMESTAMP" or part_col == "_PARTITIONTIME":
            return f"`{part_col}`", "TIMESTAMP"
        if ftype == "DATETIME":
            return f"`{part_col}`", "DATETIME"
        if ftype in ("DATE", ""):  # DATE, or unknown -> assume DATE-compatible
            return f"`{part_col}`", "DATE"
        # Unhandled partition type (e.g. integer range): fall through to the
        # configured column rather than risk a type mismatch.

    date_col = _date_column()
    _check_identifier(date_col)
    if _date_is_string():
        return f"DATE(SAFE_CAST(`{date_col}` AS TIMESTAMP))", "DATE"
    return f"`{date_col}`", "DATE"


def _coerce_date(date_expr: str, kind: str) -> str:
    """Wrap a DATE-typed SQL expression so it can be compared against a column
    of the given kind (TIMESTAMP/DATETIME need an explicit constructor)."""
    if kind == "TIMESTAMP":
        return f"TIMESTAMP({date_expr})"
    if kind == "DATETIME":
        return f"DATETIME({date_expr})"
    return date_expr


# Filter on ORDER DATE (day the order was placed) so KPIs reconcile with
# order-date sales reports. order_date is an ISO string -> cast to DATE.
# partition_date (always >= order_date) is kept only to prune partitions cheaply.
_ORDER_DATE_SQL = "DATE(SAFE_CAST(`order_date` AS TIMESTAMP))"


def _prune(project, dataset, table, lo: str) -> str:
    """`AND partition_date >= lo - 3d` to skip partitions the window can't touch
    ('' when the table has no partition column)."""
    col, typ = _partition_column(project, dataset, table)
    if not col:
        return ""
    _check_identifier(col)
    lo_buf = f"DATE_SUB({lo}, INTERVAL 3 DAY)"
    kind = (typ or "").upper()
    if kind in ("TIMESTAMP", "DATETIME"):
        lo_buf = f"{kind}({lo_buf})"
    return f" AND `{col}` >= {lo_buf}"


def _lookback_predicate(project, dataset, table, days: int) -> str:
    """Orders placed in the last `days` days, by order_date."""
    lo = f"DATE_SUB(CURRENT_DATE(), INTERVAL {days} DAY)"
    return f"{_ORDER_DATE_SQL} >= {lo}{_prune(project, dataset, table, lo)}"


def _range_predicate(project, dataset, table, date_from, date_to) -> str:
    """Inclusive [date_from, date_to] order-date range (either bound optional)."""
    parts = []
    if date_from:
        lo = _date_literal(date_from)
        parts.append(f"{_ORDER_DATE_SQL} >= {lo}{_prune(project, dataset, table, lo)}")
    if date_to:
        parts.append(f"{_ORDER_DATE_SQL} <= {_date_literal(date_to)}")
    return " AND ".join(parts) if parts else "TRUE"


def default_lookback_days() -> int:
    return _clamp_lookback(os.environ.get("BQ_LOOKBACK_DAYS"), DEFAULT_LOOKBACK_DAYS)


def _clamp_lookback(value, fallback) -> int:
    try:
        n = int(value)
    except (TypeError, ValueError):
        return int(fallback)
    return max(1, min(MAX_LOOKBACK_DAYS, n))


def effective_lookback_days(requested) -> int:
    """The actual number of days pulled after clamping (matches the SQL filter,
    including the BQ_MAX_LOOKBACK_DAYS cap)."""
    return _clamp_lookback(requested, default_lookback_days())


def lookback_window(requested) -> dict:
    """The date window the query covers: partition_date >= today - N days.
    Returns ISO strings so the UI can show one authoritative date range."""
    days = effective_lookback_days(requested)
    today = date.today()
    return {
        "from": (today - timedelta(days=days)).isoformat(),
        "to": today.isoformat(),
        "days": days,
    }


def is_configured() -> bool:
    # The table has built-in defaults, so it's always resolvable.
    return all(_table_ref())


def build_query(lookback_days: int | None = None, limit: int | None = None,
                date_from: str | None = None, date_to: str | None = None) -> str:
    """Compose a safe SELECT aliasing each column to its logical field name.

    Filtered either to an explicit [date_from, date_to] calendar range (when
    either bound is given) or, otherwise, to the last `lookback_days` days.
    """
    project, dataset, table = _table_ref()
    _check_identifier(project, dotted=True)
    _check_identifier(dataset)
    _check_identifier(table)

    select_parts = []
    for logical, column in _column_map().items():
        _check_identifier(str(column))
        select_parts.append(f"`{column}` AS `{logical}`")

    if date_from or date_to:
        where = _range_predicate(project, dataset, table, date_from, date_to)
    else:
        days = _clamp_lookback(lookback_days, default_lookback_days())
        where = _lookback_predicate(project, dataset, table, days)
    sql = (
        f"SELECT {', '.join(select_parts)} "
        f"FROM `{project}.{dataset}.{table}` "
        f"WHERE {where}"
    )
    if limit:
        sql += f" LIMIT {int(limit)}"
    return sql


# OAuth scope. We deliberately do NOT use bigquery.readonly here: running a
# query through the Python client creates a job via the jobs.insert API, and
# jobs.insert rejects the read-only scope — so a read-only token would make
# every query fail with an insufficient-scope error. "View access only" is
# instead enforced by IAM: grant this service account ONLY the BigQuery Data
# Viewer + Job User roles (never Data Editor/Admin). With those roles it can
# run read queries but physically cannot create, modify, or delete data.
_SCOPES = ["https://www.googleapis.com/auth/cloud-platform"]

# Logical service-account field -> environment variable holding its value.
_SA_ENV_KEYS = {
    "type": "GOOGLE_SA_TYPE",
    "project_id": "GOOGLE_SA_PROJECT_ID",
    "private_key_id": "GOOGLE_SA_PRIVATE_KEY_ID",
    "private_key": "GOOGLE_SA_PRIVATE_KEY",
    "client_email": "GOOGLE_SA_CLIENT_EMAIL",
    "client_id": "GOOGLE_SA_CLIENT_ID",
    "auth_uri": "GOOGLE_SA_AUTH_URI",
    "token_uri": "GOOGLE_SA_TOKEN_URI",
    "auth_provider_x509_cert_url": "GOOGLE_SA_AUTH_PROVIDER_X509_CERT_URL",
    "client_x509_cert_url": "GOOGLE_SA_CLIENT_X509_CERT_URL",
    "universe_domain": "GOOGLE_SA_UNIVERSE_DOMAIN",
}

# Defaults for fields that are identical across every Google SA key, so only
# the secret-ish fields must actually be present in the environment.
_SA_DEFAULTS = {
    "type": "service_account",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "universe_domain": "googleapis.com",
}

# Fields with no useful default — required to build a valid credentials object.
_SA_REQUIRED = ("project_id", "private_key", "client_email")


def _service_account_info() -> dict | None:
    """Assemble the service-account credentials object (dict) from env vars.

    Returns None when the required fields aren't set, so the caller can fall
    back to other credential sources. The private key may contain literal
    "\\n" sequences (common when stored in a single-line env var); those are
    restored to real newlines so the PEM parses correctly.
    """
    info = dict(_SA_DEFAULTS)
    for field, env_name in _SA_ENV_KEYS.items():
        val = os.environ.get(env_name)
        if val:
            info[field] = val
    if not all(info.get(f) for f in _SA_REQUIRED):
        return None
    info["private_key"] = info["private_key"].replace("\\n", "\n")
    return info


@functools.lru_cache(maxsize=1)
def _client():
    """Return a cached BigQuery client (view-only access enforced via IAM, see
    _SCOPES).

    The client (and the credentials object it wraps, which parses the PEM
    private key) is built once per process and reused. fetch_records and the
    partition probe both call this, so without caching the relatively expensive
    credential/client construction happened multiple times per load. The
    underlying credentials auto-refresh their access token, so a long-lived
    client is safe. lru_cache(maxsize=1) gives a per-process singleton.

    Credentials are resolved in this order:
      1. A credentials object assembled from individual GOOGLE_SA_* env vars.
      2. Back-compat: full key JSON in GOOGLE_SERVICE_ACCOUNT_JSON / GCP_SA_KEY.
      3. Application Default Credentials — the Cloud Run runtime service
         account, or a GOOGLE_APPLICATION_CREDENTIALS key file locally.
    """
    from google.cloud import bigquery  # lazy import

    project, _, _ = _table_ref()

    info = _service_account_info()
    if info is None:
        raw_key = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON") or os.environ.get("GCP_SA_KEY")
        if raw_key:
            info = json.loads(raw_key)

    if info is not None:
        from google.oauth2 import service_account
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=_SCOPES
        )
        return bigquery.Client(
            project=project or info.get("project_id"), credentials=creds
        )

    return bigquery.Client(project=project)


def _records_from_arrow(query_job) -> list | None:
    """Fast path: download results columnar via the BigQuery Storage API + Arrow.

    This is far quicker than the default paginated REST row download for large
    result sets (90-day / 1-year windows). Returns None if the optional deps
    (pyarrow, google-cloud-bigquery-storage) aren't installed or anything goes
    wrong, so the caller can fall back to plain row iteration.
    """
    result = query_job.result()
    table = None
    # Tier 1 — BigQuery Storage Read API (true columnar streaming, the fastest).
    # Only works if the API is enabled for the service account; when it isn't,
    # to_arrow can return a silent 0-row table, so we only accept a non-empty
    # result here and otherwise fall through to the REST-based Arrow download.
    try:
        t = result.to_arrow(create_bqstorage_client=True)
        if t.num_rows > 0:
            table = t
    except Exception:  # noqa: BLE001 - Storage API/deps unavailable -> next tier
        table = None
    # Tier 2 — REST-based Arrow (needs only pyarrow, NOT the Storage API). Still
    # parses far faster than iterating one Python Row object per record. If
    # pyarrow isn't installed, to_arrow raises and we fall back to row iteration.
    if table is None:
        try:
            table = result.to_arrow(create_bqstorage_client=False)
        except Exception:  # noqa: BLE001 -> caller uses plain row iteration
            return None

    # A genuinely empty window returns [] (no double-query); None means the fast
    # path is unavailable so the caller iterates rows over REST instead.
    if table.num_rows == 0:
        return []

    # Materialize each aliased column once (bulk C->Python), then build records
    # by row index through a single reused getter (avoids allocating a closure
    # per row). Column names are the logical aliases, so they feed build_record.
    columns = {name: table.column(name).to_pylist() for name in table.schema.names}
    records = []
    append = records.append
    build = kpi.build_record
    _state = {"i": 0}

    def _get(key, _cols=columns, _s=_state):
        col = _cols.get(key)
        return col[_s["i"]] if col is not None else None

    for i in range(table.num_rows):
        _state["i"] = i
        rec = build(_get)
        if rec is not None:
            append(rec)
    return records


# Opt-in in-process cache of fetched records, keyed by (lookback_days, limit).
# The dashboard scans ~half a GB per load on an unpartitioned table, so a short
# TTL turns repeated "Load from BigQuery" clicks (and warm serverless re-hits)
# into instant, zero-cost responses. Disabled by default (TTL 0) so a load
# always returns fresh data unless BQ_CACHE_TTL_SECONDS is set.
_RESULT_CACHE: dict[tuple, tuple] = {}


def _cache_ttl() -> int:
    try:
        return max(0, int(os.environ.get("BQ_CACHE_TTL_SECONDS")))
    except (TypeError, ValueError):
        return 0


def _page_size() -> int | None:
    """Rows per page for the REST download. A larger page means fewer HTTP
    round-trips when streaming a big result set (the API still caps each page's
    byte size). None lets the client choose its default."""
    raw = os.environ.get("BQ_PAGE_SIZE")
    if raw and raw.isdigit() and int(raw) > 0:
        return int(raw)
    return 100_000


def fetch_awb_product_map(date_from: str, date_to: str,
                          limit: int | None = None,
                          awbs: set | None = None,
                          carriers: set | None = None) -> tuple[dict, dict, dict, dict, dict]:
    """Return (awb2prod, order2prod, awb2value, order2value, awb2order) for
    shipments in the [date_from, date_to] window (YYYY-MM-DD), filtered on the
    configured date/partition column. The *value maps hold each shipment's
    declared order value (invoice_value = product selling price); awb2order maps
    each AWB to its order id for order-level de-duplication.

    Each value is (category, subcategory, sku, item_name), with category /
    subcategory derived from the item names via kpi._product_category — the same
    rules the dashboard uses. AWBs are normalized (upper-cased, trailing ".0"
    stripped) so they line up with the invoice parser's keys. Only four columns
    are scanned (awb, product_sku_code, items, order_id) to keep cost low.

    Performance: the result is downloaded columnar via the BigQuery Storage API +
    Arrow (far faster than paginated REST for the large window), and when `awbs`
    (the set of normalized invoice AWBs) is given we only build/derive entries for
    those AWBs — so category derivation runs on the handful of billed shipments
    instead of every shipment in the window. Category lookups are memoized per
    distinct item-name string.

    This is the "search BigQuery by invoice date" path: the caller derives the
    window from the uploaded invoices' service month, so enrichment matches the
    billed shipments regardless of what lookback window is otherwise loaded.
    """
    project, dataset, table = _table_ref()
    _check_identifier(project, dotted=True)
    _check_identifier(dataset)
    _check_identifier(table)

    cmap = _column_map()
    awb_col = cmap.get("awb", "awb")
    sku_col = cmap.get("sku", "product_sku_code")
    items_col = cmap.get("item_names", "items")
    order_col = cmap.get("order_id", "order_id")
    val_col = cmap.get("order_value", "invoice_value")
    for c in (awb_col, sku_col, items_col, order_col, val_col):
        _check_identifier(str(c))

    where = _range_predicate(project, dataset, table, date_from, date_to)

    # Restrict to the invoice's carrier(s) so BigQuery scans/returns only those
    # shipments (e.g. BlueDart) instead of every courier in the window. Tokens
    # are sanitised to [a-z0-9] so they can't inject SQL; matched as a prefix on
    # the carrier column (so "BlueDart" also catches "Bluedart Reverse").
    carrier_pred = ""
    if carriers:
        car_col = cmap.get("carrier", "courier_partner")
        _check_identifier(str(car_col))
        toks = set()
        for c in carriers:
            tok = re.sub(r"[^a-z0-9]", "", re.split(r"[ _\-/]", str(c).strip().lower())[0])
            if tok:
                toks.add(tok)
        if toks:
            likes = " OR ".join(f"LOWER(`{car_col}`) LIKE '{t}%'" for t in sorted(toks))
            carrier_pred = f" AND ({likes})"

    # Push the invoice's AWB set INTO the query so BigQuery returns only those
    # waybills, not every shipment for the carrier across the whole window (the
    # bulk of the transfer/latency). The AWB column is normalised in SQL exactly
    # as it is in Python below — CAST -> TRIM -> UPPER -> strip trailing ".0" —
    # so the match is identical; the set is passed as a query PARAMETER (array)
    # to stay injection-safe and avoid a giant inline IN list. The Python-side
    # `want` filter is kept as a belt-and-braces guard. Skipped when the set is
    # empty or very large (fall back to the carrier-only scan + Python filter).
    want = awbs or None
    use_awb_pushdown = bool(want) and len(want) <= 45000
    awb_pred = ""
    if use_awb_pushdown:
        awb_pred = (f" AND REGEXP_REPLACE(UPPER(TRIM(CAST(`{awb_col}` AS STRING))), "
                    f"r'\\.0$', '') IN UNNEST(@wanted_awbs)")
    sql = (
        f"SELECT `{awb_col}` AS awb, `{sku_col}` AS sku, "
        f"`{items_col}` AS items, `{order_col}` AS order_id, "
        f"`{val_col}` AS order_value "
        f"FROM `{project}.{dataset}.{table}` "
        f"WHERE {where} AND `{awb_col}` IS NOT NULL{carrier_pred}{awb_pred}"
    )
    if limit:
        sql += f" LIMIT {int(limit)}"

    _norm_re = re.compile(r"\.0$")
    _cat_memo = {}

    def _cat(text):
        c = _cat_memo.get(text)
        if c is None:
            c = kpi._product_category(text)
            _cat_memo[text] = c
        return c

    client = _client()
    if use_awb_pushdown:
        from google.cloud import bigquery  # lazy import (matches _client)
        job_config = bigquery.QueryJobConfig(query_parameters=[
            bigquery.ArrayQueryParameter("wanted_awbs", "STRING", sorted(want))
        ])
        job = client.query(sql, job_config=job_config)
    else:
        job = client.query(sql)
    awb2prod, order2prod = {}, {}
    # AWB/order -> declared order value (invoice_value = the product's selling
    # price), so invoice shipping cost can be expressed as a % of item value.
    awb2value, order2value = {}, {}
    # AWB -> order id, so invoice lines (which often carry no order id of their
    # own, e.g. BlueDart) can be grouped by order to de-duplicate selling value.
    awb2order = {}

    def _emit(awb_raw, sku_raw, items_raw, order_raw, value_raw):
        if awb_raw is None:
            return
        awb = _norm_re.sub("", str(awb_raw).strip().upper())
        in_want = (want is None) or (awb in want)
        # Skip shipments not on the invoice entirely (no AWB match and, for the
        # AWB-filtered path, no order fallback needed) — saves all the per-row
        # string/category work for the ~70% of window rows we don't bill.
        if want is not None and not in_want:
            return
        items_text = str(items_raw or "").strip()
        sku = str(sku_raw or "").strip()
        cat, sub = _cat(items_text)
        prod = (cat, sub, sku, items_text)
        try:
            val = float(value_raw) if value_raw not in (None, "") else None
        except (TypeError, ValueError):
            val = None
        oid = str(order_raw or "").strip().upper()
        if awb and in_want:
            awb2prod[awb] = prod
            if val is not None:
                awb2value[awb] = val
            if oid:
                awb2order[awb] = oid
        if oid:
            order2prod.setdefault(oid, prod)
            if val is not None:
                order2value.setdefault(oid, val)

    # Consume the result columnar via REST-based Arrow (pyarrow only, NOT the
    # Storage API) — materialising 5 columns in bulk and iterating by index is
    # far faster than allocating one Python Row object per record. Falls back to
    # plain REST row iteration when pyarrow is unavailable. (We intentionally do
    # NOT use create_bqstorage_client=True here: when the Storage API isn't
    # enabled it can silently yield an empty table and drop all enrichment.)
    cols = None
    try:
        table = job.result().to_arrow(create_bqstorage_client=False)
        cols = {n: table.column(n).to_pylist() for n in
                ("awb", "sku", "items", "order_id", "order_value")}
    except Exception:  # noqa: BLE001 - pyarrow missing/incompatible -> row iteration
        cols = None
    if cols is not None:
        awb_c, sku_c = cols["awb"], cols["sku"]
        items_c, ord_c, val_c = cols["items"], cols["order_id"], cols["order_value"]
        for i in range(len(awb_c)):
            _emit(awb_c[i], sku_c[i], items_c[i], ord_c[i], val_c[i])
    else:
        for row in job.result(page_size=_page_size()):
            _emit(row.get("awb"), row.get("sku"), row.get("items"),
                  row.get("order_id"), row.get("order_value"))
    return awb2prod, order2prod, awb2value, order2value, awb2order


def fetch_records(lookback_days: int | None = None,
                  limit: int | None = None,
                  date_from: str | None = None,
                  date_to: str | None = None) -> list[dict]:
    """Query BigQuery and return normalized record dicts (same shape as CSV parse).

    Pass date_from/date_to (YYYY-MM-DD) for an explicit calendar range;
    otherwise the last `lookback_days` days are pulled.
    """
    if limit is None:
        env_cap = os.environ.get("BQ_MAX_ROWS")
        limit = int(env_cap) if env_cap and env_cap.isdigit() else None

    if date_from or date_to:
        cache_key = ("range", date_from, date_to, limit)
    else:
        cache_key = (_clamp_lookback(lookback_days, default_lookback_days()), limit)
    ttl = _cache_ttl()
    if ttl:
        import time
        hit = _RESULT_CACHE.get(cache_key)
        if hit is not None and (time.monotonic() - hit[0]) < ttl:
            return hit[1]

    client = _client()
    sql = build_query(lookback_days=lookback_days, limit=limit,
                      date_from=date_from, date_to=date_to)
    query_job = client.query(sql)

    fast = _records_from_arrow(query_job)
    if fast is not None:
        records = fast
    else:
        # Fallback: stream rows over the REST API. Build directly off each
        # BigQuery Row (which supports .get by alias) to avoid allocating an
        # intermediate dict per row, and use a large page size to cut the
        # number of HTTP round-trips for big windows.
        records = []
        for row in query_job.result(page_size=_page_size()):
            rec = kpi.build_record(lambda key, _r=row: _r.get(key))
            if rec is not None:
                records.append(rec)

    if ttl:
        import time
        _RESULT_CACHE[cache_key] = (time.monotonic(), records)
    return records
