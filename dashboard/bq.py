from __future__ import annotations

import functools
import json
import os
import re
from datetime import date, timedelta

from . import kpi

# --- Target table (override via env) ----------------------------------------


# Hard upper bound on the lookback window. Configurable so memory/timeout-
# constrained hosts (e.g. Vercel, where each request re-queries because the
# in-memory cache doesn't persist) can clamp it small, e.g. BQ_MAX_LOOKBACK_DAYS=7.
def _max_lookback_days() -> int:
    raw = os.environ.get("BQ_MAX_LOOKBACK_DAYS")
    try:
        return max(1, int(raw))
    except (TypeError, ValueError):
        return 366


MAX_LOOKBACK_DAYS = _max_lookback_days()
# Fallback window (days) used when BQ_LOOKBACK_DAYS is unset or invalid.
DEFAULT_LOOKBACK_DAYS = 30

DEFAULT_COLUMN_MAP = {
    "carrier": "courier_partner",
    "account": "account_code",       
    "weight": "shipment_weight",
    "payment": "payment_mode",
    "pickup_pin": "pickup_pincode",
    "drop_pin": "drop_pincode",
    "drop_city": "drop_city",
    "pickup_ts": "pickup_date",
    "delivery_ts": "delivery_date",
    "ofd1_ts": "out_for_delivery_1st_attempt",
    "zone": "zone",
    "delivery_type": "shipment_type",
    "status": "clickpost_unified_status",
    "attempts": "out_for_delivery_attempts",
    "item_names": "items",
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
    project = os.environ.get("BQ_PROJECT")
    dataset = os.environ.get("BQ_DATASET")
    table = os.environ.get("BQ_TABLE")
    return project, dataset, table


def _date_column() -> str:
    return os.environ.get("BQ_DATE_COLUMN")


def _date_is_string() -> bool:
    """Whether the lookback column is stored as a STRING rather than a DATE/
    TIMESTAMP. Set BQ_DATE_COLUMN_IS_STRING=1 when the column holds text like
    '2024-05-11' or '2024-05-11 10:30:00' (BigQuery can't compare STRING >= DATE
    directly, so we cast it; this disables partition pruning, so prefer a real
    DATE partition column when one is available)."""
    return os.environ.get("BQ_DATE_COLUMN_IS_STRING", "0").strip().lower() in (
        "1", "true", "yes", "on",
    )


def _probe_disabled() -> bool:
    """Whether to skip the get_table metadata probe in _partition_column.

    The probe costs one extra BigQuery API round-trip. On serverless hosts
    (Vercel) every request is a cold process, so the lru_cache below never
    survives between loads and that round-trip is paid on EVERY load. When the
    target table isn't partitioned (or you've set BQ_PARTITION_COLUMN explicitly),
    the probe can never help, so set BQ_DISABLE_PARTITION_PROBE=1 to skip it and
    fall straight through to the configured BQ_DATE_COLUMN predicate.
    """
    return os.environ.get("BQ_DISABLE_PARTITION_PROBE", "0").strip().lower() in (
        "1", "true", "yes", "on",
    )


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
        ptype = os.environ.get("BQ_PARTITION_COLUMN_TYPE", "DATE")
        return (explicit, ptype)
    if _probe_disabled():
        return (None, None)

    try:
        client = _client()
        tbl = client.get_table(f"{project}.{dataset}.{table}")
    except Exception:  # noqa: BLE001 - metadata is best-effort; degrade gracefully
        return (explicit, _column_type_lookup(None, explicit)) if explicit else (None, None)

    tp = getattr(tbl, "time_partitioning", None)
    rp = getattr(tbl, "range_partitioning", None)
    field = explicit or (tp.field if tp else None) or (rp.field if rp else None)
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


def _lookback_predicate(project, dataset, table, days: int) -> str:
    """SQL WHERE predicate constraining the query to the lookback window.

    Prefers the partition column (prunes partitions) with a type-correct
    comparison. Falls back to the configured BQ_DATE_COLUMN, casting it when
    it's a STRING (BQ_DATE_COLUMN_IS_STRING). The cast path scans the whole
    table, so configuring/auto-detecting a partition column is what makes the
    query fast.
    """
    cutoff = f"DATE_SUB(CURRENT_DATE(), INTERVAL {days} DAY)"
    part_col, part_type = _partition_column(project, dataset, table)
    if part_col:
        _check_identifier(part_col)
        ftype = (part_type or "").upper()
        if ftype == "TIMESTAMP" or part_col == "_PARTITIONTIME":
            return f"`{part_col}` >= TIMESTAMP({cutoff})"
        if ftype == "DATETIME":
            return f"`{part_col}` >= DATETIME({cutoff})"
        if ftype in ("DATE", ""):  # DATE, or unknown type -> assume DATE-compatible
            return f"`{part_col}` >= {cutoff}"
        # Unhandled partition type (e.g. integer range): fall through to the
        # configured column rather than risk a type mismatch.

    date_col = _date_column()
    _check_identifier(date_col)
    if _date_is_string():
        return f"DATE(SAFE_CAST(`{date_col}` AS TIMESTAMP)) >= {cutoff}"
    return f"`{date_col}` >= {cutoff}"


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


def build_query(lookback_days: int | None = None, limit: int | None = None) -> str:
    """Compose a safe SELECT aliasing each column to its logical field name,
    filtered to the lookback window on the partition column."""
    project, dataset, table = _table_ref()
    _check_identifier(project, dotted=True)
    _check_identifier(dataset)
    _check_identifier(table)

    cmap = _column_map()
    select_parts = []
    for logical, column in cmap.items():
        _check_identifier(str(column))
        select_parts.append(f"`{column}` AS `{logical}`")

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
    try:
        # create_bqstorage_client=True uses the Storage API when available and
        # transparently falls back to REST (still needs pyarrow for to_arrow).
        table = query_job.result().to_arrow(create_bqstorage_client=True)
    except Exception:  # noqa: BLE001 - any import/runtime issue -> use REST path
        return None

    # Materialize each aliased column once; build records by row index. Column
    # names are the logical aliases, so they feed kpi.build_record directly.
    columns = {name: table.column(name).to_pylist() for name in table.schema.names}
    records = []
    for i in range(table.num_rows):
        rec = kpi.build_record(lambda key, _i=i: _col_value(columns, key, _i))
        if rec is not None:
            records.append(rec)
    return records


def _col_value(columns: dict, key: str, i: int):
    col = columns.get(key)
    return col[i] if col is not None else None


# Opt-in in-process cache of fetched records, keyed by (lookback_days, limit).
# The dashboard scans ~half a GB per load on an unpartitioned table, so a short
# TTL turns repeated "Load from BigQuery" clicks (and warm serverless re-hits)
# into instant, zero-cost responses. Disabled by default (TTL 0) so a load
# always returns fresh data unless BQ_CACHE_TTL_SECONDS is set.
_RESULT_CACHE: dict[tuple, tuple] = {}


def _cache_ttl() -> int:
    raw = os.environ.get("BQ_CACHE_TTL_SECONDS")
    try:
        return max(0, int(raw))
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


def fetch_records(lookback_days: int | None = None,
                  limit: int | None = None) -> list[dict]:
    """Query BigQuery and return normalized record dicts (same shape as CSV parse)."""
    if limit is None:
        env_cap = os.environ.get("BQ_MAX_ROWS")
        limit = int(env_cap) if env_cap and env_cap.isdigit() else None

    days = _clamp_lookback(lookback_days, default_lookback_days())
    cache_key = (days, limit)
    ttl = _cache_ttl()
    if ttl:
        import time
        hit = _RESULT_CACHE.get(cache_key)
        if hit is not None and (time.monotonic() - hit[0]) < ttl:
            return hit[1]

    client = _client()
    sql = build_query(lookback_days=lookback_days, limit=limit)
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