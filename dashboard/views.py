import csv
import datetime
import io
import json
import logging
import os

from django.http import HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from . import auth, bq, invoices, tat
from .kpi import build_report, filter_records, parse_workbook, reclassify

logger = logging.getLogger(__name__)

# Persisted default item master (SKU -> volumetric/billable weight + category)
# used for the weight-dispute comparison, so it doesn't have to be re-uploaded
# each session. Replaceable via the "Update item master" upload (master_config).
_MASTER_PATH = os.path.join(os.path.dirname(__file__), "data", "item_master.xlsx")
_MASTER_CACHE = {"mtime": None, "maps": None}


def _default_master():
    """Parse the persisted item master (cached by file mtime) into
    {awb2cat, sku2cat, sku2vol, sku_count, updated_at}. Returns empty maps when
    no master file is present."""
    empty = {"awb2cat": {}, "sku2cat": {}, "sku2vol": {}, "sku_count": 0, "updated_at": None}
    try:
        st = os.stat(_MASTER_PATH)
    except OSError:
        _MASTER_CACHE["mtime"] = None
        _MASTER_CACHE["maps"] = None
        return empty
    if _MASTER_CACHE["maps"] is not None and _MASTER_CACHE["mtime"] == st.st_mtime:
        return _MASTER_CACHE["maps"]
    maps = dict(empty)
    try:
        with open(_MASTER_PATH, "rb") as fh:
            data = fh.read()
        kind, payload = invoices.ingest(data, os.path.basename(_MASTER_PATH))
        if kind == "master":
            maps["awb2cat"] = payload.get("awb2cat", {})
            maps["sku2cat"] = payload.get("sku2cat", {})
            maps["sku2vol"] = payload.get("sku2vol", {})
    except Exception:  # noqa: BLE001 - a bad master file must not break the report
        logger.exception("Default item master parse failed")
    maps["sku_count"] = len(maps["sku2vol"])
    maps["updated_at"] = datetime.datetime.fromtimestamp(st.st_mtime).strftime("%d %b %Y %H:%M")
    _MASTER_CACHE["mtime"] = st.st_mtime
    _MASTER_CACHE["maps"] = maps
    return maps

# client can re-filter without re-uploading. Single-process dev use.
# "window" is the date range the loaded data covers (BigQuery load window),
# or None for uploaded files; it persists across re-filter requests.
_CACHE = {"records": None, "window": None}

# Cache of the latest parsed invoice line items + master enrichment maps
# (Carrier Cost Analysis). awb2cat / sku2cat come from uploaded master files
# (weights / SKU masters) and are used to recover category/SKU on invoices that
# don't carry them.
_INVOICE_CACHE = {"items": [], "files": [], "awb2cat": {}, "sku2cat": {},
                  "version": 0, "report_cache": {},
                  # BigQuery product enrichment fetched by invoice date window
                  # (awb/order -> (category, subcategory, sku, item_name)).
                  "awb2prod": {}, "order2prod": {}, "prod_window": None,
                  # Item master SKU -> volumetric/billable weight (kg), from an
                  # uploaded master file, for weight over-charge detection.
                  "sku2vol": {}}

# Cached lane maps (AWB/order -> lane), rebuilt only when the shipment record
# set changes — not on every invoice filter/refresh.
_LANE_CACHE = {"records": None, "maps": None}

# Cached product maps (AWB/order -> (category, subcategory, sku, item_name)),
# built from the loaded BigQuery shipment data so invoice AWBs can be enriched
# with sub-category / SKU. Rebuilt only when the record set changes.
_PROD_CACHE = {"records": None, "maps": None}


def _reset_invoice_cache():
    _INVOICE_CACHE["items"] = []
    _INVOICE_CACHE["files"] = []
    _INVOICE_CACHE["awb2cat"] = {}
    _INVOICE_CACHE["sku2cat"] = {}
    _INVOICE_CACHE["version"] += 1
    _INVOICE_CACHE["report_cache"] = {}
    _INVOICE_CACHE["awb2prod"] = {}
    _INVOICE_CACHE["order2prod"] = {}
    _INVOICE_CACHE["prod_window"] = None
    _INVOICE_CACHE["sku2vol"] = {}


def _refresh_invoice_product_map():
    """Search BigQuery by the uploaded invoices' service month and cache an
    AWB/order -> product (category, sub-category, SKU, item) map for enrichment.

    The window is derived from the invoice dates (service month) padded a little
    each side, so the shipments that the invoices bill are captured regardless of
    any separate BigQuery lookback the user has loaded. No-op when BigQuery isn't
    configured or no invoice carries a service-month/date."""
    from datetime import date as _date, timedelta as _td

    items = _INVOICE_CACHE["items"]
    if not (items and bq.is_configured()):
        return
    keys = sorted({it.get("service_month_key") for it in items
                   if it.get("service_month_key")})
    if not keys:
        return

    def _month_start(k):
        return _date(int(k[:4]), int(k[5:7]), 1)

    def _month_end(k):
        y, m = int(k[:4]), int(k[5:7])
        nxt = _date(y + 1, 1, 1) if m == 12 else _date(y, m + 1, 1)
        return nxt - _td(days=1)

    lo = _month_start(keys[0]) - _td(days=15)
    hi = _month_end(keys[-1]) + _td(days=45)
    # Only the AWBs actually on the invoices need enriching — pass them so the
    # BigQuery loader skips category derivation for every other shipment in the
    # window (most of the per-row cost).
    inv_awbs = {it["awb"] for it in items if it.get("awb")}
    inv_carriers = {it.get("carrier") for it in items if it.get("carrier")}
    try:
        awb2prod, order2prod = bq.fetch_awb_product_map(
            lo.isoformat(), hi.isoformat(), awbs=inv_awbs, carriers=inv_carriers)
    except Exception:  # noqa: BLE001 - enrichment is best-effort; never block upload
        logger.exception("Invoice BigQuery enrichment failed")
        return
    _INVOICE_CACHE["awb2prod"] = awb2prod
    _INVOICE_CACHE["order2prod"] = order2prod
    _INVOICE_CACHE["prod_window"] = {"from": lo.isoformat(), "to": hi.isoformat(),
                                     "awbs": len(awb2prod)}


def _invoice_product_maps():
    """Prefer the BigQuery-by-invoice-date map; fall back to the loaded shipment
    record set when no invoice-date enrichment has been fetched."""
    if _INVOICE_CACHE.get("awb2prod"):
        return _INVOICE_CACHE["awb2prod"], _INVOICE_CACHE.get("order2prod") or {}
    return _product_maps()


def _invoice_report(carrier_filter=None):
    """Build the cost report from the cached line items, optionally restricted
    to a single carrier. The full carrier list is always returned (all_carriers)
    so the frontend dropdown stays populated even while a filter is active."""
    items = _INVOICE_CACHE["items"]
    active = carrier_filter if (carrier_filter and carrier_filter != "all") else ""

    # Cache the built report by (carrier, invoice version, shipment-set id) so a
    # plain refresh or repeated carrier filter is an instant cache hit.
    ckey = (active, _INVOICE_CACHE["version"], id(_CACHE.get("records")))
    cached = _INVOICE_CACHE["report_cache"].get(ckey)
    if cached is not None:
        return cached

    # Full carrier list (by spend desc) — drives the dropdown options.
    spend_by = {}
    for it in items:
        spend_by[it["carrier"]] = spend_by.get(it["carrier"], 0.0) + it["amount"]
    all_carriers = [c for c, _ in sorted(spend_by.items(), key=lambda kv: -kv[1])]

    selected = items
    if active:
        selected = [it for it in items if it["carrier"] == active]

    awb2lane, order2lane = _lane_maps()
    awb2prod, order2prod = _invoice_product_maps()
    # Merge the persisted default item master with any master uploaded this
    # session (session entries override), so weight-dispute comparison works
    # without re-uploading the master every time.
    dm = _default_master()
    eff_awb2cat = dict(dm["awb2cat"]); eff_awb2cat.update(_INVOICE_CACHE.get("awb2cat") or {})
    eff_sku2cat = dict(dm["sku2cat"]); eff_sku2cat.update(_INVOICE_CACHE.get("sku2cat") or {})
    eff_sku2vol = dict(dm["sku2vol"]); eff_sku2vol.update(_INVOICE_CACHE.get("sku2vol") or {})
    report = invoices.build_cost_report(
        selected, _INVOICE_CACHE["files"],
        eff_awb2cat, eff_sku2cat,
        awb2lane=awb2lane, order2lane=order2lane,
        awb2prod=awb2prod, order2prod=order2prod,
        sku2vol=eff_sku2vol)
    report["master"] = {
        "sku_count": len(eff_sku2vol),
        "updated_at": dm["updated_at"],
        "source": "uploaded" if _INVOICE_CACHE.get("sku2vol") else "default",
    }
    report["all_carriers"] = all_carriers
    report["carrier_filter"] = active
    report["prod_window"] = _INVOICE_CACHE.get("prod_window")
    # Cross-carrier comparison is built from ALL loaded invoices (not the
    # filtered subset) so it stays the same regardless of the carrier dropdown.
    report["carrier_comparison"] = invoices.build_carrier_comparison(items)
    _INVOICE_CACHE["report_cache"][ckey] = report
    return report


def _lane_maps():
    """Build AWB -> (pickup_pin, drop_pin) and order-id -> lane maps from the
    loaded shipment data, so invoice lines can be matched to their lane. Uses
    the same AWB normalization as the invoice parser so keys line up."""
    recs = _CACHE.get("records") or []
    # Rebuild only when the shipment record set changes (identity check); avoids
    # rescanning tens of thousands of rows on every invoice filter/refresh.
    if _LANE_CACHE["records"] is recs and _LANE_CACHE["maps"] is not None:
        return _LANE_CACHE["maps"]
    awb2lane, order2lane = {}, {}
    for r in recs:
        pin = r.get("pickup_pin") or ""
        drop = r.get("drop_pin") or ""
        if not (pin and drop):
            continue
        lane = (pin, drop)
        awb = r.get("awb")
        if awb:
            awb2lane[invoices._norm_awb(awb)] = lane
        oid = (r.get("order_id") or "").strip().upper()
        if oid:
            order2lane.setdefault(oid, lane)
    _LANE_CACHE["records"] = recs
    _LANE_CACHE["maps"] = (awb2lane, order2lane)
    return awb2lane, order2lane


def _product_maps():
    """Build AWB -> (category, subcategory, sku, item_name) and order-id -> same
    from the loaded shipment data, so invoice lines can be enriched with the
    product sub-category / SKU pulled from BigQuery. Uses the same AWB
    normalization as the invoice parser so keys line up."""
    recs = _CACHE.get("records") or []
    if _PROD_CACHE["records"] is recs and _PROD_CACHE["maps"] is not None:
        return _PROD_CACHE["maps"]
    awb2prod, order2prod = {}, {}
    for r in recs:
        sub = (r.get("subcategory") or "").strip()
        sku = (r.get("sku") or "").strip()
        cat = (r.get("category") or "").strip()
        name = (r.get("item_name") or "").strip()
        if not (sub or sku or cat or name):
            continue
        prod = (cat, sub, sku, name)
        awb = r.get("awb")
        if awb:
            awb2prod[invoices._norm_awb(awb)] = prod
        oid = (r.get("order_id") or "").strip().upper()
        if oid:
            order2prod.setdefault(oid, prod)
    _PROD_CACHE["records"] = recs
    _PROD_CACHE["maps"] = (awb2prod, order2prod)
    return awb2prod, order2prod


def login_view(request):
    """Internal-team sign-in. GET renders the form, POST validates credentials."""
    if auth.is_authenticated(request):
        return redirect("index")

    next_url = request.GET.get("next") or request.POST.get("next") or ""
    # Only allow safe local redirects.
    if not next_url.startswith("/"):
        next_url = ""

    if request.method == "POST":
        ip = auth.client_ip(request)
        if auth.is_locked_out(ip):
            mins = max(1, auth.seconds_until_unlock(ip) // 60)
            return render(request, "dashboard/login.html", {
                "error": f"Too many attempts. Try again in about {mins} minute(s).",
                "next": next_url,
            }, status=429)

        username = request.POST.get("username", "")
        password = request.POST.get("password", "")
        if auth.verify_credentials(username, password):
            auth.clear_failures(ip)
            auth.login_session(request, username.strip())
            return redirect(next_url or "index")

        auth.record_failure(ip)
        return render(request, "dashboard/login.html", {
            "error": "Invalid username or password.",
            "username": username,
            "next": next_url,
        }, status=401)

    return render(request, "dashboard/login.html", {"next": next_url})


def logout_view(request):
    auth.logout_session(request)
    return redirect("login")


@auth.team_required
def index(request):
    return render(request, "dashboard/index.html", {
        "frido_user": request.session.get(auth.SESSION_KEY, ""),
        # Lets the frontend auto-load the default window from BigQuery on entry
        # (and skip the manual upload screen) only when BQ is actually set up.
        "bq_configured": bq.is_configured(),
    })


def _filter_kwargs(request):
    """Read the multi-select filter fields from a POST into build_report kwargs.

    Each categorical filter is a multi-select checkbox dropdown, so each arrives
    as zero or more repeated form fields. getlist() collects them; an empty list
    falls back to the default ("all" = no constraint). Delivery type also
    defaults to "all" so both forward and reverse-pickup shipments are visible
    unless the user narrows it.
    """
    return {
        "delivery_type": request.POST.getlist("delivery_type") or "all",
        "tier": request.POST.getlist("tier") or "all",
        "zone": request.POST.getlist("zone") or "all",
        "payment": request.POST.getlist("payment") or "all",
        "warehouse": request.POST.getlist("warehouse") or "all",
        "account": request.POST.getlist("account") or "all",
        "weight": request.POST.getlist("weight") or "all",
        "slot": request.POST.getlist("slot") or "all",
        "date_from": request.POST.get("date_from", ""),
        "date_to": request.POST.get("date_to", ""),
    }


def _report_response(request, empty_msg):
    """Build the report from the cached records using the request's filters."""
    records = _CACHE["records"]
    if not records:
        return JsonResponse({"error": empty_msg}, status=400)
    report = build_report(records, **_filter_kwargs(request))
    # The single authoritative date range for the loaded data (BigQuery load
    # window). None for uploaded files; the frontend falls back to pickup span.
    report["load_window"] = _CACHE.get("window")
    return JsonResponse(report)


# Columns written to the shipment-level CSV export, in order: (record key, header).
_EXPORT_FIELDS = [
    ("awb", "AWB"), ("order_id", "Order ID"),
    ("carrier", "Carrier"), ("account", "Account"),
    ("delivery_type", "Delivery type"),
    ("pickup_pin", "Pickup pincode"), ("warehouse", "Warehouse"), ("city", "Pickup city"),
    ("drop_pin", "Drop pincode"), ("drop_city", "Drop city"), ("tier", "City tier"),
    ("payment", "Payment"), ("pickup_date", "Pickup date"),
    ("status", "Latest status"), ("outcome", "Outcome"), ("delivered", "Delivered"),
    ("o2s", "Order->Pickup O2S (hrs)"),
    ("p2o", "Pickup->OFD1 (hrs)"), ("p2d", "Pickup->Delivery (hrs)"),
    ("promised_tat", "Promised TAT"), ("tat_status", "TAT status"), ("tat_margin", "TAT margin"),
]


@auth.team_required
@require_POST
def export_shipments(request):
    """Download the currently-filtered shipments as CSV so the summary cards
    (Shipments / In TAT / Out of TAT) can be traced to the underlying rows.

    Honours every active filter (same as the report). An optional `tat_status`
    multi-value field limits the export to one or more compliance buckets
    ("In TAT", "Out of TAT", "No rule", "Pending")."""
    records = _CACHE["records"]
    if not records:
        return JsonResponse({"error": "Load data first, then export."}, status=400)

    rows = filter_records(records, **_filter_kwargs(request))
    status_sel = set(request.POST.getlist("tat_status"))
    if status_sel:
        rows = [r for r in rows if (r.get("tat_status") or "No rule") in status_sel]

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow([label for _, label in _EXPORT_FIELDS])
    for r in rows:
        writer.writerow([r.get(key, "") for key, _ in _EXPORT_FIELDS])

    suffix = ("_" + "_".join(sorted(status_sel)).replace(" ", "")) if status_sel else ""
    resp = HttpResponse(buf.getvalue(), content_type="text/csv; charset=utf-8")
    resp["Content-Disposition"] = f'attachment; filename="shipments{suffix}.csv"'
    return resp


def _sla_header_index(header):
    """Locate the Warehouse-pincode / (destination) Pincode / TAT columns by
    (loose) header name. The origin column is the warehouse/pickup pincode."""
    h = [str(c or "").strip().lower() for c in (header or [])]

    def find(names, exclude=()):
        for i, c in enumerate(h):
            if i in exclude:
                continue
            if any(n in c for n in names):
                return i
        return None

    origin = find(["warehouse pin", "pickup pin", "origin pin",
                   "warehouse", "origin", "hub"])
    excl = {origin} if origin is not None else set()
    dest = find(["destination pin", "drop pin", "delivery pin", "dest pin"], exclude=excl)
    if dest is None:
        dest = find(["pincode", "pin code", "pin"], exclude=excl)
    return {
        "origin_pin": origin,
        "pincode": dest,
        "tat": find(["tat", "days", "threshold", "sla"]),
    }


def _parse_sla_rows(upload):
    """Parse an uploaded SLA file into [{warehouse, pincode, tat}] rows.

    Accepts .xlsx/.xlsm/.csv/.tsv with columns Warehouse pincode, Pincode, TAT
    (warehouse pincode optional -> defaults to 'ANY')."""
    name = (upload.name or "").lower()
    data = upload.read()
    records = []

    if name.endswith((".xlsx", ".xlsm")):
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
        it = wb[wb.sheetnames[0]].iter_rows(values_only=True)
        rows = list(it)
    elif name.endswith((".csv", ".tsv")):
        text = data.decode("utf-8", "replace")
        delim = "\t" if name.endswith(".tsv") else ","
        rows = list(csv.reader(io.StringIO(text), delimiter=delim))
    else:
        raise ValueError("Upload a .xlsx or .csv with Warehouse, Pincode, TAT columns.")

    if not rows:
        raise ValueError("The file is empty.")
    idx = _sla_header_index(rows[0])
    if idx["pincode"] is None or idx["tat"] is None:
        raise ValueError("Could not find 'Pincode' and 'TAT' columns in the header.")

    def cell(r, key):
        i = idx.get(key)
        return r[i] if (i is not None and i < len(r)) else None

    for r in rows[1:]:
        if not r:
            continue
        pin, tat = cell(r, "pincode"), cell(r, "tat")
        if pin is None or tat is None:
            continue
        records.append({"origin_pin": cell(r, "origin_pin") or "ANY",
                        "pincode": pin, "tat": tat})
    if not records:
        raise ValueError("No valid rows found (need Pincode + TAT in each row).")
    return records


@auth.team_required
@require_POST
def sla_config(request):
    """View/replace a carrier's promised-TAT (SLA) table.

    actions: list | get | save (rows JSON) | upload (file) | clear.
    A save/upload/clear re-scores the loaded data and returns a fresh report."""
    action = request.POST.get("action", "get")
    carrier = (request.POST.get("carrier") or "").strip()

    if action == "list":
        return JsonResponse({"carriers": tat.carriers_meta()})
    if action == "get":
        return JsonResponse({"carrier": carrier, "carriers": tat.carriers_meta(),
                             "rows": tat.get_override_rows(carrier)})

    try:
        if action == "clear":
            tat.clear_override(carrier)
        elif action == "save":
            rows = json.loads(request.POST.get("rows", "[]"))
            tat.set_override(carrier, rows)
        elif action == "upload":
            upload = request.FILES.get("file")
            if upload is None:
                return JsonResponse({"error": "No file uploaded."}, status=400)
            tat.set_override(carrier, _parse_sla_rows(upload))
        else:
            return JsonResponse({"error": "Unknown action."}, status=400)
    except ValueError as exc:
        return JsonResponse({"error": str(exc)}, status=400)

    out = {"ok": True, "carrier": carrier, "carriers": tat.carriers_meta(),
           "rows": tat.get_override_rows(carrier)}
    records = _CACHE["records"]
    if records:
        reclassify(records)
        report = build_report(records, **_filter_kwargs(request))
        report["load_window"] = _CACHE.get("window")
        out["report"] = report
    return JsonResponse(out)


@auth.team_required
@require_POST
def process_upload(request):
    """Accept a new file upload, or a re-filter request on cached data."""
    upload = request.FILES.get("file")
    if upload is not None:
        name = upload.name.lower()
        if not name.endswith((".xlsx", ".xlsm", ".csv", ".tsv")):
            return JsonResponse(
                {"error": "Please upload a .xlsx or .csv file in the standard export format."},
                status=400,
            )
        try:
            records = parse_workbook(upload, filename=upload.name)
        except ValueError as exc:
            return JsonResponse({"error": str(exc)}, status=400)
        except Exception as exc:  # noqa: BLE001 - surface any parse failure cleanly
            return JsonResponse({"error": f"Could not read file: {exc}"}, status=400)

        if not records:
            return JsonResponse(
                {"error": "No data rows found in the file."}, status=400
            )
        _CACHE["records"] = records
        _CACHE["window"] = None  # uploaded file has no BigQuery load window

    return _report_response(
        request, "No data loaded yet. Load from BigQuery or upload a file first."
    )


def _ingest_into_cache(name, data, errors):
    """Parse one invoice/master file's bytes and merge it into _INVOICE_CACHE.
    Shared by the browser upload and the Google Drive import. Appends a message
    to `errors` on failure; returns True when something was ingested."""
    if not str(name).lower().endswith((".xlsx", ".xlsm", ".xlsb", ".csv", ".tsv")):
        errors.append(f"{name}: unsupported file type")
        return None
    try:
        kind, payload = invoices.ingest(data or b"", name)
    except ValueError as exc:
        errors.append(f"{name}: {exc}")
        return None
    except Exception as exc:  # noqa: BLE001
        errors.append(f"{name}: could not read ({exc})")
        return None

    if kind == "master":
        _INVOICE_CACHE["awb2cat"].update(payload.get("awb2cat", {}))
        _INVOICE_CACHE["sku2cat"].update(payload.get("sku2cat", {}))
        _INVOICE_CACHE["sku2vol"].update(payload.get("sku2vol", {}))
        entries = (len(payload.get("awb2cat", {})) + len(payload.get("sku2cat", {}))
                   + len(payload.get("sku2vol", {})))
        _INVOICE_CACHE["files"].append(
            {"name": name, "carrier": "Reference / master", "lines": entries,
             "spend": 0, "kind": "master"})
    else:  # invoice
        spend = round(sum(i["amount"] for i in payload), 1)
        carrier = payload[0]["carrier"] if payload else "Unknown"
        # Each file is one billing period; tag every line with the month derived
        # from its file name so the cost report can compare periods.
        mkey, mlabel = invoices.month_from_filename(name)
        for it in payload:
            it["month"] = mkey
            it["month_label"] = mlabel
        _INVOICE_CACHE["items"].extend(payload)
        _INVOICE_CACHE["files"].append(
            {"name": name, "carrier": carrier, "lines": len(payload),
             "spend": spend, "kind": "invoice", "month": mlabel})
    return kind


@auth.team_required
@require_POST
def import_from_drive(request):
    """Import invoice/master files from a shared Google Drive folder, parse them
    server-side, and return the report. Bypasses request-body size limits (the
    server downloads the files) and supports monthly auto-import.

    Params: `folder` (optional folder ID/link; else GDRIVE_INVOICE_FOLDER env),
    `append=1` to add to the current set instead of replacing."""
    from . import gdrive
    if not gdrive.is_configured():
        return JsonResponse(
            {"error": "Google Drive isn't configured (no service-account "
                      "credentials). See setup notes."}, status=400)
    fid = gdrive.folder_id(request.POST.get("folder"))
    if not fid:
        return JsonResponse(
            {"error": "No Drive folder set. Pass a folder link or set "
                      "GDRIVE_INVOICE_FOLDER."}, status=400)
    try:
        files = gdrive.list_files(fid)
    except Exception as exc:  # noqa: BLE001
        logger.exception("Drive listing failed")
        return JsonResponse(
            {"error": f"Could not read the Drive folder (is it shared with the "
                      f"service account, and the Drive API enabled?): {exc}"}, status=502)

    invoice_files = [f for f in files if gdrive.supported(f.get("name"))]
    if not invoice_files:
        return JsonResponse(
            {"error": "No .xlsx / .xlsb / .csv files found in that Drive folder."},
            status=400)

    if request.POST.get("append") != "1":
        _reset_invoice_cache()

    errors = []
    master_saved = False
    for f in invoice_files:
        name = f.get("name")
        try:
            data = gdrive.download(f["id"], f.get("mimeType"))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{name}: download failed ({exc})")
            continue
        kind = _ingest_into_cache(name, data, errors)
        # A master file found in Drive becomes the persisted default item master
        # (so it's remembered across restarts, like the "Update item master"
        # upload). Best-effort: a read-only filesystem just skips this.
        if kind == "master":
            try:
                os.makedirs(os.path.dirname(_MASTER_PATH), exist_ok=True)
                with open(_MASTER_PATH, "wb") as fh:
                    fh.write(data)
                _MASTER_CACHE["mtime"] = None
                _MASTER_CACHE["maps"] = None
                master_saved = True
            except OSError:
                pass

    if not _INVOICE_CACHE["items"]:
        if master_saved:
            dm = _default_master()
            return JsonResponse({"ok": True, "master_only": True,
                                 "master": {"sku_count": dm["sku_count"],
                                            "updated_at": dm["updated_at"]},
                                 "warnings": errors})
        msg = ("No billed invoice lines found in the Drive files. "
               + (" · ".join(errors) if errors else ""))
        return JsonResponse({"error": msg}, status=400)

    _INVOICE_CACHE["version"] += 1
    _INVOICE_CACHE["report_cache"] = {}
    _refresh_invoice_product_map()
    report = _invoice_report()
    report["imported_from_drive"] = len(invoice_files)
    if errors:
        report["warnings"] = errors
    return JsonResponse(report)


@auth.team_required
@require_POST
def process_invoices(request):
    """Parse one or more carrier invoice (or master) files and return the
    cost-analysis report.

    Each file is auto-detected: an invoice's billed lines are aggregated into
    spend by carrier x category x SKU; a master/reference file (weights, SKU
    master) feeds AWB/SKU -> category maps used to enrich invoices that don't
    carry product info. `reset=1` (no files) clears everything; `append=1` adds
    to the current set instead of replacing it.
    """
    if request.POST.get("reset") == "1" and not request.FILES.getlist("files"):
        _reset_invoice_cache()
        return JsonResponse(_invoice_report())

    uploads = request.FILES.getlist("files") or request.FILES.getlist("file")
    if not uploads:
        # No new files: a plain refresh or a carrier-filter request on the
        # already-cached invoice lines.
        return JsonResponse(_invoice_report(request.POST.get("carrier")))

    if request.POST.get("append") != "1":
        _reset_invoice_cache()

    errors = []
    for up in uploads:
        try:
            up.seek(0)
        except Exception:  # noqa: BLE001
            pass
        _ingest_into_cache(up.name, up.read() or b"", errors)

    if not _INVOICE_CACHE["items"]:
        msg = ("No billed invoice lines found. "
               + (" · ".join(errors) if errors else
                  "Uploaded file(s) had no amounts — add an invoice with charges."))
        return JsonResponse({"error": msg}, status=400)

    # New data ingested (incl. append): invalidate the cached reports.
    _INVOICE_CACHE["version"] += 1
    _INVOICE_CACHE["report_cache"] = {}

    # Search BigQuery by the invoices' service month and cache the AWB -> SKU /
    # sub-category map, so the sub-category analysis is populated automatically.
    _refresh_invoice_product_map()

    # A fresh upload always shows the full (unfiltered) picture; the frontend
    # resets its carrier dropdown to "All carriers" to match.
    report = _invoice_report()
    if errors:
        report["warnings"] = errors
    return JsonResponse(report)


@auth.team_required
@require_POST
def export_invoice_awbs(request):
    """Download every billed AWB line as CSV, enriched with the SKU /
    sub-category joined from the loaded BigQuery shipment data. This is the full
    per-shipment detail behind the per-invoice reconciliation rows (the UI shows
    only counts and the sub-category rollup)."""
    items = _INVOICE_CACHE["items"]
    if not items:
        return JsonResponse({"error": "Upload invoices first, then export."}, status=400)

    # Enrich in place using the same map the report uses (BigQuery searched by
    # invoice date, falling back to the loaded shipment records).
    awb2prod, order2prod = _invoice_product_maps()
    invoices._attach_products(items, awb2prod, order2prod)

    carrier = (request.POST.get("carrier") or "").strip()
    rows = items
    if carrier and carrier != "all":
        rows = [it for it in items if it["carrier"] == carrier]

    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["AWB", "Invoice Number", "Service Month", "Carrier", "Direction",
                "SKU", "Sub-category", "Category", "Item Name",
                "Charged Wt (kg)", "Amount with GST", "Amount ex GST",
                "Matched from BigQuery"])
    for it in rows:
        w.writerow([
            it.get("awb", ""), it.get("invoice_number", ""), it.get("service_month", ""),
            it.get("carrier", ""), it.get("direction", ""),
            it.get("sku", ""), it.get("subcategory", ""), it.get("category", ""),
            it.get("product") or it.get("item_name") or it.get("sku_name") or "",
            it.get("weight_kg") or "",
            round(it.get("amount") or 0.0, 2), round(it.get("amount_ex_gst") or 0.0, 2),
            "yes" if it.get("prod_matched") else "no",
        ])

    resp = HttpResponse(buf.getvalue(), content_type="text/csv; charset=utf-8")
    resp["Content-Disposition"] = 'attachment; filename="invoice_awb_detail.csv"'
    return resp


@auth.team_required
@require_POST
def master_config(request):
    """View or replace the persisted item master (SKU -> volumetric weight) used
    for the weight-dispute comparison.

    actions: get (status) | upload (replace with a new .xlsx/.csv). The uploaded
    file is validated as a recognised master (must yield SKU volumetric weights)
    before it replaces the stored one, and the invoice report cache is cleared so
    the new weights apply immediately."""
    action = request.POST.get("action", "get")
    if action == "upload":
        upload = request.FILES.get("file")
        if upload is None:
            return JsonResponse({"error": "No file uploaded."}, status=400)
        if not upload.name.lower().endswith((".xlsx", ".xlsm", ".csv", ".tsv")):
            return JsonResponse({"error": "Upload a .xlsx or .csv item master."}, status=400)
        data = upload.read() or b""
        try:
            kind, payload = invoices.ingest(data, upload.name)
        except Exception as exc:  # noqa: BLE001
            return JsonResponse({"error": f"Could not read the file: {exc}"}, status=400)
        if kind != "master" or not payload.get("sku2vol"):
            return JsonResponse(
                {"error": "This file isn't a recognised item master — it needs a "
                          "SKU column and a volumetric/billable weight column."},
                status=400)
        try:
            os.makedirs(os.path.dirname(_MASTER_PATH), exist_ok=True)
            with open(_MASTER_PATH, "wb") as fh:
                fh.write(data)
        except OSError as exc:
            return JsonResponse({"error": f"Could not save the master: {exc}"}, status=500)
        _MASTER_CACHE["mtime"] = None
        _MASTER_CACHE["maps"] = None
        # New weights -> invalidate cached invoice reports.
        _INVOICE_CACHE["version"] += 1
        _INVOICE_CACHE["report_cache"] = {}

    dm = _default_master()
    return JsonResponse({"ok": True, "loaded": dm["sku_count"] > 0,
                         "sku_count": dm["sku_count"], "updated_at": dm["updated_at"]})


@auth.team_required
@require_POST
def load_bigquery(request):
    """Fetch a lookback window from BigQuery into the cache, then return the report."""
    if not bq.is_configured():
        return JsonResponse(
            {"error": "BigQuery is not configured on the server "
                      "(set BQ_PROJECT, BQ_DATASET and BQ_TABLE)."},
            status=400,
        )
    # Either an explicit calendar range (win_from/win_to, from the date pickers)
    # or a lookback window (how many days back to pull). The range takes
    # precedence when present.
    win_from = (request.POST.get("win_from") or "").strip()
    win_to = (request.POST.get("win_to") or "").strip()
    use_range = bool(win_from or win_to)

    lookback = request.POST.get("lookback_days")
    lookback_days = int(lookback) if (lookback or "").isdigit() else None

    try:
        records = bq.fetch_records(
            lookback_days=lookback_days,
            date_from=win_from or None,
            date_to=win_to or None,
        )
    except ValueError as exc:  # bad date input -> client error, not a 502
        return JsonResponse({"error": str(exc)}, status=400)
    except Exception as exc:  # noqa: BLE001 - surface any BQ/auth failure cleanly
        # Log the full traceback to the server console so the real cause is
        # visible (the client only sees the short message below).
        logger.exception("BigQuery load failed")
        return JsonResponse({"error": f"BigQuery load failed: {exc}"}, status=502)

    if not records:
        return JsonResponse(
            {"error": "BigQuery returned no rows for the selected date range."},
            status=400,
        )
    _CACHE["records"] = records
    # The authoritative window the UI shows: the picked range, or the lookback.
    if use_range:
        _CACHE["window"] = {"from": win_from or None, "to": win_to or None, "days": None}
    else:
        _CACHE["window"] = bq.lookback_window(lookback_days)

    return _report_response(request, "No data loaded.")