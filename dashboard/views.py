import csv
import io
import json
import logging

from django.http import HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from . import auth, bq, invoices, tat
from .kpi import build_report, filter_records, parse_workbook, reclassify

logger = logging.getLogger(__name__)

# client can re-filter without re-uploading. Single-process dev use.
# "window" is the date range the loaded data covers (BigQuery load window),
# or None for uploaded files; it persists across re-filter requests.
_CACHE = {"records": None, "window": None}

# Cache of the latest parsed invoice line items + master enrichment maps
# (Carrier Cost Analysis). awb2cat / sku2cat come from uploaded master files
# (weights / SKU masters) and are used to recover category/SKU on invoices that
# don't carry them.
_INVOICE_CACHE = {"items": [], "files": [], "awb2cat": {}, "sku2cat": {},
                  "version": 0, "report_cache": {}}

# Cached lane maps (AWB/order -> lane), rebuilt only when the shipment record
# set changes — not on every invoice filter/refresh.
_LANE_CACHE = {"records": None, "maps": None}


def _reset_invoice_cache():
    _INVOICE_CACHE["items"] = []
    _INVOICE_CACHE["files"] = []
    _INVOICE_CACHE["awb2cat"] = {}
    _INVOICE_CACHE["sku2cat"] = {}
    _INVOICE_CACHE["version"] += 1
    _INVOICE_CACHE["report_cache"] = {}


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
    report = invoices.build_cost_report(
        selected, _INVOICE_CACHE["files"],
        _INVOICE_CACHE["awb2cat"], _INVOICE_CACHE["sku2cat"],
        awb2lane=awb2lane, order2lane=order2lane)
    report["all_carriers"] = all_carriers
    report["carrier_filter"] = active
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
        name = up.name
        if not name.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".csv", ".tsv")):
            errors.append(f"{name}: unsupported file type")
            continue
        try:
            up.seek(0)
        except Exception:  # noqa: BLE001
            pass
        data = up.read() or b""
        try:
            kind, payload = invoices.ingest(data, name)
        except ValueError as exc:
            errors.append(f"{name}: {exc}")
            continue
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{name}: could not read ({exc})")
            continue

        if kind == "master":
            _INVOICE_CACHE["awb2cat"].update(payload.get("awb2cat", {}))
            _INVOICE_CACHE["sku2cat"].update(payload.get("sku2cat", {}))
            entries = len(payload.get("awb2cat", {})) + len(payload.get("sku2cat", {}))
            _INVOICE_CACHE["files"].append(
                {"name": name, "carrier": "Reference / master", "lines": entries,
                 "spend": 0, "kind": "master"})
        else:  # invoice
            spend = round(sum(i["amount"] for i in payload), 1)
            carrier = payload[0]["carrier"] if payload else "Unknown"
            # Each file is one billing period; tag every line with the month
            # derived from its file name so the cost report can compare periods.
            mkey, mlabel = invoices.month_from_filename(name)
            for it in payload:
                it["month"] = mkey
                it["month_label"] = mlabel
            _INVOICE_CACHE["items"].extend(payload)
            _INVOICE_CACHE["files"].append(
                {"name": name, "carrier": carrier, "lines": len(payload),
                 "spend": spend, "kind": "invoice", "month": mlabel})

    if not _INVOICE_CACHE["items"]:
        msg = ("No billed invoice lines found. "
               + (" · ".join(errors) if errors else
                  "Uploaded file(s) had no amounts — add an invoice with charges."))
        return JsonResponse({"error": msg}, status=400)

    # New data ingested (incl. append): invalidate the cached reports.
    _INVOICE_CACHE["version"] += 1
    _INVOICE_CACHE["report_cache"] = {}

    # A fresh upload always shows the full (unfiltered) picture; the frontend
    # resets its carrier dropdown to "All carriers" to match.
    report = _invoice_report()
    if errors:
        report["warnings"] = errors
    return JsonResponse(report)


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