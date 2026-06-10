import json
import logging

from django.http import JsonResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from . import auth, bq
from .kpi import build_report, parse_workbook

logger = logging.getLogger(__name__)

# client can re-filter without re-uploading. Single-process dev use.
# "window" is the date range the loaded data covers (BigQuery load window),
# or None for uploaded files; it persists across re-filter requests.
_CACHE = {"records": None, "window": None}


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
    })


def _filter_kwargs(request):
    """Read the multi-select filter fields from a POST into build_report kwargs.

    Each categorical filter is a multi-select checkbox dropdown, so each arrives
    as zero or more repeated form fields. getlist() collects them; an empty list
    falls back to the default ("all" = no constraint). Delivery type keeps its
    "Forward" default so the initial view isn't polluted by reverse-pickup
    carriers (see README).
    """
    return {
        "delivery_type": request.POST.getlist("delivery_type") or "Forward",
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
def load_bigquery(request):
    """Fetch a lookback window from BigQuery into the cache, then return the report."""
    if not bq.is_configured():
        return JsonResponse(
            {"error": "BigQuery is not configured on the server "
                      "(set BQ_PROJECT, BQ_DATASET and BQ_TABLE)."},
            status=400,
        )
    # How many days back to pull (partition-pruned). Defaults server-side.
    lookback = request.POST.get("lookback_days")
    lookback_days = int(lookback) if (lookback or "").isdigit() else None

    try:
        records = bq.fetch_records(lookback_days=lookback_days)
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
    _CACHE["window"] = bq.lookback_window(lookback_days)

    return _report_response(request, "No data loaded.")