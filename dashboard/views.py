import json

from django.http import JsonResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from . import auth
from .kpi import build_report, parse_workbook

# The parsed records for the most recent upload, kept in memory so the
# client can re-filter without re-uploading. Single-process dev use.
_CACHE = {"records": None}


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


@auth.team_required
@require_POST
def process_upload(request):
    """Accept either a new file upload or a re-filter request on cached data."""
    # Every categorical filter is now a multi-select checkbox dropdown, so each
    # arrives as zero or more repeated form fields. getlist() collects them; an
    # empty list falls back to the default ("all" = no constraint). Delivery
    # type keeps its "Forward" default so the initial view isn't polluted by
    # reverse-pickup carriers (see README).
    delivery_type = request.POST.getlist("delivery_type") or "Forward"
    zone = request.POST.getlist("zone") or "all"
    payment = request.POST.getlist("payment") or "all"
    warehouse = request.POST.getlist("warehouse") or "all"
    account = request.POST.getlist("account") or "all"
    weight = request.POST.getlist("weight") or "all"
    slot = request.POST.getlist("slot") or "all"
    date_from = request.POST.get("date_from", "")
    date_to = request.POST.get("date_to", "")

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

    records = _CACHE["records"]
    if records is None:
        return JsonResponse(
            {"error": "No file loaded yet. Upload a workbook first."}, status=400
        )

    report = build_report(
        records,
        delivery_type=delivery_type,
        zone=zone,
        payment=payment,
        warehouse=warehouse,
        account=account,
        weight=weight,
        slot=slot,
        date_from=date_from,
        date_to=date_to,
    )
    return JsonResponse(report)