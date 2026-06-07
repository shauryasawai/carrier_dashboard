"""Lightweight, database-free authentication for the internal dashboard.

Credentials live in settings.INTERNAL_USERS as username -> PBKDF2 hash. We
verify with Django's password hashers (constant-time, salted) and record the
signed-in user in the signed-cookie session. A small in-memory throttle slows
brute-force attempts per client IP.
"""

from __future__ import annotations

import time
from functools import wraps

from django.conf import settings
from django.contrib.auth.hashers import check_password
from django.http import JsonResponse
from django.shortcuts import redirect

SESSION_KEY = "frido_user"

# --- Brute-force throttle (per IP, in-memory; single-process dev use) --------
# After MAX_FAILS failed attempts within WINDOW seconds, further attempts are
# blocked until the window rolls off.
MAX_FAILS = 6
WINDOW = 15 * 60  # 15 minutes
_FAILS: dict[str, list[float]] = {}


def client_ip(request) -> str:
    xff = request.META.get("HTTP_X_FORWARDED_FOR", "")
    if xff:
        return xff.split(",")[0].strip()
    return request.META.get("REMOTE_ADDR", "unknown")


def _recent_fails(ip: str) -> int:
    now = time.time()
    hits = [t for t in _FAILS.get(ip, []) if now - t < WINDOW]
    if hits:
        _FAILS[ip] = hits
    else:
        _FAILS.pop(ip, None)
    return len(hits)


def is_locked_out(ip: str) -> bool:
    return _recent_fails(ip) >= MAX_FAILS


def seconds_until_unlock(ip: str) -> int:
    hits = _FAILS.get(ip, [])
    if not hits:
        return 0
    return max(0, int(WINDOW - (time.time() - min(hits))))


def record_failure(ip: str) -> None:
    _FAILS.setdefault(ip, []).append(time.time())


def clear_failures(ip: str) -> None:
    _FAILS.pop(ip, None)


def verify_credentials(username: str, password: str) -> bool:
    """Return True iff the username exists and the password matches its hash.

    Always runs a hash comparison (even for unknown users) so response timing
    doesn't leak whether a username exists.
    """
    users = getattr(settings, "INTERNAL_USERS", {}) or {}
    encoded = users.get((username or "").strip())
    if not encoded:
        # Dummy check against a throwaway hash to equalise timing.
        check_password(password or "", _DUMMY_HASH)
        return False
    return check_password(password or "", encoded)


# A valid hash of a random value, used only to keep timing constant for the
_DUMMY_HASH = (
    "pbkdf2_sha256$600000$gPUpwppnEXTy$"
    "EhVj58AYbq6rRM8qjB1D5bEm2pgLHRGCbd0g5lIXY8o="
)


def is_authenticated(request) -> bool:
    return bool(request.session.get(SESSION_KEY))


def login_session(request, username: str) -> None:
    request.session[SESSION_KEY] = username
    # Rotate the session key on login to thwart session fixation.
    request.session.cycle_key()
    request.session[SESSION_KEY] = username


def logout_session(request) -> None:
    request.session.pop(SESSION_KEY, None)
    request.session.flush()


def team_required(view):
    """Gate a view behind login.

    HTML views redirect to the login page (with ?next=); API/POST views get a
    401 JSON response so the frontend can redirect.
    """
    @wraps(view)
    def wrapper(request, *args, **kwargs):
        if is_authenticated(request):
            return view(request, *args, **kwargs)
        accepts_json = (
            request.headers.get("X-Requested-With") == "fetch"
            or "application/json" in request.headers.get("Accept", "")
            or request.method == "POST"
        )
        if accepts_json:
            return JsonResponse(
                {"error": "Your session has expired. Please sign in again.",
                 "auth": False},
                status=401,
            )
        return redirect(f"{settings.LOGIN_URL}?next={request.path}")

    return wrapper
