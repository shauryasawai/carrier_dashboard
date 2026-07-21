"""Database-free auth for the internal dashboard: username -> PBKDF2 hash in
settings.INTERNAL_USERS, verified with Django hashers, with a per-IP throttle.

SERVERLESS CAVEAT: the per-IP failure counter (`_FAILS`) lives in process
memory, so it is per-instance and resets on cold starts. On a single long-lived
process (Docker/Cloud Run with one worker) it throttles effectively; on Vercel's
serverless functions, requests can fan out across instances, so treat this as a
best-effort speed bump only. For hard brute-force protection there, add a shared
store (e.g. Vercel KV / Redis) behind these helpers, or enable platform-level
rate limiting / WAF on the login route."""
from __future__ import annotations

import time
from functools import wraps

from django.conf import settings
from django.contrib.auth.hashers import check_password
from django.http import JsonResponse
from django.shortcuts import redirect

SESSION_KEY = "frido_user"
MAX_FAILS = 6
WINDOW = 15 * 60
_FAILS: dict[str, list[float]] = {}

# Valid hash of a random value; used only to keep verify timing constant.
_DUMMY_HASH = (
    "pbkdf2_sha256$600000$gPUpwppnEXTy$"
    "EhVj58AYbq6rRM8qjB1D5bEm2pgLHRGCbd0g5lIXY8o="
)


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
    return max(0, int(WINDOW - (time.time() - min(hits)))) if hits else 0


def record_failure(ip: str) -> None:
    _FAILS.setdefault(ip, []).append(time.time())


def clear_failures(ip: str) -> None:
    _FAILS.pop(ip, None)


def verify_credentials(username: str, password: str) -> bool:
    # Always hash-compare, even for unknown users, so timing doesn't leak
    # whether a username exists.
    users = getattr(settings, "INTERNAL_USERS", {}) or {}
    encoded = users.get((username or "").strip())
    if not encoded:
        check_password(password or "", _DUMMY_HASH)
        return False
    return check_password(password or "", encoded)


def is_authenticated(request) -> bool:
    return bool(request.session.get(SESSION_KEY))


def current_user(request) -> str:
    """The signed-in username (empty string when not authenticated)."""
    return (request.session.get(SESSION_KEY) or "").strip()


def is_superuser(request) -> bool:
    """True when the signed-in user is a super user (see settings.SUPERUSERS).
    Super users can view the restricted sections of the dashboard; everyone
    else is a regular team user."""
    supers = getattr(settings, "SUPERUSERS", set()) or set()
    return current_user(request) in supers


def login_session(request, username: str) -> None:
    request.session[SESSION_KEY] = username
    request.session.cycle_key()  # thwart session fixation
    request.session[SESSION_KEY] = username


def logout_session(request) -> None:
    request.session.pop(SESSION_KEY, None)
    request.session.flush()


def team_required(view):
    # HTML views redirect to login; API/POST views get 401 JSON.
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


def super_required(view):
    """Restrict an API view to super users. Assumes team_required has already
    established the session (stack it as the inner decorator). Non-super users
    get a 403 JSON response so the frontend can surface it cleanly."""
    @wraps(view)
    def wrapper(request, *args, **kwargs):
        if is_superuser(request):
            return view(request, *args, **kwargs)
        return JsonResponse(
            {"error": "This section is restricted to super users.",
             "forbidden": True},
            status=403,
        )

    return wrapper
