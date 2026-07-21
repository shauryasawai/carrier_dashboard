import os
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")
SECRET_KEY = os.environ.get("DJANGO_SECRET_KEY")

if not SECRET_KEY:
    raise ValueError("DJANGO_SECRET_KEY environment variable is not set.")

DEBUG = os.environ.get("DJANGO_DEBUG", "False").lower() == "true"
# Comma-separated hostnames from the env; blanks dropped so a stray comma can't
# introduce an empty ('') host. In local DEBUG we fall back to localhost so the
# dev server works without extra config; in production DJANGO_ALLOWED_HOSTS MUST
# be set (an empty list makes Django reject every request — fail closed).
ALLOWED_HOSTS = [h.strip() for h in
                 os.environ.get("DJANGO_ALLOWED_HOSTS", "").split(",") if h.strip()]
if DEBUG and not ALLOWED_HOSTS:
    ALLOWED_HOSTS = ["localhost", "127.0.0.1"]
CSRF_TRUSTED_ORIGINS = [
    o.strip()
    for o in os.environ.get("DJANGO_CSRF_TRUSTED_ORIGINS", "").split(",")
    if o.strip()
]

INSTALLED_APPS = [
    "django.contrib.staticfiles",
    "dashboard",
]

# django_extensions is a developer convenience (shell_plus, etc.) and is not a
# deploy dependency. Load it only when it's actually installed, so production
# hosts (Vercel/Cloud Run, where it isn't in requirements.txt) don't crash.
try:
    import django_extensions  # noqa: F401
    INSTALLED_APPS.append("django_extensions")
except ImportError:
    pass

STATIC_URL = "static/"
STATIC_ROOT = BASE_DIR / "staticfiles"
# WhiteNoise: compress + hash static files at collectstatic time.
STORAGES = {
    "default": {"BACKEND": "django.core.files.storage.FileSystemStorage"},
    "staticfiles": {
        "BACKEND": "whitenoise.storage.CompressedManifestStaticFilesStorage"
    },
}

# On Vercel the filesystem is read-only and `collectstatic` doesn't run during
# the build, so the hashed manifest won't exist. Serve static files directly
# from the app's static dirs at request time instead (no manifest needed).
if os.environ.get("VERCEL"):
    STORAGES["staticfiles"]["BACKEND"] = (
        "whitenoise.storage.CompressedStaticFilesStorage"
    )
    WHITENOISE_USE_FINDERS = True

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    # WhiteNoise serves collected static files in production (right after
    # SecurityMiddleware, before everything else).
    "whitenoise.middleware.WhiteNoiseMiddleware",
    # Compress dynamic responses. The report JSON is large and is re-fetched on
    # every filter/date change; gzip cuts it several-fold. Placed after WhiteNoise
    # so static assets keep serving their own pre-compressed copies — this only
    # affects dynamic responses. No secrets/reflected input in the body (no BREACH risk).
    "django.middleware.gzip.GZipMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

SESSION_ENGINE = "django.contrib.sessions.backends.signed_cookies"
LOGIN_URL = "login"

# Cookie hardening. Cookies are HttpOnly + SameSite=Lax; mark them Secure
# (HTTPS-only) automatically once DEBUG is off. Set DJANGO_DEBUG=0 in prod.
SESSION_COOKIE_HTTPONLY = True
SESSION_COOKIE_SAMESITE = "Lax"
SESSION_COOKIE_SECURE = not DEBUG
CSRF_COOKIE_SAMESITE = "Lax"
CSRF_COOKIE_SECURE = not DEBUG
SESSION_COOKIE_AGE = 60 * 60 * 8
SESSION_EXPIRE_AT_BROWSER_CLOSE = True

# Behind-HTTPS / proxy security headers (no-ops in local http dev).
SECURE_CONTENT_TYPE_NOSNIFF = True
SECURE_REFERRER_POLICY = "same-origin"
X_FRAME_OPTIONS = "DENY"
if not DEBUG:
    SECURE_SSL_REDIRECT = os.environ.get("DJANGO_SSL_REDIRECT", "1") == "1"
    SECURE_PROXY_SSL_HEADER = ("HTTP_X_FORWARDED_PROTO", "https")
    # HTTP Strict Transport Security: tell browsers to only ever use HTTPS.
    # Defaults to 1 year with subdomains + preload; override/relax via env while
    # you validate that every subdomain is HTTPS-ready (preload is hard to undo).
    SECURE_HSTS_SECONDS = int(os.environ.get("DJANGO_HSTS_SECONDS", 60 * 60 * 24 * 365))
    SECURE_HSTS_INCLUDE_SUBDOMAINS = (
        os.environ.get("DJANGO_HSTS_INCLUDE_SUBDOMAINS", "1") == "1")
    SECURE_HSTS_PRELOAD = os.environ.get("DJANGO_HSTS_PRELOAD", "1") == "1"


import os as _os
import json as _json
import logging as _logging

logger = _logging.getLogger(__name__)

INTERNAL_USERS = {}

_env_users = _os.environ.get("FRIDO_USERS")
if _env_users:
    try:
        INTERNAL_USERS = _json.loads(_env_users)
    except ValueError:
        logger.warning(
            "FRIDO_USERS env var contains invalid JSON — falling back to empty INTERNAL_USERS."
        )

# Role-based access: usernames listed here (comma-separated in FRIDO_SUPERUSERS)
# are "super users" and can see the restricted sections of the dashboard
# (AI business summary, Revenue / order value, Destination city tiers, Payment
# mode performance, Orders placed per day, Product breakdown). Everyone else is
# a regular team user and sees the rest of the dashboard only. If the env var is
# unset, no one is treated as a super user (fail closed).
SUPERUSERS = {
    u.strip()
    for u in _os.environ.get("FRIDO_SUPERUSERS", "").split(",")
    if u.strip()
}

ROOT_URLCONF = "config.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [],
        "APP_DIRS": True,
        "OPTIONS": {"context_processors": []},
    },
]

WSGI_APPLICATION = "config.wsgi.application"

# No database needed - this app is stateless and processes uploads in-memory.
DATABASES = {}

LANGUAGE_CODE = "en-us"
TIME_ZONE = "Asia/Kolkata"
USE_I18N = True
USE_TZ = True


DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

# Cap uploads at 50 MB to avoid memory blowups on huge files.
DATA_UPLOAD_MAX_MEMORY_SIZE = 50 * 1024 * 1024
FILE_UPLOAD_MAX_MEMORY_SIZE = 50 * 1024 * 1024
