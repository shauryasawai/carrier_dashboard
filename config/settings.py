import os
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")
SECRET_KEY = os.environ.get("DJANGO_SECRET_KEY")

if not SECRET_KEY:
    raise ValueError("DJANGO_SECRET_KEY environment variable is not set.")

DEBUG = os.environ.get("DJANGO_DEBUG", "False").lower() == "true"
ALLOWED_HOSTS = os.environ.get("DJANGO_ALLOWED_HOSTS", "").split(",")
# HTTPS origins Django trusts for CSRF (Cloud Run serves an https URL). Comma-
# separated, scheme included, e.g. https://carrier-dash-xxxxxx-uc.a.run.app
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
