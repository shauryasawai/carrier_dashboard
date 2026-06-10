"""Vercel serverless entrypoint.

Vercel's @vercel/python runtime imports this module and serves the WSGI
callable named `app`. All routes are sent here (see vercel.json); WhiteNoise
serves the static files in-process.
"""

import os

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "config.settings")

from django.core.wsgi import get_wsgi_application  # noqa: E402

app = get_wsgi_application()
