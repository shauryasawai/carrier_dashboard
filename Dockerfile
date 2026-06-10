# Cloud Run container for the Frido Carrier Console.
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=8080 \
    DJANGO_DEBUG=0

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Collect static files at build time. settings.py requires a SECRET_KEY just to
# import, so a throwaway one is supplied here — it is NOT used at runtime (the
# real key comes from the DJANGO_SECRET_KEY env var on the service).
RUN DJANGO_SECRET_KEY=build-time-only python manage.py collectstatic --noinput

# Cloud Run routes traffic to $PORT. ONE worker only (the parsed dataset lives
# in process memory and is shared across threads); threads give concurrency.
CMD ["sh", "-c", "gunicorn config.wsgi:application --bind :$PORT --workers 1 --threads 8 --timeout 120"]
