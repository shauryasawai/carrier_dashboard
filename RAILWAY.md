# Deploying to Railway

Railway runs the same container as Cloud Run (it auto-detects the `Dockerfile`
and injects `$PORT`). The only extra step versus Cloud Run is **BigQuery
credentials**: Railway isn't on GCP, so you provide a service-account key via an
environment variable instead of relying on an automatic identity.

Keep it a **single instance / single worker** (the Dockerfile already uses
`--workers 1`) so the in-memory dataset cache is shared and persists.

---

## 1. Create a BigQuery service account + key (in GCP)

```bash
PROJECT=frido-429506
gcloud iam service-accounts create carrier-dashboard \
  --project "$PROJECT" --display-name "Carrier Dashboard (Railway)"

SA="carrier-dashboard@${PROJECT}.iam.gserviceaccount.com"
gcloud projects add-iam-policy-binding "$PROJECT" \
  --member="serviceAccount:${SA}" --role="roles/bigquery.dataViewer"
gcloud projects add-iam-policy-binding "$PROJECT" \
  --member="serviceAccount:${SA}" --role="roles/bigquery.jobUser"

# Download the key (you'll paste its CONTENTS into Railway, then delete it)
gcloud iam service-accounts keys create sa-key.json --iam-account "$SA"
```

> Treat `sa-key.json` as a secret — **never commit it**. Add `sa-key.json`
> (or `*.json` keys) to `.gitignore`, and delete the local copy once you've
> pasted its contents into Railway.

## 2. Create the Railway service

Either:
- **Dashboard:** New Project → Deploy from GitHub repo (Railway detects the
  Dockerfile and builds it), or
- **CLI:** `npm i -g @railway/cli`, then `railway login` and `railway up` from
  the project folder.

## 3. Set environment variables (Railway → Variables)

| Variable | Value |
|---|---|
| `DJANGO_SECRET_KEY` | a long random string (`python -c "import secrets;print(secrets.token_urlsafe(64))"`) |
| `DJANGO_DEBUG` | `0` |
| `DJANGO_ALLOWED_HOSTS` | your Railway domain, e.g. `carrier-dashboard-production.up.railway.app` |
| `DJANGO_CSRF_TRUSTED_ORIGINS` | `https://carrier-dashboard-production.up.railway.app` |
| `FRIDO_USERS` | `{"admin":"pbkdf2_sha256$..."}` (from `python scripts/make_hash.py`) |
| `GOOGLE_SA_PROJECT_ID` | `frido-429506` (from the key's `project_id`) |
| `GOOGLE_SA_CLIENT_EMAIL` | the key's `client_email` |
| `GOOGLE_SA_PRIVATE_KEY` | the key's `private_key` — paste it with literal `\n` on one line, in quotes |
| `GOOGLE_SA_PRIVATE_KEY_ID` / `GOOGLE_SA_CLIENT_ID` | optional, from the key |
| `BQ_DATE_COLUMN` | `partition_date` (a real DATE; equals created_at's date) |
| `BQ_DATE_COLUMN_IS_STRING` | `0` |
| `BQ_LOOKBACK_DAYS` | `30` (optional; default window) |
| `BQ_PROJECT` / `BQ_DATASET` / `BQ_TABLE` | optional — already default to `frido-429506` / `production` / `Clickpost_Shipment_Tracking_Report` |

> **Credentials:** the app assembles a service-account object from the
> individual `GOOGLE_SA_*` variables (no key file on disk; the `\n` in the
> private key is restored automatically). As a simpler alternative you can
> instead set a single `GOOGLE_SERVICE_ACCOUNT_JSON` to the entire contents of
> `sa-key.json` — both paths work and both are scoped read-only (view access)
> via the BigQuery Data Viewer + Job User IAM roles above.

## 4. Domain → fix the host/CSRF chicken-and-egg

On first deploy you may not know the domain yet. Either:
- Go to **Settings → Networking → Generate Domain**, copy it, then set
  `DJANGO_ALLOWED_HOSTS` and `DJANGO_CSRF_TRUSTED_ORIGINS` to it and redeploy; or
- Temporarily set `DJANGO_ALLOWED_HOSTS=.up.railway.app` for the first boot,
  then pin to the exact domain.

## 5. Open the domain → sign in → **Load from BigQuery**

Pick a date range (default last 30 days) and load.

---

## Notes

- **HTTPS:** Railway terminates TLS and forwards `X-Forwarded-Proto`, which the
  settings already trust (`SECURE_PROXY_SSL_HEADER`), so secure-cookie + SSL
  redirect behave correctly. No redirect loop.
- **Cost:** Hobby is $5/mo with $5 of usage credit; an intermittently-used
  internal tool typically stays near that. Keep the service at **1 replica**.
- **Cloud Run vs Railway:** the code now supports both. On Cloud Run you skip
  `GOOGLE_SERVICE_ACCOUNT_JSON` (it uses the runtime service account); on Railway
  you provide the key. Everything else is identical.
