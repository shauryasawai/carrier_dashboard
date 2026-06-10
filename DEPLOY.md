# Deploying to Google Cloud Run

This app fetches data from BigQuery (no upload), keeps the parsed dataset in
process memory, and is served by gunicorn. Run it as a **single worker** with a
**minimum of 1 warm instance** so the in-memory cache survives between requests.

---

## 0. Prerequisites

- The [gcloud CLI](https://cloud.google.com/sdk/docs/install) installed and
  `gcloud auth login` done.
- A GCP project with billing enabled. Set it once:

  ```bash
  gcloud config set project YOUR_PROJECT_ID
  ```

- Your shipment data already in a BigQuery table.

## 1. Enable the APIs

```bash
gcloud services enable \
  run.googleapis.com \
  cloudbuild.googleapis.com \
  artifactregistry.googleapis.com \
  bigquery.googleapis.com
```

## 2. Let the app read BigQuery

Cloud Run uses a runtime service account. Simplest is the default compute SA;
grant it read + query access (scope to the dataset in production if you prefer):

```bash
PROJECT_ID=$(gcloud config get-value project)
PROJECT_NUMBER=$(gcloud projects describe "$PROJECT_ID" --format='value(projectNumber)')
RUNTIME_SA="${PROJECT_NUMBER}-compute@developer.gserviceaccount.com"

gcloud projects add-iam-policy-binding "$PROJECT_ID" \
  --member="serviceAccount:${RUNTIME_SA}" --role="roles/bigquery.dataViewer"
gcloud projects add-iam-policy-binding "$PROJECT_ID" \
  --member="serviceAccount:${RUNTIME_SA}" --role="roles/bigquery.jobUser"
```

No key files needed — the container authenticates as this SA automatically.

## 3. Generate your secrets

```bash
# Django signing key
python -c "import secrets; print(secrets.token_urlsafe(64))"

# A password hash for each internal user
python scripts/make_hash.py "the-password"
```

Build the users JSON, e.g.
`{"admin":"pbkdf2_sha256$...","asha":"pbkdf2_sha256$..."}`.

## 4. First deploy

`--source .` builds the image from the Dockerfile with Cloud Build. Use
`.run.app` as an allowed host for the first boot (you don't know the exact URL
yet), and 2 GiB RAM for the full dataset.

```bash
gcloud run deploy carrier-dashboard \
  --source . \
  --region asia-south1 \
  --allow-unauthenticated \
  --memory 2Gi --cpu 1 \
  --min-instances 1 --max-instances 1 \
  --set-env-vars "DJANGO_DEBUG=0" \
  --set-env-vars "DJANGO_ALLOWED_HOSTS=.run.app" \
  --set-env-vars "DJANGO_SECRET_KEY=PASTE_YOUR_KEY" \
  --set-env-vars 'FRIDO_USERS={"admin":"pbkdf2_sha256$..."}' \
  --set-env-vars "BQ_PROJECT=frido-429506" \
  --set-env-vars "BQ_DATASET=production" \
  --set-env-vars "BQ_TABLE=Clickpost_Shipment_Tracking_Report" \
  --set-env-vars "BQ_LOOKBACK_DAYS=30"
```

> The table id `frido-429506.production.Clickpost_Shipment_Tracking_Report` and
> the column mapping are already the defaults in `dashboard/bq.py`, so the
> `BQ_*` vars above are optional — set them only to point at a different table.
> Run the app in the **same project** (`frido-429506`) or grant its runtime
> service account read access to that project's BigQuery.

> **Data window:** the table holds ~1 year. Every load filters on the
> `partition_date` partition with a lookback window (default 30 days, chosen on
> the load screen: 7 / 30 / 90 / 180 / 365). This keeps each query fast and
> well inside the BigQuery free tier. `BQ_LOOKBACK_DAYS` sets the default.

> `--min-instances 1` keeps one instance warm so the cached dataset persists
> (and avoids cold starts). `--max-instances 1` guarantees a single shared
> cache. Drop to `--min-instances 0` to save money if you accept a reload after
> idle scale-down.

## 5. Lock down the host + CSRF

The deploy prints a **Service URL** like
`https://carrier-dashboard-xxxxxxxxxx-el.a.run.app`. Pin it:

```bash
gcloud run services update carrier-dashboard --region asia-south1 \
  --update-env-vars "DJANGO_ALLOWED_HOSTS=carrier-dashboard-xxxxxxxxxx-el.a.run.app" \
  --update-env-vars "DJANGO_CSRF_TRUSTED_ORIGINS=https://carrier-dashboard-xxxxxxxxxx-el.a.run.app"
```

Open the URL → you'll hit the login page → sign in → click **Load from
BigQuery**.

## 6. (Recommended) Move secrets into Secret Manager

```bash
echo -n "YOUR_KEY"  | gcloud secrets create django-secret-key --data-file=-
echo -n '{"admin":"pbkdf2_sha256$..."}' | gcloud secrets create frido-users --data-file=-

gcloud run services update carrier-dashboard --region asia-south1 \
  --update-secrets "DJANGO_SECRET_KEY=django-secret-key:latest" \
  --update-secrets "FRIDO_USERS=frido-users:latest"
```

(Grant the runtime SA `roles/secretmanager.secretAccessor`.)

---

## Environment variables reference

| Variable | Required | Notes |
|---|---|---|
| `DJANGO_SECRET_KEY` | yes | Long random string; signs the session cookie. |
| `DJANGO_DEBUG` | — | `0` in production (default). |
| `DJANGO_ALLOWED_HOSTS` | yes | Comma-separated hosts. `.run.app` for first boot; pin to the exact host after. |
| `DJANGO_CSRF_TRUSTED_ORIGINS` | yes | `https://<your-run-url>` (scheme included). |
| `FRIDO_USERS` | yes | JSON map of username → PBKDF2 hash (from `scripts/make_hash.py`). |
| `BQ_PROJECT` / `BQ_DATASET` / `BQ_TABLE` | yes | The BigQuery table to read. |
| `BQ_COLUMN_MAP` | if columns differ | JSON overriding the logical-field → column map in `dashboard/bq.py`. |
| `BQ_MAX_ROWS` | no | Safety cap on rows fetched. |

## Notes

- **Column mapping:** `dashboard/bq.py` assumes snake_case columns matching the
  standard export. If your table's columns differ, set `BQ_COLUMN_MAP` (JSON) or
  edit `DEFAULT_COLUMN_MAP` in that file.
- **Cost:** at this data size you're well within BigQuery's free tier. To keep
  it that way as data grows, partition the table by pickup date and select only
  the mapped columns (the query already does the latter).
- **Redeploy** after code changes: re-run the `gcloud run deploy` command from
  step 4 (env vars persist across deploys).
