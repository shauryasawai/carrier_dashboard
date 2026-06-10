# Deploy the Carrier Dashboard to Google Cloud Run — Step by Step

A beginner-friendly walkthrough. Follow it top to bottom. Each step says **what
to do**, **what it means**, and **what you should see**.

End result: your dashboard runs at a public `https://...run.app` URL, protected
by your existing login, reading BigQuery directly. It scales to zero when nobody
is using it, so it stays within Google's free tier.

> **Time:** ~20–30 minutes the first time.
> **Cost:** free for light internal use (see the last section).

---

## What you need before starting

1. **A Google account** that can access the `frido-429506` project (the same
   project your BigQuery data lives in). If you can open
   [BigQuery in the console](https://console.cloud.google.com/bigquery?project=frido-429506),
   you're good.
2. **Billing enabled** on the project. Cloud Run's free tier *requires* a billing
   account to be attached, but you won't be charged under the free limits. Check
   at [Billing](https://console.cloud.google.com/billing) — if it says a billing
   account is linked, you're set. (Ask whoever owns the GCP account if unsure.)
3. **The project code on your computer** — it's already in this folder:
   `Desktop\carrier_dashboard`.

---

## Step 1 — Install the Google Cloud CLI ("gcloud")

`gcloud` is Google's command-line tool. We use it to build and deploy.

1. Download the Windows installer:
   <https://cloud.google.com/sdk/docs/install#windows>
2. Run it, accept the defaults. When it finishes, tick **"Start Google Cloud SDK
   Shell"** and **"Run gcloud init"**.
3. From now on, do everything in the **"Google Cloud SDK Shell"** (search for it
   in the Start menu). It's a black terminal window with `gcloud` ready to use.

**Check it worked** — type this and press Enter:

```
gcloud version
```

You should see a list of versions (Google Cloud SDK, bq, core, etc.).

---

## Step 2 — Sign in and select the project

In the Cloud SDK Shell:

```
gcloud auth login
```

This opens your browser. Pick the Google account that can access `frido-429506`
and allow access. The shell will say *"You are now logged in as ..."*.

Then point gcloud at the project:

```
gcloud config set project frido-429506
```

**What you should see:** `Updated property [core/project].`

---

## Step 3 — Turn on the Google services we need

These are off by default on new projects. This enables Cloud Run (hosting),
Cloud Build (turns your code into a container), Artifact Registry (stores the
container), and BigQuery.

```
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com bigquery.googleapis.com
```

**What you should see:** it thinks for ~30 seconds, then returns to the prompt
with no error. (If it says "billing must be enabled", finish the billing check
in the "What you need" section first.)

---

## Step 4 — Let the app read BigQuery (no key file needed)

On Cloud Run, your app proves its identity using a built-in "service account"
instead of a key file. We just grant that account permission to **read** and
**query** BigQuery — view access only, never edit.

Run these three lines one at a time:

```
gcloud projects describe frido-429506 --format="value(projectNumber)"
```

That prints a number (your project number). Copy it. Now build the account name
and grant the two roles — **replace `NUMBER` with the number you just copied**:

```
gcloud projects add-iam-policy-binding frido-429506 --member="serviceAccount:NUMBER-compute@developer.gserviceaccount.com" --role="roles/bigquery.dataViewer"
```

```
gcloud projects add-iam-policy-binding frido-429506 --member="serviceAccount:NUMBER-compute@developer.gserviceaccount.com" --role="roles/bigquery.jobUser"
```

**What you should see:** each command prints `Updated IAM policy for project`.

> **Why this matters:** because the app uses this account, you do **not** need
> the service-account key from your `.env` on Cloud Run. (And the key you pasted
> earlier should still be rotated/deleted in the console for safety.)

---

## Step 5 — Create your settings file (`env.yaml`)

This file holds the app's settings and secrets. We use a file (not long
commands) so you don't fight with quotes on Windows. It's already in
`.gitignore`, so it won't be committed or uploaded.

1. In this folder (`Desktop\carrier_dashboard`), create a new text file named
   exactly **`env.yaml`**.
2. Paste the block below into it.
3. Fill in the two secret values **from your existing `.env` file**:
   - `DJANGO_SECRET_KEY` — copy the value after `DJANGO_SECRET_KEY=` in `.env`.
   - `FRIDO_USERS` — copy the whole `{...}` value after `FRIDO_USERS=` in `.env`.

```yaml
DJANGO_DEBUG: "0"
DJANGO_ALLOWED_HOSTS: ".run.app"
DJANGO_SECRET_KEY: "PASTE_THE_KEY_FROM_YOUR_.env"
FRIDO_USERS: '{"admin":"pbkdf2_sha256$PASTE_THE_REST_FROM_YOUR_.env"}'
BQ_PROJECT: "frido-429506"
BQ_DATASET: "production"
BQ_TABLE: "Clickpost_Shipment_Tracking_Report"
BQ_DATE_COLUMN: "partition_date"
BQ_DATE_COLUMN_IS_STRING: "0"
BQ_LOOKBACK_DAYS: "30"
```

Notes:
- Keep the quotes exactly as shown. The `FRIDO_USERS` line uses **single quotes**
  on the outside and double quotes inside — that's intentional.
- Do **not** add any `GOOGLE_SA_*` or `GOOGLE_SERVICE_ACCOUNT_JSON` here. On
  Cloud Run the app uses the service account from Step 4 automatically.

---

## Step 6 — Deploy

Make sure the shell is "in" the project folder. Type:

```
cd "C:\Users\ShauryamanSawai\OneDrive - Arcatron Mobility Pvt Ltd\Desktop\carrier_dashboard"
```

Then deploy (this is one long command — copy the whole line):

```
gcloud run deploy carrier-dashboard --source . --region asia-south1 --allow-unauthenticated --memory 2Gi --cpu 1 --min-instances 0 --max-instances 1 --env-vars-file env.yaml
```

**What happens now:**
- It uploads your code and builds the container with Cloud Build. **This takes
  3–6 minutes the first time** — lots of log lines scroll by. That's normal.
- If it asks *"Allow unauthenticated invocations?"* answer **y** (your own login
  page protects the app).
- When it finishes you'll see a green **Service URL** like
  `https://carrier-dashboard-xxxxxxxxxx-el.a.run.app`. **Copy that URL.**

**Flags explained:**
- `--source .` — build from the Dockerfile in this folder.
- `--region asia-south1` — Mumbai, same region as your data.
- `--memory 2Gi` — enough RAM to hold a 45-day load comfortably.
- `--min-instances 0` — scale to zero when idle = stays free. (Trade-off: after
  it's been unused a while, the next visit waits a few seconds to "wake up", and
  you re-click "Load from BigQuery".)
- `--max-instances 1` — one instance, so there's a single shared data cache.

---

## Step 7 — Lock the web address (host + CSRF)

For security, Django only answers on hosts you approve. We set a temporary
`.run.app` in Step 5; now pin it to your exact URL. **Replace the URL below with
the one you copied.**

```
gcloud run services update carrier-dashboard --region asia-south1 --update-env-vars "DJANGO_ALLOWED_HOSTS=carrier-dashboard-xxxxxxxxxx-el.a.run.app" --update-env-vars "DJANGO_CSRF_TRUSTED_ORIGINS=https://carrier-dashboard-xxxxxxxxxx-el.a.run.app"
```

Note: `DJANGO_ALLOWED_HOSTS` is the host **without** `https://`; the CSRF one
**includes** `https://`.

**What you should see:** `Service [carrier-dashboard] revision ... has been
deployed`.

---

## Step 8 — Open it and test

1. Open the Service URL in your browser.
2. You'll see the **login page** — sign in with the username/password whose hash
   is in `FRIDO_USERS` (e.g. `admin`).
3. Pick a date range and click **Load from BigQuery**. The first load takes a
   few seconds; after that, changing filters is instant.

That's it — you're live.

---

## If something goes wrong

| What you see | Likely cause | Fix |
|---|---|---|
| Browser shows **"Bad Request (400)"** | The host isn't allowed | Re-check Step 7 — the URL must match exactly (no trailing slash). |
| Login page works but **"BigQuery load failed"** | Permissions or settings | Re-run Step 4 (roles). Confirm `BQ_DATE_COLUMN: "partition_date"` is in `env.yaml`. |
| **"Memory limit exceeded"** in logs | Load too big for RAM | Lower `BQ_LOOKBACK_DAYS` in `env.yaml`, or raise `--memory` to `4Gi` and redeploy. |
| Build fails during deploy | Code/dependency error | Read the last red lines; usually a typo in `env.yaml`. Fix and re-run Step 6. |
| Page is slow to first load after idle | Scaled to zero (expected on free) | Normal. To remove it, set `--min-instances 1` — but that's no longer free. |

**See live logs** (very useful for the 502/"load failed" case):

```
gcloud run services logs read carrier-dashboard --region asia-south1 --limit 50
```

---

## Updating the app later

After you change code or settings, just deploy again from the project folder:

```
gcloud run deploy carrier-dashboard --source . --region asia-south1 --memory 2Gi --min-instances 0 --max-instances 1 --env-vars-file env.yaml
```

Your URL and settings stay the same.

---

## Keeping it free (and the one real speed-up)

- **Free tier:** Cloud Run gives 2M requests + 360,000 GB-seconds of memory per
  month free. With `--min-instances 0`, you only consume time while the app is
  actually handling requests (plus ~15 min of idle before it sleeps). Light
  internal use stays free. Heavy all-day use could exceed it.
- **Secrets (optional, recommended):** for stronger security, move
  `DJANGO_SECRET_KEY` and `FRIDO_USERS` into Google Secret Manager instead of
  `env.yaml` — see Step 6 in `DEPLOY.md`.
- **Speed:** each load scans ~425 MB because the source BigQuery table isn't
  partitioned. The date range only changes how many rows come back, not the scan
  size. The real fix is to **cluster/partition the table on `partition_date`** —
  ask if you'd like help setting that up.
