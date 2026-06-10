# Deploying the Carrier Dashboard to Vercel

> **Read this first.** Vercel is *serverless*, which fights this app's design in
> two ways you should accept before starting:
>
> 1. **No persistent memory between requests.** The app's speed normally comes
>    from loading data once and re-filtering it from RAM. On Vercel each request
>    can land on a fresh instance, so the cache often won't be there — which
>    means **every filter change may re-query BigQuery** (a few seconds each,
>    and BigQuery cost each time). That's why we hard-cap the window small.
> 2. **A 10-second timeout on the free (Hobby) plan.** A BigQuery load scans
>    ~425 MB and can exceed 10 s. For anything but the smallest window you'll
>    need the **Pro plan (60 s)** — so Vercel is effectively *not free* here.
>
> If those trade-offs aren't worth it, deploy to **Cloud Run** instead (free,
> persistent, fast) — see `CLOUD_RUN_STEP_BY_STEP.md`. This guide is for when you
> specifically need Vercel.

The repo is already prepared for Vercel: `vercel.json`, `api/index.py` (the
entrypoint), a slimmed `requirements.txt` (pyarrow/bigquery-storage removed to
fit Vercel's 250 MB limit), `.vercelignore`, and static files served by
WhiteNoise at request time.

---

## Step 0 — Rotate the BigQuery key first

Unlike Cloud Run, Vercel has **no Google identity**, so you must give the app a
service-account key via environment variables. The key you pasted in chat
earlier is compromised — in the Google Cloud console go to **IAM & Admin →
Service Accounts → your account → Keys**, delete the old key, and **create a new
JSON key**. You'll copy values from that new file below.

(Grant that service account **BigQuery Data Viewer + Job User** on project
`frido-429506` if it doesn't have them already.)

---

## Step 1 — Put the code on GitHub

Vercel deploys from a Git repo. From the project folder:

```
git add .
git commit -m "Prepare for Vercel"
git push
```

If the repo isn't on GitHub yet: create an empty repo on github.com, then follow
its "push an existing repository" instructions. **Before pushing, double-check
`git status` does NOT list `.env`, `env.yaml`, or any `*.json` key** — they're
in `.gitignore`, but verify, because they contain secrets.

---

## Step 2 — Import the project into Vercel

1. Go to <https://vercel.com> and sign in (use "Continue with GitHub").
2. **Add New… → Project**, pick this repository, click **Import**.
3. Framework preset: leave as **Other**. Don't set a build command. Click
   **Deploy** — the first deploy will succeed but the app won't work yet until
   you add the environment variables (next step). That's expected.

---

## Step 3 — Add environment variables

In the Vercel project: **Settings → Environment Variables**. Add each of these
(Name = left column). Set them for the **Production** environment.

| Name | Value |
|---|---|
| `DJANGO_SECRET_KEY` | the long random string from your `.env` |
| `DJANGO_DEBUG` | `0` |
| `DJANGO_ALLOWED_HOSTS` | `.vercel.app` |
| `DJANGO_CSRF_TRUSTED_ORIGINS` | `https://<your-project>.vercel.app` (fill in after first deploy) |
| `FRIDO_USERS` | the full `{"admin":"pbkdf2_sha256$..."}` from your `.env` |
| `BQ_PROJECT` | `frido-429506` |
| `BQ_DATASET` | `production` |
| `BQ_TABLE` | `Clickpost_Shipment_Tracking_Report` |
| `BQ_DATE_COLUMN` | `partition_date` |
| `BQ_DATE_COLUMN_IS_STRING` | `0` |
| `BQ_LOOKBACK_DAYS` | `7` |
| `BQ_MAX_LOOKBACK_DAYS` | `7` |
| `GOOGLE_SA_PROJECT_ID` | `frido-429506` (from the new key's `project_id`) |
| `GOOGLE_SA_CLIENT_EMAIL` | the new key's `client_email` |
| `GOOGLE_SA_PRIVATE_KEY` | the new key's `private_key` — paste it with the literal `\n` sequences, on one line |
| `GOOGLE_SA_PRIVATE_KEY_ID` | the new key's `private_key_id` (optional) |
| `GOOGLE_SA_CLIENT_ID` | the new key's `client_id` (optional) |

Notes:
- `BQ_MAX_LOOKBACK_DAYS=7` is the important one — it hard-caps every load to 7
  days (~45K rows) so it fits the timeout/memory. The date dropdown still shows
  larger options, but the server clamps them to 7.
- `VERCEL` is set automatically by Vercel — that switches the app to serve
  static files without a build step. You don't add it.
- For `GOOGLE_SA_PRIVATE_KEY`, the value looks like
  `-----BEGIN PRIVATE KEY-----\nMIIE...\n-----END PRIVATE KEY-----\n`. Paste it
  exactly, including the `\n`s; the app converts them to real line breaks.

---

## Step 4 — Redeploy and pin the URL

1. After saving the variables, go to **Deployments → ⋯ → Redeploy** on the
   latest deployment (env vars only apply to new deployments).
2. Note your URL, e.g. `https://carrier-dashboard.vercel.app`.
3. Update `DJANGO_CSRF_TRUSTED_ORIGINS` to `https://carrier-dashboard.vercel.app`
   (your exact URL) and redeploy once more.

---

## Step 5 — (If loads time out) upgrade the timeout

On the free Hobby plan, functions stop at **10 seconds**. If "Load from
BigQuery" fails with a timeout:

1. Upgrade the project to the **Pro** plan.
2. Open `vercel.json` and raise the limit, then push:

   ```json
   "functions": {
     "api/index.py": { "memory": 1024, "maxDuration": 60 }
   }
   ```

   (Leave `maxDuration` out on Hobby — values above 10 are rejected there.)

---

## Step 6 — Open it

Visit your Vercel URL → sign in → **Load from BigQuery** (7-day window). The
first load takes a few seconds. Remember: changing a filter may reload from
BigQuery again, so it won't feel as instant as the Cloud Run version.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Build error: function exceeds 250 MB | A heavy dependency slipped back in | Ensure `pyarrow` / `google-cloud-bigquery-storage` are NOT in `requirements.txt`. |
| "Bad Request (400)" | Host not allowed | `DJANGO_ALLOWED_HOSTS=.vercel.app` and redeploy. |
| Timeout / 504 after ~10 s | Hobby timeout too short | Lower `BQ_MAX_LOOKBACK_DAYS` further, or go Pro + `maxDuration` (Step 5). |
| "BigQuery load failed" | Bad/missing credentials | Re-check the `GOOGLE_SA_*` vars; the private key must keep its `\n`s. |
| Filters are slow / "No data loaded" after idle | Serverless cache didn't persist | Inherent to Vercel; click Load again. This is the core trade-off. |
| CSS/JS missing | Static not served | Confirm the deploy picked up `VERCEL` (it's automatic) — the app serves static via WhiteNoise. |

---

## Why Cloud Run is still the better fit (for reference)

On Cloud Run the instance stays warm, so the cache persists and filtering is
instant; you can use the full 7/15/30/45-day windows; you get 1–2 GB RAM; and it
authenticates with no key file. It's free for light use. Vercel works, but
you're trading speed and a paid tier for it.
