"""Google Drive ingestion for carrier invoices.

Lets the app pull invoice files straight from a shared Google Drive folder,
server-side, instead of the browser uploading them. This bypasses request-body
size limits (e.g. Vercel's 4.5 MB cap — the file is downloaded by the server,
not posted through it) and enables monthly auto-import: drop the month's
invoices in the folder and import them in one click.

Auth: reuses the same service-account credentials as BigQuery (the GOOGLE_SA_*
env vars, assembled by bq._service_account_info), scoped to Drive read-only.
Falls back to Application Default Credentials when no SA env is set.

Setup required (one-time):
  1. Enable the Google Drive API on the GCP project.
  2. Share the invoice Drive folder with the service account's email
     (GOOGLE_SA_CLIENT_EMAIL) — same as sharing with a person.
  3. Set GDRIVE_INVOICE_FOLDER to that folder's ID or share link.
"""
from __future__ import annotations

import io
import os
import re

from . import bq

# Read-only is enough — we only list and download.
_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

# Invoice/master file types we know how to parse.
SUPPORTED_EXT = (".xlsx", ".xlsm", ".xlsb", ".csv", ".tsv")

# Google-native Sheets are exported to .xlsx on download.
_GOOGLE_SHEET = "application/vnd.google-apps.spreadsheet"
_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_FOLDER_MIME = "application/vnd.google-apps.folder"


def folder_id(value=None):
    """Resolve a Drive folder ID from a raw ID, a /folders/<id> link, or an
    ?id=<id> link. Falls back to the GDRIVE_INVOICE_FOLDER env var."""
    v = (value or os.environ.get("GDRIVE_INVOICE_FOLDER") or "").strip()
    if not v:
        return ""
    m = re.search(r"/folders/([A-Za-z0-9_-]+)", v)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([A-Za-z0-9_-]+)", v)
    if m:
        return m.group(1)
    return v  # assume it's already a bare folder ID


def is_configured():
    """True when Drive credentials are available (a folder can still be passed
    per-request even without the env default)."""
    if bq._service_account_info() is not None:
        return True
    return bool(os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"))


def _service():
    from googleapiclient.discovery import build

    info = bq._service_account_info()
    if info is not None:
        from google.oauth2 import service_account
        creds = service_account.Credentials.from_service_account_info(info, scopes=_SCOPES)
    else:
        import google.auth
        creds, _ = google.auth.default(scopes=_SCOPES)
    try:
        return build("drive", "v3", credentials=creds,
                     cache_discovery=False, static_discovery=True)
    except TypeError:
        return build("drive", "v3", credentials=creds, cache_discovery=False)


def supported(name):
    return (name or "").lower().endswith(SUPPORTED_EXT)


def list_files(folder, recursive=True):
    """List non-trashed files inside `folder`, descending into sub-folders when
    `recursive` (so the folder can be organised carrier-wise: BlueDart/, Swift/,
    …). Returns [{id, name, mimeType, size, modifiedTime, folder}] where `folder`
    is the name of the sub-folder the file was found in ("" for the top level)."""
    svc = _service()
    out = []
    _collect(svc, folder, "", recursive, out, 0)
    return out


def _collect(svc, folder_id_, folder_name, recursive, out, depth):
    q = "'%s' in parents and trashed = false" % folder_id_
    page = None
    while True:
        resp = svc.files().list(
            q=q, spaces="drive",
            fields="nextPageToken, files(id, name, mimeType, size, modifiedTime)",
            pageSize=200, pageToken=page, orderBy="modifiedTime desc",
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        for f in resp.get("files", []):
            if f.get("mimeType") == _FOLDER_MIME:
                if recursive and depth < 6:   # descend into carrier sub-folders
                    _collect(svc, f["id"], f.get("name", ""), recursive, out, depth + 1)
            else:
                f["folder"] = folder_name
                out.append(f)
        page = resp.get("nextPageToken")
        if not page:
            break


def download(file_id, mime=None, svc=None):
    """Download a Drive file's bytes. Google-native Sheets are exported to xlsx;
    everything else is fetched as-is.

    Pass an existing `svc` (from `_service()`) to avoid rebuilding the Drive
    client — building it re-creates credentials and is wasteful when downloading
    many files."""
    from googleapiclient.http import MediaIoBaseDownload

    if svc is None:
        svc = _service()
    if mime and mime.startswith("application/vnd.google-apps"):
        request = svc.files().export_media(fileId=file_id, mimeType=_XLSX_MIME)
    else:
        request = svc.files().get_media(fileId=file_id, supportsAllDrives=True)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


def new_service():
    """Build a Drive API client to reuse across several download() calls in one
    request. Rebuilding the client per file (new credentials + a discovery fetch)
    is the main per-file overhead, so callers downloading many files should build
    it once here and pass it to each download().

    Note: downloads are kept sequential (one reused client, no threads). Parallel
    downloads segfaulted the Python worker on Vercel's serverless runtime, where
    native SSL under the API client is not safe across threads."""
    return _service()
