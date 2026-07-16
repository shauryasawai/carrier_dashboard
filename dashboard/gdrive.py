"""Google Drive ingestion for carrier invoices (server-side download + parse)."""
from __future__ import annotations

import io
import os
import re

from . import bq

_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
SUPPORTED_EXT = (".xlsx", ".xlsm", ".xlsb", ".csv", ".tsv")
_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_FOLDER_MIME = "application/vnd.google-apps.folder"


def folder_id(value=None):
    v = (value or os.environ.get("GDRIVE_INVOICE_FOLDER") or "").strip()
    if not v:
        return ""
    m = re.search(r"/folders/([A-Za-z0-9_-]+)", v) or re.search(r"[?&]id=([A-Za-z0-9_-]+)", v)
    return m.group(1) if m else v


def is_configured():
    return bq._service_account_info() is not None or bool(os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"))


def _service():
    from googleapiclient.discovery import build
    info = bq._service_account_info()
    if info is not None:
        from google.oauth2 import service_account
        creds = service_account.Credentials.from_service_account_info(info, scopes=_SCOPES)
    else:
        import google.auth
        creds, _ = google.auth.default(scopes=_SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def supported(name):
    return (name or "").lower().endswith(SUPPORTED_EXT)


def list_files(folder, recursive=True):
    out = []
    _collect(_service(), folder, "", recursive, out, 0)
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
                if recursive and depth < 6:
                    _collect(svc, f["id"], f.get("name", ""), recursive, out, depth + 1)
            else:
                f["folder"] = folder_name
                out.append(f)
        page = resp.get("nextPageToken")
        if not page:
            break


def download(file_id, mime=None, svc=None):
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


def download_many(files, max_workers=None):
    # Parallel download. httplib2 isn't thread-safe so each worker keeps its own
    # service via thread-local. Env GDRIVE_DOWNLOAD_WORKERS (default 6).
    import concurrent.futures
    import threading
    files = list(files)
    if not files:
        return []
    raw = os.environ.get("GDRIVE_DOWNLOAD_WORKERS", "6") if max_workers is None else str(max_workers)
    n = int(raw) if str(raw).isdigit() and int(raw) > 0 else 6
    workers = max(1, min(n, len(files)))
    tl = threading.local()

    def _svc():
        svc = getattr(tl, "svc", None)
        if svc is None:
            svc = tl.svc = _service()
        return svc

    def _one(f):
        try:
            return (f, download(f["id"], f.get("mimeType"), svc=_svc()), None)
        except Exception as exc:  # noqa: BLE001
            return (f, None, str(exc))

    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as ex:
        return list(ex.map(_one, files))
