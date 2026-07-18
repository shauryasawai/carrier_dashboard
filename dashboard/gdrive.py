"""Google Drive ingestion for carrier invoices (server-side download + parse)."""
from __future__ import annotations

import io
import json
import os
import re

from . import bq

# Read/write: onboarding a new carrier creates a folder, uploads the reference
# invoice and stores the shared carrier-config file, so the read-only scope is
# no longer enough. Full "drive" also covers every read the import path needs.
_SCOPES = ["https://www.googleapis.com/auth/drive"]
SUPPORTED_EXT = (".xlsx", ".xlsm", ".xlsb", ".csv", ".tsv")
_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_FOLDER_MIME = "application/vnd.google-apps.folder"

# File name (inside the root invoice folder) that holds the saved per-carrier
# column-mapping configs — the shared source of truth read by every deployment.
CONFIG_FILE_NAME = "carrier_configs.json"


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


# --------------------------------------------------------------------------
# Write helpers (new-carrier onboarding): create a folder, upload the reference
# invoice, and read/write the shared carrier-config JSON. All best-effort — the
# caller surfaces any failure to the user.
# --------------------------------------------------------------------------
def find_file(name, parent_id, svc=None):
    """Return the file metadata (id, name, mimeType) for `name` directly inside
    `parent_id`, or None. Matches on exact (case-insensitive) name."""
    svc = svc or _service()
    safe = (name or "").replace("'", "\\'")
    q = ("name = '%s' and '%s' in parents and trashed = false" % (safe, parent_id))
    resp = svc.files().list(
        q=q, spaces="drive", fields="files(id, name, mimeType)",
        pageSize=10, supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def find_or_create_folder(name, parent_id, svc=None):
    """Return the id of the sub-folder `name` under `parent_id`, creating it if
    it doesn't exist yet. Used to give each onboarded carrier its own folder."""
    svc = svc or _service()
    existing = find_file(name, parent_id, svc=svc)
    if existing and existing.get("mimeType") == _FOLDER_MIME:
        return existing["id"]
    meta = {"name": name, "mimeType": _FOLDER_MIME, "parents": [parent_id]}
    folder = svc.files().create(
        body=meta, fields="id", supportsAllDrives=True).execute()
    return folder["id"]


def upload_bytes(name, data, parent_id, mime=None, svc=None):
    """Upload `data` as a file named `name` into `parent_id`. If a file with the
    same name already exists there it is replaced (new revision). Returns file id."""
    from googleapiclient.http import MediaIoBaseUpload
    svc = svc or _service()
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime or "application/octet-stream",
                              resumable=False)
    existing = find_file(name, parent_id, svc=svc)
    if existing:
        f = svc.files().update(
            fileId=existing["id"], media_body=media, fields="id",
            supportsAllDrives=True).execute()
        return f["id"]
    meta = {"name": name, "parents": [parent_id]}
    f = svc.files().create(
        body=meta, media_body=media, fields="id", supportsAllDrives=True).execute()
    return f["id"]


def read_configs(root_folder_id, svc=None):
    """Read the shared carrier-config JSON from the root invoice folder. Returns
    a dict keyed by carrier name (empty when the file is absent or unreadable)."""
    svc = svc or _service()
    meta = find_file(CONFIG_FILE_NAME, root_folder_id, svc=svc)
    if not meta:
        return {}
    try:
        raw = download(meta["id"], meta.get("mimeType"), svc=svc)
        data = json.loads(raw.decode("utf-8-sig"))
        return data if isinstance(data, dict) else {}
    except Exception:  # noqa: BLE001 - a corrupt config must not break imports
        return {}


def write_configs(root_folder_id, configs, svc=None):
    """Persist the carrier-config dict as JSON into the root invoice folder."""
    svc = svc or _service()
    blob = json.dumps(configs, ensure_ascii=False, indent=2).encode("utf-8")
    return upload_bytes(CONFIG_FILE_NAME, blob, root_folder_id,
                        mime="application/json", svc=svc)
