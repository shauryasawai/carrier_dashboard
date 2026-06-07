"""
KPI computation engine for carrier partner efficiency.

Reads a shipment workbook (same column layout as the standard export) and
produces per-carrier KPIs plus a weighted 0-100 efficiency score, along with
overall summary numbers and business-mix breakdowns.

This version adds WAREHOUSE-LEVEL CARRIER TRACKING: in addition to the global
per-carrier scoreboard and the per-warehouse scoreboard, it computes a
warehouse x carrier matrix so you can see how each partner performs out of a
specific warehouse (pickup pincode). Warehouses are labelled "City - Pincode"
using a pincode-prefix lookup so the UI reads naturally.

All timestamp columns may arrive either as real datetimes (openpyxl converts
Excel serials automatically) or as raw numeric serials; both are handled.
"""

from __future__ import annotations

from datetime import datetime, timedelta
from io import BytesIO

from openpyxl import load_workbook

# Logical field -> list of candidate header names (first match wins). Matching
# is case/space-insensitive, so minor header drift won't break parsing, and any
# export that is a SUBSET of the full 151-column standard format works: missing
# optional columns simply yield blank/None for that field.
COLUMNS = {
    "carrier": ["Carrier Partner Name"],
    "account": ["Carrier Partner Account Name"],
    "weight": ["Shipment Weight"],
    "payment": ["Payment Mode"],
    "pickup_pin": ["Pickup Pincode"],
    "drop_pin": ["Drop Pin Code", "Drop Pincode"],
    # Destination city as shipped (free text). Optional; we fall back to a
    # pincode-derived region when it's blank.
    "drop_city": ["Drop City", "Destination City"],
    "pickup_ts": ["Pickup Timestamp"],
    "delivery_ts": ["Delivery Timestamp"],
    "ofd1_ts": ["OFD1 Timestamp"],
    # The standard export uses "Shipment Zone"; older trimmed exports used "Zone".
    "zone": ["Zone", "Shipment Zone", "Pricing Zone"],
    "delivery_type": ["Delivery Type"],
    "status": ["Latest Status"],
    "attempts": ["Number of Delivery Attempts"],
    # Product info (used to derive category / subcategory). Optional.
    "item_names": ["Item Names", "Item Name", "Product Name"],
    "sku": ["Product SKU Codes", "SKU", "Product SKU Code"],
}

# Fields that are required; if none of their candidate headers are present the
# file is rejected with a clear error. Everything else is optional.
REQUIRED_FIELDS = ["carrier"]

# Weights for the composite efficiency score (must sum to 1.0).
WEIGHTS = {
    "p2o": 0.30,   # Pickup -> OFD1 TAT       (lower is better)
    "p2d": 0.25,   # Pickup -> Delivery TAT   (lower is better)
    "succ": 0.20,  # Delivery success rate    (higher is better)
    "fa": 0.15,    # First-attempt strike     (higher is better)
    "att": 0.10,   # Avg delivery attempts    (lower is better)
}

EXCEL_EPOCH = datetime(1899, 12, 30)

# ---------------------------------------------------------------------------
# Pincode -> city resolution.
#
# The export has no pickup-city column (Drop City is the destination, not the
# warehouse), so we derive a readable warehouse name from the pickup pincode.
# We use the first 3 digits of the Indian PIN, which maps to a sorting
# district / region. The table below covers the high-volume warehouses seen in
# the data plus common metro prefixes; anything unmatched falls back to the
# 2-digit postal-circle name, and finally to "PIN <prefix>xx".
# ---------------------------------------------------------------------------

# 3-digit prefix -> city/region label (highest priority).
PIN3_CITY = {
    "412": "Pune (Maval)",
    "411": "Pune",
    "413": "Solapur",
    "122": "Gurugram",
    "121": "Faridabad",
    "110": "Delhi",
    "562": "Bengaluru (Rural)",
    "560": "Bengaluru",
    "561": "Bengaluru (Rural)",
    "712": "Howrah",
    "711": "Howrah",
    "700": "Kolkata",
    "501": "Hyderabad (RR)",
    "500": "Hyderabad",
    "421": "Thane",
    "400": "Mumbai",
    "401": "Palghar",
    "440": "Nagpur",
    "302": "Jaipur",
    "380": "Ahmedabad",
    "600": "Chennai",
    "632": "Vellore",
    "590": "Belagavi",
    "682": "Kochi",
    "457": "Ratlam",
    "470": "Sagar",
    "532": "Srikakulam",
    "232": "Ghazipur",
}

# 2-digit prefix -> postal circle / broad region (fallback).
PIN2_CIRCLE = {
    "11": "Delhi", "12": "Haryana", "13": "Punjab", "14": "Punjab",
    "16": "Chandigarh", "17": "Himachal", "18": "J&K", "19": "J&K",
    "20": "UP (West)", "21": "UP", "22": "UP", "23": "UP", "24": "UP",
    "25": "UP", "26": "UP", "27": "UP", "28": "UP",
    "30": "Rajasthan", "31": "Rajasthan", "32": "Rajasthan", "33": "Rajasthan",
    "34": "Rajasthan",
    "36": "Gujarat", "37": "Gujarat", "38": "Gujarat", "39": "Gujarat",
    "40": "Mumbai/MH", "41": "Maharashtra", "42": "Maharashtra",
    "43": "Maharashtra", "44": "Maharashtra",
    "45": "MP", "46": "MP", "47": "MP", "48": "MP",
    "49": "Chhattisgarh",
    "50": "Telangana", "51": "Andhra", "52": "Andhra", "53": "Andhra",
    "56": "Karnataka", "57": "Karnataka", "58": "Karnataka", "59": "Karnataka",
    "60": "Chennai/TN", "61": "Tamil Nadu", "62": "Tamil Nadu",
    "63": "Tamil Nadu", "64": "Tamil Nadu",
    "67": "Kerala", "68": "Kerala", "69": "Kerala",
    "70": "Kolkata/WB", "71": "West Bengal", "72": "West Bengal",
    "73": "West Bengal", "74": "West Bengal",
    "75": "Odisha", "76": "Odisha", "77": "Odisha",
    "78": "Assam", "79": "North East",
    "80": "Bihar", "81": "Bihar", "82": "Jharkhand", "83": "Jharkhand",
    "84": "Bihar", "85": "Bihar",
}


def city_for_pincode(pin: str) -> str:
    """Return a readable city/region for a (string) pincode, or '' if blank."""
    if not pin:
        return ""
    digits = "".join(ch for ch in str(pin) if ch.isdigit())
    if len(digits) < 2:
        return ""
    p3 = digits[:3]
    if p3 in PIN3_CITY:
        return PIN3_CITY[p3]
    p2 = digits[:2]
    if p2 in PIN2_CIRCLE:
        return PIN2_CIRCLE[p2]
    return "PIN " + p3 + "xx"


def warehouse_label(pin: str) -> str:
    """'City - Pincode' label used as the warehouse display name."""
    if not pin:
        return ""
    city = city_for_pincode(pin)
    return (city + " \u00b7 " + str(pin)) if city else str(pin)


def _clean_city(text: str) -> str:
    """Normalize a free-text city name so casing/spacing variants merge.

    The export's Drop City is raw uppercase free text (e.g. 'NEW DELHI',
    ' mumbai '); collapsing whitespace and title-casing folds the obvious
    duplicates together for a cleaner destination rollup.
    """
    if not text:
        return ""
    return " ".join(str(text).split()).title()


def _norm_pincode(raw) -> str:
    """Normalize a pincode cell to a clean digit string ('560037.0' -> '560037')."""
    if raw is None or str(raw).strip() == "":
        return ""
    try:
        return str(int(float(raw)))
    except (ValueError, TypeError):
        return str(raw).strip()


def _norm_header(value) -> str:
    return str(value).strip().lower() if value is not None else ""


def _to_datetime(value):
    """Coerce a cell value to datetime, or None if not parseable/blank."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)):
        try:
            return EXCEL_EPOCH + timedelta(days=float(value))
        except (ValueError, OverflowError):
            return None
    s = str(value).strip()
    if not s:
        return None
    # ISO first (real datetimes / clean exports).
    try:
        return datetime.fromisoformat(s)
    except (ValueError, TypeError):
        pass
    # CSV exports of this dataset use US-style M/D/YYYY with an optional time,
    # e.g. "6/2/2026 23:58" or "6/2/2026". Try the common explicit formats.
    for fmt in (
        "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%m/%d/%Y",
        "%m/%d/%y %H:%M:%S", "%m/%d/%y %H:%M", "%m/%d/%y",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _hours_between(start, end):
    """Positive elapsed hours, or None if either end is missing/negative."""
    if start is None or end is None:
        return None
    delta = (end - start).total_seconds() / 3600.0
    return delta if delta >= 0 else None


def _to_float(value):
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def _weight_bucket(grams):
    if grams is None:
        return "Unknown"
    if grams <= 500:
        return "0-500g"
    if grams <= 1000:
        return "500g-1kg"
    if grams <= 5000:
        return "1-5kg"
    return "5kg+"


# Coarse 3-way weight class used for the filter dropdown and its breakdown.
# Weights in the export are in GRAMS, so the kg thresholds are *1000.
#   Light  : 0 - 3 kg
#   Medium : 3 - 10 kg
#   Heavy  : 10 kg and above
WEIGHT_LIGHT_MAX = 3000     # <= 3 kg
WEIGHT_MEDIUM_MAX = 10000   # 3 kg - 10 kg


def _weight_class(grams):
    if grams is None:
        return "Unknown"
    if grams <= WEIGHT_LIGHT_MAX:
        return "Light"
    if grams <= WEIGHT_MEDIUM_MAX:
        return "Medium"
    return "Heavy"


# Time-of-day slots for the pickup timestamp. Boundaries are by hour:
# Time-of-day slots for the pickup timestamp. Hourly between 12pm and 5pm,
# with open buckets before noon and after 5pm:
#   Before 12pm : hour < 12
#   12-1pm      : hour == 12
#   1-2pm       : hour == 13
#   2-3pm       : hour == 14
#   3-4pm       : hour == 15
#   4-5pm       : hour == 16
#   After 5pm   : hour >= 17
# The slot labels double as the filter values (kept stable for the frontend).
PICKUP_SLOTS = [
    "Before 12pm", "12\u20131pm", "1\u20132pm", "2\u20133pm",
    "3\u20134pm", "4\u20135pm", "After 5pm",
]

# Map the hour (12..16) to its label for the hourly band.
_HOURLY_SLOT = {
    12: "12\u20131pm", 13: "1\u20132pm", 14: "2\u20133pm",
    15: "3\u20134pm", 16: "4\u20135pm",
}


def _pickup_slot(dt):
    if dt is None:
        return "Unknown"
    h = dt.hour
    if h < 12:
        return "Before 12pm"
    if h >= 17:
        return "After 5pm"
    return _HOURLY_SLOT[h]


# ---------------------------------------------------------------------------
# Product categorization (derived from the Item Names text, since the export
# has no category column). Rules are ordered: the FIRST matching (category,
# subcategory) whose keyword appears in the item name wins, so put more
# specific rules before broader ones. Matching is case-insensitive on whole
# words/substrings of the product name.
# ---------------------------------------------------------------------------
# Each rule: (category, subcategory, [keywords]). First keyword hit wins, so
# ORDER MATTERS - more specific categories come before broader ones (a
# "pregnancy pillow" must hit Maternity before the generic Pillows rule; an
# "arch support insole" must hit Insoles before Footwears/Orthotics).
#
# The ten top-level categories are fixed (the Frido catalogue); subcategories
# are a convenience grouping and can be tuned freely. Anything unmatched falls
# to ("Others", "Other"). Keywords match as case-insensitive substrings.
PRODUCT_RULES = [
    # --- Maternity & Baby Care (first, so a pregnancy/maternity pillow lands
    # here instead of under Pillows) ---
    ("Maternity & Baby Care", "Pregnancy Pillow", ["pregnancy", "maternity"]),
    ("Maternity & Baby Care", "Baby Care", ["baby", "infant", "nursing", "feeding pillow", "kids"]),

    # --- Insoles (before Footwears/Orthotics; an insole is its own category) ---
    ("Insoles", "Insoles", ["insole", "shoe insert", "foot insert"]),

    # --- Socks ---
    ("Socks", "Socks", ["sock"]),

    # --- Mattress ---
    ("Mattress", "Mattress", ["mattress"]),
    ("Mattress", "Topper", ["topper"]),

    # --- Footwears ---
    ("Footwears", "Sandals", ["sandal"]),
    ("Footwears", "Slippers", ["slipper", "flip flop", "flipflop", "clog"]),
    ("Footwears", "Shoes", ["shoe", "sneaker", "footwear"]),

    # --- Pillows ---
    ("Pillows", "Neck Pillow", ["neck pillow", "cervical", "neck contour", "travel pillow"]),
    ("Pillows", "Wedge Pillow", ["wedge"]),
    ("Pillows", "Sleep Pillow", ["sleep pillow", "cozy pillow", "memory foam pillow", "bed pillow", "pillow"]),

    # --- Cushions ---
    ("Cushions", "Seat Cushion", ["seat cushion", "donut", "coccyx", "seat"]),
    ("Cushions", "Backrest", ["backrest", "back rest", "lumbar cushion", "lumbar"]),
    ("Cushions", "Cushion", ["cushion"]),

    # --- Frido Orthotics (posture, braces, joint & foot supports) ---
    ("Frido Orthotics", "Posture Corrector", ["posture"]),
    ("Frido Orthotics", "Knee & Joint Support", ["knee", "ankle", "elbow", "wrist", "shoulder"]),
    ("Frido Orthotics", "Braces & Wraps", ["brace", "wrap", "lumbo sacral", "sacral", "compression", "support belt", "belt", "support"]),
    ("Frido Orthotics", "Foot Care", ["bunion", "heel", "plantar", "toe", "arch", "orthotic", "foot"]),

    # --- Personal Care (therapy, masks, nasal, massage, pain relief) ---
    ("Personal Care", "Masks", ["eye mask", "sleep mask", "mask"]),
    ("Personal Care", "Hot/Cold Therapy", ["therapy", "hot & cold", "cold & hot", "heating pad", "heat pad"]),
    ("Personal Care", "Nasal Care", ["nasal", "nose"]),
    ("Personal Care", "Massage & Relief", ["massager", "massage", "roller", "pain relief", "pain-relief"]),
]


def _product_category(item_name):
    """Return (category, subcategory) for a product name, or ('Others','Other')."""
    if not item_name:
        return ("Unknown", "Unknown")
    text = item_name.lower()
    for category, subcategory, keywords in PRODUCT_RULES:
        for kw in keywords:
            if kw in text:
                return (category, subcategory)
    return ("Others", "Other")


# ---------------------------------------------------------------------------
# Delivery-outcome categorization (drives the status matrix).
#
# Latest Status is bucketed into four mutually exclusive outcomes that mirror
# the ops pivot: Delivered / FWD Pendency / RTO / Cancelled. FWD Pendency is
# the catch-all for shipments still moving through the forward pipeline (not
# yet delivered, not returned, not cancelled). The four outcome buckets are a
# strict partition, so per-carrier shares sum to 100%.
# ---------------------------------------------------------------------------
OUTCOMES = ["Delivered", "FWD Pendency", "RTO", "Cancelled"]

# Latest Status values treated as terminal cancellations (not pendency).
_CANCELLED_STATUSES = {
    "cancelled", "notserviceable", "lost", "damaged", "nostatusexist",
}


def _outcome(status):
    s = (status or "").strip()
    low = s.lower()
    if low == "delivered":
        return "Delivered"
    # Any RTO-* / RTO ... status is a return-to-origin outcome.
    if low.startswith("rto-") or low.startswith("rto "):
        return "RTO"
    if low in _CANCELLED_STATUSES:
        return "Cancelled"
    return "FWD Pendency"


# The FWD-Pendency sub-states we surface as columns in the pendency matrix.
# These are the Latest Status values that fall under FWD Pendency, matching the
# operational checklist. Any pendency status not in this list rolls into
# "Other".
PENDENCY_STATES = [
    "InTransit", "OutForDelivery", "FailedDelivery", "ShipmentDelayed",
    "OutForPickup", "PickupFailed", "PickupPending", "ShipmentHeld",
    "DestinationHubIn", "OriginCityIn", "OriginCityOut", "PickedUp",
    "ContactCustomerCare", "OrderPlaced", "Awb Registered",
]
_PENDENCY_SET = {s.lower() for s in PENDENCY_STATES}


def _pendency_state(status):
    """Sub-bucket for a FWD-Pendency shipment's Latest Status."""
    s = (status or "").strip()
    if s.lower() in _PENDENCY_SET:
        # Return the canonical spelling from PENDENCY_STATES.
        for canon in PENDENCY_STATES:
            if canon.lower() == s.lower():
                return canon
    return "Other"


def _read_all_bytes(file_obj) -> bytes:
    """Return the full byte content of an upload, regardless of stream state.

    Django's UploadedFile may have already had its position moved (by a size
    check, middleware, or a prior read), which would make a plain .read()
    return empty bytes and openpyxl raise 'File is not a zip file'. We rewind
    when possible and fall back to .chunks() for large/temporary uploads.
    """
    try:
        file_obj.seek(0)
    except (AttributeError, OSError):
        pass

    data = b""
    chunks = getattr(file_obj, "chunks", None)
    if callable(chunks):
        try:
            data = b"".join(chunks())
        except Exception:  # noqa: BLE001 - fall through to .read()
            data = b""
    if not data:
        try:
            file_obj.seek(0)
        except (AttributeError, OSError):
            pass
        data = file_obj.read() or b""

    if isinstance(data, str):
        data = data.encode("utf-8", "ignore")
    return data


def _rows_from_bytes(data: bytes, filename: str = ""):
    """Yield rows (tuples) from either an .xlsx or a .csv byte payload.

    Returns an iterator over rows where the first row is the header. Raises
    ValueError with a clear message if the bytes are neither.
    """
    name = (filename or "").lower()
    is_xlsx = data[:2] == b"PK"

    if is_xlsx:
        wb = load_workbook(filename=BytesIO(data), read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        return ws.iter_rows(values_only=True)

    # Not a ZIP/xlsx. Old .xls (OLE) can't be read here.
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise ValueError(
            "This looks like an old-format .xls file. Please re-save it as .xlsx "
            "or export as CSV, then upload again."
        )

    # Treat as CSV/TSV text. Decode tolerantly (handles the BOM and stray bytes).
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("latin-1", "replace")

    # If it looks like HTML, reject with a clear message.
    stripped = text.lstrip()
    if stripped[:1] == "<" or stripped[:5].lower() == "<?xml":
        raise ValueError(
            "This file looks like HTML/XML, not a workbook or CSV. Open it in "
            "Excel and Save As 'CSV' or 'Excel Workbook (.xlsx)', then upload."
        )

    import csv as _csv
    import io as _io
    # Sniff the delimiter (comma vs tab vs semicolon); default to comma.
    sample = text[:4096]
    try:
        dialect = _csv.Sniffer().sniff(sample, delimiters=",\t;|")
        delim = dialect.delimiter
    except _csv.Error:
        delim = "\t" if name.endswith(".tsv") else ","
    reader = _csv.reader(_io.StringIO(text), delimiter=delim)
    return reader


def parse_workbook(file_obj, filename: str = "") -> list[dict]:
    """Read the first worksheet/CSV into a list of normalized row dicts.

    Accepts .xlsx/.xlsm workbooks and .csv/.tsv files (the standard export is
    a 151-column CSV). 'filename' is an optional hint used only for delimiter
    defaults; content sniffing is primary.
    """
    data = _read_all_bytes(file_obj)
    if not data:
        raise ValueError(
            "The uploaded file was empty (no bytes received). Please re-select "
            "the file and try again."
        )

    rows_iter = _rows_from_bytes(data, filename)
    try:
        header = next(rows_iter)
    except StopIteration:
        return []

    # Map our logical names to actual column indices. Each field has a list of
    # candidate header names; the first one present in this file wins. This is
    # what lets any subset of the standard format work.
    lookup = {_norm_header(h): i for i, h in enumerate(header)}
    idx = {}
    for key, candidates in COLUMNS.items():
        for name in candidates:
            i = lookup.get(_norm_header(name))
            if i is not None:
                idx[key] = i
                break

    missing_required = [k for k in REQUIRED_FIELDS if k not in idx]
    if missing_required:
        wanted = ", ".join(COLUMNS[k][0] for k in missing_required)
        raise ValueError(
            "The file is missing required column(s): " + wanted + ". "
            "Check that the uploaded file matches the standard export format."
        )

    def cell(row, key):
        i = idx.get(key)
        return row[i] if i is not None and i < len(row) else None

    records = []
    for row in rows_iter:
        if row is None:
            continue
        carrier = cell(row, "carrier")
        if carrier is None or str(carrier).strip() == "":
            continue

        pickup = _to_datetime(cell(row, "pickup_ts"))
        ofd1 = _to_datetime(cell(row, "ofd1_ts"))
        delivery = _to_datetime(cell(row, "delivery_ts"))
        attempts = _to_float(cell(row, "attempts"))
        weight = _to_float(cell(row, "weight"))
        status = str(cell(row, "status") or "").strip()
        payment = str(cell(row, "payment") or "").strip().upper()

        # Pincodes can arrive as ints or floats (560037.0) - normalize to a
        # clean string without decimals.
        pickup_pin = _norm_pincode(cell(row, "pickup_pin"))
        drop_pin = _norm_pincode(cell(row, "drop_pin"))

        pickup_city = city_for_pincode(pickup_pin)
        # Destination: prefer the explicit Drop City column; fall back to a
        # region derived from the drop pincode, then to the bare pincode.
        drop_city = _clean_city(str(cell(row, "drop_city") or "").strip())
        if not drop_city:
            drop_city = city_for_pincode(drop_pin) or drop_pin
        # Lane label only when both endpoints resolve to something readable.
        lane = (pickup_city + " → " + drop_city) if (pickup_city and drop_city) else ""

        item_name = str(cell(row, "item_names") or "").strip()
        category, subcategory = _product_category(item_name)

        records.append({
            "carrier": str(carrier).strip(),
            "account": str(cell(row, "account") or "").strip(),
            "delivery_type": str(cell(row, "delivery_type") or "").strip(),
            "zone": str(cell(row, "zone") or "").strip(),
            "pickup_pin": pickup_pin,
            "warehouse": warehouse_label(pickup_pin),
            "city": pickup_city,
            "drop_pin": drop_pin,
            "drop_city": drop_city,
            "lane": lane,
            "payment": payment,
            "weight": weight,
            "weight_class": _weight_class(weight),
            "pickup_date": pickup.date().isoformat() if pickup else "",
            "pickup_slot": _pickup_slot(pickup),
            "status": status,
            "outcome": _outcome(status),
            "pendency_state": _pendency_state(status),
            "item_name": item_name,
            "category": category,
            "subcategory": subcategory,
            "picked": pickup is not None,
            "delivered": status.lower() == "delivered",
            "attempts": attempts,
            "p2o": _hours_between(pickup, ofd1),
            "p2d": _hours_between(pickup, delivery),
        })
    return records


def _mean(values):
    clean = [v for v in values if v is not None]
    return sum(clean) / len(clean) if clean else None


def _filter_set(val):
    """Normalize a filter argument to a set of accepted values, or None for 'all'.

    Each categorical filter may arrive as 'all'/''/None (no filtering), a single
    string value, or a list/tuple/set of values (multi-select). Returning None
    means 'accept everything'; a set means 'accept only these'.
    """
    if val is None or val == "" or val == "all":
        return None
    if isinstance(val, (list, tuple, set)):
        picked = {v for v in val if v not in (None, "", "all")}
        return picked or None
    return {val}


def filter_records(records, delivery_type="Forward", zone="all", payment="all",
                   warehouse="all", account="all", weight="all",
                   slot="all", date_from="", date_to=""):
    # Every categorical filter supports single value, list (multi-select) or
    # "all". An empty selection (None) means no constraint on that field.
    dt_set = _filter_set(delivery_type)
    zone_set = _filter_set(zone)
    pay_set = _filter_set(payment)
    wh_set = _filter_set(warehouse)
    acct_set = _filter_set(account)
    wt_set = _filter_set(weight)
    slot_set = _filter_set(slot)

    out = []
    for r in records:
        if dt_set is not None and r["delivery_type"] not in dt_set:
            continue
        if zone_set is not None and r["zone"] not in zone_set:
            continue
        if pay_set is not None and r["payment"] not in pay_set:
            continue
        if wh_set is not None and r["pickup_pin"] not in wh_set:
            continue
        if acct_set is not None and r["account"] not in acct_set:
            continue
        if wt_set is not None and r["weight_class"] not in wt_set:
            continue
        if slot_set is not None and r["pickup_slot"] not in slot_set:
            continue
        # Date range is inclusive on both ends; rows with no pickup date are
        # dropped only when a bound is set.
        if date_from:
            if not r["pickup_date"] or r["pickup_date"] < date_from:
                continue
        if date_to:
            if not r["pickup_date"] or r["pickup_date"] > date_to:
                continue
        out.append(r)
    return out


def aggregate_by(records, key_field, label_field="group", extra_fields=None) -> list[dict]:
    """Group records by any field and compute the standard KPI block.

    extra_fields: optional list of (out_name, record_key) to carry the first
    seen value of a record field onto each group row (e.g. carry the city
    label onto a warehouse group).
    """
    extra_fields = extra_fields or []
    groups: dict[str, dict] = {}
    for r in records:
        key = r.get(key_field) or ""
        if key == "":
            continue
        g = groups.get(key)
        if g is None:
            g = {
                label_field: key, "n": 0, "picked": 0, "delivered": 0,
                "first_attempt": 0, "p2o": [], "p2d": [], "att": [],
            }
            for out_name, rec_key in extra_fields:
                g[out_name] = r.get(rec_key, "")
            groups[key] = g
        g["n"] += 1
        if r["picked"]:
            g["picked"] += 1
        if r["delivered"]:
            g["delivered"] += 1
            if r["attempts"] == 1:
                g["first_attempt"] += 1
        if r["p2o"] is not None:
            g["p2o"].append(r["p2o"])
        if r["p2d"] is not None:
            g["p2d"].append(r["p2d"])
        if r["attempts"] is not None:
            g["att"].append(r["attempts"])

    result = []
    for g in groups.values():
        row = {
            label_field: g[label_field],
            "n": g["n"],
            "picked": g["picked"],
            "delivered": g["delivered"],
            "success_rate": (g["delivered"] / g["picked"] * 100) if g["picked"] else None,
            "first_attempt_rate": (g["first_attempt"] / g["delivered"] * 100) if g["delivered"] else None,
            "p2o": _mean(g["p2o"]),
            "p2d": _mean(g["p2d"]),
            "avg_attempts": _mean(g["att"]),
        }
        for out_name, _ in extra_fields:
            row[out_name] = g.get(out_name, "")
        result.append(row)
    return result


def aggregate_by_carrier(records) -> list[dict]:
    return aggregate_by(records, "carrier", "carrier")


def _score_metric(value, all_values, lower_is_better):
    """Min-max normalize one metric to 0-100 across the carrier set."""
    pool = [v for v in all_values if v is not None]
    if value is None or len(pool) < 2:
        return None
    lo, hi = min(pool), max(pool)
    if hi == lo:
        return 100.0
    s = (value - lo) / (hi - lo)
    if lower_is_better:
        s = 1 - s
    return s * 100.0


def attach_scores(agg: list[dict], min_n: int = 0) -> list[dict]:
    # Only rows meeting the volume threshold participate in the normalization
    # pool and receive a score - this keeps a 3-shipment warehouse from
    # producing a meaningless rank.
    scorable = [a for a in agg if a["n"] >= min_n]
    pools = {
        "p2o": [a["p2o"] for a in scorable],
        "p2d": [a["p2d"] for a in scorable],
        "succ": [a["success_rate"] for a in scorable],
        "fa": [a["first_attempt_rate"] for a in scorable],
        "att": [a["avg_attempts"] for a in scorable],
    }
    lower_better = {"p2o": True, "p2d": True, "succ": False, "fa": False, "att": True}

    for a in agg:
        if a["n"] < min_n:
            a["subscores"] = {k: None for k in WEIGHTS}
            a["score"] = None
            continue
        raw = {
            "p2o": a["p2o"], "p2d": a["p2d"], "succ": a["success_rate"],
            "fa": a["first_attempt_rate"], "att": a["avg_attempts"],
        }
        sub = {k: _score_metric(raw[k], pools[k], lower_better[k]) for k in WEIGHTS}
        a["subscores"] = {k: (round(v, 1) if v is not None else None) for k, v in sub.items()}

        total_w, acc = 0.0, 0.0
        for k, w in WEIGHTS.items():
            if sub[k] is not None:
                acc += sub[k] * w
                total_w += w
        a["score"] = round(acc / total_w, 1) if total_w > 0 else None
    return agg


def _round(v, dp=1):
    return round(v, dp) if v is not None else None


def _round_metrics(agg: list[dict]) -> None:
    """Round the numeric KPI fields in place for JSON transport."""
    for a in agg:
        a["success_rate"] = _round(a["success_rate"])
        a["first_attempt_rate"] = _round(a["first_attempt_rate"])
        a["p2o"] = _round(a["p2o"])
        a["p2d"] = _round(a["p2d"])
        a["avg_attempts"] = _round(a["avg_attempts"], 2)


# Warehouses with fewer than this many shipments still appear in the table but
# are not assigned an efficiency score (too little data to rank fairly).
WAREHOUSE_MIN_N = 20

# Warehouse x carrier cells need fewer shipments to be meaningful than a whole
# warehouse, but we still suppress scores for very thin cells.
WH_CARRIER_MIN_N = 10

# Only the top-N warehouses by shipment volume are shown in the breakdown table
# (scores are still computed across ALL warehouses, so each rank is fair; the
# table is just truncated). Set to None to show every warehouse.
TOP_N_WAREHOUSES = 10

# Destination (drop city/region) and lane (pickup -> drop) breakdowns. Like the
# warehouse table, scores are computed across ALL rows then the table is
# truncated to the top-N by volume so each rank stays fair.
TOP_N_DESTINATIONS = 12
TOP_N_LANES = 15
# Destinations reuse the warehouse volume threshold; lanes are finer-grained so
# they get a slightly lower bar before a score is assigned.
DESTINATION_MIN_N = WAREHOUSE_MIN_N
LANE_MIN_N = 15


def aggregate_warehouse_carrier(records) -> list[dict]:
    """Warehouse x carrier matrix.

    One row per (pickup pincode, carrier) pair. Each row carries the warehouse
    label, city, carrier name and the standard KPI block. Scores are computed
    PER WAREHOUSE: within each warehouse the carriers are normalized against
    each other, so the score answers 'which partner is best out of THIS
    warehouse', which is exactly the comparison an ops lead wants.
    """
    rows = []
    # Bucket records by warehouse first.
    by_wh: dict[str, list] = {}
    for r in records:
        if not r["pickup_pin"]:
            continue
        by_wh.setdefault(r["pickup_pin"], []).append(r)

    for pin, recs in by_wh.items():
        agg = aggregate_by(
            recs, "carrier", "carrier",
            extra_fields=[("warehouse", "warehouse"), ("city", "city")],
        )
        # Score carriers relative to each other within this warehouse.
        attach_scores(agg, min_n=WH_CARRIER_MIN_N)
        _round_metrics(agg)
        wh_total = sum(a["n"] for a in agg)
        for a in agg:
            a["pickup_pin"] = pin
            a["wh_total"] = wh_total
            a["wh_share"] = _round(a["n"] / wh_total * 100 if wh_total else None)
        agg.sort(key=lambda a: (a["score"] is None, -(a["score"] or 0), -a["n"]))
        rows.extend(agg)
    return rows


def aggregate_status_matrix(records, key_field="account") -> list[dict]:
    """Per-carrier-account outcome matrix (mirrors the ops pivot).

    One row per carrier account, with the count and % share of each of the
    four outcomes (Delivered / FWD Pendency / RTO / Cancelled). Shares are of
    that account's own total, so each row's four percentages sum to 100.
    """
    groups: dict[str, dict] = {}
    for r in records:
        key = r.get(key_field) or ""
        if key == "":
            continue
        g = groups.get(key)
        if g is None:
            g = {"key": key, "n": 0, "counts": {o: 0 for o in OUTCOMES}}
            groups[key] = g
        g["n"] += 1
        g["counts"][r["outcome"]] += 1

    result = []
    for g in groups.values():
        n = g["n"]
        row = {"account": g["key"], "n": n, "counts": dict(g["counts"]), "pct": {}}
        for o in OUTCOMES:
            row["pct"][o] = _round(g["counts"][o] / n * 100 if n else None)
        result.append(row)
    result.sort(key=lambda a: -a["n"])
    return result


def aggregate_pendency_matrix(records, key_field="account") -> dict:
    """FWD-Pendency sub-state matrix, per carrier account.

    Only shipments whose outcome is FWD Pendency are counted. Columns are the
    PENDENCY_STATES (plus "Other"); each row's percentages are of that
    account's pendency total. Returns the rows plus the ordered column list so
    the frontend can render a stable header.
    """
    cols = PENDENCY_STATES + ["Other"]
    groups: dict[str, dict] = {}
    for r in records:
        if r["outcome"] != "FWD Pendency":
            continue
        key = r.get(key_field) or ""
        if key == "":
            continue
        g = groups.get(key)
        if g is None:
            g = {"key": key, "n": 0, "counts": {c: 0 for c in cols}}
            groups[key] = g
        g["n"] += 1
        g["counts"][r["pendency_state"]] += 1

    rows = []
    for g in groups.values():
        n = g["n"]
        row = {"account": g["key"], "n": n, "counts": dict(g["counts"]), "pct": {}}
        for c in cols:
            row["pct"][c] = _round(g["counts"][c] / n * 100 if n else None)
        rows.append(row)
    rows.sort(key=lambda a: -a["n"])
    # Drop columns that are entirely zero across all accounts to keep the table
    # readable (e.g. RiderAllocated rarely appears).
    active_cols = [c for c in cols
                   if any(rw["counts"].get(c, 0) > 0 for rw in rows)]
    return {"columns": active_cols, "rows": rows}


def build_report(records, delivery_type="Forward", zone="all", payment="all",
                 warehouse="all", account="all", weight="all",
                 slot="all", date_from="", date_to="") -> dict:
    """Top-level entry: filter, aggregate, score, and assemble the payload."""
    rows = filter_records(records, delivery_type, zone, payment,
                          warehouse, account, weight,
                          slot, date_from, date_to)
    agg = attach_scores(aggregate_by_carrier(rows))
    agg.sort(key=lambda a: (a["score"] is None, -(a["score"] or 0)))

    # Warehouse-wise breakdown, keyed on pickup pincode but labelled by city.
    # Same KPI block and scoring, with a volume threshold so the long tail of
    # tiny pincodes doesn't generate noisy ranks.
    wh = attach_scores(
        aggregate_by(rows, "pickup_pin", "pickup_pin",
                     extra_fields=[("warehouse", "warehouse"), ("city", "city")]),
        min_n=WAREHOUSE_MIN_N,
    )
    wh.sort(key=lambda a: -a["n"])  # warehouses ranked by volume
    _round_metrics(wh)
    wh_total_count = len(wh)
    if TOP_N_WAREHOUSES is not None:
        wh = wh[:TOP_N_WAREHOUSES]

    # Warehouse x carrier matrix (the new partner-per-warehouse view).
    wh_carrier = aggregate_warehouse_carrier(rows)

    # Status-outcome matrix per carrier account (the ops pivot), plus the
    # FWD-Pendency sub-state breakdown.
    status_matrix = aggregate_status_matrix(rows, "account")
    pendency_matrix = aggregate_pendency_matrix(rows, "account")

    # City-level rollup (a coarser breakdown than pincode).
    city_agg = attach_scores(
        aggregate_by(rows, "city", "city"), min_n=WAREHOUSE_MIN_N
    )
    city_agg.sort(key=lambda a: -a["n"])
    _round_metrics(city_agg)

    # Destination rollup: same KPI block + scoring, grouped by the drop city
    # (where the parcel is going), ranked by volume and truncated to top-N.
    dest_agg = attach_scores(
        aggregate_by(rows, "drop_city", "drop_city"), min_n=DESTINATION_MIN_N
    )
    dest_agg.sort(key=lambda a: -a["n"])
    _round_metrics(dest_agg)
    dest_total = len(dest_agg)

    # Lane rollup: pickup city -> drop city pairs. Highest-volume lanes first.
    lane_agg = attach_scores(
        aggregate_by(rows, "lane", "lane"), min_n=LANE_MIN_N
    )
    lane_agg.sort(key=lambda a: -a["n"])
    _round_metrics(lane_agg)
    lane_total = len(lane_agg)

    # The FULL destination/lane lists are sent (not truncated) so the client
    # can search across every entry; the UI shows the top-N by volume until the
    # user types in the panel's search box. subscores aren't shown in these
    # tables, so drop them to keep the payload lean (these can be 1000s of rows).
    for _row in dest_agg:
        _row.pop("subscores", None)
    for _row in lane_agg:
        _row.pop("subscores", None)

    total = len(rows)
    delivered = sum(1 for r in rows if r["delivered"])
    picked = sum(1 for r in rows if r["picked"])

    # Business mix breakdowns.
    def count_by(key_fn):
        out: dict[str, int] = {}
        for r in rows:
            k = key_fn(r)
            if k:
                out[k] = out.get(k, 0) + 1
        return out

    weight_mix = count_by(lambda r: _weight_class(r["weight"]))
    weight_order = ["Light", "Medium", "Heavy", "Unknown"]
    weight_mix = {k: weight_mix[k] for k in weight_order if k in weight_mix}

    # Product breakdown: category -> subcategory volume tree, plus which carrier
    # accounts ship each category. Only meaningful when the file carries an Item
    # Names column; otherwise everything is "Unknown".
    product_tree: dict[str, dict] = {}
    has_products = False
    for r in rows:
        cat = r.get("category") or "Unknown"
        sub = r.get("subcategory") or "Unknown"
        acct = r.get("account") or "(unknown)"
        if r.get("item_name"):
            has_products = True
        c = product_tree.setdefault(
            cat, {"category": cat, "n": 0, "subs": {}, "accts": {}}
        )
        c["n"] += 1
        c["subs"][sub] = c["subs"].get(sub, 0) + 1
        c["accts"][acct] = c["accts"].get(acct, 0) + 1
    products = []
    for c in sorted(product_tree.values(), key=lambda x: -x["n"]):
        subs = [
            {"subcategory": s, "n": n}
            for s, n in sorted(c["subs"].items(), key=lambda kv: -kv[1])
        ]
        accts = [
            {"account": a, "n": n}
            for a, n in sorted(c["accts"].items(), key=lambda kv: -kv[1])
        ]
        products.append({
            "category": c["category"], "n": c["n"],
            "subs": subs, "accounts": accts,
        })

    _round_metrics(agg)

    # Filter option lists are built from the FULL record set (not the filtered
    # rows) so selecting one value doesn't empty out the other dropdowns.
    zones = sorted({r["zone"] for r in records if r["zone"]})
    # Carrier-account options for the dropdown, ordered by volume.
    acct_counts: dict[str, int] = {}
    for r in records:
        a = r["account"]
        if a:
            acct_counts[a] = acct_counts.get(a, 0) + 1
    accounts = [
        {"value": a, "label": a, "n": cnt}
        for a, cnt in sorted(acct_counts.items(), key=lambda kv: -kv[1])
    ]
    # Pickup date range (ISO strings) to bound the date pickers.
    pickup_dates = [r["pickup_date"] for r in records if r["pickup_date"]]
    date_min = min(pickup_dates) if pickup_dates else ""
    date_max = max(pickup_dates) if pickup_dates else ""
    # Warehouse options: pincode value + city label, ordered by volume.
    wh_counts: dict[str, int] = {}
    wh_labels: dict[str, str] = {}
    for r in records:
        if r["pickup_pin"]:
            wh_counts[r["pickup_pin"]] = wh_counts.get(r["pickup_pin"], 0) + 1
            wh_labels[r["pickup_pin"]] = r["warehouse"]
    warehouses_opts = [
        {"value": pin, "label": wh_labels[pin], "n": cnt}
        for pin, cnt in sorted(wh_counts.items(), key=lambda kv: -kv[1])
    ]
    # Cap the filter-bar dropdown to the same top-N as the breakdown table.
    if TOP_N_WAREHOUSES is not None:
        warehouses_opts = warehouses_opts[:TOP_N_WAREHOUSES]

    return {
        "summary": {
            "total": total,
            "picked": picked,
            "delivered": delivered,
            "success_rate": _round(delivered / picked * 100 if picked else None),
            "avg_p2o": _round(_mean([r["p2o"] for r in rows])),
            "avg_p2d": _round(_mean([r["p2d"] for r in rows])),
            "carriers": len(agg),
            "warehouses": wh_total_count,
        },
        "carriers": agg,
        "products": products,
        "has_products": has_products,
        "warehouses": wh,
        "warehouse_total": wh_total_count,
        "warehouse_top_n": TOP_N_WAREHOUSES,
        "warehouse_carrier": wh_carrier,
        "status_matrix": status_matrix,
        "pendency_matrix": pendency_matrix,
        "outcomes": OUTCOMES,
        "cities": city_agg,
        "destinations": dest_agg,
        "destination_total": dest_total,
        "destination_top_n": TOP_N_DESTINATIONS,
        "lanes": lane_agg,
        "lane_total": lane_total,
        "lane_top_n": TOP_N_LANES,
        "lane_min_n": LANE_MIN_N,
        "warehouse_min_n": WAREHOUSE_MIN_N,
        "wh_carrier_min_n": WH_CARRIER_MIN_N,
        "mix": {
            "weight": weight_mix,
            "payment": count_by(lambda r: r["payment"]),
            "delivery_type": count_by(lambda r: r["delivery_type"]),
        },
        "filters": {
            "zones": zones,
            "accounts": accounts,
            "warehouses": warehouses_opts,
            "weight_classes": ["Light", "Medium", "Heavy"],
            "pickup_slots": PICKUP_SLOTS,
            "date_min": date_min,
            "date_max": date_max,
        },
        "weights": {k: int(v * 100) for k, v in WEIGHTS.items()},
    }