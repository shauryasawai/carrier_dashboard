"""
Promised-TAT (SLA) lookup and In-TAT / Out-of-TAT classification.

Carriers publish a serviceability / TAT master that gives, for an (origin hub,
destination pincode) pair, the promised transit time. We consolidate those
masters into compact lookups (``data/<carrier>_tat.json.gz``) and use them to
label each delivered shipment:

    In TAT      - delivered within the promised TAT
    Out of TAT  - delivered, but slower than the promised TAT
    No rule     - no promised TAT for this lane (origin/dest not mapped)
    Pending     - not yet delivered, so compliance can't be evaluated

Two SLA "units" are supported, because carriers express TAT differently:
    hours  - compare elapsed pickup->delivery hours (``p2d``) vs promised hours
    days   - compare calendar days (delivery date - pickup date) vs promised days

---------------------------------------------------------------------------
Carriers wired up
---------------------------------------------------------------------------
BLUE DART (Apex air, B2C)  - unit = hours
    Matched by carrier name (any Blue Dart account). Lookup keyed
    warehouse -> payment -> {dest_pincode: tat_hours}. COD/Prepaid TAT only
    differ for Hyderabad. Source: the "BD Air" Apex per-warehouse files.

DELHIVERY NDD (express)    - unit = days
    Matched by carrier name "Delhivery" AND account containing "NDD" (so
    Surface / Reverse / heavy Delhivery accounts are NOT caught). Lookup keyed
    warehouse -> {dest_pincode: tat_days}. NDD-serviceable pincodes = 1 day;
    longer lanes use the EXPRESS TAT (days). Source: the DLV NDD pincode +
    "NDD +1 & 2" files. Covers Pune (CAH), Bangalore (BLR), Gurugram (GGN).

Add a carrier: drop its consolidated gz in data/ and add an entry to _CARRIERS.
"""

from __future__ import annotations

import gzip
import json
import os
from functools import lru_cache

_DATA_DIR = os.path.join(os.path.dirname(__file__), "data")

# Compliance status labels (order = display order).
TAT_STATUSES = ["In TAT", "Out of TAT", "No rule", "Pending"]


# ---------------------------------------------------------------------------
# Carrier registration. First matching entry wins. A shipment matches an entry
# when its carrier name contains any `carrier_contains` token AND, if
# `account_contains` is set, its account contains one of those tokens too.
#   unit          : "hours" or "days" (how the lookup's TAT values are read)
#   payment_split : lookup is warehouse -> payment -> {pin: tat} (else warehouse -> {pin: tat})
# ---------------------------------------------------------------------------
_CARRIERS = [
    {
        "name": "Bluedart", "file": "bd_tat.json.gz",
        "carrier_contains": ["blue", "dart"], "account_contains": None,
        "exclude_contains": ["reverse", "rvp"],
        "unit": "hours", "payment_split": True,
    },
    {
        "name": "Delhivery NDD", "file": "dlv_tat.json.gz",
        "carrier_contains": ["delhivery"], "account_contains": ["ndd"],
        "exclude_contains": ["reverse"],
        "unit": "days", "payment_split": False,
    },
    {
        # GoSwift, all forward categories (GoSwift Forward, Swift_NDD, Swift
        # Mobility). Reverse accounts are excluded. SLA from Swift_TAT_lanes.
        "name": "GoSwift", "file": "swift_tat.json.gz",
        "carrier_contains": ["swift", "goswift"], "account_contains": None,
        "exclude_contains": ["reverse", "revers"],
        "unit": "days", "payment_split": False,
    },
]


@lru_cache(maxsize=8)
def _load_lookup(filename: str) -> dict:
    """Load and cache a consolidated TAT lookup. Returns {} if missing/unreadable
    so the dashboard degrades to 'No rule' rather than erroring."""
    path = os.path.join(_DATA_DIR, filename)
    try:
        with gzip.open(path, "rb") as fh:
            payload = json.loads(fh.read().decode("utf-8"))
        return payload.get("lookup", {})
    except (OSError, ValueError):
        return {}


@lru_cache(maxsize=256)
def _match_carrier(carrier: str, account: str):
    """Return the _CARRIERS entry (as a tuple key) that applies, or None."""
    cl = (carrier or "").strip().lower()
    al = (account or "").strip().lower()
    for idx, ent in enumerate(_CARRIERS):
        if not any(tok in cl for tok in ent["carrier_contains"]):
            continue
        acc_toks = ent.get("account_contains")
        if acc_toks and not any(tok in al for tok in acc_toks):
            continue
        excl = ent.get("exclude_contains")
        if excl and any(tok in cl or tok in al for tok in excl):
            continue
        return idx
    return None


# ---------------------------------------------------------------------------
# Pickup pincode -> warehouse code (shared across carriers).
# Explicit entries cover the exact pickup pincodes seen in the data; the
# 3-digit prefix map is a safety net. Codes not present in a carrier's lookup
# (e.g. Blue Dart has no BLR/MAA) yield "No rule".
# ---------------------------------------------------------------------------
WAREHOUSE_FOR_PIN = {
    "412106": "CAH", "412105": "CAH", "411062": "CAH",   # Pune / Chakan
    "421302": "BOM", "410507": "BOM",                    # Mumbai / Bhiwandi
    "122001": "GGN", "122503": "GGN", "122506": "GGN",   # Gurugram
    "712223": "HOG",                                     # Howrah / Hooghly
    "501401": "HYD", "500078": "HYD",                    # Hyderabad
    "562114": "BLR", "562123": "BLR",                    # Bengaluru
    "600052": "MAA",                                     # Chennai
}

WAREHOUSE_PREFIX = {
    "411": "CAH", "412": "CAH",
    "400": "BOM", "401": "BOM", "410": "BOM", "421": "BOM",
    "121": "GGN", "122": "GGN",
    "711": "HOG", "712": "HOG",
    "500": "HYD", "501": "HYD", "502": "HYD",
    "560": "BLR", "561": "BLR", "562": "BLR",
    "600": "MAA", "601": "MAA", "602": "MAA", "603": "MAA",
}


@lru_cache(maxsize=100_000)
def warehouse_for_pin(pin: str) -> str:
    """Resolve a pickup pincode to its Frido warehouse code (or '' if unknown)."""
    if not pin:
        return ""
    digits = "".join(ch for ch in str(pin) if ch.isdigit())
    if len(digits) < 3:
        return ""
    if digits in WAREHOUSE_FOR_PIN:
        return WAREHOUSE_FOR_PIN[digits]
    return WAREHOUSE_PREFIX.get(digits[:3], "")


def _promised(entry_idx: int, pickup_pin: str, payment: str, drop_pin: str):
    """Promised TAT (in the carrier's unit) for this lane, or None."""
    ent = _CARRIERS[entry_idx]
    lookup = _load_lookup(ent["file"])
    wh = warehouse_for_pin(pickup_pin)
    wh_node = lookup.get(wh)
    if not wh_node:
        return None

    if ent["payment_split"]:
        pay = (payment or "").strip().upper()
        table = wh_node.get(pay) or wh_node.get("PREPAID")
        if table is None:
            table = next(iter(wh_node.values()), None)
    else:
        table = wh_node
    if not table:
        return None

    drop = "".join(ch for ch in str(drop_pin or "") if ch.isdigit())
    return table.get(drop)


def classify(carrier, account, pickup_pin, payment, drop_pin,
             p2d_hours, transit_days, delivered):
    """Return (tat_status, promised, margin).

    `promised` is in the matched carrier's unit (hours or days). `margin` is
    promised - actual in that same unit (positive => on time / early).
      - hours carriers compare elapsed pickup->delivery hours (p2d_hours)
      - days  carriers compare calendar days (transit_days)
    """
    idx = _match_carrier(carrier or "", account or "")
    if idx is None:
        return ("No rule", None, None)

    promised = _promised(idx, pickup_pin, payment, drop_pin)
    if promised is None:
        return ("No rule", None, None)

    unit = _CARRIERS[idx]["unit"]
    actual = p2d_hours if unit == "hours" else transit_days
    if not delivered or actual is None:
        return ("Pending", promised, None)

    margin = promised - actual
    status = "In TAT" if actual <= promised else "Out of TAT"
    return (status, promised, margin)
