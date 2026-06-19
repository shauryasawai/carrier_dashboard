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
BLUE DART (Surface, B2C)   - unit = hours
    Matched by carrier name (any Blue Dart account). Destination-based flat
    lookup {dest_pincode: tat_hours} from the Surface serviceability list
    (origin TWH / Talegaon); surface transit is typically 3-8 days. Replaces
    the earlier Apex-Air ~24h lookup, which mis-scored surface volume.

DELHIVERY NDD (express)    - unit = days
    Matched by carrier name "Delhivery" AND account containing "NDD" (so
    Surface / Reverse / heavy Delhivery accounts are NOT caught). Lookup keyed
    warehouse -> {dest_pincode: tat_days}. NDD-serviceable pincodes = 1 day;
    longer lanes use the EXPRESS TAT (days). Source: the DLV NDD pincode +
    "NDD +1 & 2" files. Covers Pune (CAH), Bangalore (BLR), Gurugram (GGN).
    Any matched Delhivery NDD lane NOT in the lookup falls back to a 3-day
    promise (default_tat=3) rather than "No rule".

ELASTIC RUN (SDD/NDD)      - unit = days, DESTINATION-based
    Matched by ACCOUNT containing "elasticrun" (the carrier column is generic,
    the account code carries the identity). Hyperlocal same-day/next-day, so the
    promised TAT is a property of the DROP pincode, not the pickup hub: the
    lookup is a FLAT {dest_pincode: tat_days} map (warehouse_agnostic=True, the
    pickup->warehouse step is skipped). Every serviceable pincode (SDD or NDD)
    promises 1 day from pickup (it's a next-day account). Covers all 10 ER cities
    (Bangalore, Mumbai, Delhi, Pune, Kolkata, Hyderabad, Surat, Jaipur,
    Ahmedabad, Chennai) plus the intercity lanes. Source: the ER UPDATED
    PINCODE MASTER file.

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
# when its carrier name contains any `carrier_contains` token (skipped if that
# is None) AND, if `account_contains` is set, its account contains one of those
# tokens too. At least one of the two must be set.
#   unit               : "hours" or "days" (how the lookup's TAT values are read)
#   payment_split      : lookup is warehouse -> payment -> {pin: tat} (else warehouse -> {pin: tat})
#   warehouse_agnostic : lookup is a FLAT {dest_pin: tat} keyed by drop pincode
#                        only; the pickup->warehouse step is skipped (for
#                        hyperlocal carriers whose SLA depends on destination).
#   default_tat        : fallback promised TAT (in the carrier's unit) applied
#                        when the lookup has no rule for a lane, so a matched
#                        carrier's unmapped lanes are still scored instead of
#                        falling to "No rule".
#   default_file       : a secondary warehouse-keyed {pin: tat} lookup consulted
#                        when the primary lookup misses a lane (e.g. GoSwift's
#                        B2C city matrix), before falling through to No rule.
#   city_default_file  : last-resort {pickup_city: {drop_city: tat}} lookup used
#                        when even the pincode lookups miss (e.g. a pickup that
#                        never resolved to a warehouse hub). Keyed on city names.
# ---------------------------------------------------------------------------
_CARRIERS = [
    {
        # Bluedart SURFACE (ground). Destination-based: promised TAT (hours) per
        # drop pincode from the Surface serviceability list (origin TWH /
        # Talegaon). Replaces the old Apex-Air ~24h lookup, which wrongly scored
        # multi-day surface volume as almost entirely Out of TAT.
        "name": "Bluedart", "file": "bd_surface_tat.json.gz",
        "carrier_contains": ["blue", "dart"], "account_contains": None,
        "exclude_contains": ["reverse", "rvp"],
        "unit": "hours", "payment_split": False, "warehouse_agnostic": True,
    },
    {
        # Delhivery NDD (express). Lanes in the lookup use the EXPRESS TAT
        # (days); any matched Delhivery NDD lane NOT in the lookup falls back to
        # a 3-day promise (3 calendar days from pickup) instead of "No rule".
        "name": "Delhivery NDD", "file": "dlv_tat.json.gz",
        "carrier_contains": ["delhivery"], "account_contains": ["ndd"],
        "exclude_contains": ["reverse"],
        "unit": "days", "payment_split": False, "default_tat": 3,
    },
    {
        # Swift NDD (GoSwift next-day). More specific than generic GoSwift, so
        # it must come first: matched when carrier is Swift/GoSwift AND the
        # account contains "ndd". Every serviceable lane promises 1 day; lanes
        # not in the NDD master are "No rule" (not an NDD lane).
        "name": "Swift NDD", "file": "swift_ndd_tat.json.gz",
        "carrier_contains": ["swift", "goswift"], "account_contains": ["ndd"],
        "exclude_contains": ["reverse", "revers"],
        "unit": "days", "payment_split": False,
    },
    {
        # GoSwift, remaining forward categories (GoSwift Forward, Swift
        # Mobility) after Swift NDD is peeled off above. Reverse excluded.
        # SLA from Swift_TAT_lanes;
        # lanes missing there fall back to the B2C city-matrix (default_file)
        # so far fewer GoSwift lanes end up "No rule".
        "name": "GoSwift", "file": "swift_tat.json.gz",
        "carrier_contains": ["swift", "goswift"], "account_contains": None,
        "exclude_contains": ["reverse", "revers"],
        "unit": "days", "payment_split": False,
        "default_file": "swift_default_tat.json.gz",
        "city_default_file": "swift_city_tat.json.gz",
    },
    {
        # Elastic Run, SDD/NDD next-day account. Identified by the ACCOUNT code
        # (carrier column is generic), so carrier_contains is None and the match
        # is account-only. SLA from the ER pincode master: every serviceable
        # pincode (SDD or NDD) promises 1 day from pickup.
        "name": "Elastic Run", "file": "er_tat.json.gz",
        "carrier_contains": None, "account_contains": ["elasticrun", "elastic run"],
        "exclude_contains": ["reverse", "revers", "rvp"],
        "unit": "days", "payment_split": False, "warehouse_agnostic": True,
    },
    {
        # Skye Air, drone SAME-DAY delivery. Account-matched ("skye"); like
        # Elastic Run it is destination-based (warehouse_agnostic). Every
        # serviceable pincode promises 0 days (delivered the same calendar day).
        "name": "Skye Air", "file": "skye_tat.json.gz",
        "carrier_contains": None, "account_contains": ["skye", "sky air"],
        "exclude_contains": ["reverse", "revers", "rvp"],
        "unit": "days", "payment_split": False, "warehouse_agnostic": True,
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
        car_toks = ent.get("carrier_contains")
        acc_toks = ent.get("account_contains")
        # A carrier-name match is required only when carrier_contains is set;
        # otherwise the entry is matched purely on its account tokens.
        if car_toks and not any(tok in cl for tok in car_toks):
            continue
        if acc_toks and not any(tok in al for tok in acc_toks):
            continue
        if not car_toks and not acc_toks:
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


def _lookup_tat(lookup, ent, pickup_pin, payment, drop_pin):
    """Resolve a promised TAT from one consolidated lookup, or None."""
    if not lookup:
        return None
    if ent.get("warehouse_agnostic"):
        # Flat {dest_pin: tat} keyed by drop pincode; pickup hub is irrelevant.
        table = lookup
    else:
        wh_node = lookup.get(warehouse_for_pin(pickup_pin))
        if not wh_node:
            return None
        if ent.get("payment_split"):
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


# ---------------------------------------------------------------------------
# User SLA overrides. The dashboard's SLA editor can REPLACE a carrier's entire
# promised-TAT table at runtime, either by typing rows or uploading a file of
# (Warehouse, Pincode, TAT). Overrides live in memory and are persisted
# best-effort to data/overrides/<key>.json.gz so they survive a restart.
# Shape: {carrier_key: {"unit": str, "rows": [{warehouse,pincode,tat}],
#                       "lookup": {WAREHOUSE: {pincode: tat}}}}.
# A warehouse of "ANY" matches every origin.
# ---------------------------------------------------------------------------
_OVERRIDE_DIR = os.path.join(_DATA_DIR, "overrides")
_OVERRIDES: dict = {}
_OVERRIDES_LOADED = False


def carrier_key(name: str) -> str:
    """Stable filesystem-safe key for a carrier name."""
    return "".join(c if c.isalnum() else "_" for c in str(name).lower()).strip("_")


def _ensure_overrides_loaded():
    global _OVERRIDES_LOADED
    if _OVERRIDES_LOADED:
        return
    _OVERRIDES_LOADED = True
    try:
        files = os.listdir(_OVERRIDE_DIR)
    except OSError:
        return
    for fn in files:
        if not fn.endswith(".json.gz"):
            continue
        try:
            with gzip.open(os.path.join(_OVERRIDE_DIR, fn), "rb") as fh:
                _OVERRIDES[fn[:-len(".json.gz")]] = json.loads(fh.read().decode("utf-8"))
        except (OSError, ValueError):
            continue


def _norm_tat(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return int(f) if f == int(f) else f


def _override_lookup_from_rows(rows):
    """Build ({origin_pin: {dest_pin: tat}}, cleaned_rows) from raw rows of
    (origin_pin, pincode, tat). origin_pin is the exact pickup/warehouse pincode
    (6 digits) or 'ANY' to apply to every origin."""
    lk, clean = {}, []
    for row in rows:
        raw = str(row.get("origin_pin") or "ANY").strip().upper()
        if raw in ("", "ANY"):
            origin = "ANY"
        else:
            d = "".join(ch for ch in raw if ch.isdigit())
            origin = d if len(d) == 6 else None
        pin = "".join(ch for ch in str(row.get("pincode") or "") if ch.isdigit())
        tat = _norm_tat(row.get("tat"))
        if origin is None or len(pin) != 6 or tat is None:
            continue
        lk.setdefault(origin, {})[pin] = tat
        clean.append({"origin_pin": origin, "pincode": pin, "tat": tat})
    return lk, clean


def set_override(name: str, rows) -> int:
    """Replace a carrier's SLA with (warehouse, pincode, tat) rows. Returns count."""
    ent = next((e for e in _CARRIERS if e["name"] == name), None)
    if ent is None:
        raise ValueError("Unknown carrier: " + str(name))
    _ensure_overrides_loaded()
    key = carrier_key(name)
    lk, clean = _override_lookup_from_rows(rows)
    _OVERRIDES[key] = {"unit": ent["unit"], "rows": clean, "lookup": lk}
    try:
        os.makedirs(_OVERRIDE_DIR, exist_ok=True)
        with gzip.open(os.path.join(_OVERRIDE_DIR, key + ".json.gz"), "wb") as fh:
            fh.write(json.dumps(_OVERRIDES[key], separators=(",", ":")).encode("utf-8"))
    except OSError:
        pass  # in-memory override still applies for this process
    return len(clean)


def clear_override(name: str):
    """Drop a carrier's override and revert to the built-in SLA."""
    _ensure_overrides_loaded()
    key = carrier_key(name)
    _OVERRIDES.pop(key, None)
    try:
        os.remove(os.path.join(_OVERRIDE_DIR, key + ".json.gz"))
    except OSError:
        pass


def get_override_rows(name: str) -> list:
    _ensure_overrides_loaded()
    ov = _OVERRIDES.get(carrier_key(name))
    return list(ov["rows"]) if ov else []


def _override_for(entry_idx: int):
    _ensure_overrides_loaded()
    return _OVERRIDES.get(carrier_key(_CARRIERS[entry_idx]["name"]))


def carriers_meta() -> list:
    """Per-carrier info for the SLA editor: name, key, unit, override state."""
    _ensure_overrides_loaded()
    out = []
    for e in _CARRIERS:
        ov = _OVERRIDES.get(carrier_key(e["name"]))
        out.append({
            "name": e["name"], "key": carrier_key(e["name"]), "unit": e["unit"],
            "has_override": ov is not None,
            "rule_count": len(ov["rows"]) if ov else 0,
        })
    return out


def _promised(entry_idx: int, pickup_pin: str, payment: str, drop_pin: str):
    """Promised TAT (in the carrier's unit) for this lane, or None.

    If a user override exists for the carrier it REPLACES the built-in SLA
    entirely. Otherwise tries the primary lookup, then a `default_file` matrix.
    """
    ent = _CARRIERS[entry_idx]

    ov = _override_for(entry_idx)
    if ov is not None:
        lk = ov["lookup"]
        drop = "".join(ch for ch in str(drop_pin or "") if ch.isdigit())
        pk = "".join(ch for ch in str(pickup_pin or "") if ch.isdigit())
        # Match the exact pickup pincode, then any "ANY"-origin rules.
        for w in (pk, "ANY"):
            node = lk.get(w)
            if node and drop in node:
                return node[drop]
        return None  # replace mode: lanes not in the override are unscored

    val = _lookup_tat(_load_lookup(ent["file"]), ent, pickup_pin, payment, drop_pin)
    if val is not None:
        return val

    dfile = ent.get("default_file")
    if dfile:
        # The default matrix is plain warehouse -> {pin: tat} (no payment split,
        # not destination-only), regardless of the primary lookup's shape.
        plain = {"unit": ent["unit"]}
        return _lookup_tat(_load_lookup(dfile), plain, pickup_pin, payment, drop_pin)
    return None


# City-name normalisation for the city-to-city fallback. Pickup cities resolved
# from a pincode (e.g. "Bengaluru", "Howrah") are aliased to the names the lane
# table uses ("BANGALORE", "KOLKATA", NCR hubs -> "DELHI").
_PICKUP_CITY_ALIAS = {
    "BENGALURU": "BANGALORE", "BENGALURU (RURAL)": "BANGALORE",
    "GURUGRAM": "DELHI", "GURGAON": "DELHI", "NOIDA": "DELHI",
    "GHAZIABAD": "DELHI", "FARIDABAD": "DELHI", "NCR": "DELHI",
    "HOWRAH": "KOLKATA",
    "THANE": "MUMBAI", "NAVI MUMBAI": "MUMBAI", "BHIWANDI": "MUMBAI",
}


def _norm_city(s: str) -> str:
    return " ".join(str(s or "").strip().upper().split())


def _city_promised(entry_idx: int, pickup_city: str, drop_city: str):
    """Last-resort TAT from a {pickup_city: {drop_city: tat}} lane table."""
    ent = _CARRIERS[entry_idx]
    cfile = ent.get("city_default_file")
    if not cfile or not pickup_city or not drop_city:
        return None
    pu = _norm_city(pickup_city)
    pu = _PICKUP_CITY_ALIAS.get(pu, pu)
    node = _load_lookup(cfile).get(pu)
    if not node:
        return None
    return node.get(_norm_city(drop_city))


def classify(carrier, account, pickup_pin, payment, drop_pin,
             p2d_hours, transit_days, delivered,
             age_hours=None, age_days=None, forward_pending=False,
             rto=False, ofd1_hours=None, ofd1_days=None,
             pickup_city=None, drop_city=None):
    """Return (tat_status, promised, margin).

    `promised` is in the matched carrier's unit (hours or days). `margin` is
    promised - actual in that same unit (positive => on time / early).
      - hours carriers compare elapsed pickup->delivery hours (p2d_hours)
      - days  carriers compare calendar days (transit_days)

    Late forward pendency (all carriers): an undelivered shipment that is still
    an active forward-pendency (in-transit / out-for-delivery / delayed, i.e.
    NOT delivered, RTO or cancelled) and whose elapsed age since pickup has
    ALREADY exceeded the promised TAT is a confirmed breach -> "Out of TAT",
    even though it hasn't been delivered yet. `age_hours` / `age_days` are the
    pickup->now elapsed time (hours for hour-based carriers, calendar days for
    day-based ones) and `forward_pending` marks active forward pendency.
    Undelivered shipments still inside the promised window stay "Pending".

    RTO (all carriers): a returned shipment never reaches the customer, but the
    carrier DID attempt delivery, so its TAT is judged on pickup->OFD1 (first
    out-for-delivery attempt) vs the promise: In TAT if the first attempt was on
    time, else Out of TAT. `ofd1_hours` / `ofd1_days` carry that elapsed time;
    if it's missing the RTO is left "Pending" (can't be judged).
    """
    idx = _match_carrier(carrier or "", account or "")
    if idx is None:
        return ("No rule", None, None)

    promised = _promised(idx, pickup_pin, payment, drop_pin)
    # Built-in fallbacks (default TAT, city-to-city) apply only when the carrier
    # is NOT under a user override; an override replaces the SLA entirely.
    overridden = _override_for(idx) is not None
    if promised is None and not overridden:
        # No per-lane rule; fall back to the carrier's default TAT if it has one
        # (e.g. Delhivery NDD -> 3 days), otherwise the lane is unscored.
        promised = _CARRIERS[idx].get("default_tat")
    if promised is None and not overridden:
        # Still nothing: try the city-to-city lane table (e.g. GoSwift) so a
        # pickup that never resolved to a warehouse can still be scored.
        promised = _city_promised(idx, pickup_city, drop_city)
    if promised is None:
        return ("No rule", None, None)

    unit = _CARRIERS[idx]["unit"]
    actual = p2d_hours if unit == "hours" else transit_days
    if delivered and actual is not None:
        margin = promised - actual
        status = "In TAT" if actual <= promised else "Out of TAT"
        return (status, promised, margin)

    # RTO: returned, never delivered. Score the carrier on its first delivery
    # attempt (pickup->OFD1) vs the promise instead of leaving it unscored.
    if rto:
        ofd = ofd1_hours if unit == "hours" else ofd1_days
        if ofd is not None:
            status = "In TAT" if ofd <= promised else "Out of TAT"
            return (status, promised, promised - ofd)
        return ("Pending", promised, None)

    # Not yet delivered. If it's active forward pendency already past its SLA,
    # the breach is certain regardless of eventual delivery -> Out of TAT.
    if forward_pending:
        elapsed = age_hours if unit == "hours" else age_days
        if elapsed is not None and elapsed > promised:
            return ("Out of TAT", promised, promised - elapsed)

    return ("Pending", promised, None)
