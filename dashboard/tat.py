"""
Promised-TAT (SLA) lookup and In-TAT / Out-of-TAT classification.

Carriers publish a serviceability / TAT master that gives, for an (origin hub,
destination pincode) pair, the promised transit time. We consolidate those
masters into compact lookups (``data/<carrier>_tat.json.gz``) and use them to
label each delivered shipment:

Compliance is measured, for every carrier, against that carrier's OWN committed
delivery date (EDD = expected_delivery_date_by_courier_partner):

    In TAT      - delivered on/before the committed EDD
    Out of TAT  - delivered after the EDD (or already past EDD while undelivered)
    No rule     - shipment carrier isn't a tracked carrier
    Pending     - not delivered yet and still within the EDD window

When a shipment has no EDD (e.g. Skye Air) or its carrier is under a manual SLA
override, scoring falls back to the carrier's committed-TAT lookup files
(pickup->OFD1, or pickup->delivery for measure_on="delivery"); those use either
hours or calendar days per the carrier's unit.

---------------------------------------------------------------------------
Carriers wired up
---------------------------------------------------------------------------
BLUE DART (Surface, B2C)   - unit = hours
    Matched by carrier name (any Blue Dart account). Destination-based flat
    lookup {dest_pincode: tat_hours} from the Surface serviceability list
    (origin TWH / Talegaon); surface transit is typically 3-8 days. Replaces
    the earlier Apex-Air ~24h lookup, which mis-scored surface volume.

DELHIVERY NDD              - unit = days, scored pickup->DELIVERY (NSL basis)
    Matched by carrier name "Delhivery" AND account containing "NDD" (so
    Surface / Reverse / heavy Delhivery accounts are NOT caught). SLA = the
    PITCHED per-lane EXPRESS TAT (days), keyed warehouse -> {dest_pincode: days}
    from "NDD +1 & 2 for PUNE & BLR" + "DLV NDD Pin Codes" gap-fill. EDD is
    ignored (Delhivery pads it). Lanes off the master fall back to 3 days.

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


def _digits(value) -> str:
    """Digits-only string of `value` (drops any non-numeric characters)."""
    return "".join(ch for ch in str(value or "") if ch.isdigit())


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
#   measure_on         : "ofd1" (default) scores pickup->first-attempt (OFD1);
#                        "delivery" scores pickup->delivery (e.g. Delhivery NDD,
#                        to match its NSL on-time metric).
#   ignore_edd         : skip the universal committed-EDD basis for this carrier
#                        and use its lookup files instead (e.g. Delhivery NDD,
#                        whose courier EDD is padded ~100%; its zone SLA tracks
#                        NSL ~90%).
#   ofd1_grace_hours   : (OFD1 basis) also In TAT if OFD1 is within this many
#                        elapsed hours of pickup, in addition to the promised
#                        rule (e.g. Skye Air: same day OR within 15h).
#   tat_offset_days    : add N days to every promised TAT for this carrier
#                        (e.g. Shadowfax "test" held to NDD+1 -> +1).
# NOTE: In TAT is scored against the carrier's committed EDD by default; the
# lookup files (pickup->OFD1, or ->delivery for measure_on="delivery") are the
# fallback when EDD is missing, overridden, or ignore_edd is set (see classify()).
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
        # Delhivery NDD. SLA = Delhivery's PITCHED per-lane EXPRESS TAT (days),
        # keyed warehouse -> {dest_pincode: days}, from "NDD +1 & 2 for PUNE &
        # BLR" (authoritative) + "DLV NDD Pin Codes" (next-day gap-fill). Scored
        # pickup->DELIVERY. EDD is ignored here because Delhivery pads it (~3d vs
        # ~1.7d actual -> a meaningless ~100%); the pitched lane TAT is the
        # contractual commitment. Lanes off the master fall back to 3 days.
        "name": "Delhivery NDD", "file": "dlv_tat.json.gz",
        "carrier_contains": ["delhivery"], "account_contains": ["ndd"],
        "exclude_contains": ["reverse"],
        "unit": "days", "payment_split": False,
        "default_tat": 3, "measure_on": "delivery", "ignore_edd": True,
    },
    {
        # Swift NDD (GoSwift next-day). Matched before generic GoSwift (carrier
        # Swift/GoSwift AND account contains "ndd"). SLA = pitched zone-committed
        # TAT (A=1, B=2, C=3, D=4 days) keyed warehouse -> {dest_pincode: days}
        # from "NDD-GoSwift.xlsx". Scored pickup->DELIVERY; EDD ignored because
        # GoSwift pads it (~6d vs ~2.1d actual -> a meaningless ~98%). Lanes off
        # the master fall back to 3 days.
        "name": "Swift NDD", "file": "swift_ndd_tat.json.gz",
        "carrier_contains": ["swift", "goswift"], "account_contains": ["ndd"],
        "exclude_contains": ["reverse", "revers"],
        "unit": "days", "payment_split": False,
        "default_tat": 3, "measure_on": "delivery", "ignore_edd": True,
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
        # Elastic Run, SDD/NDD next-day account (account-matched; carrier is
        # generic). SLA from the ER pincode master: every serviceable pincode
        # promises 1 day (next day) from pickup. Scored pickup->DELIVERY; EDD
        # ignored because ER pads it (~2d vs ~1.3d actual). Non-serviceable
        # pincodes -> No rule (not an ER lane).
        "name": "Elastic Run", "file": "er_tat.json.gz",
        "carrier_contains": None, "account_contains": ["elasticrun", "elastic run"],
        "exclude_contains": ["reverse", "revers", "rvp"],
        "unit": "days", "payment_split": False, "warehouse_agnostic": True,
        "measure_on": "delivery", "ignore_edd": True,
    },
    {
        # Skye Air, drone SAME-DAY delivery (account-matched; destination-based).
        # SLA: OFD1 the same calendar day as pickup, OR within 15 elapsed hours
        # of pickup (ofd1_grace_hours). Scored pickup->OFD1. (Customer order_date
        # is date-only ~2 days before pickup, so "order received" is taken as
        # Skye's handover = pickup.)
        "name": "Skye Air", "file": "skye_tat.json.gz",
        "carrier_contains": None, "account_contains": ["skye", "sky air"],
        "exclude_contains": ["reverse", "revers", "rvp"],
        "unit": "days", "payment_split": False, "warehouse_agnostic": True,
        "ofd1_grace_hours": 15,
    },
    {
        # Shadowfax LARGE (3kg+; account "Shadofax_NDD_Large"). Pitched per-lane
        # TAT (NDD=1, NDD+1=2, NDD+2=3 days) keyed warehouse -> {dest_pin: days}
        # from "Shadowfax Large 3kg+". Must precede the general Shadowfax entry.
        "name": "Shadowfax Large", "file": "shadowfax_large_tat.json.gz",
        "carrier_contains": ["shadowfax", "shadofax"], "account_contains": ["large"],
        "exclude_contains": ["reverse", "revers", "rvp"],
        "unit": "days", "payment_split": False,
        "default_tat": 3, "measure_on": "delivery", "ignore_edd": True,
    },
    {
        # Shadowfax Prime/Small (the "test" account, ~1.4kg avg). Pitched per-lane
        # TAT (days; 0 = same-day intracity) from "Shadowfax Prime Small upto
        # 1kg", keyed warehouse -> {dest_pin: days}. Held to NDD+1 (tat_offset
        # +1d) since "test" parcels exceed the <=1kg Prime band. Catches all
        # other Shadowfax accounts.
        "name": "Shadowfax", "file": "shadowfax_small_tat.json.gz",
        "carrier_contains": ["shadowfax", "shadofax"], "account_contains": None,
        "exclude_contains": ["reverse", "revers", "rvp"],
        "unit": "days", "payment_split": False,
        "default_tat": 3, "measure_on": "delivery", "ignore_edd": True,
        "tat_offset_days": 1,
    },
    {
        # Shiprocket (Shiprocket-Shopify). Scored against the carrier's own
        # committed EDD, which here is accurate (~2.9d vs ~2.67d actual) and is
        # the fair commitment -- Shiprocket ships ~2-3 day, not the aspirational
        # Apex-air 24h. The Apex TAT (hours) stays as a fallback for the rare
        # lines without an EDD. Reverse/RVP excluded.
        "name": "Shiprocket", "file": "shiprocket_tat.json.gz",
        "carrier_contains": ["shiprocket"], "account_contains": None,
        "exclude_contains": ["reverse", "revers", "rvp"],
        "unit": "hours", "payment_split": False, "warehouse_agnostic": True,
        "default_tat": 72, "measure_on": "delivery",
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
    digits = _digits(pin)
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
    return table.get(_digits(drop_pin))


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
            d = _digits(raw)
            origin = d if len(d) == 6 else None
        pin = _digits(row.get("pincode"))
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
        key = carrier_key(e["name"])
        ov = _OVERRIDES.get(key)
        out.append({
            "name": e["name"], "key": key, "unit": e["unit"],
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
        drop = _digits(drop_pin)
        pk = _digits(pickup_pin)
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


# Reverse-logistics (return pickup / RVP) SLA: reverse shipments are excluded
# from every forward carrier rule above, so they're scored on a single flat
# window instead — a return has REVERSE_TAT_DAYS calendar days (pickup -> back
# at destination) to complete.
REVERSE_TAT_DAYS = 45


def _classify_reverse(transit_days, delivered, age_days):
    """SLA for a reverse (return-pickup) shipment, measured pickup -> delivery:

        In TAT      - returned within REVERSE_TAT_DAYS days
        Out of TAT  - returned later than that, OR still in transit past it
        Pending     - still in transit but within the window
    """
    limit = REVERSE_TAT_DAYS
    if delivered:
        if transit_days is None:
            return ("Pending", limit, None)
        return ("In TAT" if transit_days <= limit else "Out of TAT",
                limit, limit - transit_days)
    # Not delivered yet: a breach only once it's been in transit beyond the
    # window; still within it counts as pendency.
    if age_days is not None and age_days > limit:
        return ("Out of TAT", limit, limit - age_days)
    return ("Pending", limit, None)


def classify(carrier, account, pickup_pin, payment, drop_pin,
             p2d_hours, transit_days, delivered,
             age_hours=None, age_days=None, forward_pending=False,
             rto=False, ofd1_hours=None, ofd1_days=None,
             pickup_city=None, drop_city=None, edd_days=None,
             reverse=False):
    """Return (tat_status, promised, margin).

    UNIVERSAL basis (default): score against the carrier's OWN committed
    delivery date (EDD). When `edd_days` (calendar days pickup->EDD) is present
    and the carrier is not under a manual override, In TAT = delivered within
    that window (transit_days <= edd_days); RTO uses the first attempt (OFD1)
    vs EDD; an undelivered shipment already past EDD is an "Out of TAT" breach;
    otherwise "Pending".

    FALLBACK (no EDD, e.g. Skye Air, or a manual override): the carrier's
    committed-TAT lookup is used instead, measured pickup->OFD1 by default or
    pickup->delivery for carriers with measure_on="delivery". `promised` is in
    the carrier's unit (hours or days); `margin` is promised - actual.

    REVERSE (return pickups / RVP): scored on the flat REVERSE_TAT_DAYS window
    instead of any forward carrier rule (see _classify_reverse).
    """
    if reverse:
        return _classify_reverse(transit_days, delivered, age_days)

    idx = _match_carrier(carrier or "", account or "")
    if idx is None:
        return ("No rule", None, None)

    # Universal SLA: score against the carrier's OWN committed delivery date
    # (EDD). `edd_days` = calendar days pickup->EDD; In TAT when the actual
    # transit is within that. A manual override (SLA editor) takes precedence;
    # shipments with no EDD (e.g. Skye Air) fall through to the lookup logic.
    if (edd_days is not None and _override_for(idx) is None
            and not _CARRIERS[idx].get("ignore_edd")):
        if delivered and transit_days is not None:
            return ("In TAT" if transit_days <= edd_days else "Out of TAT",
                    edd_days, edd_days - transit_days)
        if rto and ofd1_days is not None:
            return ("In TAT" if ofd1_days <= edd_days else "Out of TAT",
                    edd_days, edd_days - ofd1_days)
        if forward_pending and age_days is not None and age_days > edd_days:
            return ("Out of TAT", edd_days, edd_days - age_days)
        return ("Pending", edd_days, None)

    promised = _promised(idx, pickup_pin, payment, drop_pin)
    # Built-in fallbacks (default TAT, city-to-city) apply only when the carrier
    # is NOT under a user override; an override replaces the SLA entirely.
    overridden = _override_for(idx) is not None
    ent = _CARRIERS[idx]
    if promised is None and not overridden:
        # No per-lane rule; fall back to the carrier's default TAT if it has one
        # (e.g. Delhivery NDD -> 3 days), otherwise the lane is unscored.
        promised = ent.get("default_tat")
    if promised is None and not overridden:
        # Still nothing: try the city-to-city lane table (e.g. GoSwift) so a
        # pickup that never resolved to a warehouse can still be scored.
        promised = _city_promised(idx, pickup_city, drop_city)
    if promised is None:
        return ("No rule", None, None)

    # Optional per-carrier offset added to the promised TAT (e.g. Shadowfax
    # "test" parcels run ~1.4kg, above the <=1kg Prime band, so they're held to
    # NDD+1 -> tat_offset_days=1). A manual override is exact, so no offset then.
    offset = ent.get("tat_offset_days")
    if offset and not overridden:
        promised += offset

    unit = ent["unit"]
    basis = ent.get("measure_on", "ofd1")

    if basis == "delivery":
        # Score pickup -> DELIVERY (e.g. Delhivery NDD, to track its NSL metric).
        actual = p2d_hours if unit == "hours" else transit_days
        if delivered and actual is not None:
            return ("In TAT" if actual <= promised else "Out of TAT",
                    promised, promised - actual)
        # Returned (never delivered): judge on the first attempt (OFD1).
        if rto:
            ofd = ofd1_hours if unit == "hours" else ofd1_days
            if ofd is not None:
                return ("In TAT" if ofd <= promised else "Out of TAT",
                        promised, promised - ofd)
            return ("Pending", promised, None)
        # Not delivered yet: a forward-pendency already past its SLA is a breach.
        if forward_pending:
            elapsed = age_hours if unit == "hours" else age_days
            if elapsed is not None and elapsed > promised:
                return ("Out of TAT", promised, promised - elapsed)
        return ("Pending", promised, None)

    # Default basis: pickup -> OFD1 (first out-for-delivery attempt). A shipment
    # is scored as soon as it was attempted (whether it ends Delivered or RTO),
    # using OFD1 elapsed time (hours or calendar days per the carrier's unit).
    actual = ofd1_hours if unit == "hours" else ofd1_days
    if actual is not None:
        # Optional grace: also In TAT if OFD1 is within N elapsed hours of pickup
        # (e.g. Skye Air: same calendar day OR within 15h).
        grace = ent.get("ofd1_grace_hours")
        on_time = (actual <= promised) or (
            grace is not None and ofd1_hours is not None and ofd1_hours <= grace)
        return ("In TAT" if on_time else "Out of TAT", promised, promised - actual)
    if forward_pending:
        elapsed = age_hours if unit == "hours" else age_days
        if elapsed is not None and elapsed > promised:
            return ("Out of TAT", promised, promised - elapsed)
    return ("Pending", promised, None)
