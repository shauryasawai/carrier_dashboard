"""Carrier invoice parser + cost aggregation (v1: BlueDart, SkyAir, Frido Prime).

Each carrier ships a wildly different layout, so we use small per-carrier
"adapters" that detect their own format and normalize every billed line into a
common record:
    {carrier, awb, order_id, sku, sku_name, product, category,
     weight_kg, zone, amount, shipments}
Amounts are the GST-inclusive billed total (what the carrier charges us).
Only openpyxl (already a dependency) + the stdlib csv module are used.
"""
from __future__ import annotations

import csv as _csv
import io as _io
from openpyxl import load_workbook

# --- Canonical Frido categories -------------------------------------------
# Ordered specific -> general. First keyword hit wins.
CATEGORY_RULES = [
    ("Maternity & Baby Care", ["pregnancy", "maternity", "baby", "infant", "nursing", "feeding pillow"]),
    ("Workspace", ["standing desk", "desk converter", "desk", "laptop stand", "monitor stand", "footrest"]),
    ("Mobility & Chairs", ["wheelchair", "commode", "recliner", "ergo chair", "ergonomic chair",
                            "executive chair", "gaming chair", "office chair", "ergoluxe", "ergo luxe",
                            "posture plus chair", "chair", "walker", "rollator", "scooter", "transfer lift",
                            "bed rail", "grab bar", "guardrail", "safety rail", "joystick", "ramp",
                            "crutch", "mobility"]),
    ("Insoles", ["insole", "arch support", "arch cushion", "shoe insert", "foot insert", "arch sports"]),
    ("Footwears", ["sock shoe", "barefoot", "sandal", "slipper", "chappal", "flip flop", "flipflop",
                   "clog", "sneaker", "footwear", "shoe"]),
    ("Socks", ["sock"]),
    ("Mattress", ["mattress", "topper"]),
    ("Pillows", ["neck pillow", "cervical", "wedge", "travel pillow", "memory foam pillow",
                 "sleep pillow", "cozy pillow", "pillow"]),
    ("Cushions", ["seat cushion", "coccyx", "donut", "backrest", "back rest", "lumbar", "cushion", "seat"]),
    ("Frido Orthotics", ["posture", "orthotic", "knee", "ankle", "elbow", "wrist", "shoulder",
                         "brace", "wrap", "support belt", "lumbo", "sacral", "bunion", "heel",
                         "plantar", "toe", "compression", "belt", "support"]),
    ("Personal Care", ["eye mask", "sleep mask", "mask", "therapy", "heating pad", "heat pad",
                       "nasal", "nose", "massager", "massage", "roller", "pain relief",
                       "kinesiology", "tape"]),
    ("Home & Furnishing", ["bath mat", "doormat", "bed sheet", "blanket", "curtain", "furnishing"]),
    ("Accessories", ["cap", "cover", "bottle", "pouch", "bag", "strap", "glove", "accessor"]),
]

# Frido-Prime-style explicit product_cat -> canonical.
EXPLICIT_MAP = {
    "orthotics": "Frido Orthotics", "orthotic": "Frido Orthotics",
    "footwear": "Footwears", "footwears": "Footwears",
    "insole": "Insoles", "insoles": "Insoles",
    "pillows": "Pillows", "pillow": "Pillows",
    "cushion": "Cushions", "cushions": "Cushions",
    "mattress": "Mattress", "topper": "Mattress",
    "socks": "Socks", "sock": "Socks",
    "cap": "Accessories", "covers": "Accessories", "cover": "Accessories",
    "accessories": "Accessories", "accessory": "Accessories",
    "eye mask": "Personal Care", "mask": "Personal Care", "masks": "Personal Care",
    "furnishing": "Home & Furnishing", "home": "Home & Furnishing",
    "maternity": "Maternity & Baby Care", "baby": "Maternity & Baby Care",
    "chairs": "Mobility & Chairs", "chair": "Mobility & Chairs", "mobility": "Mobility & Chairs",
}


def resolve_category(explicit: str, name_text: str) -> str:
    ex = (explicit or "").strip()
    if ex:
        first = ex.split("|")[0].strip().lower()
        if first in EXPLICIT_MAP:
            return EXPLICIT_MAP[first]
        # try keyword on the explicit text too
    text = (name_text or "").lower()
    for cat, kws in CATEGORY_RULES:
        for kw in kws:
            if kw in text:
                return cat
    if ex:
        for cat, kws in CATEGORY_RULES:
            for kw in kws:
                if kw in ex.lower():
                    return cat
    return "Others"


def _f(v):
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(",", "").replace("₹", "").strip())
    except (ValueError, TypeError):
        return None


def _s(v):
    return "" if v is None else str(v).strip()


_JUNK = {"", "#n/a", "n/a", "na", "nan", "0", "none", "null", "-"}


def _clean_name(v):
    s = _s(v)
    return "" if s.lower() in _JUNK else s


def _norm_zone(z):
    z = _s(z)
    return z.title() if z else ""


# --- low-level readers -----------------------------------------------------
def _xlsx_sheets(data: bytes):
    wb = load_workbook(filename=_io.BytesIO(data), read_only=True, data_only=True)
    for name in wb.sheetnames:
        ws = wb[name]
        rows = [r for r in ws.iter_rows(values_only=True)]
        yield name, rows


def _csv_rows(data: bytes):
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("latin-1", "replace")
    return list(_csv.reader(_io.StringIO(text)))


def _find_header(rows, signatures, scan=8):
    """Return (header_index, colmap) for the first row containing all signature
    headers (case-insensitive). colmap maps lowercased header -> col index."""
    sig = [s.lower() for s in signatures]
    for i in range(min(scan, len(rows))):
        cells = [(_s(c)).lower() for c in (rows[i] or [])]
        if all(any(s == c for c in cells) for s in sig):
            cmap = {}
            for j, c in enumerate(rows[i] or []):
                key = _s(c).lower()
                if key and key not in cmap:
                    cmap[key] = j
            return i, cmap
    return None, None


def _get(row, cmap, header, default=None):
    j = cmap.get(header.lower())
    if j is None or j >= len(row):
        return default
    return row[j]


# --- per-carrier adapters --------------------------------------------------
def _parse_bluedart(sheets):
    for name, rows in sheets:
        hi, cmap = _find_header(rows, ["awb no.", "skus", "net"])
        if hi is None:
            hi, cmap = _find_header(rows, ["cawbno", "ntotalamt"])
        if hi is None:
            continue
        out = []
        for row in rows[hi + 1:]:
            if not row:
                continue
            awb = _s(_get(row, cmap, "awb no.") or _get(row, cmap, "cawbno"))
            amt = _f(_get(row, cmap, "net")) or _f(_get(row, cmap, "total "))
            if not awb:
                continue
            name_txt = _clean_name(_get(row, cmap, "sub_cat"))
            out.append({
                "carrier": "BlueDart", "awb": awb, "order_id": "",
                "sku": _s(_get(row, cmap, "skus")), "sku_name": name_txt,
                "product": name_txt,
                "category": resolve_category("", name_txt),
                "weight_kg": _f(_get(row, cmap, "nchrgwt")),
                "zone": _norm_zone(_get(row, cmap, "zone")),
                "amount": amt or 0.0, "shipments": 1,
            })
        return out
    return None


def _parse_skyair(sheets):
    for name, rows in sheets:
        hi, cmap = _find_header(rows, ["awb", "sku name", "total"])
        if hi is None:
            continue
        out = []
        for row in rows[hi + 1:]:
            if not row:
                continue
            awb = _s(_get(row, cmap, "awb"))
            amt = _f(_get(row, cmap, "total"))
            if not awb:
                continue
            sku_name = _clean_name(_get(row, cmap, "sku name"))
            out.append({
                "carrier": "SkyAir", "awb": awb, "order_id": "",
                "sku": _s(_get(row, cmap, "sku code")), "sku_name": sku_name,
                "product": sku_name,
                "category": resolve_category("", sku_name),
                "weight_kg": _f(_get(row, cmap, "round weight")) or _f(_get(row, cmap, "sky air weight")),
                "zone": "",
                "amount": amt or 0.0, "shipments": 1,
            })
        return out
    return None


def _parse_frido_prime(sheets):
    for name, rows in sheets:
        hi, cmap = _find_header(rows, ["awb_number", "product_cat", "total_charges"])
        if hi is None:
            continue
        out = []
        for row in rows[hi + 1:]:
            if not row:
                continue
            awb = _s(_get(row, cmap, "awb_number"))
            amt = _f(_get(row, cmap, "total_charges"))
            if not awb:
                continue
            desc = _clean_name(_get(row, cmap, "product_desc"))
            cat = _s(_get(row, cmap, "product_cat"))
            out.append({
                "carrier": "Frido Prime", "awb": awb,
                "order_id": _s(_get(row, cmap, "client_order_id")),
                "sku": "", "sku_name": desc, "product": desc,
                "category": resolve_category(cat, desc),
                "weight_kg": _f(_get(row, cmap, "weight")),
                "zone": _norm_zone(_get(row, cmap, "zone")),
                "amount": amt or 0.0, "shipments": 1,
            })
        return out
    return None


ADAPTERS = [_parse_frido_prime, _parse_bluedart, _parse_skyair]


def parse_invoice(data: bytes, filename: str = ""):
    name = (filename or "").lower()
    if data[:2] == b"PK":  # xlsx
        sheets = list(_xlsx_sheets(data))
    else:
        sheets = [("csv", _csv_rows(data))]
    for adapter in ADAPTERS:
        try:
            items = adapter(list(sheets))
        except Exception:
            items = None
        if items:
            return items
    raise ValueError("Unrecognised invoice format: " + (filename or "file"))


# --- Frido category images (CDN) + icon fallback --------------------------
_CDN = "https://cdn.shopify.com/s/files/1/0553/0419/2034/files/"
CATEGORY_IMAGES = {
    "Frido Orthotics": _CDN + "Category_1_2_1.jpg?v=1769695217&width=400",
    "Insoles": _CDN + "Everyday_Insole_Combo_55a98dda-d383-4a27-9f57-39c4a5e8bcab.png?v=1770808742&width=400",
    "Footwears": _CDN + "WS-01_8b6068d3-4e48-4b4c-8d37-07fd0338d7cb.jpg?v=1753101892&width=400",
    "Pillows": _CDN + "CNP-Black-01B_4997708e-d0e7-49d3-af3f-fe7d51f06423.jpg?v=1739886430&width=400",
    "Cushions": _CDN + "Wedge-plus_color-picker-tab_2edb45fe-649c-4966-b706-0d148f3de35d.png?v=1747912121&width=400",
    "Mattress": _CDN + "UMT_2.jpg?v=1720264311&width=400",
    "Mobility & Chairs": _CDN + "Ergo-chair-01_f909b535-6d3d-4bc9-b4bd-0c3ff7ecdeee.jpg?v=1729333804&width=400",
}
# Emoji fallback used by the UI when no CDN image exists for a category.
CATEGORY_ICONS = {
    "Frido Orthotics": "🩺", "Insoles": "👣", "Footwears": "🥿", "Pillows": "🛏️",
    "Cushions": "🪑", "Mattress": "🛌", "Mobility & Chairs": "♿", "Socks": "🧦",
    "Maternity & Baby Care": "🤰", "Personal Care": "💆", "Accessories": "🎒",
    "Workspace": "🖥️", "Home & Furnishing": "🏠", "Others": "📦",
}


def _avg(spend, n):
    return round(spend / n, 1) if n else None


def build_cost_report(items, files=None):
    """Aggregate normalized invoice line items into the cost-analysis payload."""
    cur = "₹"
    total_spend = sum(i["amount"] for i in items)
    total_ship = sum(i["shipments"] for i in items)

    # --- by carrier ---
    cb = {}
    for i in items:
        g = cb.setdefault(i["carrier"], {"carrier": i["carrier"], "spend": 0.0, "shipments": 0})
        g["spend"] += i["amount"]; g["shipments"] += i["shipments"]
    carriers = []
    for g in sorted(cb.values(), key=lambda x: -x["spend"]):
        carriers.append({"carrier": g["carrier"], "spend": round(g["spend"], 1),
                         "shipments": g["shipments"], "avg_cost": _avg(g["spend"], g["shipments"]),
                         "share": round(g["spend"] / total_spend * 100, 1) if total_spend else 0})
    carrier_names = [c["carrier"] for c in carriers]

    # --- by category (with per-carrier + top SKUs) ---
    cat = {}
    for i in items:
        c = i["category"]
        g = cat.setdefault(c, {"category": c, "spend": 0.0, "shipments": 0,
                               "weight": 0.0, "carriers": {}, "skus": {}})
        g["spend"] += i["amount"]; g["shipments"] += i["shipments"]
        if i["weight_kg"]:
            g["weight"] += i["weight_kg"]
        cc = g["carriers"].setdefault(i["carrier"], {"carrier": i["carrier"], "spend": 0.0, "shipments": 0})
        cc["spend"] += i["amount"]; cc["shipments"] += i["shipments"]
        skey = i["sku"] or i["product"] or "(unmapped)"
        sk = g["skus"].setdefault(skey, {"sku": i["sku"], "name": i["product"] or i["sku_name"] or skey,
                                          "spend": 0.0, "shipments": 0})
        sk["spend"] += i["amount"]; sk["shipments"] += i["shipments"]

    categories = []
    for g in sorted(cat.values(), key=lambda x: -x["spend"]):
        clist = sorted(g["carriers"].values(), key=lambda x: -x["spend"])
        carrier_rows = [{"carrier": c["carrier"], "spend": round(c["spend"], 1),
                         "shipments": c["shipments"], "avg_cost": _avg(c["spend"], c["shipments"])}
                        for c in clist]
        skus = sorted(g["skus"].values(), key=lambda x: -x["spend"])
        top_skus = [{"sku": s["sku"], "name": s["name"], "spend": round(s["spend"], 1),
                     "shipments": s["shipments"], "avg_cost": _avg(s["spend"], s["shipments"])}
                    for s in skus[:8]]
        categories.append({
            "category": g["category"],
            "image": CATEGORY_IMAGES.get(g["category"], ""),
            "icon": CATEGORY_ICONS.get(g["category"], "📦"),
            "spend": round(g["spend"], 1), "shipments": g["shipments"],
            "avg_cost": _avg(g["spend"], g["shipments"]),
            "avg_per_kg": _avg(g["spend"], g["weight"]) if g["weight"] else None,
            "share": round(g["spend"] / total_spend * 100, 1) if total_spend else 0,
            "carriers": carrier_rows, "top_skus": top_skus, "sku_count": len(g["skus"]),
        })

    # --- carrier x category matrix (avg cost + spend per cell) ---
    cells = {}
    for c in categories:
        cells[c["category"]] = {cr["carrier"]: {"spend": cr["spend"], "shipments": cr["shipments"],
                                                "avg_cost": cr["avg_cost"]} for cr in c["carriers"]}
    matrix = {"carriers": carrier_names, "categories": [c["category"] for c in categories], "cells": cells}

    # --- full SKU list (for search) ---
    sk = {}
    for i in items:
        skey = (i["sku"] or i["product"] or "(unmapped)")
        g = sk.setdefault(skey, {"sku": i["sku"], "name": i["product"] or i["sku_name"] or skey,
                                 "category": i["category"], "spend": 0.0, "shipments": 0, "carriers": {}})
        g["spend"] += i["amount"]; g["shipments"] += i["shipments"]
        cc = g["carriers"].setdefault(i["carrier"], {"carrier": i["carrier"], "spend": 0.0, "shipments": 0})
        cc["spend"] += i["amount"]; cc["shipments"] += i["shipments"]
    skus = []
    for g in sorted(sk.values(), key=lambda x: -x["spend"]):
        crows = sorted(g["carriers"].values(), key=lambda x: -x["spend"])
        skus.append({"sku": g["sku"], "name": g["name"], "category": g["category"],
                     "spend": round(g["spend"], 1), "shipments": g["shipments"],
                     "avg_cost": _avg(g["spend"], g["shipments"]),
                     "carriers": [{"carrier": c["carrier"], "spend": round(c["spend"], 1),
                                   "shipments": c["shipments"], "avg_cost": _avg(c["spend"], c["shipments"])}
                                  for c in crows]})

    sku_total = len(skus)
    skus = skus[:500]   # cap payload; table is searchable within the top SKUs

    return {
        "currency": cur,
        "summary": {"total_spend": round(total_spend, 1), "shipments": total_ship,
                    "avg_cost": _avg(total_spend, total_ship),
                    "categories": len(categories), "skus": sku_total, "carriers": len(carriers)},
        "carriers": carriers, "categories": categories, "matrix": matrix,
        "skus": skus, "sku_total": sku_total, "files": files or [],
    }
