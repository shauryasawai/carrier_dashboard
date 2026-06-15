"""Carrier invoice parser + cost aggregation.

Handles the heterogeneous carrier invoice/billing files Frido receives and
normalizes every billed line into a common record:
    {carrier, awb, order_id, sku, sku_name, product, category,
     weight_kg, zone, amount, shipments}

Supported inputs
----------------
* Frido "Working" billing template (Invoice Amt (₹) + Order ID/AWB) — used by
  ElasticRun, SPS, Swift, BlueDart pre-billing, etc. One generic adapter.
* Urban Bolt billing (Invoice Value + Awb Number).
* BlueDart B2B (NET + SKUs + Sub_Cat).
* SkyAir (Total + SKU Name/Code).
* Frido Prime forward CSV (total_charges + product_cat).
* Master files (Arcatron Weights, Required wts, Swift SKU master) — no charges;
  used to enrich SKU/category onto invoices by AWB / SKU code.

Carrier is read from a Courier column when present, else inferred from the file
name. Reads .xlsx (openpyxl), .xlsb (pyxlsb) and .csv/.tsv (stdlib csv).
"""
from __future__ import annotations

import csv as _csv
import io as _io
import re
from openpyxl import load_workbook

# --------------------------------------------------------------------------
# Category resolution
# --------------------------------------------------------------------------
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
    ("Home & Furnishing", ["bath mat", "doormat", "bed sheet", "blanket", "curtain", "furnishing", "bottle"]),
    ("Accessories", ["cap", "cover", "pouch", "bag", "strap", "glove", "accessor", "combo"]),
]

EXPLICIT_MAP = {
    "orthotics": "Frido Orthotics", "orthotic": "Frido Orthotics",
    "footwear": "Footwears", "footwears": "Footwears",
    "insole": "Insoles", "insoles": "Insoles",
    "pillows": "Pillows", "pillow": "Pillows",
    "cushion": "Cushions", "cushions": "Cushions",
    "mattress": "Mattress", "topper": "Mattress",
    "socks": "Socks", "sock": "Socks",
    "cap": "Accessories", "covers": "Accessories", "cover": "Accessories",
    "accessories": "Accessories", "accessory": "Accessories", "combo": "Accessories", "combos": "Accessories",
    "eye mask": "Personal Care", "mask": "Personal Care", "masks": "Personal Care",
    "furnishing": "Home & Furnishing", "home": "Home & Furnishing",
    "maternity": "Maternity & Baby Care", "baby": "Maternity & Baby Care",
    "chairs": "Mobility & Chairs", "chair": "Mobility & Chairs", "mobility": "Mobility & Chairs",
}


def resolve_category(explicit, name_text):
    ex = (explicit or "").strip()
    if ex:
        first = re.split(r"[|/]", ex)[0].strip().lower()
        if first in EXPLICIT_MAP:
            return EXPLICIT_MAP[first]
    text = (name_text or "").lower()
    for cat, kws in CATEGORY_RULES:
        for kw in kws:
            if kw in text:
                return cat
    if ex:
        low = ex.lower()
        for cat, kws in CATEGORY_RULES:
            for kw in kws:
                if kw in low:
                    return cat
    return "Others"


# --------------------------------------------------------------------------
# value helpers
# --------------------------------------------------------------------------
def _f(v):
    if v is None or v == "":
        return None
    try:
        return float(re.sub(r"[,₹\s]", "", str(v)))
    except (ValueError, TypeError):
        return None


def _s(v):
    return "" if v is None else str(v).strip()


_JUNK = {"", "#n/a", "n/a", "na", "nan", "0", "none", "null", "-", "0.0"}


def _clean_name(v):
    s = _s(v)
    return "" if s.lower() in _JUNK else s


def _norm_awb(v):
    s = _s(v).upper()
    s = re.sub(r"\.0$", "", s)
    return s


def _norm_zone(z):
    z = _s(z)
    return z.title() if z else ""


def _nh(h):
    """Normalize a header cell for matching."""
    s = "" if h is None else str(h)
    s = s.replace("\n", " ").lower()
    s = s.replace("₹", " ")
    s = re.sub(r"\(.*?\)", " ", s)   # drop "(₹)", "(kg)", "(incl gst)" etc.
    s = re.sub(r"\s+", " ", s).strip()
    return s


# --------------------------------------------------------------------------
# low-level readers
# --------------------------------------------------------------------------
def _xlsx_sheets(data):
    wb = load_workbook(filename=_io.BytesIO(data), read_only=True, data_only=True)
    out = []
    for name in wb.sheetnames:
        ws = wb[name]
        out.append((name, [r for r in ws.iter_rows(values_only=True)]))
    return out


def _xlsb_sheets(data):
    from pyxlsb import open_workbook
    out = []
    with open_workbook(_io.BytesIO(data)) as wb:
        for name in wb.sheets:
            rows = []
            with wb.get_sheet(name) as sheet:
                for row in sheet.rows():
                    rows.append([c.v for c in row])
            out.append((name, rows))
    return out


def _csv_rows(data):
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("latin-1", "replace")
    return list(_csv.reader(_io.StringIO(text)))


def _read_sheets(data, filename=""):
    name = (filename or "").lower()
    if name.endswith(".xlsb"):
        return _xlsb_sheets(data)
    if data[:2] == b"PK":
        return _xlsx_sheets(data)
    return [("csv", _csv_rows(data))]


# --------------------------------------------------------------------------
# header location + field maps
# --------------------------------------------------------------------------
AMOUNT_KEYS = ["invoice amt", "invoice value", "net amount", "net",
               "grand total", "total amount", "cost", "total freight", "total"]
ORDER_KEYS = ["order id", "order number", "order_id", "client_order_id", "reference number"]
AWB_KEYS = ["awb number", "awb no.", "awbno", "awb", "tracking_id", "global_tracking_id",
            "awb_number", "waybill number", "transaction_id", "swift id", "cawbno", "wbn"]
ZONE_KEYS = ["zone charged", "zone", "shipment zone", "pricing zone"]
WEIGHT_KEYS = ["weight", "charge weight", "round weight", "billedwt", "frido final wt in kg"]
SKU_KEYS = ["skus", "sku code", "sku codes", "product sku codes", "sku", "sku_list", "product code"]
NAME_KEYS = ["sku name", "product name", "product_desc", "sub_cat", "item names", "sub_category"]
CARRIER_COL_KEYS = ["courier", "courier name", "carrier", "carrier partner name"]


def _find_header(rows, need_amount=True, scan=8):
    """Return (idx, colmap) for the first row that has an amount column (when
    need_amount) plus at least one AWB/order id column."""
    for i in range(min(scan, len(rows))):
        cells = [_nh(c) for c in (rows[i] or [])]
        cset = set(cells)
        has_amount = any(k in cset for k in AMOUNT_KEYS)
        has_id = any(k in cset for k in (AWB_KEYS + ORDER_KEYS))
        if (has_amount or not need_amount) and has_id:
            cmap = {}
            for j, c in enumerate(cells):
                if c and c not in cmap:
                    cmap[c] = j
            return i, cmap
    return None, None


def _first(cmap, keys):
    for k in keys:
        if k in cmap:
            return cmap[k]
    return None


def _cell(row, idx):
    if idx is None or idx >= len(row):
        return None
    return row[idx]


# --------------------------------------------------------------------------
# carrier inference
# --------------------------------------------------------------------------
_CARRIER_FILE = [
    ("skyair", "SkyAir"), ("sky air", "SkyAir"),
    ("elasticrun", "ElasticRun"), ("elastic", "ElasticRun"), ("er_", "ElasticRun"), ("er ", "ElasticRun"),
    ("urbanbolt", "Urban Bolt"), ("urbane", "Urban Bolt"), ("ub_", "Urban Bolt"), ("ub ", "Urban Bolt"),
    ("swift", "Swift"), ("frido invoice summary", "Swift"),
    ("safexpress", "Safexpress"), ("safe express", "Safexpress"),
    ("sps", "SPS"),
    ("ekart", "Ekart"),
    ("prime large", "Frido Prime Large"), ("prime small", "Frido Prime Small"), ("prime", "Frido Prime"),
    ("bd b2b", "BlueDart"), ("bd_", "BlueDart"), ("bluedart", "BlueDart"), ("blue dart", "BlueDart"),
    ("prebilling", "BlueDart"), ("delhivery", "Delhivery"),
]


def carrier_from_filename(filename):
    n = (filename or "").lower()
    for token, label in _CARRIER_FILE:
        if token in n:
            return label
    return "Unknown carrier"


def _clean_courier(v):
    s = _s(v)
    if not s:
        return ""
    token = re.split(r"[ _\-/]", s)[0]
    return token.title() if token else s


# --------------------------------------------------------------------------
# adapters
# --------------------------------------------------------------------------
def _parse_frido_prime(sheets, filename):
    for name, rows in sheets:
        idx, cmap = None, None
        for i in range(min(4, len(rows))):
            cells = set(_nh(c) for c in (rows[i] or []))
            if "awb_number" in cells and "product_cat" in cells and "total_charges" in cells:
                idx = i
                cmap = {_nh(c): j for j, c in enumerate(rows[i] or []) if _nh(c)}
                break
        if idx is None:
            continue
        carrier = carrier_from_filename(filename) or "Frido Prime"
        if carrier == "Unknown carrier":
            carrier = "Frido Prime"
        out = []
        for row in rows[idx + 1:]:
            if not row:
                continue
            awb = _norm_awb(_cell(row, cmap.get("awb_number")))
            amt = _f(_cell(row, cmap.get("total_charges")))
            if not awb:
                continue
            desc = _clean_name(_cell(row, cmap.get("product_desc")))
            cat = _s(_cell(row, cmap.get("product_cat")))
            out.append({
                "carrier": carrier, "awb": awb,
                "order_id": _s(_cell(row, cmap.get("client_order_id"))),
                "sku": "", "sku_name": desc, "product": desc,
                "category": resolve_category(cat, desc),
                "weight_kg": _f(_cell(row, cmap.get("weight"))),
                "zone": _norm_zone(_cell(row, cmap.get("zone"))),
                "amount": amt or 0.0, "shipments": 1,
            })
        if out:
            return out
    return None


def _parse_bluedart_b2b(sheets, filename):
    for name, rows in sheets:
        idx = cmap = None
        for i in range(min(4, len(rows))):
            cells = set(_nh(c) for c in (rows[i] or []))
            if "net" in cells and "skus" in cells and ("awb no." in cells or "cawbno" in cells):
                idx = i
                cmap = {_nh(c): j for j, c in enumerate(rows[i] or []) if _nh(c)}
                break
        if idx is None:
            continue
        out = []
        for row in rows[idx + 1:]:
            if not row:
                continue
            awb = _norm_awb(_cell(row, cmap.get("awb no.")) or _cell(row, cmap.get("cawbno")))
            amt = _f(_cell(row, cmap.get("net"))) or _f(_cell(row, cmap.get("total ")))
            if not awb:
                continue
            nm = _clean_name(_cell(row, cmap.get("sub_cat")))
            out.append({
                "carrier": "BlueDart", "awb": awb, "order_id": "",
                "sku": _s(_cell(row, cmap.get("skus"))), "sku_name": nm, "product": nm,
                "category": resolve_category("", nm),
                "weight_kg": _f(_cell(row, cmap.get("nchrgwt"))),
                "zone": _norm_zone(_cell(row, cmap.get("zone"))),
                "amount": amt or 0.0, "shipments": 1,
            })
        if out:
            return out
    return None


def _parse_skyair(sheets, filename):
    for name, rows in sheets:
        idx = cmap = None
        for i in range(min(4, len(rows))):
            cells = set(_nh(c) for c in (rows[i] or []))
            if "awb" in cells and "sku name" in cells and "total" in cells:
                idx = i
                cmap = {_nh(c): j for j, c in enumerate(rows[i] or []) if _nh(c)}
                break
        if idx is None:
            continue
        out = []
        for row in rows[idx + 1:]:
            if not row:
                continue
            awb = _norm_awb(_cell(row, cmap.get("awb")))
            amt = _f(_cell(row, cmap.get("total")))
            if not awb:
                continue
            nm = _clean_name(_cell(row, cmap.get("sku name")))
            out.append({
                "carrier": "SkyAir", "awb": awb, "order_id": "",
                "sku": _s(_cell(row, cmap.get("sku code"))), "sku_name": nm, "product": nm,
                "category": resolve_category("", nm),
                "weight_kg": _f(_cell(row, cmap.get("round weight"))) or _f(_cell(row, cmap.get("sky air weight"))),
                "zone": "", "amount": amt or 0.0, "shipments": 1,
            })
        if out:
            return out
    return None


def _parse_generic(sheets, filename):
    """Frido 'Working' billing template + Urban Bolt: any sheet with an amount
    column (Invoice Amt (₹) / Invoice Value / ...) plus an AWB or Order id."""
    carrier_fb = carrier_from_filename(filename)
    for name, rows in sheets:
        idx, cmap = _find_header(rows, need_amount=True)
        if idx is None:
            continue
        amt_i = _first(cmap, AMOUNT_KEYS)
        awb_i = _first(cmap, AWB_KEYS)
        ord_i = _first(cmap, ORDER_KEYS)
        zone_i = _first(cmap, ZONE_KEYS)
        wt_i = _first(cmap, WEIGHT_KEYS)
        sku_i = _first(cmap, SKU_KEYS)
        name_i = _first(cmap, NAME_KEYS)
        car_i = _first(cmap, CARRIER_COL_KEYS)
        out = []
        for row in rows[idx + 1:]:
            if not row:
                continue
            amt = _f(_cell(row, amt_i))
            awb = _norm_awb(_cell(row, awb_i))
            order = _s(_cell(row, ord_i))
            if amt is None or (not awb and not order):
                continue
            carrier = _clean_courier(_cell(row, car_i)) or carrier_fb
            nm = _clean_name(_cell(row, name_i))
            out.append({
                "carrier": carrier or "Unknown carrier", "awb": awb, "order_id": order,
                "sku": _s(_cell(row, sku_i)), "sku_name": nm, "product": nm,
                "category": resolve_category(nm, nm),
                "weight_kg": _f(_cell(row, wt_i)),
                "zone": _norm_zone(_cell(row, zone_i)),
                "amount": amt, "shipments": 1,
            })
        if out:
            return out
    return None


_ADAPTERS = [_parse_frido_prime, _parse_bluedart_b2b, _parse_skyair, _parse_generic]


# --------------------------------------------------------------------------
# master files (no charges) -> AWB / SKU enrichment
# --------------------------------------------------------------------------
def _parse_master(sheets):
    """Detect a SKU/weight master and return {'awb2cat':{awb:(cat,product,sku)},
    'sku2cat':{sku:(cat,product)}} or None if it isn't a master."""
    awb2cat, sku2cat = {}, {}
    found = False
    for name, rows in sheets:
        idx = cmap = None
        for i in range(min(6, len(rows))):
            cells = set(_nh(c) for c in (rows[i] or []))
            has_cat = ("category" in cells)
            has_awb = any(k in cells for k in AWB_KEYS)
            has_sku = any(k in cells for k in ["sku_list", "sku code", "product code", "skus", "sku"])
            # Reached only after every invoice adapter declined, so we don't need
            # to exclude files that merely carry an (empty) amount column.
            if (has_cat or has_sku) and (has_awb or has_sku):
                idx = i
                cmap = {_nh(c): j for j, c in enumerate(rows[i] or []) if _nh(c)}
                break
        if idx is None:
            continue
        awb_i = _first(cmap, AWB_KEYS)
        cat_i = cmap.get("category")
        sub_i = cmap.get("sub_category")
        sku_i = _first(cmap, ["sku_list", "sku code", "product code", "skus", "sku"])
        nm_i = _first(cmap, ["sku name", "product name", "product_desc", "sub_cat"])
        for row in rows[idx + 1:]:
            if not row:
                continue
            cat_raw = _s(_cell(row, cat_i))
            nm = _clean_name(_cell(row, nm_i)) or _clean_name(_cell(row, sub_i))
            sku = _s(_cell(row, sku_i))
            if not (cat_raw or nm or sku):
                continue
            cat = resolve_category(cat_raw, nm)
            awb = _norm_awb(_cell(row, awb_i)) if awb_i is not None else ""
            if awb:
                awb2cat[awb] = (cat, nm or sku, sku)
                found = True
            if sku and (nm or cat != "Others"):
                sku2cat[sku.upper()] = (cat, nm or sku)
                found = True
    return {"awb2cat": awb2cat, "sku2cat": sku2cat} if found else None


def ingest(data, filename=""):
    """Classify and parse a file. Returns ('invoice', items) or ('master', maps)."""
    sheets = _read_sheets(data, filename)
    for adapter in _ADAPTERS:
        try:
            items = adapter(list(sheets), filename)
        except Exception:
            items = None
        if items:
            return ("invoice", items)
    master = _parse_master(sheets)
    if master:
        return ("master", master)
    raise ValueError("Unrecognised file format: " + (filename or "file"))


def parse_invoice(data, filename=""):
    kind, payload = ingest(data, filename)
    if kind != "invoice":
        raise ValueError("This looks like a master/reference file, not an invoice: " + (filename or "file"))
    return payload


# --------------------------------------------------------------------------
# Frido category images (CDN) + icon fallback
# --------------------------------------------------------------------------
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
CATEGORY_ICONS = {
    "Frido Orthotics": "🩺", "Insoles": "👣", "Footwears": "🥿", "Pillows": "🛏️",
    "Cushions": "🪑", "Mattress": "🛌", "Mobility & Chairs": "♿", "Socks": "🧦",
    "Maternity & Baby Care": "🤰", "Personal Care": "💆", "Accessories": "🎒",
    "Workspace": "🖥️", "Home & Furnishing": "🏠", "Others": "📦",
}


def _avg(spend, n):
    return round(spend / n, 1) if n else None


def _enrich(items, awb2cat, sku2cat):
    """Fill category / product / sku for weak rows using the master maps. A
    master category only overrides when it is a real (non-'Others') value."""
    filled = 0
    for it in items:
        weak = it["category"] in ("Others", "Unknown", "")
        if not (weak or not it["product"]):
            continue
        ent = awb2cat.get(it["awb"]) if it["awb"] else None
        if ent:
            cat, prod, sku = ent
            if weak and cat and cat != "Others":
                it["category"] = cat; weak = False
            if not it["product"] and prod:
                it["product"] = prod
            if not it["sku"] and sku:
                it["sku"] = sku
        skey = (it["sku"] or "").upper()
        if weak and skey and skey in sku2cat:
            cat, prod = sku2cat[skey]
            if cat and cat != "Others":
                it["category"] = cat; weak = False
            if not it["product"] and prod:
                it["product"] = prod
        if not weak:
            filled += 1
    return filled


def build_cost_report(items, files=None, awb2cat=None, sku2cat=None):
    if awb2cat or sku2cat:
        _enrich(items, awb2cat or {}, sku2cat or {})

    total_spend = sum(i["amount"] for i in items)
    total_ship = sum(i["shipments"] for i in items)

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

    cat = {}
    for i in items:
        c = i["category"] or "Others"
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

    cells = {}
    for c in categories:
        cells[c["category"]] = {cr["carrier"]: {"spend": cr["spend"], "shipments": cr["shipments"],
                                                "avg_cost": cr["avg_cost"]} for cr in c["carriers"]}
    matrix = {"carriers": carrier_names, "categories": [c["category"] for c in categories], "cells": cells}

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
    skus = skus[:500]

    return {
        "currency": "₹",
        "summary": {"total_spend": round(total_spend, 1), "shipments": total_ship,
                    "avg_cost": _avg(total_spend, total_ship),
                    "categories": len(categories), "skus": sku_total, "carriers": len(carriers)},
        "carriers": carriers, "categories": categories, "matrix": matrix,
        "skus": skus, "sku_total": sku_total, "files": files or [],
    }
