import json
import re

CODES_DB_PATH = "codes_db.json"

RP_PREFIXES = [
    ("RMS", "HMX Series"),
    ("RSD", "Duralec Series"),
    ("RS",  "HP API Series"),
    ("HMX", "HMX Series"),
    ("RP",  "RP Synthetic"),
]

SKIP_CODES = {"11722", "18000"}


def _load_db():
    try:
        with open(CODES_DB_PATH) as f:
            return json.load(f)
    except Exception:
        return {}


def _save_db(db):
    with open(CODES_DB_PATH, "w") as f:
        json.dump(db, f, indent=2)


def _get_all_known_codes(db):
    known = set()
    for series in db.get("rp_products", {}).values():
        for sku in series.get("skus", []):
            known.add(sku.get("code", "").upper())
    for brand in db.get("competitor_brands", []):
        for c in brand.get("codes", []):
            known.add(c.get("code", "").upper())
    return known


def _extract_letter_prefix(code):
    m = re.match(r"^([A-Za-z]+)", code)
    return m.group(1).upper() if m else ""


def _guess_rp_series(code, db):
    c = code.upper()
    rp_products = db.get("rp_products", {})
    for prefix, series_hint in RP_PREFIXES:
        if c.startswith(prefix):
            # Exact match first
            if series_hint in rp_products:
                return series_hint, prefix
            # Partial match: DB key starts with our hint
            for db_key in rp_products:
                if db_key.startswith(series_hint):
                    return db_key, prefix
            # Fallback: first RP series
            for db_key in rp_products:
                return db_key, prefix
    return None, None


def _guess_competitor_brand(code, db):
    c = code.upper()
    letter_pfx = _extract_letter_prefix(c)
    if not letter_pfx:
        return None

    best_brand = None
    best_match_len = 0

    for brand_entry in db.get("competitor_brands", []):
        brand_name = brand_entry.get("brand") or brand_entry.get("name", "")
        for existing in brand_entry.get("codes", []):
            ec = existing.get("code", "").upper()
            ec_pfx = _extract_letter_prefix(ec)
            if not ec_pfx:
                continue
            match_len = 0
            for a, b in zip(letter_pfx, ec_pfx):
                if a == b:
                    match_len += 1
                else:
                    break
            if match_len > 0 and match_len > best_match_len:
                best_match_len = match_len
                best_brand = brand_name

    return best_brand if best_match_len >= 1 else None


def auto_classify_code(code, db):
    c = code.upper().strip()

    if not c or c in SKIP_CODES:
        return {"type": "skip"}

    if not re.match(r"^[A-Za-z]", c):
        return {"type": "skip"}

    series, prefix = _guess_rp_series(c, db)
    if series:
        viscosity = c[len(prefix):]
        return {
            "type": "rp",
            "label": "Royal Purple",
            "series": series,
            "viscosity": viscosity,
            "prefix": prefix,
        }

    brand = _guess_competitor_brand(c, db)
    if brand:
        return {
            "type": "competitor",
            "label": brand,
            "brand": brand,
        }

    return {"type": "unknown", "label": "Unrecognized"}


def detect_new_codes(stores, db=None):
    if db is None:
        db = _load_db()

    known = _get_all_known_codes(db)

    code_stats = {}
    for store in stores:
        for pb in store.get("productBreakdown", []):
            code = str(pb.get("code", "")).strip()
            if not code or code.upper() in SKIP_CODES:
                continue
            if code.upper() in known:
                continue
            if code not in code_stats:
                code_stats[code] = {"store_names": set(), "line_count": 0, "revenue": 0.0}
            code_stats[code]["store_names"].add(store.get("name", store.get("storeName", "")))
            code_stats[code]["line_count"] += pb.get("lineCount", 0)
            code_stats[code]["revenue"] += pb.get("revenue", 0.0)

    results = []
    for code, stats in sorted(code_stats.items()):
        cl = auto_classify_code(code, db)
        if cl["type"] == "skip":
            continue
        results.append({
            "code": code,
            "classification": cl,
            "store_count": len(stats["store_names"]),
            "line_count": stats["line_count"],
            "revenue": round(stats["revenue"], 2),
        })

    return results, db


def add_new_codes_to_db(confirmed_items, db=None):
    if db is None:
        db = _load_db()

    added_rp = 0
    added_comp = 0
    skipped = 0

    for item in confirmed_items:
        code = item["code"]
        cl = item["classification"]
        override_series = item.get("override_series")
        override_brand = item.get("override_brand")

        if cl["type"] == "rp":
            series_name = override_series or cl.get("series")
            if not series_name or series_name not in db.get("rp_products", {}):
                series_name = next(iter(db.get("rp_products", {})), None)
            if series_name:
                skus = db["rp_products"][series_name].setdefault("skus", [])
                if not any(s.get("code", "").upper() == code.upper() for s in skus):
                    skus.append({
                        "code": code,
                        "viscosity": cl.get("viscosity", ""),
                        "notes": "Auto-detected from report",
                    })
                    added_rp += 1
                else:
                    skipped += 1
            else:
                skipped += 1

        elif cl["type"] == "competitor":
            brand_name = override_brand or cl.get("brand") or cl.get("label")
            matched = False
            for brand in db.get("competitor_brands", []):
                if (brand.get("brand") or brand.get("name", "")) == brand_name:
                    codes_list = brand.setdefault("codes", [])
                    if not any(c.get("code", "").upper() == code.upper() for c in codes_list):
                        codes_list.append({
                            "code": code,
                            "name": code,
                            "viscosity": "",
                            "notes": "Auto-detected from report",
                        })
                        added_comp += 1
                    else:
                        skipped += 1
                    matched = True
                    break
            if not matched:
                skipped += 1

        else:
            skipped += 1

    _save_db(db)
    return added_rp, added_comp, skipped
