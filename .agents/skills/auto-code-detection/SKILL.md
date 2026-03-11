---
name: auto-code-detection
description: Auto-detect and classify new product codes from uploaded Royal Purple Excel reports. Use when modifying the code detection system, adding new prefix rules, changing classification logic, or debugging why codes are misclassified.
---

# Auto Code Detection

Scans parsed Excel report data for product codes not yet in `codes_db.json`, auto-classifies them as Royal Purple or competitor, and lets the user add them to the database in one click.

## Architecture

### Files

- **`code_detector.py`** — detection engine (standalone module, no Streamlit dependency)
- **`app.py`** — UI integration (report generator page, after `parse_excel()` succeeds)
- **`codes_db.json`** — the live product code database

### Key Functions in `code_detector.py`

| Function | Purpose |
|---|---|
| `detect_new_codes(stores, db=None)` | Scans parsed stores for unknown codes, returns `(results_list, db)` |
| `auto_classify_code(code, db)` | Classifies a single code → `{type, label, series/brand, ...}` |
| `add_new_codes_to_db(items, db=None)` | Saves confirmed codes to `codes_db.json`, returns `(added_rp, added_comp, skipped)` |

### Classification Types

| `type` | Meaning | Destination |
|---|---|---|
| `rp` | Royal Purple product | Added to an RP series in `codes_db.json → rp_products` |
| `competitor` | Competitor brand product | Added to a competitor brand in `codes_db.json → competitor_brands` |
| `unknown` | No prefix match | Flagged for manual review (not auto-added) |
| `skip` | Ancillary/non-product (e.g., `11722`, `18000`) | Silently ignored |

## RP Prefix Classification

Prefixes are checked in order — longer/more specific prefixes come first to avoid false matches (e.g., `RSD` before `RS`):

```python
RP_PREFIXES = [
    ("HPS", "HPS Series"),       # High Performance Street
    ("HMX", "HMX Series"),      # High Mileage
    ("RMS", "HMX Series"),       # High Mileage (alternate prefix)
    ("RSD", "Duralec Series"),   # Diesel
    ("RS",  "HP API Series"),    # High Performance API-licensed
    ("RP",  "RP Synthetic"),    # Standard Synthetic
]
```

Series names in the DB include descriptions (e.g., `"RS Series — High Performance Synthetic"`), so matching uses `startswith()` on the DB key.

## Competitor Prefix Classification

Competitor codes are matched by comparing the letter prefix of the unknown code against existing codes in each competitor brand. The brand with the longest matching letter prefix wins.

Example: Code `S10W30` → letter prefix `S` → matches existing CAM2 codes `S0W20`, `S5W20` → classified as CAM2.

## Adding New Prefix Rules

To add a new RP prefix:
1. Add the entry to `RP_PREFIXES` in `code_detector.py`
2. Place it **before** any shorter prefix it could shadow (e.g., `RSD` before `RS`)
3. The series hint must match (or start-match) a key in `codes_db.json → rp_products`

To add a new competitor brand:
- Just add the brand and its first code in the Admin panel — future codes with a matching letter prefix will auto-classify to that brand

## UI Flow in `app.py`

1. After `parse_excel()` succeeds, `detect_new_codes(stores)` runs
2. Results are cached in `st.session_state[f"detected_codes_{filename}"]`
3. If new codes found, an expandable section shows:
   - Table of auto-classified codes (RP/competitor) with destination series/brand
   - Table of unrecognized codes (manual review needed)
   - "Add N recognized codes to database" primary button
   - "Dismiss" button to skip
4. On add: calls `add_new_codes_to_db()`, clears `load_codes_db` cache, reruns
5. Section auto-hides after add or dismiss (tracked via `dismissed_codes_{filename}` session state key)

## Important Constants

- `SKIP_CODES = {"11722", "18000"}` — MC code and excluded code, always ignored
- `CODES_DB_PATH = "codes_db.json"` — hardcoded path to the database
- Codes are compared case-insensitively (`.upper()`)
- Added codes get `notes: "Auto-detected from report"`
