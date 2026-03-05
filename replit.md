# Royal Purple Partnership Hub

## Overview
"The Royal Purple Partnership Hub by ThrottlePro" — a Streamlit web app for Royal Purple installer partners with three sections: Report Generator (Excel→PPTX), interactive Distribution Map (Plotly choropleth), and Product Reference.

## Architecture
- **app.py** — Streamlit frontend with 3-page sidebar nav, interactive US map, report generation with Max-Clean analytics display
- **report_generator.py** — PPTX engine + adaptive Excel parser with invoice deduplication and Max-Clean attachment analysis
- **distribution_data.py** — STATE_DISTRIBUTORS mapping (50 states + DC), DISTRIBUTOR_COLORS, ALL_DISTRIBUTORS
- **assets/** — Royal Purple logos and branding images

## Excel Parsing (Fully Adaptive)
- Scans first 10 rows for best header match (keyword scoring, min 2 keywords)
- Column detection: exact match first, then substring match
- Supports two layouts:
  - **Multi-sheet**: Each worksheet = one store (sheet name = store name)
  - **Consolidated**: Single sheet with a Store/Location column → auto-splits into per-store data
- Revenue fallback: if no "revenue" header found, heuristically detects numeric columns with largest totals
- Currency stripping: handles $, commas in values
- Date parsing: datetime objects, string formats ("Month Day, Year", "MM/DD/YYYY"), and sheet title scanning
- Case-insensitive product prefix matching (longest-prefix-first)

## Invoice Deduplication
- RP POS exports duplicate invoice totals across every RP product line on the same ticket
- `_group_invoices()` groups data rows by Invoice # column (if present) or (date, revenue, vehicle) proxy key
- Revenue per invoice = shared invoice total (counted once, not summed across product lines)
- Applied consistently in both `_parse_single_store_sheet` and `_parse_consolidated_sheet`

## Max-Clean Analytics
- MC_CODE = "11722"; RP_OIL_PREFIXES = ("RP", "RS", "HMX", "RMS", "RSD"); 18000 (Max-Atomizer) excluded
- Per-store metrics: total MC invoices, withRpOil, withNonRpOil, soloInData, attachmentRate, avgTicket, nonMcAvgTicket, ticketLift
- "Solo" Max-Clean invoices = non-RP oil changes where MC was added as upsell (not standalone retail)
- Network-level aggregation shown in app with MC Attachment Analysis section
- Store Rankings table includes MC Rate and MC Lift columns
- Dedicated "Max-Clean by Store" tab with detailed breakdown

## Interactive Distribution Map
- Plotly go.Choropleth with USA scope
- 6 distributors color-coded: Texas Enterprises (green), American Lube Supply (blue), Avery Oil (salmon), Arnold Oil (silver), Dennis K Burke (gold), Brennan Oil (crimson)
- Filterable by distributor via multiselect
- ABE Legend with state counts
- State detail selector with colored accent bars

## Logo Assets (assets/)
- `Royal Purple White Logo.png` — DO NOT USE in PPTX (old blue/yellow text logo)
- `RPMO_logo_BF_Outline.png` — Full logo with checkered flag, Streamlit sidebar only
- `rp_synthetic_expert_white.png` — White "Synthetic Expert" for dark backgrounds, PPTX footer
- `RP_Synthetic_Expert_Logo_Black_Text.png` — Black text, PPTX header badge
- `25-RYP-02147 Employee LinkedIn Thumbnails P1-6.jpg` — "NEVER SETTLE" background

## PPTX Slide Design (Clean Minimal)
- Content slides: off-white background, thin purple top bar, RP badge top-right, gold accent bar next to title
- Metrics use `_add_metric()` — text-only (no card backgrounds, no gold accent bars)
- Sections separated by `_add_thin_divider()` — hairline gray lines
- Rankings/Matrix use native python-pptx Table objects (not individual cell shapes)
- Observations/Next Steps use numbered text lists (no card containers)
- Deep dives: metrics row → thin divider → bar chart left + text notes right (no colored panels)
- Cover, section dividers, closing: dark background with NEVER SETTLE image overlay

## PPTX Slide Structure (dynamic)
Cover → TOC → Exec Summary → Exec Observations → Revenue Overview → Rankings → Matrix → Product Mix → Product Deep Dives Section Divider → Product Deep Dives (per category) → [Distribution Maps] → Store Deep Dives Section Divider → Store Deep Dives → Next Steps → Closing

## Dependencies
- Python 3.11
- streamlit, python-pptx, openpyxl, Pillow, plotly

## Running
```
streamlit run app.py --server.port 5000
```
