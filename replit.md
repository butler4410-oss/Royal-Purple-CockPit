# Royal Purple Partnership Hub

## Overview
"The Royal Purple Partnership Hub by ThrottlePro" — a Streamlit web app for Royal Purple installer partners with four sections: Report Generator (Excel→PPTX), interactive Distribution Map (Plotly choropleth), Customer Map (Leaflet.js interactive marker map), and Product Reference.

## Architecture
- **app.py** — Streamlit frontend with 4-page sidebar nav, interactive US map, customer map, report generation with Max-Clean analytics display
- **report_generator.py** — PPTX engine + adaptive Excel parser with invoice deduplication and Max-Clean attachment analysis
- **distribution_data.py** — STATE_DISTRIBUTORS mapping (50 states + DC), DISTRIBUTOR_COLORS, ALL_DISTRIBUTORS
- **customer_map.py** — Leaflet.js map builder: loads customers.json, parses CSV uploads, generates embedded HTML map component
- **c4c_report_generator.py** — Excel report generator: C4C gap analysis → 10-sheet .xlsx workbook
- **customers.json** — Geocoded installer accounts (3,516 total: 2,093 Promo Only, 233 On Both Lists, 1,190 C4C Only)
- **assets/** — Royal Purple branding images

## Customer Map (Leaflet.js)
- Embedded via st.components.v1.html() — full Leaflet map with marker clusters
- **Global coverage**: US (3,498), Costa Rica (38), Canada (2), Puerto Rico (1) = 3,539 total
- 3 account categories with distinct colors:
  - **Promo Only (Not on C4C)** — Red (#DC2626) — 2,097 accounts from promo list not found in C4C
  - **On Both Lists** — Green (#16A34A) — 234 accounts matched on both promo and C4C lists
  - **C4C Only** — Blue (#2563EB) — 1,208 accounts on C4C but not on promo list
- Filters: search bar, country dropdown, region/state dropdown, account type dropdown
- Marker clustering via leaflet.markercluster for zoomed-out views
- Clickable markers with popup showing store name, address, city/state/country, type badge
- Stats bar with inline per-category counts
- Collapsible sidebar list synced with visible markers; click to fly to location
- Optional CSV upload to replace default customers.json data
- CARTO light basemap tiles
- Geocoding: pgeocode for US, CR, CA zip codes; fallback province coordinates for Costa Rica
- Distributor info removed — do NOT re-add until user provides real data

## C4C Report (10 Sheets)
1. Executive Summary — totals, gap %, key findings
2. State Breakdown — per-state counts (Not on C4C vs C4C Matched)
3. Not on C4C Full List — 2,093 promo-only accounts
4. C4C Matched Full List — 1,423 accounts (On Both Lists + C4C Only)
5. Top Priority States — ranked by gap
6. Reconciliation — cross-reference summary
7. C4C Duplicates — 52 duplicate entries in C4C list
8. Promo Duplicates — 123 duplicate entries in promo list
9. On C4C Only — 1,170 accounts on C4C but not matched to promo
10. Failed to Geolocate — 62 accounts with invalid zip codes

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

## Logo/Branding Rules
- NO logo images in PPTX slides — gold text "ROYAL PURPLE" badge only
- Sidebar is text-only (no logo)
- NEVER SETTLE background image stays on cover/dividers/closing slides only
- assets/25-RYP-02147 Employee LinkedIn Thumbnails P1-6.jpg — background image

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
- streamlit, python-pptx, openpyxl, Pillow, plotly, pgeocode

## Running
```
streamlit run app.py --server.port 5000
```
