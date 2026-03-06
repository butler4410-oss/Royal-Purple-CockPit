# Royal Purple Partnership Hub

## Overview
"The Royal Purple Partnership Hub by ThrottlePro" — a Streamlit web app for Royal Purple installer partners with three sections: Report Generator (Excel→PPTX), Customer Map (Leaflet.js interactive marker map with county-level data), and Product Reference.

## Architecture
- **app.py** — Streamlit frontend with 3-page sidebar nav: Report Generator, Customer Map, Product Reference
- **report_generator.py** — PPTX engine + adaptive Excel parser with invoice deduplication and Max-Clean attachment analysis
- **customer_map.py** — Leaflet.js map builder: loads customers.json, parses CSV uploads, generates embedded HTML map component
- **c4c_report_generator.py** — Excel report generator: C4C gap analysis → 10-sheet .xlsx workbook
- **customers.json** — Geocoded installer accounts (3,539 total with county data for US accounts)
- **distribution_data.py** — Legacy file (not imported by active code). Do NOT re-add distributor references until user provides real data
- **assets/** — Royal Purple branding images

## Customer Map (Leaflet.js)
- Embedded via st.components.v1.html() — full Leaflet map with marker clusters
- **3,539 accounts** across US (3,498), Costa Rica (38), Canada (2), Puerto Rico (1)
- **708 unique US counties** mapped via pgeocode zip-to-county lookup
- 3 account categories with distinct colors:
  - **Promo Only (Not on C4C)** — Red (#DC2626) — 2,097 accounts
  - **On Both Lists** — Green (#16A34A) — 234 accounts
  - **C4C Only** — Blue (#2563EB) — 1,208 accounts
- Filters: search bar, state dropdown, county dropdown (cascading from state), account type dropdown
- County shown in marker popups and sidebar list items
- State filter dynamically updates county dropdown options
- Marker clustering via leaflet.markercluster for zoomed-out views
- Stats bar with inline per-category counts
- Collapsible sidebar list synced with visible markers
- **Export Map Data**: CSV download with all account fields including county
- **Export C4C Report**: Excel report generation button
- CARTO light basemap tiles
- Geocoding: pgeocode for US/CR/CA zip codes; fallback coordinates for Costa Rica provinces

## C4C Report (10 Sheets)
1. Executive Summary — totals, gap %, key findings
2. State Breakdown — per-state counts (Not on C4C vs C4C Matched)
3. Not on C4C Full List — promo-only accounts
4. C4C Matched Full List — On Both Lists + C4C Only accounts
5. Top Priority States — ranked by gap
6. Reconciliation — cross-reference summary
7. C4C Duplicates — duplicate entries in C4C list
8. Promo Duplicates — duplicate entries in promo list
9. On C4C Only — accounts on C4C but not matched to promo
10. Failed to Geolocate — accounts with invalid zip codes

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
Cover → TOC → Exec Summary → Exec Observations → Revenue Overview → Rankings → Matrix → Product Mix → Product Deep Dives Section Divider → Product Deep Dives (per category) → [Map Images] → Store Deep Dives Section Divider → Store Deep Dives → Next Steps → Closing

## Distributor Info
- ALL distributor data has been removed from active code paths
- distribution_data.py exists but is NOT imported anywhere
- Do NOT re-add distributor references until user provides real ABE distributor data

## Dependencies
- Python 3.11
- streamlit, python-pptx, openpyxl, Pillow, plotly, pgeocode

## Running
```
streamlit run app.py --server.port 5000
```
