# Butler Performance Partnership Hub

## Overview
"The Butler Performance Partnership Hub by ThrottlePro" — a Streamlit web app for Butler Performance installer partners with three sections: Report Generator (Excel→PPTX), Customer Map (Leaflet.js interactive marker map), and Product Reference.

## Architecture
- **app.py** — Streamlit frontend with 6-page sidebar nav: Home, Report Generator, Customer Map, Product Reference, Profit Calculator, Admin
- **profit_calculator.py** — Interactive RP vs competitor profit comparison tool with volume/pricing inputs and real-time results dashboard
- **report_generator.py** — PPTX engine + adaptive Excel parser with invoice deduplication and Max-Clean attachment analysis
- **customer_map.py** — Leaflet.js map builder: loads customers.json + distributors.json, generates embedded HTML map with 8 marker categories
- **c4c_report_generator.py** — Excel report generator: C4C gap analysis → 10-sheet .xlsx workbook
- **map_data_exporter.py** — Excel export generator: branded workbook with Dashboard, per-state tabs, All Accounts, County Summary, Distributors
- **customers.json** — All geocoded accounts (4,658 total: installers, powersports, international, Canada)
- **distributors.json** — Geocoded distributor locations (161 total)
- **distribution_data.py** — Legacy file (not imported by active code)
- **assets/** — Butler Performance branding images

## Data Sources
- **InstallerRack_RP Excel** — 5-sheet workbook:
  - Rack Installers USA (581): cross-referenced against promo/C4C lists — 331 matched existing, 240 new geocoded as "Rack Installer"
  - Powersports/Motorsports: 797 geocoded (includes 4 Canadian dealers)
  - International: 43 geocoded across 32 countries
  - Canada: 22 geocoded
  - Distributors: 161 geocoded, stored in distributors.json
- **Promo List** (Installer_Promotion_Participation): 2,097 promo-only accounts
- **C4C List** (Royal_Purple_C4C_Installer_List): 1,208 C4C-only + 234 on both lists
- **rack_installer flag**: 554 accounts across all types flagged as having RP display racks

## Customer Map (Leaflet.js)
- Embedded via st.components.v1.html() — full Leaflet map with marker clusters
- **4,819 total locations**: 4,658 customer accounts + 161 distributors
- 8 marker categories with distinct colors/icons:
  - **Promo Only (Not on C4C)** — Red (#DC2626) — 2,113 accounts
  - **On Both Lists** — Green (#16A34A) — 235 accounts
  - **C4C Only** — Blue (#2563EB) — 1,208 accounts
  - **Rack Installer** — Purple (#7C3AED) — 240 accounts (new from rack list, not on promo/C4C)
  - **Distributor** — Gold star (#F59E0B) — 161 locations (larger pin with star icon)
  - **Powersports/Motorsports** — Rose (#E11D48) — 797 accounts
  - **International** — Indigo (#4F46E5) — 43 locations (32 countries)
  - **Canada** — Emerald (#059669) — 22 locations
- Filters: search bar, state dropdown, county dropdown (cascading), type dropdown (all 8 types)
- Compact horizontal legend with all 8 categories
- Stats bar with per-type colored counts
- Metrics: two-row layout (Total Locations, Installer Accounts, Distributors, Powersports / Promo Only, On Both, C4C Only, Rack Installer)
- **Export Map Data**: Branded Excel workbook (62 sheets) with per-state tabs, county breakdown, distributor tab
- **Export C4C Report**: 19-sheet comprehensive account intelligence workbook
- ESRI English-only basemap tiles

## RPO Autocare C4C Gap Analysis
- **rpo_autocare_processed.json** — 4,125 RPO Autocare 2025 installer accounts cross-referenced against C4C, Promo, and Rack lists
- Results: 701 On C4C (17.0%), 836 Promo Only, 65 Rack Only, 2,523 Not in System → **3,424 total not on C4C**
- Clean names (ID numbers stripped from "187612 | Name" format)
- Filterable by C4C status, sortable by sales/name/district/region
- CSV export for any filtered view
- Located on Customer Map page below the map and export sections

## C4C Report (19 Sheets)
1. Dashboard — C4C explanation, network overview (all 8 types), gap summary
2. All Accounts — 4,683 master list with Type, County, Rack status (filterable)
3. State Breakdown — all types per state: Promo, Both, C4C, Rack, Dist, Powersports, Gap%, C4C Rate
4. County Breakdown — 1,088 counties with all type counts and gap analysis
5. Not on C4C — 2,097 accounts needing C4C onboarding (with county, rack status)
6. C4C Matched — 1,442 accounts on C4C (with C4C status type, rack status)
7. Top Priority States — ranked by volume of accounts not on C4C
8. Top Priority Counties — top 200 counties by onboarding need
9. Distributors — 58 locations with full details
10. Rack Installers — 554 flagged accounts with C4C status color-coding
11. Powersports — 804 powersports/motorsports accounts
12. International — 28 global partner locations
13. Canada — 14 Canadian locations
14. Distributor Coverage — installers per distributor ratio, identifies distribution gaps
15. Reconciliation — C4C vs Promo list cross-reference and source file totals
16. C4C Duplicates — duplicate entries in C4C source list
17. Promo Duplicates — duplicate entries in promo source list
18. On C4C Only — accounts on C4C but not matched to promo list
19. Failed to Geolocate — accounts with invalid addresses
- Every data sheet has auto-filters on all columns + frozen header rows
- Color-coded by account type throughout (type-specific fill colors)

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

## Product Reference Database (codes_db.json)
- **11 RP product lines, 66 total SKUs**: HP API (11), HMX (6), HPS (5), RP Synthetic (6), Duralec (6), XPR Racing (10), Max-Cycle (2), Break-In Oil (1), HP 2-C (1), Snow 2-C (1), Additives/Specialty (17)
- **6 competitor brands**: CAM2, Valvoline, Mobil 1, Castrol, Pennzoil, Chase's Oil
- **SKU Lookup tab**: Search any code for full product reference card with catalog descriptions, application notes, and cross-series recommendations
- **Code detector** (code_detector.py): Auto-classifies new codes from uploads by prefix — XPR, HPS, HMX, RMS, RSD, RS, RP
- **Admin panel**: Login-protected CRUD for codes_db.json; cache cleared after saves via `load_codes_db.clear()`
- Source data: 2025 RP Consumer Products Catalog + Consumer Brochure (pdfplumber extracted)

## Dependencies
- Python 3.11
- streamlit, python-pptx, openpyxl, Pillow, plotly, pgeocode, pdfplumber

## Running
```
streamlit run app.py --server.port 5000
```
