# Royal Purple Partnership Hub

## Overview
"The Royal Purple Partnership Hub by ThrottlePro" — a Streamlit web app for Royal Purple installer partners with three sections: Report Generator (Excel→PPTX), interactive Distribution Map (Plotly choropleth), and Product Reference.

## Architecture
- **app.py** — Streamlit frontend with 3-page sidebar nav, interactive US map, report generation
- **report_generator.py** — PPTX engine + adaptive Excel parser (header auto-detection, consolidated sheet support, currency stripping, revenue heuristics)
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

## Interactive Distribution Map
- Plotly go.Choropleth with USA scope
- 6 distributors color-coded: Texas Enterprises (green), American Lube Supply (blue), Avery Oil (salmon), Arnold Oil (silver), Dennis K Burke (gold), Brennan Oil (crimson)
- Filterable by distributor via multiselect
- ABE Legend with state counts
- State detail selector with colored accent bars

## Logo Assets (assets/)
- `Royal Purple White Logo.png` — Content slide badges and footer bars
- `RPMO_logo_BF_Outline.png` — Full logo, Streamlit sidebar
- `rp_synthetic_expert_white.png` — White "Synthetic Expert" for dark backgrounds
- `RP_Synthetic_Expert_Logo_Black_Text.png` — Black text, Streamlit header
- `25-RYP-02147 Employee LinkedIn Thumbnails P1-6.jpg` — "NEVER SETTLE" background

## PPTX Slide Structure (dynamic)
Cover → TOC → Exec Summary → Exec Observations → Revenue Overview → Rankings → Matrix → Product Mix → [Distribution Maps] → Section Divider → Store Deep Dives → Next Steps → Closing

## Dependencies
- Python 3.11
- streamlit, python-pptx, openpyxl, Pillow, plotly

## Running
```
streamlit run app.py --server.port 5000
```
