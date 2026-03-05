# Royal Purple Partnership Hub

## Overview
"The Royal Purple Partnership Hub by ThrottlePro" — a Streamlit web application for Royal Purple installer partners. Generates branded PowerPoint presentations from Excel reports, with adaptive parsing that handles varying column layouts across different regions and distributors. Supports optional distribution map image uploads.

## Architecture
- **app.py** — Streamlit frontend with sidebar navigation (Report Generator, Product Reference). Dual file upload (Excel + optional map images), data preview with KPI metrics, tabbed store rankings/details, and PPTX download.
- **report_generator.py** — Core PPTX engine using python-pptx. Adaptive Excel parser detects columns by header name patterns (not hardcoded indices). Produces branded slides with Royal Purple design system.
- **assets/** — Royal Purple logo files for both Streamlit UI and PPTX slides

## Excel Parsing (Adaptive)
- Columns detected by header name patterns via HEADER_PATTERNS dict
- Primary fields: date, product, invoices, revenue, avg_rev, vehicles
- Fallback fields: amount→revenue, qty→invoices (used only when primary not found)
- Skipped sheets: Report Summary, Summary, Totals, Notes, Instructions, Template, Info
- Header row auto-detected in first 5 rows by keyword scanning
- Handles varying column counts, names, and orders
- Robust avg_rev calculation with guard against <$1 values and missing-column fallbacks

## Logo Assets (assets/)
- `Royal Purple White Logo.png` — Content slide badges and footer bars
- `RPMO_logo_BF_Outline.png` — Full logo, Streamlit sidebar
- `rp_synthetic_expert_white.png` — White "Synthetic Expert" logo for dark backgrounds
- `RP_Synthetic_Expert_Logo_Black_Text.png` — Black text variant, Streamlit header
- `25-RYP-02147 Employee LinkedIn Thumbnails P1-6.jpg` — "NEVER SETTLE" background

## Slide Structure (dynamic count)
1. Cover (Never Settle bg, expert logo, "NEVER SETTLE", stat box)
2. Table of Contents
3. Executive Summary KPIs
4. Executive Observations
5. Revenue Overview
6+. Store Performance Ranking (auto-paginating)
7+. Store Performance Matrix (auto-paginating)
8+. Product Mix Analysis
9+. Distribution Map slides (0+, from uploaded images)
10+. Section Divider
11+. Store Deep Dives (one per store)
N-1. Next Steps & Recommendations
N. Closing / Thank You

## Product Category Mapping
- Longest-prefix-first matching (RSD before RS)
- Categories: Royal Purple Syn, High Mileage, Duralec, Max-Clean, Max-Atomizer, Other

## Dependencies
- Python 3.11
- streamlit, python-pptx, openpyxl, Pillow

## Design System
- Purple: #4B2D8A, Gold: #C8973A, Dark: #1E1035, Off-White: #F8F5FF

## Running
```
streamlit run app.py --server.port 5000
```
