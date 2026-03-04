# Royal Purple Report Generator

## Overview
A Streamlit web application that generates branded PowerPoint presentations from Royal Purple installer program Excel reports. Features adaptive Excel parsing (works with varying column layouts), optional distribution map slides, official Royal Purple logos, and "NEVER SETTLE" branding throughout.

## Architecture
- **app.py** — Streamlit frontend: dual file upload (Excel report + optional map images), data preview with metrics, report generation + download. Shows RP logos in sidebar and header, "Never Settle" banner in sidebar.
- **report_generator.py** — Core PPTX generation engine using python-pptx. Adaptive Excel parser detects columns by header name patterns. Produces branded slides with Royal Purple design system (deep purple + gold palette, Calibri typography, stat cards, bar charts, ranking tables). Supports distribution map image slides.
- **assets/** — Royal Purple logo files used in both the Streamlit UI and PPTX slides

## Excel Parsing (Adaptive)
- Columns detected by header name patterns, not hardcoded indices
- Recognized patterns: invoice date, operation code/product, revenue, invoices, avg rev, vehicles
- Skipped sheets: Report Summary, Summary, Totals, Notes, Instructions, Template, Info
- Header row auto-detected in first 5 rows by scanning for keywords
- Handles varying column counts, names, and orders across different Royal Purple report formats

## Logo Assets (assets/)
- `Royal Purple White Logo.png` — Content slide badges (top-right) and footer bars
- `RPMO_logo_BF_Outline.png` — Full Synthetic Oil logo, Streamlit sidebar
- `rp_synthetic_expert_white.png` — White text "Synthetic Expert" logo for dark backgrounds (cover, divider, closing)
- `RP_Synthetic_Expert_Logo_Black_Text.png` — Black text variant, Streamlit header
- `25-RYP-02147 Employee LinkedIn Thumbnails P1-6.jpg` — "NEVER SETTLE" background (cover, divider, closing, sidebar)

## Slide Structure (dynamic count)
1. Cover (Never Settle background, expert white logo, "NEVER SETTLE" tagline, stat box)
2. Table of Contents
3. Executive Summary KPIs
4. Executive Observations
5. Revenue Overview
6+. Store Performance Ranking (auto-paginating)
7+. Store Performance Matrix (auto-paginating)
8+. Product Mix Analysis
9+. Distribution Map slides (0 or more, from uploaded images)
10+. Section Divider (Never Settle background)
11+. Store Deep Dives (one per store, ranked by revenue)
N-1. Next Steps & Recommendations
N. Closing / Thank You

## Product Category Mapping
- Prefix matching uses longest-prefix-first order (RSD before RS)
- Categories: Royal Purple Syn (RS*), High Mileage (HMX*/RMS*), Duralec (RSD*), Max-Clean (11722), Max-Atomizer (18000), Other

## Dependencies
- Python 3.11
- streamlit, python-pptx, openpyxl, Pillow

## Design System Colors
- Primary Purple: #4B2D8A
- Gold: #C8973A
- Dark: #1E1035
- Off-White Background: #F8F5FF

## Running
```
streamlit run app.py --server.port 5000
```
