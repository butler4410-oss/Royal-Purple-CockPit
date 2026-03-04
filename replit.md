# Royal Purple Report Generator

## Overview
A Streamlit web application that generates branded 23-slide PowerPoint presentations from Royal Purple installer program Excel reports.

## Architecture
- **app.py** — Streamlit frontend: file upload, data preview, report generation + download
- **report_generator.py** — Core PPTX generation engine using python-pptx. Parses Excel with openpyxl, produces branded slides with Royal Purple design system (deep purple + gold palette, Calibri typography, stat cards, bar charts, ranking tables).

## Slide Structure (23 slides)
1. Cover (branded, with stat box)
2. Table of Contents
3. Executive Summary KPIs
4. Executive Observations
5. Revenue Overview
6. Store Performance Ranking (auto-paginating)
7. Store Performance Matrix (auto-paginating)
8. Product Mix Analysis
9. Section Divider
10–21. Store Deep Dives (one per store, ranked by revenue)
22. Next Steps & Recommendations
23. Closing / Thank You

## Dependencies
- Python 3.11
- streamlit
- python-pptx
- openpyxl

## Design System Colors
- Primary Purple: #4B2D8A
- Gold: #C8973A
- Dark: #1E1035
- Off-White Background: #F8F5FF

## Running
```
streamlit run app.py --server.port 5000
```
