# Royal Purple Report Generator

## Overview
A Streamlit web application that generates branded 23-slide PowerPoint presentations from Royal Purple installer program Excel reports. Features official Royal Purple logos throughout both the web app and the generated PPTX.

## Architecture
- **app.py** — Streamlit frontend: file upload, data preview with metrics, report generation + download. Shows RP logos in sidebar and header.
- **report_generator.py** — Core PPTX generation engine using python-pptx. Parses Excel with openpyxl, produces branded slides with Royal Purple design system (deep purple + gold palette, Calibri typography, stat cards, bar charts, ranking tables). Embeds logos on every slide.
- **assets/** — Royal Purple logo files used in both the Streamlit UI and PPTX slides

## Logo Assets (assets/)
- `Royal Purple White Logo.png` — Used in content slide badges (top-right), footer bars, and section dividers
- `RPMO_logo_BF_Outline.png` — Full Synthetic Oil logo, used in Streamlit sidebar
- `RP_Synthetic_Expert_Logo_Yellow_Text.png` — Yellow text variant, used on cover and closing slides (dark backgrounds)
- `RP_Synthetic_Expert_Logo_Black_Text.png` — Black text variant, used in Streamlit app header
- `25-RYP-02147 Employee LinkedIn Thumbnails P1-6.jpg` — Purple honeycomb background used on cover, section divider, and closing slides
- `Better Oil Starts Here.png` — Marketing asset (available for future use)

## Slide Structure (23 slides)
1. Cover (branded background image, expert logo, stat box)
2. Table of Contents
3. Executive Summary KPIs
4. Executive Observations
5. Revenue Overview
6. Store Performance Ranking (auto-paginating)
7. Store Performance Matrix (auto-paginating)
8. Product Mix Analysis
9. Section Divider (background image, centered logo)
10-21. Store Deep Dives (one per store, ranked by revenue)
22. Next Steps & Recommendations
23. Closing / Thank You (background image, expert logo, stat box)

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
