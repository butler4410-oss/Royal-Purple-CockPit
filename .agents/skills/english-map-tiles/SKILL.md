---
name: english-map-tiles
description: Configure Leaflet.js maps with English-only ESRI tile layers. Use when the map shows labels in foreign languages, when switching tile providers, or when setting up a new Leaflet map that needs English-only labels.
---

# English Map Tiles

Replace multilingual tile layers (CartoDB, OpenStreetMap) with ESRI's English-only tile layers for clean, US-focused Leaflet maps.

## The Problem

Default CartoDB/OSM tile layers (`light_all`, `voyager`, etc.) render labels in each region's local language — "AFRIKA" in Arabic script, "AMERICA DO SUL" in Portuguese, etc. This looks unprofessional for US-focused business tools.

## The Solution

ESRI's World Light Gray Canvas tiles are English-only at every zoom level, have a clean gray aesthetic, and require no API key.

### Two-Layer Setup

Use both the base layer (gray background, borders) and the reference layer (English labels):

```javascript
// Base — gray landmass, borders, water
L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}', {
    attribution: 'Tiles &copy; Esri',
    maxZoom: 16
}).addTo(map);

// Reference — English-only labels on top
L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Reference/MapServer/tile/{z}/{y}/{x}', {
    attribution: '',
    maxZoom: 16
}).addTo(map);
```

### In Python f-strings (Streamlit)

Braces must be double-escaped inside f-strings:

```python
L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{{z}}/{{y}}/{{x}}', {{
    attribution: 'Tiles &copy; Esri',
    maxZoom: 16
}}).addTo(map);

L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Reference/MapServer/tile/{{z}}/{{y}}/{{x}}', {{
    attribution: '',
    maxZoom: 16,
    opacity: 1
}}).addTo(map);
```

## ESRI Tile URL Format

**Important:** ESRI uses `{z}/{y}/{x}` order (y before x), unlike OSM/CartoDB which use `{z}/{x}/{y}`.

## Alternative ESRI Layers

| Layer | URL Path | Style |
|---|---|---|
| Light Gray Base | `Canvas/World_Light_Gray_Base` | Minimal gray, no labels |
| Light Gray Reference | `Canvas/World_Light_Gray_Reference` | Labels only (overlay) |
| World Street Map | `World_Street_Map` | Colorful, English labels built in |
| World Topo Map | `World_Topo_Map` | Topographic, English labels |
| Dark Gray Base | `Canvas/World_Dark_Gray_Base` | Dark theme, no labels |
| Dark Gray Reference | `Canvas/World_Dark_Gray_Reference` | Dark theme labels (overlay) |

All URLs follow the pattern:
```
https://server.arcgisonline.com/ArcGIS/rest/services/{LAYER}/MapServer/tile/{z}/{y}/{x}
```

## Layer Ordering

Add tile layers **before** marker layers. In Leaflet's z-index stack:
- `tilePane` (z-index 200) — tile layers render here
- `overlayPane` (z-index 400) — markers render here

So markers always appear above tile labels automatically.

## Map Configuration

For US-focused maps, pair with:

```javascript
const map = L.map('map', {
    zoomControl: true,
    scrollWheelZoom: true,
    attributionControl: false    // removes Leaflet attribution flag
}).setView([39.8283, -98.5795], 4);  // center of contiguous US, zoom 4
```

## Current Implementation

The map is in `customer_map.py` in the `build_leaflet_html()` function, around line 510.
