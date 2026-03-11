import json
import os
import csv
import io

CUSTOMERS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "customers.json")
DISTRIBUTORS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "distributors.json")

def load_customers():
    customers = []
    if os.path.exists(CUSTOMERS_PATH):
        with open(CUSTOMERS_PATH, "r") as f:
            customers = json.load(f)
    return customers

def load_distributors():
    if os.path.exists(DISTRIBUTORS_PATH):
        with open(DISTRIBUTORS_PATH, "r") as f:
            raw = json.load(f)
        return [{
            "store_name": d.get("name", ""),
            "address": d.get("address", ""),
            "city": d.get("city", ""),
            "state": d.get("state", ""),
            "zip": d.get("zip", ""),
            "county": d.get("county", ""),
            "country": "US",
            "latitude": d.get("latitude"),
            "longitude": d.get("longitude"),
            "type": "Distributor",
        } for d in raw if d.get("latitude") is not None and d.get("longitude") is not None]
    return []

def parse_csv_customers(csv_content):
    reader = csv.DictReader(io.StringIO(csv_content))
    customers = []
    for row in reader:
        try:
            customers.append({
                "store_name": row.get("store_name", "Unknown"),
                "address": row.get("address", ""),
                "city": row.get("city", ""),
                "state": row.get("state", ""),
                "zip": row.get("zip", ""),
                "latitude": float(row.get("latitude", 0)),
                "longitude": float(row.get("longitude", 0)),
                "type": row.get("type", "Retail"),
            })
        except (ValueError, TypeError):
            continue
    return customers

def get_states(customers):
    states = sorted(set(c["state"] for c in customers if c.get("state")))
    return states

def build_leaflet_html(customers, height=700):
    customers_json = json.dumps(customers)

    html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }}

#map-container {{ position: relative; width: 100%; height: {height}px; }}
#map {{ width: 100%; height: 100%; border-radius: 8px; }}

/* ── Toolbar (top-left search + dropdowns) ── */
#toolbar {{
    position: absolute;
    top: 10px;
    left: 60px;
    z-index: 1000;
    display: flex;
    gap: 6px;
    align-items: center;
    flex-wrap: wrap;
    max-width: calc(100% - 180px);
}}

.search-wrap {{
    position: relative;
    display: flex;
    align-items: center;
}}

#search-input {{
    padding: 8px 32px 8px 12px;
    border: 2px solid #E2E8F0;
    border-radius: 6px;
    font-size: 14px;
    background: white;
    box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    outline: none;
    width: 220px;
    transition: border-color 0.2s;
}}
#search-input:focus {{ border-color: #4B2D8A; }}

#search-clear {{
    position: absolute;
    right: 8px;
    background: none;
    border: none;
    cursor: pointer;
    color: #94A3B8;
    font-size: 16px;
    line-height: 1;
    display: none;
    padding: 0;
}}
#search-clear:hover {{ color: #4B2D8A; }}

#toolbar select {{
    padding: 8px 10px;
    border: 2px solid #E2E8F0;
    border-radius: 6px;
    font-size: 14px;
    background: white;
    box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    outline: none;
    min-width: 130px;
    transition: border-color 0.2s;
    cursor: pointer;
}}
#toolbar select:focus {{ border-color: #4B2D8A; }}

/* ── Type filter pills (below toolbar) ── */
#type-pills {{
    position: absolute;
    top: 52px;
    left: 60px;
    z-index: 1000;
    display: flex;
    gap: 5px;
    flex-wrap: wrap;
    max-width: calc(100% - 80px);
}}

.type-pill {{
    display: flex;
    align-items: center;
    gap: 5px;
    padding: 4px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
    border: 2px solid transparent;
    background: white;
    box-shadow: 0 1px 4px rgba(0,0,0,0.15);
    transition: all 0.15s;
    user-select: none;
    white-space: nowrap;
}}

.type-pill .pill-dot {{
    width: 10px;
    height: 10px;
    border-radius: 50%;
    flex-shrink: 0;
}}

.type-pill .pill-count {{
    font-size: 11px;
    font-weight: 700;
    opacity: 0.8;
}}

.type-pill:hover {{
    box-shadow: 0 2px 8px rgba(0,0,0,0.25);
    transform: translateY(-1px);
}}

.type-pill.active {{
    color: white !important;
    border-color: transparent;
    box-shadow: 0 2px 8px rgba(0,0,0,0.3);
}}

.type-pill.dimmed {{
    opacity: 0.35;
}}

/* ── Stats bar (bottom-left) ── */
#stats-bar {{
    position: absolute;
    bottom: 10px;
    left: 10px;
    z-index: 1000;
    background: rgba(30, 20, 60, 0.88);
    color: white;
    padding: 7px 14px;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 500;
    box-shadow: 0 2px 8px rgba(0,0,0,0.25);
    backdrop-filter: blur(4px);
}}

/* ── Legend (bottom-right) ── */
#legend {{
    position: absolute;
    bottom: 10px;
    right: 10px;
    z-index: 1000;
    background: white;
    padding: 10px 14px;
    border-radius: 10px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.15);
    font-size: 12px;
    min-width: 200px;
}}

#legend-title {{
    font-size: 11px;
    font-weight: 700;
    color: #4B2D8A;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 8px;
    padding-bottom: 6px;
    border-bottom: 1px solid #E2E8F0;
}}

.legend-row {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 3px 0;
}}

.legend-left {{
    display: flex;
    align-items: center;
    gap: 7px;
}}

.legend-swatch {{
    width: 14px;
    height: 14px;
    border-radius: 50%;
    flex-shrink: 0;
}}

.legend-label {{
    color: #374151;
    font-size: 12px;
    font-weight: 500;
}}

.legend-count {{
    font-size: 12px;
    font-weight: 700;
    color: #6B7280;
    min-width: 36px;
    text-align: right;
}}

/* ── Popup ── */
.leaflet-popup-content-wrapper {{
    border-radius: 10px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.18);
    padding: 0;
    overflow: hidden;
}}
.leaflet-popup-content {{ margin: 0; }}

.popup-header {{
    padding: 10px 14px 8px;
    border-left: 5px solid #4B2D8A;
}}

.popup-content h3 {{
    margin: 0 0 2px 0;
    color: #1E293B;
    font-size: 14px;
    font-weight: 700;
    line-height: 1.3;
}}

.popup-addr {{
    color: #64748B;
    font-size: 12px;
    margin: 1px 0;
}}

.popup-footer {{
    padding: 6px 14px 10px;
    display: flex;
    align-items: center;
    gap: 6px;
    flex-wrap: wrap;
}}

.popup-type-badge {{
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: 700;
    color: white;
}}

.popup-county {{
    font-size: 11px;
    color: #94A3B8;
}}

/* ── Sidebar (results list) ── */
#sidebar {{
    position: absolute;
    top: 10px;
    right: 10px;
    z-index: 1000;
    width: 300px;
    max-height: {height - 70}px;
    background: white;
    border-radius: 10px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.18);
    display: none;
    flex-direction: column;
    overflow: hidden;
}}

#sidebar-header {{
    padding: 12px 14px;
    background: #4B2D8A;
    color: white;
    font-weight: 700;
    font-size: 14px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    cursor: pointer;
    flex-shrink: 0;
}}

#sidebar-count {{
    font-size: 12px;
    font-weight: 400;
    opacity: 0.85;
}}

#sidebar-list {{
    overflow-y: auto;
    flex: 1;
}}

.sidebar-item {{
    display: flex;
    align-items: stretch;
    border-bottom: 1px solid #F1F5F9;
    cursor: pointer;
    transition: background 0.12s;
}}

.sidebar-item:hover {{ background: #F8F5FF; }}

.sidebar-color-bar {{
    width: 5px;
    flex-shrink: 0;
}}

.sidebar-item-body {{
    padding: 9px 12px;
    flex: 1;
    min-width: 0;
}}

.si-name {{
    font-weight: 600;
    color: #1E293B;
    font-size: 13px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}}

.si-loc {{
    color: #64748B;
    font-size: 11px;
    margin-top: 2px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}}

.si-type-badge {{
    display: inline-block;
    padding: 1px 7px;
    border-radius: 8px;
    font-size: 10px;
    font-weight: 700;
    color: white;
    margin-top: 3px;
}}

/* ── List toggle button ── */
#toggle-sidebar {{
    position: absolute;
    top: 10px;
    right: 10px;
    z-index: 999;
    background: white;
    border: 2px solid #E2E8F0;
    border-radius: 6px;
    padding: 7px 12px;
    cursor: pointer;
    box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    font-size: 13px;
    font-weight: 600;
    color: #4B2D8A;
    transition: all 0.15s;
}}
#toggle-sidebar:hover {{ background: #F8F5FF; border-color: #4B2D8A; }}

/* ── Clear filters button ── */
#clear-filters {{
    display: none;
    padding: 6px 11px;
    background: #4B2D8A;
    color: white;
    border: none;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
    box-shadow: 0 1px 4px rgba(0,0,0,0.2);
    white-space: nowrap;
}}
#clear-filters:hover {{ background: #3B1F7A; }}
</style>
</head>
<body>
<div id="map-container">
    <div id="map"></div>

    <div id="toolbar">
        <div class="search-wrap">
            <input type="text" id="search-input" placeholder="&#128269; Search store name..." />
            <button id="search-clear" onclick="clearSearch()" title="Clear search">&times;</button>
        </div>
        <select id="state-filter">
            <option value="">All States</option>
        </select>
        <select id="county-filter">
            <option value="">All Counties</option>
        </select>
        <button id="clear-filters" onclick="clearAllFilters()">&#10005; Clear Filters</button>
    </div>

    <div id="type-pills"></div>

    <div id="stats-bar">Loading...</div>

    <div id="legend">
        <div id="legend-title">Account Types</div>
        <div id="legend-rows"></div>
    </div>

    <button id="toggle-sidebar" onclick="toggleSidebar()">&#9776; List</button>

    <div id="sidebar">
        <div id="sidebar-header" onclick="toggleSidebar()">
            <span>Locations <span id="sidebar-count"></span></span>
            <span style="font-size:18px;">&times;</span>
        </div>
        <div id="sidebar-list"></div>
    </div>
</div>

<script>
const customers = {customers_json};

const TYPE_CONFIG = [
    {{ key: 'Promo Only (Not on C4C)', label: 'Promo Only',          color: '#DC2626' }},
    {{ key: 'On Both Lists',            label: 'On Both Lists',       color: '#16A34A' }},
    {{ key: 'C4C Only',                 label: 'C4C Only',            color: '#2563EB' }},
    {{ key: 'Rack Installer',           label: 'Rack Installer',      color: '#7C3AED' }},
    {{ key: 'Distributor',              label: 'Distributor',         color: '#F59E0B' }},
    {{ key: 'Powersports/Motorsports',  label: 'Powersports',         color: '#E11D48' }},
    {{ key: 'International',            label: 'International',       color: '#4F46E5' }},
    {{ key: 'Canada',                   label: 'Canada',              color: '#059669' }},
];

const TYPE_COLOR_MAP = {{}};
TYPE_CONFIG.forEach(t => TYPE_COLOR_MAP[t.key] = t.color);

const map = L.map('map', {{
    zoomControl: true,
    scrollWheelZoom: true,
    attributionControl: false
}}).setView([39.8283, -98.5795], 4);

L.tileLayer('https://{{s}}.basemaps.cartocdn.com/light_all/{{z}}/{{x}}/{{y}}{{r}}.png', {{
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OSM</a> &copy; <a href="https://carto.com/">CARTO</a>',
    subdomains: 'abcd',
    maxZoom: 19
}}).addTo(map);

const markerClusterGroup = L.markerClusterGroup({{
    maxClusterRadius: 50,
    spiderfyOnMaxZoom: true,
    showCoverageOnHover: false,
    iconCreateFunction: function(cluster) {{
        const count = cluster.getChildCount();
        return L.divIcon({{
            html: '<div style="background:#4B2D8A;color:white;border-radius:50%;width:40px;height:40px;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;box-shadow:0 2px 6px rgba(0,0,0,0.3);">' + count + '</div>',
            className: 'custom-cluster',
            iconSize: L.point(40, 40)
        }});
    }}
}});

function createIcon(type) {{
    const color = TYPE_COLOR_MAP[type] || '#7C3AED';
    let svg;
    if (type === 'Distributor') {{
        svg = `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="44" viewBox="0 0 32 44">
            <path d="M16 0C7.16 0 0 7.16 0 16c0 12 16 28 16 28s16-16 16-28C32 7.16 24.84 0 16 0z" fill="${{color}}" stroke="#D97706" stroke-width="1.5"/>
            <polygon points="16,6 18.5,12.5 25,13 20,17.5 21.5,24 16,20.5 10.5,24 12,17.5 7,13 13.5,12.5" fill="white"/>
        </svg>`;
        return L.divIcon({{ html: svg, className: '', iconSize: [32, 44], iconAnchor: [16, 44], popupAnchor: [0, -40] }});
    }}
    svg = `<svg xmlns="http://www.w3.org/2000/svg" width="28" height="40" viewBox="0 0 28 40">
        <path d="M14 0C6.27 0 0 6.27 0 14c0 10.5 14 26 14 26s14-15.5 14-26C28 6.27 21.73 0 14 0z" fill="${{color}}"/>
        <circle cx="14" cy="14" r="6" fill="white"/>
    </svg>`;
    return L.divIcon({{ html: svg, className: '', iconSize: [28, 40], iconAnchor: [14, 40], popupAnchor: [0, -36] }});
}}

let allMarkers = [];

customers.forEach(function(c) {{
    if (!c.latitude || !c.longitude) return;
    const color = TYPE_COLOR_MAP[c.type] || '#7C3AED';
    const countyLine = c.county ? `<span class="popup-county">${{c.county}} County</span>` : '';
    const popupHtml = `
        <div>
            <div class="popup-header" style="border-left-color:${{color}}">
                <h3 class="popup-content">${{c.store_name}}</h3>
                <div class="popup-addr">${{c.address || ''}}</div>
                <div class="popup-addr">${{c.city}}, ${{c.state}} ${{c.zip || ''}}</div>
            </div>
            <div class="popup-footer">
                <span class="popup-type-badge" style="background:${{color}}">${{c.type || 'Unknown'}}</span>
                ${{countyLine}}
            </div>
        </div>
    `;
    const marker = L.marker([c.latitude, c.longitude], {{ icon: createIcon(c.type || 'Unknown') }}).bindPopup(popupHtml);
    marker._customerData = c;
    allMarkers.push(marker);
}});

allMarkers.forEach(m => markerClusterGroup.addLayer(m));
map.addLayer(markerClusterGroup);

const states = [...new Set(customers.map(c => c.state).filter(Boolean))].sort();
const stateSelect = document.getElementById('state-filter');
states.forEach(function(s) {{
    const opt = document.createElement('option');
    opt.value = s; opt.textContent = s;
    stateSelect.appendChild(opt);
}});

const counties = [...new Set(customers.map(c => c.county).filter(Boolean))].sort();
const countySelect = document.getElementById('county-filter');
counties.forEach(function(co) {{
    const opt = document.createElement('option');
    opt.value = co; opt.textContent = co + ' County';
    countySelect.appendChild(opt);
}});

document.getElementById('state-filter').addEventListener('change', function() {{
    const sv = this.value;
    const cf = document.getElementById('county-filter');
    cf.innerHTML = '<option value="">All Counties</option>';
    const filtered = sv
        ? [...new Set(customers.filter(c => c.state === sv).map(c => c.county).filter(Boolean))].sort()
        : counties;
    filtered.forEach(function(co) {{
        const opt = document.createElement('option');
        opt.value = co; opt.textContent = co + ' County';
        cf.appendChild(opt);
    }});
    filterMarkers();
}});

/* ── Type pill filter state ── */
let activeTypeFilter = '';

function buildTypePills() {{
    const container = document.getElementById('type-pills');
    container.innerHTML = '';
    TYPE_CONFIG.forEach(function(tc) {{
        const pill = document.createElement('div');
        pill.className = 'type-pill';
        pill.dataset.key = tc.key;
        pill.innerHTML =
            `<div class="pill-dot" style="background:${{tc.color}}"></div>` +
            `<span class="pill-label">${{tc.label}}</span>` +
            `<span class="pill-count" id="pill-count-${{tc.key.replace(/[^a-z0-9]/gi,'_')}}">0</span>`;
        pill.style.color = tc.color;
        pill.style.borderColor = tc.color + '44';
        pill.addEventListener('click', function() {{
            if (activeTypeFilter === tc.key) {{
                activeTypeFilter = '';
            }} else {{
                activeTypeFilter = tc.key;
            }}
            filterMarkers();
        }});
        container.appendChild(pill);
    }});
}}

function updatePillStyles(typeCounts) {{
    document.querySelectorAll('.type-pill').forEach(function(pill) {{
        const key = pill.dataset.key;
        const cfg = TYPE_CONFIG.find(t => t.key === key);
        const countEl = document.getElementById('pill-count-' + key.replace(/[^a-z0-9]/gi,'_'));
        if (countEl) countEl.textContent = (typeCounts[key] || 0).toLocaleString();

        if (activeTypeFilter === '') {{
            pill.classList.remove('active', 'dimmed');
            pill.style.background = 'white';
            pill.style.color = cfg.color;
        }} else if (activeTypeFilter === key) {{
            pill.classList.add('active');
            pill.classList.remove('dimmed');
            pill.style.background = cfg.color;
            pill.style.color = 'white';
        }} else {{
            pill.classList.add('dimmed');
            pill.classList.remove('active');
            pill.style.background = 'white';
            pill.style.color = cfg.color;
        }}
    }});
}}

function buildLegend() {{
    const container = document.getElementById('legend-rows');
    container.innerHTML = '';
    TYPE_CONFIG.forEach(function(tc) {{
        const row = document.createElement('div');
        row.className = 'legend-row';
        row.innerHTML =
            `<div class="legend-left">` +
            `<div class="legend-swatch" style="background:${{tc.color}}"></div>` +
            `<span class="legend-label">${{tc.label}}</span>` +
            `</div>` +
            `<span class="legend-count" id="leg-count-${{tc.key.replace(/[^a-z0-9]/gi,'_')}}">–</span>`;
        container.appendChild(row);
    }});
}}

function updateLegendCounts(typeCounts) {{
    TYPE_CONFIG.forEach(function(tc) {{
        const el = document.getElementById('leg-count-' + tc.key.replace(/[^a-z0-9]/gi,'_'));
        if (el) el.textContent = (typeCounts[tc.key] || 0).toLocaleString();
    }});
}}

function filterMarkers() {{
    const searchTerm = document.getElementById('search-input').value.toLowerCase().trim();
    const stateVal = document.getElementById('state-filter').value;
    const countyVal = document.getElementById('county-filter').value;

    markerClusterGroup.clearLayers();
    let visibleCount = 0;
    const typeCounts = {{}};

    allMarkers.forEach(function(marker) {{
        const c = marker._customerData;
        let show = true;
        if (searchTerm && !c.store_name.toLowerCase().includes(searchTerm)) show = false;
        if (stateVal && c.state !== stateVal) show = false;
        if (countyVal && (c.county || '') !== countyVal) show = false;
        if (activeTypeFilter && c.type !== activeTypeFilter) show = false;
        if (show) {{
            markerClusterGroup.addLayer(marker);
            visibleCount++;
            typeCounts[c.type] = (typeCounts[c.type] || 0) + 1;
        }}
    }});

    updatePillStyles(typeCounts);
    updateLegendCounts(typeCounts);

    const hasFilters = searchTerm || stateVal || countyVal || activeTypeFilter;
    document.getElementById('clear-filters').style.display = hasFilters ? 'inline-block' : 'none';
    document.getElementById('search-clear').style.display = searchTerm ? 'block' : 'none';

    document.getElementById('stats-bar').innerHTML =
        '<strong>' + visibleCount.toLocaleString() + '</strong> of ' + allMarkers.length.toLocaleString() + ' locations shown';

    updateSidebar(visibleCount);

    if (visibleCount > 0 && hasFilters) {{
        const bounds = markerClusterGroup.getBounds();
        if (bounds.isValid()) map.fitBounds(bounds, {{ padding: [60, 60] }});
    }}
}}

function clearSearch() {{
    document.getElementById('search-input').value = '';
    filterMarkers();
}}

function clearAllFilters() {{
    document.getElementById('search-input').value = '';
    document.getElementById('state-filter').value = '';
    document.getElementById('county-filter').innerHTML = '<option value="">All Counties</option>';
    counties.forEach(function(co) {{
        const opt = document.createElement('option');
        opt.value = co; opt.textContent = co + ' County';
        document.getElementById('county-filter').appendChild(opt);
    }});
    activeTypeFilter = '';
    filterMarkers();
    map.setView([39.8283, -98.5795], 4);
}}

document.getElementById('search-input').addEventListener('input', filterMarkers);
document.getElementById('county-filter').addEventListener('change', filterMarkers);

function updateSidebar(visibleCount) {{
    const list = document.getElementById('sidebar-list');
    const countEl = document.getElementById('sidebar-count');
    list.innerHTML = '';

    const visible = [];
    markerClusterGroup.eachLayer(function(marker) {{ visible.push(marker); }});
    visible.sort(function(a, b) {{ return a._customerData.store_name.localeCompare(b._customerData.store_name); }});

    countEl.textContent = '(' + (visibleCount || visible.length).toLocaleString() + ')';

    visible.forEach(function(marker) {{
        const c = marker._customerData;
        const color = TYPE_COLOR_MAP[c.type] || '#999';
        const typeCfg = TYPE_CONFIG.find(t => t.key === c.type);
        const label = typeCfg ? typeCfg.label : (c.type || 'Unknown');
        const countyInfo = c.county ? c.county + ' Co. · ' : '';

        const item = document.createElement('div');
        item.className = 'sidebar-item';
        item.innerHTML =
            `<div class="sidebar-color-bar" style="background:${{color}}"></div>` +
            `<div class="sidebar-item-body">` +
            `<div class="si-name">${{c.store_name}}</div>` +
            `<div class="si-loc">${{countyInfo}}${{c.city}}, ${{c.state}}</div>` +
            `<span class="si-type-badge" style="background:${{color}}">${{label}}</span>` +
            `</div>`;
        item.onclick = function() {{
            map.setView([c.latitude, c.longitude], 15);
            marker.openPopup();
        }};
        list.appendChild(item);
    }});
}}

let sidebarOpen = false;
function toggleSidebar() {{
    sidebarOpen = !sidebarOpen;
    document.getElementById('sidebar').style.display = sidebarOpen ? 'flex' : 'none';
    document.getElementById('toggle-sidebar').style.display = sidebarOpen ? 'none' : 'block';
    if (sidebarOpen) updateSidebar();
}}

buildTypePills();
buildLegend();
filterMarkers();
</script>
</body>
</html>
"""
    return html
