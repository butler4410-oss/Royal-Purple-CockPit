import json
import os
import csv
import io

CUSTOMERS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "customers.json")

def load_customers():
    if os.path.exists(CUSTOMERS_PATH):
        with open(CUSTOMERS_PATH, "r") as f:
            return json.load(f)
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

#toolbar {{
    position: absolute;
    top: 10px;
    left: 60px;
    z-index: 1000;
    display: flex;
    gap: 8px;
    align-items: center;
    flex-wrap: wrap;
}}

#toolbar input, #toolbar select {{
    padding: 8px 12px;
    border: 2px solid #E2E8F0;
    border-radius: 6px;
    font-size: 14px;
    background: white;
    box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    outline: none;
    transition: border-color 0.2s;
}}

#toolbar input:focus, #toolbar select:focus {{
    border-color: #4B2D8A;
}}

#toolbar input {{ width: 220px; }}
#toolbar select {{ min-width: 140px; }}

#stats-bar {{
    position: absolute;
    bottom: 10px;
    left: 10px;
    z-index: 1000;
    background: rgba(75, 45, 138, 0.9);
    color: white;
    padding: 6px 14px;
    border-radius: 6px;
    font-size: 13px;
    font-weight: 500;
    box-shadow: 0 2px 6px rgba(0,0,0,0.2);
}}

#legend {{
    position: absolute;
    bottom: 10px;
    right: 10px;
    z-index: 1000;
    background: white;
    padding: 10px 14px;
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    font-size: 13px;
}}

#legend h4 {{
    margin: 0 0 6px 0;
    font-size: 13px;
    color: #4B2D8A;
}}

.legend-item {{
    display: flex;
    align-items: center;
    gap: 6px;
    margin-bottom: 3px;
}}

.legend-dot {{
    width: 12px;
    height: 12px;
    border-radius: 50%;
    flex-shrink: 0;
}}

.leaflet-popup-content-wrapper {{
    border-radius: 8px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}}

.popup-content {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    min-width: 180px;
}}

.popup-content h3 {{
    margin: 0 0 4px 0;
    color: #4B2D8A;
    font-size: 15px;
}}

.popup-content p {{
    margin: 2px 0;
    color: #475569;
    font-size: 13px;
}}

.popup-type {{
    display: inline-block;
    padding: 2px 8px;
    border-radius: 10px;
    font-size: 11px;
    font-weight: 600;
    margin-top: 4px;
}}

.type-installer {{ background: #F0FDF4; color: #166534; }}
.type-retail {{ background: #F5F3FF; color: #4B2D8A; }}
.type-distributor {{ background: #EFF6FF; color: #1E40AF; }}

#sidebar {{
    position: absolute;
    top: 10px;
    right: 10px;
    z-index: 1000;
    width: 280px;
    max-height: {height - 60}px;
    background: white;
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    display: none;
    overflow: hidden;
}}

#sidebar-header {{
    padding: 10px 14px;
    background: #4B2D8A;
    color: white;
    font-weight: 600;
    font-size: 14px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    cursor: pointer;
}}

#sidebar-list {{
    overflow-y: auto;
    max-height: {height - 110}px;
    padding: 0;
}}

.sidebar-item {{
    padding: 8px 14px;
    border-bottom: 1px solid #F1F5F9;
    cursor: pointer;
    font-size: 13px;
    transition: background 0.15s;
}}

.sidebar-item:hover {{ background: #F8F5FF; }}
.sidebar-item .si-name {{ font-weight: 600; color: #1E293B; }}
.sidebar-item .si-loc {{ color: #64748B; font-size: 12px; }}

#toggle-sidebar {{
    position: absolute;
    top: 10px;
    right: 10px;
    z-index: 999;
    background: white;
    border: none;
    border-radius: 6px;
    padding: 8px 12px;
    cursor: pointer;
    box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    font-size: 13px;
    font-weight: 500;
    color: #4B2D8A;
}}

#toggle-sidebar:hover {{ background: #F8F5FF; }}
</style>
</head>
<body>
<div id="map-container">
    <div id="map"></div>

    <div id="toolbar">
        <input type="text" id="search-input" placeholder="Search store name..." />
        <select id="state-filter">
            <option value="">All States</option>
        </select>
        <select id="type-filter">
            <option value="">All Types</option>
            <option value="Retail">Retail</option>
            <option value="Installer">Installer</option>
            <option value="Distributor">Distributor</option>
        </select>
    </div>

    <div id="stats-bar">Loading...</div>

    <div id="legend">
        <h4>Location Types</h4>
        <div class="legend-item"><div class="legend-dot" style="background:#7C3AED;"></div> Retail</div>
        <div class="legend-item"><div class="legend-dot" style="background:#16A34A;"></div> Installer</div>
        <div class="legend-item"><div class="legend-dot" style="background:#2563EB;"></div> Distributor</div>
    </div>

    <button id="toggle-sidebar" onclick="toggleSidebar()">&#9776; List</button>

    <div id="sidebar">
        <div id="sidebar-header" onclick="toggleSidebar()">
            <span>Locations</span>
            <span>&times;</span>
        </div>
        <div id="sidebar-list"></div>
    </div>
</div>

<script>
const customers = {customers_json};

const TYPE_COLORS = {{
    'Retail': '#7C3AED',
    'Installer': '#16A34A',
    'Distributor': '#2563EB'
}};

const map = L.map('map', {{
    zoomControl: true,
    scrollWheelZoom: true
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
        let size = 'small';
        if (count > 20) size = 'large';
        else if (count > 5) size = 'medium';
        return L.divIcon({{
            html: '<div style="background:#4B2D8A;color:white;border-radius:50%;width:40px;height:40px;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;box-shadow:0 2px 6px rgba(0,0,0,0.3);">' + count + '</div>',
            className: 'custom-cluster',
            iconSize: L.point(40, 40)
        }});
    }}
}});

function createIcon(type) {{
    const color = TYPE_COLORS[type] || '#7C3AED';
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="28" height="40" viewBox="0 0 28 40">
        <path d="M14 0C6.27 0 0 6.27 0 14c0 10.5 14 26 14 26s14-15.5 14-26C28 6.27 21.73 0 14 0z" fill="${{color}}"/>
        <circle cx="14" cy="14" r="6" fill="white"/>
    </svg>`;
    return L.divIcon({{
        html: svg,
        className: '',
        iconSize: [28, 40],
        iconAnchor: [14, 40],
        popupAnchor: [0, -36]
    }});
}}

let allMarkers = [];

customers.forEach(function(c) {{
    if (!c.latitude || !c.longitude) return;

    const typeClass = c.type ? c.type.toLowerCase() : 'retail';
    const popupHtml = `
        <div class="popup-content">
            <h3>${{c.store_name}}</h3>
            <p>${{c.address || ''}}</p>
            <p>${{c.city}}, ${{c.state}} ${{c.zip || ''}}</p>
            <span class="popup-type type-${{typeClass}}">${{c.type || 'Retail'}}</span>
        </div>
    `;

    const marker = L.marker([c.latitude, c.longitude], {{
        icon: createIcon(c.type || 'Retail')
    }}).bindPopup(popupHtml);

    marker._customerData = c;
    allMarkers.push(marker);
}});

allMarkers.forEach(m => markerClusterGroup.addLayer(m));
map.addLayer(markerClusterGroup);

const states = [...new Set(customers.map(c => c.state).filter(Boolean))].sort();
const stateSelect = document.getElementById('state-filter');
states.forEach(function(s) {{
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    stateSelect.appendChild(opt);
}});

function filterMarkers() {{
    const searchTerm = document.getElementById('search-input').value.toLowerCase();
    const stateVal = document.getElementById('state-filter').value;
    const typeVal = document.getElementById('type-filter').value;

    markerClusterGroup.clearLayers();
    let visibleCount = 0;

    allMarkers.forEach(function(marker) {{
        const c = marker._customerData;
        let show = true;

        if (searchTerm && !c.store_name.toLowerCase().includes(searchTerm)) show = false;
        if (stateVal && c.state !== stateVal) show = false;
        if (typeVal && c.type !== typeVal) show = false;

        if (show) {{
            markerClusterGroup.addLayer(marker);
            visibleCount++;
        }}
    }});

    document.getElementById('stats-bar').textContent =
        visibleCount + ' of ' + allMarkers.length + ' locations shown';

    updateSidebar();

    if (visibleCount > 0 && (stateVal || searchTerm)) {{
        const bounds = markerClusterGroup.getBounds();
        if (bounds.isValid()) map.fitBounds(bounds, {{ padding: [50, 50] }});
    }}
}}

document.getElementById('search-input').addEventListener('input', filterMarkers);
document.getElementById('state-filter').addEventListener('change', filterMarkers);
document.getElementById('type-filter').addEventListener('change', filterMarkers);

document.getElementById('stats-bar').textContent =
    allMarkers.length + ' of ' + allMarkers.length + ' locations shown';

function updateSidebar() {{
    const list = document.getElementById('sidebar-list');
    list.innerHTML = '';

    const visible = [];
    markerClusterGroup.eachLayer(function(marker) {{
        visible.push(marker);
    }});

    visible.sort(function(a, b) {{
        return a._customerData.store_name.localeCompare(b._customerData.store_name);
    }});

    visible.forEach(function(marker) {{
        const c = marker._customerData;
        const div = document.createElement('div');
        div.className = 'sidebar-item';
        div.innerHTML = '<div class="si-name">' + c.store_name + '</div>' +
            '<div class="si-loc">' + c.city + ', ' + c.state + ' &middot; ' + (c.type || 'Retail') + '</div>';
        div.onclick = function() {{
            map.setView([c.latitude, c.longitude], 15);
            marker.openPopup();
        }};
        list.appendChild(div);
    }});
}}

let sidebarOpen = false;
function toggleSidebar() {{
    sidebarOpen = !sidebarOpen;
    document.getElementById('sidebar').style.display = sidebarOpen ? 'block' : 'none';
    document.getElementById('toggle-sidebar').style.display = sidebarOpen ? 'none' : 'block';
    if (sidebarOpen) updateSidebar();
}}
</script>
</body>
</html>
"""
    return html
