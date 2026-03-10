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
    padding: 8px 12px;
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    font-size: 11px;
    display: flex;
    flex-wrap: wrap;
    gap: 4px 12px;
    max-width: 420px;
    align-items: center;
}}

#legend h4 {{
    margin: 0;
    font-size: 11px;
    color: #4B2D8A;
    width: 100%;
}}

.legend-item {{
    display: flex;
    align-items: center;
    gap: 4px;
    white-space: nowrap;
}}

.legend-dot {{
    width: 10px;
    height: 10px;
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

.type-promo-only {{ background: #FEE2E2; color: #991B1B; }}
.type-both-lists {{ background: #D1FAE5; color: #065F46; }}
.type-c4c-only {{ background: #DBEAFE; color: #1E40AF; }}
.type-distributor {{ background: #FEF3C7; color: #92400E; }}
.type-rack-installer {{ background: #F3E8FF; color: #6B21A8; }}
.type-powersports {{ background: #FFE4E6; color: #9F1239; }}
.type-international {{ background: #E0E7FF; color: #3730A3; }}
.type-canada {{ background: #ECFDF5; color: #065F46; }}

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
        <select id="county-filter">
            <option value="">All Counties</option>
        </select>
        <select id="type-filter">
            <option value="">All Types</option>
            <option value="Promo Only (Not on C4C)">Promo Only</option>
            <option value="On Both Lists">On Both Lists</option>
            <option value="C4C Only">C4C Only</option>
            <option value="Rack Installer">Rack Installer</option>
            <option value="Distributor">Distributor</option>
            <option value="Powersports/Motorsports">Powersports/Motorsports</option>
            <option value="International">International</option>
            <option value="Canada">Canada</option>
        </select>
    </div>

    <div id="stats-bar">Loading...</div>

    <div id="legend">
        <h4>Account Types</h4>
        <div class="legend-item"><div class="legend-dot" style="background:#DC2626;"></div> Promo Only</div>
        <div class="legend-item"><div class="legend-dot" style="background:#16A34A;"></div> On Both Lists</div>
        <div class="legend-item"><div class="legend-dot" style="background:#2563EB;"></div> C4C Only</div>
        <div class="legend-item"><div class="legend-dot" style="background:#7C3AED;"></div> Rack Installer</div>
        <div class="legend-item"><div class="legend-dot" style="background:#F59E0B;border:2px solid #D97706;"></div> Distributor</div>
        <div class="legend-item"><div class="legend-dot" style="background:#E11D48;"></div> Powersports</div>
        <div class="legend-item"><div class="legend-dot" style="background:#4F46E5;"></div> International</div>
        <div class="legend-item"><div class="legend-dot" style="background:#059669;"></div> Canada</div>
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
    'Promo Only (Not on C4C)': '#DC2626',
    'On Both Lists': '#16A34A',
    'C4C Only': '#2563EB',
    'Rack Installer': '#7C3AED',
    'Distributor': '#F59E0B',
    'Powersports/Motorsports': '#E11D48',
    'International': '#4F46E5',
    'Canada': '#059669'
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
    let svg;
    if (type === 'Distributor') {{
        svg = `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="44" viewBox="0 0 32 44">
            <path d="M16 0C7.16 0 0 7.16 0 16c0 12 16 28 16 28s16-16 16-28C32 7.16 24.84 0 16 0z" fill="${{color}}" stroke="#D97706" stroke-width="1.5"/>
            <polygon points="16,6 18.5,12.5 25,13 20,17.5 21.5,24 16,20.5 10.5,24 12,17.5 7,13 13.5,12.5" fill="white"/>
        </svg>`;
        return L.divIcon({{
            html: svg,
            className: '',
            iconSize: [32, 44],
            iconAnchor: [16, 44],
            popupAnchor: [0, -40]
        }});
    }}
    svg = `<svg xmlns="http://www.w3.org/2000/svg" width="28" height="40" viewBox="0 0 28 40">
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

    let typeClass = 'promo-only';
    if (c.type === 'Promo Only (Not on C4C)') typeClass = 'promo-only';
    else if (c.type === 'On Both Lists') typeClass = 'both-lists';
    else if (c.type === 'C4C Only') typeClass = 'c4c-only';
    else if (c.type === 'Rack Installer') typeClass = 'rack-installer';
    else if (c.type === 'Distributor') typeClass = 'distributor';
    else if (c.type === 'Powersports/Motorsports') typeClass = 'powersports';
    else if (c.type === 'International') typeClass = 'international';
    else if (c.type === 'Canada') typeClass = 'canada';
    const countyLine = c.county ? `<p style="margin:2px 0;color:#6B7280;font-size:12px;">${{c.county}} County</p>` : '';
    const popupHtml = `
        <div class="popup-content">
            <h3>${{c.store_name}}</h3>
            <p>${{c.address || ''}}</p>
            <p>${{c.city}}, ${{c.state}} ${{c.zip || ''}}</p>
            ${{countyLine}}
            <span class="popup-type type-${{typeClass}}">${{c.type || 'Unknown'}}</span>
        </div>
    `;

    const marker = L.marker([c.latitude, c.longitude], {{
        icon: createIcon(c.type || 'Unknown')
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

const counties = [...new Set(customers.map(c => c.county).filter(Boolean))].sort();
const countySelect = document.getElementById('county-filter');
counties.forEach(function(co) {{
    const opt = document.createElement('option');
    opt.value = co;
    opt.textContent = co + ' County';
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
        opt.value = co;
        opt.textContent = co + ' County';
        cf.appendChild(opt);
    }});
}});

function filterMarkers() {{
    const searchTerm = document.getElementById('search-input').value.toLowerCase();
    const stateVal = document.getElementById('state-filter').value;
    const countyVal = document.getElementById('county-filter').value;
    const typeVal = document.getElementById('type-filter').value;

    markerClusterGroup.clearLayers();
    let visibleCount = 0;

    allMarkers.forEach(function(marker) {{
        const c = marker._customerData;
        let show = true;

        if (searchTerm && !c.store_name.toLowerCase().includes(searchTerm)) show = false;
        if (stateVal && c.state !== stateVal) show = false;
        if (countyVal && (c.county || '') !== countyVal) show = false;
        if (typeVal && c.type !== typeVal) show = false;

        if (show) {{
            markerClusterGroup.addLayer(marker);
            visibleCount++;
        }}
    }});

    const typeCounts = {{}};
    markerClusterGroup.eachLayer(function(m) {{
        const t = m._customerData.type;
        typeCounts[t] = (typeCounts[t] || 0) + 1;
    }});

    const tc = typeCounts;
    document.getElementById('stats-bar').innerHTML =
        '<strong>' + visibleCount + '</strong> of ' + allMarkers.length + ' shown — ' +
        '<span style="color:#DC2626">&#9679;' + (tc['Promo Only (Not on C4C)']||0) + '</span> ' +
        '<span style="color:#16A34A">&#9679;' + (tc['On Both Lists']||0) + '</span> ' +
        '<span style="color:#2563EB">&#9679;' + (tc['C4C Only']||0) + '</span> ' +
        '<span style="color:#7C3AED">&#9679;' + (tc['Rack Installer']||0) + '</span> ' +
        '<span style="color:#F59E0B">&#9733;' + (tc['Distributor']||0) + '</span> ' +
        '<span style="color:#E11D48">&#9679;' + (tc['Powersports/Motorsports']||0) + '</span> ' +
        '<span style="color:#4F46E5">&#9679;' + (tc['International']||0) + '</span> ' +
        '<span style="color:#059669">&#9679;' + (tc['Canada']||0) + '</span>';

    updateSidebar();

    if (visibleCount > 0 && (stateVal || countyVal || searchTerm)) {{
        const bounds = markerClusterGroup.getBounds();
        if (bounds.isValid()) map.fitBounds(bounds, {{ padding: [50, 50] }});
    }}
}}

document.getElementById('search-input').addEventListener('input', filterMarkers);
document.getElementById('state-filter').addEventListener('change', function() {{ filterMarkers(); }});
document.getElementById('county-filter').addEventListener('change', filterMarkers);
document.getElementById('type-filter').addEventListener('change', filterMarkers);

filterMarkers();

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
        const dotColor = TYPE_COLORS[c.type] || '#999';
        const countyInfo = c.county ? ' (' + c.county + ' Co.)' : '';
        div.innerHTML = '<div class="si-name">' + c.store_name + '</div>' +
            '<div class="si-loc"><span style="color:' + dotColor + '">&#9679;</span> ' + c.city + ', ' + c.state + countyInfo + ' &middot; ' + (c.type || 'Unknown') + '</div>';
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
