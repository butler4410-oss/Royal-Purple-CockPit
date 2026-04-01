import streamlit as st
import streamlit.components.v1 as components
import tempfile
import os
import json
from report_generator import (
    generate_report, parse_excel, fmt_currency, fmt_number,
    PRODUCT_DESCRIPTIONS, get_product_display_name,
)
from customer_map import load_customers, load_distributors, parse_csv_customers, build_leaflet_html, get_states
from c4c_report_generator import generate_c4c_report
from map_data_exporter import generate_map_export
from code_detector import detect_new_codes, add_new_codes_to_db
import product_reference
import admin_panel
import profit_calculator

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "RP_Synthetic_Expert_Logo_Black_Text.png")
LOGO_SIDEBAR_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "RPMO_logo_BF_Outline.png")
LOGO_NEVER_SETTLE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "25-RYP-02147 Employee LinkedIn Thumbnails P1-6.jpg")
COCKPIT_HERO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "cockpit_hero.png")

st.set_page_config(
    page_title="Royal Purple CockPit | Butler Performance Analytics",
    page_icon="👑",
    layout="wide",
)

# ── Royal Purple CockPit — Glassmorphism Skin ─────────────────────
st.markdown("""
<style>
    /* ═══ GLOBAL ═══ */
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(111,45,189,0.22), transparent 28%),
            radial-gradient(circle at top right, rgba(255,212,71,0.10), transparent 20%),
            linear-gradient(180deg, #0B0712 0%, #110A1B 45%, #0B0712 100%) !important;
        color: #F4F1FF;
    }
    .block-container { padding: 1.4rem 2rem 2rem !important; max-width: 1380px; }
    header[data-testid="stHeader"] { background: rgba(11,7,18,0.85) !important; backdrop-filter: blur(14px); border-bottom: 1px solid rgba(157,107,255,0.18); }

    /* ═══ SIDEBAR ═══ */
    section[data-testid="stSidebar"] { background: linear-gradient(180deg, #0B0712 0%, #151022 100%) !important; border-right: 1px solid rgba(157,107,255,0.18); }
    section[data-testid="stSidebar"] .stRadio > div { gap: 2px; }
    section[data-testid="stSidebar"] .stRadio label { padding: 10px 16px !important; border-radius: 12px; transition: all 0.15s; margin: 1px 0; }
    section[data-testid="stSidebar"] .stRadio label:hover { background: rgba(111,45,189,0.15); }
    section[data-testid="stSidebar"] hr { border-color: rgba(157,107,255,0.18) !important; margin: 12px 0; }

    /* ═══ METRICS ═══ */
    [data-testid="stMetric"] { background: linear-gradient(180deg, rgba(28,20,48,0.90), rgba(16,10,27,0.96)) !important; border: 1px solid rgba(157,107,255,0.18); border-radius: 20px; padding: 16px 18px !important; box-shadow: 0 10px 28px rgba(0,0,0,0.28); }
    [data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 800 !important; color: #F4F1FF !important; }
    [data-testid="stMetricLabel"] { font-size: 0.82rem !important; text-transform: uppercase; letter-spacing: 0.02em; color: #B7A8DB !important; }
    [data-testid="stMetricDelta"] { font-size: 12px !important; }
    [data-testid="stMetricDelta"] svg { display: none; }

    /* ═══ TABS ═══ */
    .stTabs [data-baseweb="tab-list"] { gap: 0; border-bottom: 2px solid rgba(157,107,255,0.18); }
    .stTabs [data-baseweb="tab"] { font-size: 13px !important; font-weight: 600; padding: 12px 24px !important; color: #B7A8DB; border-bottom: 3px solid transparent; background: transparent !important; }
    .stTabs [data-baseweb="tab"]:hover { color: #F4F1FF; background: rgba(111,45,189,0.08) !important; }
    .stTabs [data-baseweb="tab"][aria-selected="true"] { color: #FFD447 !important; border-bottom-color: #FFD447 !important; background: rgba(111,45,189,0.10) !important; }
    .stTabs [data-baseweb="tab-highlight"] { background-color: #FFD447 !important; }
    .stTabs [data-baseweb="tab-border"] { display: none; }

    /* ═══ EXPANDERS ═══ */
    details[data-testid="stExpander"] { background: linear-gradient(180deg, rgba(28,20,48,0.90), rgba(16,10,27,0.96)) !important; border: 1px solid rgba(157,107,255,0.18) !important; border-radius: 20px !important; }
    .streamlit-expanderHeader { background: transparent !important; font-weight: 600; color: #B7A8DB !important; }
    .streamlit-expanderContent { border-top: 1px solid rgba(157,107,255,0.18); }

    /* ═══ DATAFRAMES ═══ */
    [data-testid="stDataFrame"] { border-radius: 16px !important; overflow: hidden; border: 1px solid rgba(157,107,255,0.18); }

    /* ═══ FILE UPLOADER ═══ */
    [data-testid="stFileUploader"] { background: rgba(255,255,255,0.03) !important; border: 2px dashed rgba(157,107,255,0.25) !important; border-radius: 20px !important; padding: 24px !important; }
    [data-testid="stFileUploader"]:hover { border-color: #9D6BFF !important; }

    /* ═══ BUTTONS ═══ */
    .stButton > button { border-radius: 14px !important; font-weight: 700 !important; padding: 10px 24px !important; transition: all 0.15s; border: 1px solid rgba(157,107,255,0.24) !important; }
    .stButton > button[kind="primary"] { background: linear-gradient(180deg, #7B35D5, #5A1FA8) !important; border: none !important; color: white !important; box-shadow: 0 10px 22px rgba(79,25,143,0.35); }
    .stButton > button[kind="primary"]:hover { background: linear-gradient(180deg, #8B45E5, #6A2FB8) !important; transform: translateY(-1px); box-shadow: 0 14px 28px rgba(79,25,143,0.45); border-color: rgba(255,212,71,0.45) !important; }
    .stButton > button[kind="secondary"]:hover { background: rgba(111,45,189,0.12) !important; border-color: #9D6BFF !important; }

    /* ═══ DOWNLOAD BUTTONS ═══ */
    .stDownloadButton > button { background: linear-gradient(180deg, #7B35D5, #5A1FA8) !important; border: none !important; border-radius: 14px !important; font-weight: 700 !important; color: white !important; box-shadow: 0 10px 22px rgba(79,25,143,0.35); }
    .stDownloadButton > button:hover { background: linear-gradient(180deg, #8B45E5, #6A2FB8) !important; transform: translateY(-1px); border-color: rgba(255,212,71,0.45) !important; }

    /* ═══ SELECT / MULTISELECT ═══ */
    [data-baseweb="select"] > div { background: rgba(255,255,255,0.03) !important; border-color: rgba(157,107,255,0.18) !important; border-radius: 14px !important; }
    [data-baseweb="select"] > div:hover { border-color: #9D6BFF !important; }
    [data-baseweb="popover"] { background: #151022 !important; border: 1px solid rgba(157,107,255,0.18) !important; border-radius: 16px !important; }
    [data-baseweb="menu"] { background: #151022 !important; }
    [data-baseweb="menu"] li:hover { background: rgba(111,45,189,0.15) !important; }

    /* ═══ TEXT / NUMBER INPUTS ═══ */
    .stTextInput input, .stNumberInput input, .stTextArea textarea { background: rgba(255,255,255,0.03) !important; border: 1px solid rgba(157,107,255,0.18) !important; border-radius: 14px !important; color: #F4F1FF !important; }
    .stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus { border-color: #9D6BFF !important; box-shadow: 0 0 0 3px rgba(157,107,255,0.15) !important; }

    /* ═══ ALERTS ═══ */
    [data-testid="stAlert"] { border-radius: 16px !important; border: 1px solid rgba(157,107,255,0.18) !important; }

    /* ═══ DIVIDERS ═══ */
    hr { border-color: rgba(157,107,255,0.18) !important; }

    /* ═══ SPINNER ═══ */
    .stSpinner > div > div { border-top-color: #9D6BFF !important; }

    /* ═══ SCROLLBAR ═══ */
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: #0B0712; }
    ::-webkit-scrollbar-thumb { background: rgba(157,107,255,0.25); border-radius: 4px; }
    ::-webkit-scrollbar-thumb:hover { background: #9D6BFF; }

    /* ═══ COLUMN GAPS ═══ */
    [data-testid="stColumn"] { padding: 0 6px; }

    /* ═══ CAPTION ═══ */
    .stCaption, .stMarkdown small { color: #B7A8DB !important; }

    /* ═══ FOOTER ═══ */
    footer { visibility: hidden; }
    footer:after { content: 'Royal Purple CockPit  ·  Powered by Butler Performance Analytics'; visibility: visible; display: block; text-align: center; padding: 12px; font-size: 11px; color: #B7A8DB; letter-spacing: 0.5px; }
</style>
""", unsafe_allow_html=True)


with st.sidebar:
    st.markdown("")
    if os.path.exists(COCKPIT_HERO_PATH):
        st.image(COCKPIT_HERO_PATH, use_container_width=True)
    st.markdown("<div style='text-align:center;padding:12px 0 4px;'><div style='font-size:18px;font-weight:900;color:#4B2D8A;letter-spacing:1px;'>ROYAL PURPLE</div><div style='width:120px;height:2px;background:linear-gradient(90deg,#4B2D8A,#C8A951);margin:6px auto;border-radius:1px;'></div><div style='font-size:11px;font-weight:700;letter-spacing:3px;color:#C8A951;'>COCKPIT</div></div>", unsafe_allow_html=True)
    st.markdown("<div style='text-align:center;padding:6px 0;'><div style='font-size:9px;letter-spacing:2px;color:#8888a8;text-transform:uppercase;'>Powered by</div><div style='font-size:11px;font-weight:700;color:#4B2D8A;letter-spacing:0.5px;'>Butler Performance Analytics</div></div>", unsafe_allow_html=True)
    st.markdown("---")

    nav = st.radio(
        "Navigation",
        ["Home", "Report Generator", "Customer Map", "Product Reference", "Installer Incremental Profit Model", "Admin"],
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown("<div style='text-align:center;padding:6px 0;'><div style='font-size:9px;letter-spacing:2px;color:#8888a8;text-transform:uppercase;'>Powered by</div><div style='font-size:11px;font-weight:700;color:#4B2D8A;letter-spacing:0.5px;'>Butler Performance Analytics</div></div>", unsafe_allow_html=True)


def page_header(title, subtitle):
    st.markdown(
        f"<h1 style='color:#4B2D8A; margin: 0;'>{title}</h1>"
        f"<p style='color:#8888a8; margin: 4px 0 0 0;'>{subtitle}</p>",
        unsafe_allow_html=True,
    )


if nav == "Home":
    import base64 as _b64

    _cockpit_b64 = ""
    if os.path.exists(COCKPIT_HERO_PATH):
        with open(COCKPIT_HERO_PATH, "rb") as _hf:
            _cockpit_b64 = _b64.b64encode(_hf.read()).decode()

    _cockpit_html = ""
    if _cockpit_b64:
        _cockpit_html = (
            f'<img src="data:image/png;base64,{_cockpit_b64}" '
            f'style="width:150px;height:150px;object-fit:cover;border-radius:12px;'
            f'border:1px solid #2a2a45;flex-shrink:0;" />'
        )

    _logo_b64 = ""
    if os.path.exists(LOGO_SIDEBAR_PATH):
        with open(LOGO_SIDEBAR_PATH, "rb") as _lf:
            _logo_b64 = _b64.b64encode(_lf.read()).decode()

    _logo_html = ""
    if _logo_b64:
        _logo_html = (
            f'<img src="data:image/png;base64,{_logo_b64}" '
            f'style="width:150px;flex-shrink:0;" />'
        )

    # ── Hero banner with logos ──
    st.markdown(
        f"""
        <div style="background:linear-gradient(120deg, rgba(111,45,189,0.30), rgba(255,212,71,0.08)),
                    linear-gradient(180deg, rgba(21,16,34,0.96), rgba(11,7,18,0.96));
                    border:1px solid rgba(157,107,255,0.22);border-radius:28px;padding:28px 36px;
                    margin-bottom:24px;position:relative;overflow:hidden;
                    box-shadow:0 20px 52px rgba(0,0,0,0.34);">
            <div style="position:absolute;inset:0;background:radial-gradient(circle at 75% 25%,rgba(157,107,255,0.20),transparent 26%);pointer-events:none;"></div>
            <div style="display:flex;align-items:center;gap:28px;position:relative;">
                {_logo_html}
                <div style="text-align:center;flex:1;">
                    <div style="font-size:36px;font-weight:900;color:#9D6BFF;letter-spacing:1px;
                                line-height:1.1;">ROYAL PURPLE</div>
                    <div style="width:140px;height:2px;background:linear-gradient(90deg,#9D6BFF,#FFD447);
                                margin:10px auto;border-radius:1px;"></div>
                    <div style="font-size:16px;font-weight:700;letter-spacing:5px;color:#B7A8DB;
                                text-transform:uppercase;margin-bottom:14px;">COCKPIT</div>
                    <div style="font-size:10px;letter-spacing:2.5px;color:#B7A8DB;
                                text-transform:uppercase;margin-bottom:2px;">Powered by</div>
                    <div style="font-size:15px;font-weight:700;color:#FFD447;letter-spacing:0.5px;">
                        Butler Performance Analytics</div>
                </div>
                {_logo_html}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    customers_data = load_customers()
    distributors_data = load_distributors()
    all_locations = customers_data + distributors_data

    type_counts = {}
    for c in all_locations:
        t = c.get("type", "Unknown")
        type_counts[t] = type_counts.get(t, 0) + 1
    us_states = set()
    unique_countries = set()
    unique_counties = set()
    us_state_abbrs = {
        "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA",
        "KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ",
        "NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT",
        "VA","WA","WV","WI","WY","DC","PR","GU","VI","AS","MP",
    }
    for c in all_locations:
        st_val = c.get("state", "").strip()
        county = c.get("county", "").strip()
        country = c.get("country", "").strip()
        if st_val and st_val.upper() in us_state_abbrs:
            us_states.add(st_val.upper())
        if country:
            unique_countries.add(country if country != "US" else "United States")
        if county:
            unique_counties.add(county)
    installer_types = ["Promo Only (Not on C4C)", "C4C List", "Rack Installer"]
    installer_total = sum(type_counts.get(t, 0) for t in installer_types)

    import json as _json
    with open("codes_db.json") as _f:
        _db = _json.load(_f)
    rp_series_count = len(_db["rp_products"])
    rp_sku_count = sum(len(s["skus"]) for s in _db["rp_products"].values())
    comp_brand_count = len(_db["competitor_brands"])

    # ── Network at a Glance ──
    st.markdown(
        '<div style="font-size:0.72rem;text-transform:uppercase;letter-spacing:0.18em;color:#FFD447;font-weight:700;margin-bottom:6px;">Network at a Glance</div>',
        unsafe_allow_html=True,
    )

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Total Locations", f"{len(all_locations):,}")
    m2.metric("Installer Accounts", f"{installer_total:,}")
    m3.metric("Distributors", f"{type_counts.get('Distributor', 0):,}")
    m4.metric("States", len(us_states))
    m5.metric("Countries", len(unique_countries))
    m6.metric("Counties", f"{len(unique_counties):,}")

    st.markdown("")

    # ── Feature cards ──
    card_style = (
        "background:linear-gradient(180deg, rgba(28,20,48,0.92), rgba(16,10,27,0.98));"
        "border:1px solid rgba(157,107,255,0.18);border-radius:22px;"
        "padding:22px 20px 20px;height:100%;box-shadow:0 12px 30px rgba(0,0,0,0.24);"
    )
    icon_style = (
        "width:44px;height:44px;border-radius:12px;display:flex;"
        "align-items:center;justify-content:center;font-size:20px;margin-bottom:14px;"
    )

    col1, col2, col3 = st.columns(3, gap="medium")

    with col1:
        st.markdown(
            f"""
            <div style="{card_style}">
                <div style="{icon_style}background:rgba(111,45,189,0.2);color:#9D6BFF;">&#9889;</div>
                <div style="font-size:17px;font-weight:700;color:#F4F1FF;margin-bottom:6px;">
                    Report Generator</div>
                <div style="font-size:13px;color:#B7A8DB;line-height:1.6;margin-bottom:16px;">
                    Upload monthly Royal Purple Excel exports and generate fully branded PowerPoint
                    presentations with revenue analytics, Max-Clean attachment metrics, and per-store deep dives.</div>
                <div style="display:flex;gap:12px;flex-wrap:wrap;">
                    <span style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:999px;padding:4px 10px;font-size:11px;color:#F4F1FF;">Auto-parse</span>
                    <span style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:999px;padding:4px 10px;font-size:11px;color:#F4F1FF;">Deduplication</span>
                    <span style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:999px;padding:4px 10px;font-size:11px;color:#F4F1FF;">Branded PPTX</span>
                </div>
            </div>""", unsafe_allow_html=True)

    with col2:
        st.markdown(
            f"""
            <div style="{card_style}">
                <div style="{icon_style}background:rgba(37,99,235,0.15);color:#60a5fa;">&#127758;</div>
                <div style="font-size:17px;font-weight:700;color:#F4F1FF;margin-bottom:6px;">
                    Customer Map</div>
                <div style="font-size:13px;color:#B7A8DB;line-height:1.6;margin-bottom:16px;">
                    Interactive map of {len(all_locations):,} Royal Purple locations across {len(us_states)} states and {len(unique_countries)} countries.
                    Filter by account type, search by name, and export to branded Excel workbooks.</div>
                <div style="display:flex;gap:12px;flex-wrap:wrap;">
                    <span style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:999px;padding:4px 10px;font-size:11px;color:#F4F1FF;">{len(all_locations):,} pins</span>
                    <span style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:999px;padding:4px 10px;font-size:11px;color:#F4F1FF;">7 account types</span>
                    <span style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:999px;padding:4px 10px;font-size:11px;color:#F4F1FF;">Excel export</span>
                </div>
            </div>""", unsafe_allow_html=True)

    with col3:
        st.markdown(
            f"""
            <div style="{card_style}">
                <div style="{icon_style}background:rgba(255,212,71,0.12);color:#FFD447;">&#128218;</div>
                <div style="font-size:17px;font-weight:700;color:#F4F1FF;margin-bottom:6px;">
                    Product Reference</div>
                <div style="font-size:13px;color:#B7A8DB;line-height:1.6;margin-bottom:16px;">
                    Complete database of {rp_sku_count} Royal Purple products across {rp_series_count} product lines,
                    plus {comp_brand_count} competitor brands. Viscosities and cross-references.</div>
                <div style="display:flex;gap:12px;flex-wrap:wrap;">
                    <span style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:999px;padding:4px 10px;font-size:11px;color:#F4F1FF;">{rp_sku_count} RP products</span>
                    <span style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:999px;padding:4px 10px;font-size:11px;color:#F4F1FF;">{comp_brand_count} competitors</span>
                    <span style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:999px;padding:4px 10px;font-size:11px;color:#F4F1FF;">Admin editable</span>
                </div>
            </div>""", unsafe_allow_html=True)

    st.markdown("")

    # ── Account Type Breakdown ──
    st.markdown(
        '<div style="font-size:0.72rem;text-transform:uppercase;letter-spacing:0.18em;color:#FFD447;font-weight:700;margin-bottom:6px;">Account Type Breakdown</div>',
        unsafe_allow_html=True,
    )

    type_colors = {
        "Promo Only (Not on C4C)": "#FF5C7A",
        "C4C List": "#3DDC97",
        "Rack Installer": "#9D6BFF",
        "Distributor": "#FFD447",
        "Powersports/Motorsports": "#F5A623",
        "International": "#60a5fa",
        "Canada": "#3DDC97",
    }

    sorted_types = sorted(type_counts.items(), key=lambda x: -x[1])
    cols = st.columns(min(len(sorted_types), 4))
    for i, (ttype, count) in enumerate(sorted_types):
        color = type_colors.get(ttype, "#B7A8DB")
        with cols[i % 4]:
            st.markdown(
                f"""
                <div style="background:linear-gradient(180deg, rgba(28,20,48,0.94), rgba(15,10,25,0.98));
                            border:1px solid rgba(157,107,255,0.18);border-radius:20px;
                            padding:16px;margin-bottom:8px;">
                    <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
                        <div style="width:10px;height:10px;border-radius:50%;background:{color};"></div>
                        <span style="font-size:12px;color:#B7A8DB;font-weight:500;">{ttype}</span>
                    </div>
                    <div style="font-size:1.8rem;font-weight:800;color:#F4F1FF;line-height:1;">{count:,}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("")

    # ── RPO NAPA Accounts ──
    rpo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rpo_autocare_processed.json")
    if os.path.exists(rpo_path):
        with open(rpo_path) as _rpf:
            rpo_data = json.load(_rpf)

        if rpo_data:
            rpo_total = len(rpo_data)
            rpo_sales = sum(a.get("cytd_sales", 0) for a in rpo_data)
            rpo_c4c = sum(1 for a in rpo_data if a.get("c4c_status") == "On C4C")
            rpo_not_c4c = rpo_total - rpo_c4c
            rpo_gold = sum(1 for a in rpo_data if a.get("gold_flag"))

            st.markdown(
                '<div style="font-size:0.72rem;text-transform:uppercase;letter-spacing:0.18em;color:#FFD447;font-weight:700;margin-bottom:6px;">RPO NAPA Accounts</div>',
                unsafe_allow_html=True,
            )
            st.caption("RPO Autocare installer accounts with $1,000+ PYTD sales.")

            r1, r2, r3, r4, r5 = st.columns(5)
            r1.metric("RPO Accounts", f"{rpo_total:,}")
            r2.metric("PYTD Revenue", f"${rpo_sales:,.0f}")
            r3.metric("On C4C", f"{rpo_c4c:,}")
            r4.metric("Not on C4C", f"{rpo_not_c4c:,}")
            r5.metric("Gold Flag", f"{rpo_gold:,}")

    # ── Divider ──
    st.markdown('<div style="height:1px;background:linear-gradient(90deg,transparent,rgba(157,107,255,0.35),transparent);margin:14px 0 16px 0;"></div>', unsafe_allow_html=True)
    st.caption("Use the sidebar to navigate between pages.")

elif nav == "Customer Map":
    page_header("Customer Map", "Interactive map of Royal Purple customer locations across the United States.")
    st.markdown("")

    csv_upload = st.file_uploader(
        "Upload Customer CSV (optional)",
        type=["csv"],
        help="CSV with columns: store_name, address, city, state, zip, latitude, longitude, type",
    )

    customers = load_customers()
    distributors = load_distributors()

    if csv_upload is not None:
        try:
            csv_text = csv_upload.getvalue().decode("utf-8")
            uploaded_customers = parse_csv_customers(csv_text)
            if uploaded_customers:
                customers = uploaded_customers
                st.success(f"Loaded {len(uploaded_customers)} locations from CSV.")
            else:
                st.warning("No valid customer records found in CSV.")
        except Exception as e:
            st.error(f"Error reading CSV: {e}")

    all_map_data = customers + distributors

    if all_map_data:
        all_states = get_states(all_map_data)
        type_counts = {}
        for c in all_map_data:
            t = c.get("type", "Retail")
            type_counts[t] = type_counts.get(t, 0) + 1

        unique_counties = len(set(c.get("county", "") for c in all_map_data if c.get("county")))

        installer_types = ["Promo Only (Not on C4C)", "C4C List", "Rack Installer"]
        installer_total = sum(type_counts.get(t, 0) for t in installer_types)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Locations", len(all_map_data))
        col2.metric("Installer Accounts", installer_total)
        col3.metric("Distributors", type_counts.get("Distributor", 0))
        col4.metric("Powersports", type_counts.get("Powersports/Motorsports", 0))

        col5, col6, col7 = st.columns(3)
        col5.metric("Promo Only", type_counts.get("Promo Only (Not on C4C)", 0))
        col6.metric("C4C List", type_counts.get("C4C List", 0))
        col7.metric("Rack Installer", type_counts.get("Rack Installer", 0))
        st.markdown("")

        map_html = build_leaflet_html(all_map_data, height=650)
        components.html(map_html, height=660, scrolling=False)

        st.markdown("")
        st.caption("Use the search bar and filters on the map to find specific locations. Click the List button to see a sidebar of all locations.")

        st.markdown("---")
        exp_col1, exp_col2 = st.columns(2)

        with exp_col1:
            st.markdown("### Export Map Data")
            st.caption("Download a branded Excel workbook with per-state tabs, county breakdown, and filterable columns — ready to share with the Royal Purple team.")

            if st.button("Generate Map Data Export", type="primary", key="map_export"):
                with st.spinner("Building Excel workbook..."):
                    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                        export_path = tmp.name
                    stats = generate_map_export(export_path, all_map_data)

                    with open(export_path, "rb") as f:
                        export_data = f.read()
                    os.unlink(export_path)

                    st.success(
                        f"Export ready — {stats['sheets']} sheets: "
                        f"Dashboard + {stats['states']} state tabs + All Accounts + County Summary + Distributors | "
                        f"{stats['total']} locations across {stats['counties']} counties"
                    )

                    st.download_button(
                        label="Download Excel Workbook",
                        data=export_data,
                        file_name="RP_Installer_Account_Data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        with exp_col2:
            st.markdown("### Export C4C Report")
            st.caption("C4C (Connect for Calumet) gap analysis — identifies installer accounts not yet onboarded into Royal Purple's dealer system. Includes state breakdown, duplicates, and reconciliation.")

            if st.button("Generate C4C Report", type="primary", key="c4c_export"):
                with st.spinner("Building report..."):
                    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                        report_path = tmp.name
                    stats = generate_c4c_report(report_path)

                    with open(report_path, "rb") as f:
                        report_data = f.read()
                    os.unlink(report_path)

                    rpo_msg = ""
                    if stats.get("rpo_total"):
                        rpo_msg = f" + RPO Autocare: {stats['rpo_total']:,} accounts ({stats['rpo_not_c4c']:,} not on C4C)."
                    st.success(
                        f"Report generated — {stats['sheets']} sheets, {stats['total_accounts']} total accounts: "
                        f"{stats['not_on_c4c']} not on C4C, {stats['c4c_matched']} matched, "
                        f"{stats['distributors']} distributors, {stats['states']} states, {stats['counties']} counties."
                        f"{rpo_msg}"
                    )

                    st.download_button(
                        label="Download C4C Report",
                        data=report_data,
                        file_name="RP_C4C_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        st.markdown("---")
        st.markdown("### RPO Autocare — C4C Gap Analysis")
        st.caption("Cross-references 4,125 RPO Autocare 2025 installer accounts against C4C, Promo, and Rack lists to identify accounts not yet onboarded into C4C.")

        rpo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rpo_autocare_processed.json")
        if os.path.exists(rpo_path):
            import pandas as pd

            with open(rpo_path) as _rpf:
                rpo_data = json.load(_rpf)

            c4c_count = sum(1 for a in rpo_data if a['c4c_status'] == 'On C4C')
            promo_count = sum(1 for a in rpo_data if a['c4c_status'] == 'Promo Only')
            rack_count = sum(1 for a in rpo_data if a['c4c_status'] == 'Rack Only')
            not_in_count = sum(1 for a in rpo_data if a['c4c_status'] == 'Not in System')
            total_not_c4c = promo_count + rack_count + not_in_count
            total_sales_not_c4c = sum(a['cytd_sales'] for a in rpo_data if a['c4c_status'] != 'On C4C')

            rc1, rc2, rc3, rc4, rc5 = st.columns(5)
            rc1.metric("Total RPO Accounts", f"{len(rpo_data):,}")
            rc2.metric("On C4C", f"{c4c_count:,}")
            rc3.metric("Not on C4C", f"{total_not_c4c:,}")
            rc4.metric("C4C Rate", f"{c4c_count/len(rpo_data)*100:.1f}%")
            rc5.metric("Non-C4C Revenue", f"${total_sales_not_c4c:,.0f}")

            st.markdown("")
            rpo_filter = st.selectbox(
                "Filter by C4C Status",
                ["All Not on C4C", "All Accounts", "Promo Only", "Rack Only", "Not in System", "On C4C"],
                key="rpo_filter",
            )
            rpo_sort = st.selectbox(
                "Sort by",
                ["PYTD Sales (High to Low)", "PYTD Sales (Low to High)", "Name (A-Z)", "Name (Z-A)", "District", "Region"],
                key="rpo_sort",
            )

            filtered = rpo_data
            if rpo_filter == "All Not on C4C":
                filtered = [a for a in rpo_data if a['c4c_status'] != 'On C4C']
            elif rpo_filter == "On C4C":
                filtered = [a for a in rpo_data if a['c4c_status'] == 'On C4C']
            elif rpo_filter == "Promo Only":
                filtered = [a for a in rpo_data if a['c4c_status'] == 'Promo Only']
            elif rpo_filter == "Rack Only":
                filtered = [a for a in rpo_data if a['c4c_status'] == 'Rack Only']
            elif rpo_filter == "Not in System":
                filtered = [a for a in rpo_data if a['c4c_status'] == 'Not in System']

            if rpo_sort == "PYTD Sales (High to Low)":
                filtered.sort(key=lambda x: -x['cytd_sales'])
            elif rpo_sort == "PYTD Sales (Low to High)":
                filtered.sort(key=lambda x: x['cytd_sales'])
            elif rpo_sort == "Name (A-Z)":
                filtered.sort(key=lambda x: x['name'].upper())
            elif rpo_sort == "Name (Z-A)":
                filtered.sort(key=lambda x: x['name'].upper(), reverse=True)
            elif rpo_sort == "District":
                filtered.sort(key=lambda x: (x['district'], -x['cytd_sales']))
            elif rpo_sort == "Region":
                filtered.sort(key=lambda x: (x['region'], -x['cytd_sales']))

            st.markdown(f"**Showing {len(filtered):,} accounts**")

            df = pd.DataFrame([{
                "Installer Name": a['name'],
                "C4C Status": a['c4c_status'],
                "PYTD Sales": a['cytd_sales'],
                "Gold Flag": a['gold_flag'],
                "District": a['district'],
                "Region": a['region'],
                "Company Owned": a['company_owned'],
                "City": a['city'],
            } for a in filtered])

            if not df.empty:
                df["PYTD Sales"] = df["PYTD Sales"].apply(lambda x: f"${x:,.2f}")
                st.dataframe(df, use_container_width=True, height=500, hide_index=True)

                csv_export = df.to_csv(index=False)
                st.download_button(
                    label=f"Download {rpo_filter} ({len(filtered):,} accounts)",
                    data=csv_export,
                    file_name=f"RPO_Autocare_{rpo_filter.replace(' ', '_')}.csv",
                    mime="text/csv",
                    key="rpo_csv_download",
                )
        else:
            st.info("RPO Autocare data not yet processed. Upload the RPO Autocare Excel to cross-reference.")
    else:
        st.info("No customer data available. Upload a CSV file to get started.")

elif nav == "Product Reference":
    page_header("Product Reference", "Browse the full Royal Purple product catalog, identify competitor brands, and find the right RP alternative.")
    st.markdown("")
    product_reference.render()

elif nav == "Installer Incremental Profit Model":
    st.markdown(
        """
        <div style="background:linear-gradient(135deg,#2D1B5E 0%,#4B2D8A 60%,#6B3FA0 100%);
                    border-radius:12px;padding:32px 36px 28px;margin-bottom:8px;">
            <div style="font-size:11px;font-weight:700;letter-spacing:3px;color:#C4B5E8;
                        text-transform:uppercase;margin-bottom:8px;">Royal Purple CockPit</div>
            <div style="font-size:28px;font-weight:800;color:#FFFFFF;line-height:1.2;margin-bottom:8px;">
                Installer Installer Incremental Profit Model
            </div>
            <div style="font-size:14px;color:#C4B5E8;max-width:560px;line-height:1.6;">
                Compare Royal Purple profitability vs. your installer's current top-selling brand.
                See the incremental profit per service, per location, and total annual impact.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("")
    profit_calculator.render()

elif nav == "Admin":
    page_header("Database Editor", "Add, edit, or remove product codes and competitor brands. Changes save immediately.")
    st.markdown("")
    admin_panel.render()

elif nav == "Report Generator":
    st.markdown(
        """
        <div style="background:linear-gradient(135deg,#2D1B5E 0%,#4B2D8A 60%,#6B3FA0 100%);
                    border-radius:12px;padding:32px 36px 28px;margin-bottom:8px;">
            <div style="font-size:11px;font-weight:700;letter-spacing:3px;color:#C4B5E8;
                        text-transform:uppercase;margin-bottom:8px;">Royal Purple CockPit</div>
            <div style="font-size:28px;font-weight:800;color:#FFFFFF;line-height:1.2;margin-bottom:8px;">
                Installer Report Generator
            </div>
            <div style="font-size:14px;color:#C4B5E8;max-width:560px;line-height:1.6;">
                Upload your monthly Royal Purple Excel export to get a fully branded PowerPoint
                with network-level analytics, Max-Clean attachment metrics, and per-store deep dives.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("")

    upload_col, info_col = st.columns([3, 2], gap="large")
    with upload_col:
        uploaded_file = st.file_uploader(
            "Drop your Excel report here",
            type=["xlsx"],
            help="The app auto-detects columns, deduplicates multi-product invoices, and computes corrected revenue.",
            label_visibility="visible",
        )
    with info_col:
        st.markdown(
            """
            <div style="background:rgba(75,45,138,0.1);border-left:4px solid #4B2D8A;border-radius:0 8px 8px 0;
                        padding:16px 18px;margin-top:8px;">
                <div style="font-weight:700;color:#C4B5E8;font-size:13px;margin-bottom:8px;">
                    What this generates
                </div>
                <div style="font-size:12px;color:#8888a8;line-height:1.8;">
                    ✦ &nbsp;Network revenue &amp; invoice summary<br>
                    ✦ &nbsp;Max-Clean attachment analysis<br>
                    ✦ &nbsp;Per-store ranked deep dives<br>
                    ✦ &nbsp;Top product &amp; SKU breakdown<br>
                    ✦ &nbsp;Fully branded Royal Purple PPTX
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if uploaded_file is not None:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            stores, month_year = parse_excel(tmp_path)

            total_rev = sum(s["totalRevenue"] for s in stores)
            total_inv = sum(s["invoices"] for s in stores)
            avg_rev = total_rev / total_inv if total_inv else 0
            total_veh = sum(s["vehicles"] for s in stores)
            total_raw = sum(s.get("rawLineCount", 0) for s in stores)

            dedup_note = f"  ·  deduplicated from {fmt_number(total_raw)} raw lines" if total_raw > total_inv else ""
            st.markdown(
                f"""
                <div style="background:rgba(0,200,83,0.08);border:1px solid rgba(0,200,83,0.2);border-radius:8px;
                            padding:12px 18px;margin:12px 0 20px;">
                    <span style="color:#00c853;font-weight:700;">
                        {len(stores)} locations parsed &nbsp;·&nbsp; {month_year}
                    </span>
                    <span style="color:#66bb6a;font-size:13px;">{dedup_note}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            st.markdown(
                """
                <div style="font-size:10px;font-weight:700;letter-spacing:2.5px;color:#8888a8;
                            text-transform:uppercase;margin-bottom:4px;">Network Summary</div>
                """,
                unsafe_allow_html=True,
            )
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Revenue", fmt_currency(total_rev))
            col2.metric("Unique Invoices", fmt_number(total_inv))
            col3.metric("Avg Rev / Invoice", f"${avg_rev:.2f}")
            col4.metric("Unique Vehicles", fmt_number(total_veh))

            st.markdown("")

            network_mc = sum(s.get("maxClean", {}).get("total", 0) for s in stores)
            if network_mc > 0:
                mc_pct = network_mc / total_inv * 100 if total_inv else 0
                mc_rev = sum(
                    s.get("maxClean", {}).get("avgTicket", 0) * s.get("maxClean", {}).get("total", 0)
                    for s in stores
                )
                mc_avg = mc_rev / network_mc if network_mc else 0
                non_mc_count = total_inv - network_mc
                non_mc_rev = total_rev - mc_rev
                non_mc_avg = non_mc_rev / non_mc_count if non_mc_count else 0
                network_lift = mc_avg - non_mc_avg

                st.markdown(
                    """
                    <div style="border-top:2px solid #2a2a45;margin:8px 0 16px;">
                        <span style="display:inline-block;background:#4B2D8A;color:white;
                                     font-size:10px;font-weight:700;letter-spacing:2px;
                                     text-transform:uppercase;padding:3px 10px;border-radius:0 0 6px 6px;">
                            Max-Clean Attachment
                        </span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.caption(
                    "The RP export only shows Royal Purple products. 'Solo' Max-Clean lines represent "
                    "non-RP oil changes where Max-Clean was added as an upsell."
                )

                mc1, mc2, mc3, mc4 = st.columns(4)
                mc1.metric("MC Invoices", fmt_number(network_mc), f"{mc_pct:.1f}% attach rate")
                mc2.metric("MC Avg Ticket", f"${mc_avg:.2f}")
                mc3.metric("Non-MC Avg Ticket", f"${non_mc_avg:.2f}")
                mc4.metric("MC Ticket Lift", f"+${network_lift:.2f}", f"+{network_lift/non_mc_avg*100:.1f}%" if non_mc_avg else "")

                mc_with_rp = sum(s.get("maxClean", {}).get("withRpOil", 0) for s in stores)
                mc_non_rp = sum(s.get("maxClean", {}).get("withNonRpOil", 0) for s in stores)
                mc_solo = sum(s.get("maxClean", {}).get("soloInData", 0) for s in stores)
                st.markdown("")
                bk1, bk2, bk3 = st.columns(3)
                bk1.metric("MC + RP Oil", fmt_number(mc_with_rp), f"{mc_with_rp/network_mc*100:.1f}%" if network_mc else "")
                bk2.metric("MC + Non-RP Oil", fmt_number(mc_non_rp), f"{mc_non_rp/network_mc*100:.1f}%" if network_mc else "")
                bk3.metric("MC Solo (Non-RP OC)", fmt_number(mc_solo), f"{mc_solo/network_mc*100:.1f}%" if network_mc else "")

            st.markdown("")

            # ── New Code Detection ────────────────────────────────────────────
            detect_key = f"detected_codes_{uploaded_file.name}"
            dismiss_key = f"dismissed_codes_{uploaded_file.name}"

            if not st.session_state.get(dismiss_key, False):
                if detect_key not in st.session_state:
                    new_codes, _db_snap = detect_new_codes(stores)
                    st.session_state[detect_key] = new_codes
                    st.session_state[f"_db_snap_{uploaded_file.name}"] = _db_snap

                new_codes = st.session_state.get(detect_key, [])
                db_snap = st.session_state.get(f"_db_snap_{uploaded_file.name}", {})

                rp_items = [x for x in new_codes if x["classification"]["type"] == "rp"]
                comp_items = [x for x in new_codes if x["classification"]["type"] == "competitor"]
                unk_items = [x for x in new_codes if x["classification"]["type"] == "unknown"]
                auto_items = rp_items + comp_items

                if new_codes:
                    with st.expander(
                        f"**{len(new_codes)} new product code{'s' if len(new_codes) != 1 else ''} detected** — "
                        f"{len(rp_items)} RP, {len(comp_items)} competitor"
                        + (f", {len(unk_items)} unrecognized" if unk_items else ""),
                        expanded=True,
                    ):
                        st.caption(
                            "These codes appear in the report but are not yet in the Product Reference database. "
                            "Auto-classified codes can be added in one click."
                        )

                        if auto_items:
                            rows = []
                            for item in auto_items:
                                cl = item["classification"]
                                if cl["type"] == "rp":
                                    dest = cl.get("series", "RP")
                                    badge = "Royal Purple"
                                else:
                                    dest = cl.get("brand", cl.get("label", ""))
                                    badge = "Competitor"
                                rows.append({
                                    "Code": item["code"],
                                    "Classified As": badge,
                                    "Destination": dest,
                                    "In Reports": f"{item['store_count']} store{'s' if item['store_count'] != 1 else ''}",
                                    "Lines": item["line_count"],
                                })
                            st.dataframe(rows, use_container_width=True, hide_index=True)

                        if unk_items:
                            st.markdown("**Unrecognized codes** (no prefix match — add manually via Admin):")
                            unk_row = [{"Code": x["code"], "Lines": x["line_count"]} for x in unk_items]
                            st.dataframe(unk_row, use_container_width=True, hide_index=True)

                        btn_col, skip_col = st.columns([2, 1])
                        with btn_col:
                            if auto_items:
                                if st.button(
                                    f"Add {len(auto_items)} recognized code{'s' if len(auto_items) != 1 else ''} to database",
                                    type="primary",
                                    key=f"add_codes_{uploaded_file.name}",
                                ):
                                    added_rp, added_comp, _ = add_new_codes_to_db(auto_items, db_snap)
                                    try:
                                        from product_reference import load_codes_db
                                        load_codes_db.clear()
                                    except Exception:
                                        pass
                                    st.session_state[dismiss_key] = True
                                    st.success(
                                        f"Added {added_rp} RP SKU{'s' if added_rp != 1 else ''} "
                                        f"and {added_comp} competitor code{'s' if added_comp != 1 else ''} to the database."
                                    )
                                    st.rerun()
                        with skip_col:
                            if st.button("Dismiss", key=f"skip_codes_{uploaded_file.name}"):
                                st.session_state[dismiss_key] = True
                                st.rerun()

            # ─────────────────────────────────────────────────────────────────

            tab_rankings, tab_mc, tab_details = st.tabs(["Store Rankings", "Max-Clean by Store", "Store Details"])

            with tab_rankings:
                ranking_data = []
                for s in stores:
                    pct = s["totalRevenue"] / total_rev * 100 if total_rev else 0
                    mc = s.get("maxClean", {})
                    ranking_data.append({
                        "Rank": s["rank"],
                        "Store": s["name"],
                        "Revenue": fmt_currency(s["totalRevenue"]),
                        "Invoices": fmt_number(s["invoices"]),
                        "Avg Rev/Inv": f"${s['avgRevPerInvoice']:.2f}",
                        "Share": f"{pct:.1f}%",
                        "MC Rate": f"{mc.get('attachmentRate', 0):.0f}%",
                        "MC Lift": f"+${mc.get('ticketLift', 0):.2f}",
                    })
                st.dataframe(ranking_data, use_container_width=True, hide_index=True, key="ranking_df")

            with tab_mc:
                if network_mc > 0:
                    mc_data = []
                    for s in stores:
                        mc = s.get("maxClean", {})
                        if mc.get("total", 0) == 0:
                            continue
                        rp_pct = mc["withRpOil"] / mc["total"] * 100 if mc["total"] else 0
                        non_rp_pct = mc["withNonRpOil"] / mc["total"] * 100 if mc["total"] else 0
                        mc_data.append({
                            "Store": s["name"],
                            "MC Invoices": mc["total"],
                            "Attach Rate": f"{mc['attachmentRate']:.1f}%",
                            "With RP Oil": f"{mc['withRpOil']} ({rp_pct:.0f}%)",
                            "Non-RP Oil": f"{mc['withNonRpOil']} ({non_rp_pct:.0f}%)",
                            "MC Avg Ticket": f"${mc['avgTicket']:.2f}",
                            "Non-MC Avg": f"${mc['nonMcAvgTicket']:.2f}",
                            "Ticket Lift": f"+${mc['ticketLift']:.2f}",
                        })
                    st.dataframe(mc_data, use_container_width=True, hide_index=True, key="mc_df")

                    st.markdown("")
                    st.markdown("#### Key Insight")
                    best_lift = max(stores, key=lambda s: s.get("maxClean", {}).get("ticketLift", 0))
                    best_rate = max(stores, key=lambda s: s.get("maxClean", {}).get("attachmentRate", 0))
                    bl_mc = best_lift.get("maxClean", {})
                    br_mc = best_rate.get("maxClean", {})
                    st.info(
                        f"**{best_lift['name']}** has the highest ticket lift at "
                        f"+${bl_mc.get('ticketLift', 0):.2f} per Max-Clean invoice. "
                        f"**{best_rate['name']}** leads in attachment rate at "
                        f"{br_mc.get('attachmentRate', 0):.1f}%. "
                        f"Stores with high non-RP oil + Max-Clean rates are successfully "
                        f"selling the RP performance chemical upsell even on conventional oil changes."
                    )
                else:
                    st.info("No Max-Clean data detected in this report.")

            with tab_details:
                for s in stores:
                    with st.expander(f"#{s['rank']} — {s['name']}"):
                        dc1, dc2, dc3, dc4 = st.columns(4)
                        dc1.metric("Revenue", fmt_currency(s["totalRevenue"]))
                        dc2.metric("Invoices", fmt_number(s["invoices"]))
                        dc3.metric("Avg Rev/Inv", f"${s['avgRevPerInvoice']:.2f}")
                        mc = s.get("maxClean", {})
                        dc4.metric("MC Attach Rate", f"{mc.get('attachmentRate', 0):.1f}%")

                        if mc.get("total", 0) > 0:
                            st.caption(
                                f"Max-Clean: {mc['total']} invoices | "
                                f"With RP oil: {mc['withRpOil']} | "
                                f"Non-RP oil: {mc['withNonRpOil']} | "
                                f"Ticket lift: +${mc['ticketLift']:.2f}"
                            )

                        if s["productBreakdown"]:
                            st.caption("Top Products:")
                            for pb in s["productBreakdown"][:5]:
                                display_name = get_product_display_name(pb["code"])
                                line_ct = pb.get("lineCount", "")
                                ct_str = f" ({line_ct} lines)" if line_ct else ""
                                st.text(f"  {display_name} ({pb['category']}){ct_str}")

            st.markdown(
                f"""
                <div style="background:linear-gradient(135deg,#1E0F3C 0%,#2D1B5E 100%);
                            border-radius:10px;padding:24px 28px;margin:24px 0 8px;">
                    <div style="color:#C4B5E8;font-size:11px;font-weight:700;letter-spacing:2px;
                                text-transform:uppercase;margin-bottom:6px;">Ready to Export</div>
                    <div style="color:white;font-size:18px;font-weight:700;margin-bottom:4px;">
                        {len(stores)} stores &nbsp;·&nbsp; {month_year}
                    </div>
                    <div style="color:#8888a8;font-size:13px;">
                        Branded PowerPoint with network summary, Max-Clean analysis, and per-store slides
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button("Generate PowerPoint Report", type="primary", use_container_width=True):
                try:
                    with st.spinner("Generating branded presentation..."):
                        output_filename = f"Royal_Purple_Partnership_Report_{month_year.replace(' ', '_')}.pptx"
                        output_path = os.path.join(tempfile.gettempdir(), output_filename)
                        generate_report(tmp_path, output_path)

                        with open(output_path, "rb") as f:
                            pptx_data = f.read()

                        st.success(f"Report generated — {len(stores)} store deep dives included.")

                        st.download_button(
                            label="Download PowerPoint Report",
                            data=pptx_data,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="primary",
                            use_container_width=True,
                        )

                        os.unlink(output_path)
                except Exception as gen_err:
                    st.error(f"Error generating report: {gen_err}")

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
    else:
        st.info("Upload an Excel file to get started. The app auto-detects column layouts, deduplicates multi-product invoices, and computes corrected revenue figures.")
