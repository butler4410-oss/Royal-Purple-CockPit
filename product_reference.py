import streamlit as st

RP_PRODUCTS = {
    "RS Series — High Performance Synthetic": {
        "color": "#4B2D8A",
        "badge": "RS",
        "description": "Royal Purple's flagship full synthetic for high-performance and everyday driving. Exceeds API/ILSAC standards. Synerlec® additive technology delivers superior film strength and wear protection.",
        "application": "Modern engines, import vehicles, performance cars, daily drivers",
        "skus": [
            {"code": "RS0W20", "viscosity": "0W-20", "notes": "Honda, Toyota, Mazda — most common modern spec"},
            {"code": "RS5W20", "viscosity": "5W-20", "notes": "Ford, GM — wide-market workhorse"},
            {"code": "RS5W30", "viscosity": "5W-30", "notes": "Most common viscosity across all platforms"},
            {"code": "RS5W40", "viscosity": "5W-40", "notes": "European vehicles, Audi, BMW, Mercedes"},
            {"code": "RS0W16", "viscosity": "0W-16", "notes": "Late-model Honda/Toyota ultra-low viscosity spec"},
            {"code": "RS0W40", "viscosity": "0W-40", "notes": "European performance — Porsche, AMG, VW"},
        ]
    },
    "HMX Series — High Mileage Synthetic": {
        "color": "#7C3AED",
        "badge": "HMX",
        "description": "Engineered for engines over 75,000 miles. Enhanced seal conditioners revitalize worn gaskets, reduce oil consumption, and clean sludge deposits while maintaining full synthetic protection.",
        "application": "High-mileage vehicles (75K+ miles), vehicles with seeping gaskets or oil consumption issues",
        "skus": [
            {"code": "HMX0W20", "viscosity": "0W-20", "notes": "High-mileage Japanese/domestic, most popular HMX"},
            {"code": "HMX5W20", "viscosity": "5W-20", "notes": "High-mileage Ford/GM applications"},
            {"code": "HMX5W30", "viscosity": "5W-30", "notes": "Broadest high-mileage coverage"},
            {"code": "RMS5W20", "viscosity": "5W-20", "notes": "Royal Purple High Mileage Synthetic variant"},
            {"code": "RMS5W30", "viscosity": "5W-30", "notes": "Royal Purple High Mileage Synthetic variant"},
        ]
    },
    "Duralec Series — Diesel Synthetic": {
        "color": "#1D4ED8",
        "badge": "RSD",
        "description": "API CK-4 rated diesel engine oil for modern emission-controlled diesel engines. Compatible with DPF, EGR, and SCR after-treatment systems. Extended drain intervals.",
        "application": "Diesel trucks, work vans, fleet vehicles with DPF/EGR systems",
        "skus": [
            {"code": "RSD5W40",  "viscosity": "5W-40",  "notes": "Light-duty diesel — Cummins, Duramax, Power Stroke"},
            {"code": "RSD15W40", "viscosity": "15W-40", "notes": "Heavy-duty diesel, fleet and commercial applications"},
        ]
    },
    "RP Synthetic — Standard Full Synthetic": {
        "color": "#059669",
        "badge": "RP",
        "description": "Core Royal Purple synthetic line. Same Synerlec® technology as RS Series, positioned as the standard full synthetic offer across common viscosity grades.",
        "application": "General full synthetic upsell, vehicles not requiring specialty RS or HMX formulations",
        "skus": [
            {"code": "RP0W20", "viscosity": "0W-20", "notes": "Modern import spec"},
            {"code": "RP5W20", "viscosity": "5W-20", "notes": "Ford/GM domestic spec"},
            {"code": "RP5W30", "viscosity": "5W-30", "notes": "Universal synthetic coverage"},
            {"code": "RP5W40", "viscosity": "5W-40", "notes": "European Synthetic"},
            {"code": "RP0W16", "viscosity": "0W-16", "notes": "Ultra-low Honda/Toyota spec"},
            {"code": "RP0W40", "viscosity": "0W-40", "notes": "European performance"},
        ]
    },
    "Royal Purple Additives": {
        "color": "#D97706",
        "badge": "ADD",
        "description": "Royal Purple fuel and engine additives. Note: These are separate from oil revenue — they appear as line items on invoices alongside oil codes.",
        "application": "Upsell items — added to fuel or oil at time of service",
        "skus": [
            {"code": "18000", "viscosity": "Max-Atomizer", "notes": "Fuel injector cleaner — optimizes spray patterns"},
            {"code": "11755", "viscosity": "Max-Tane",     "notes": "Diesel fuel treatment — boosts cetane rating"},
            {"code": "11722", "viscosity": "Generic RP",   "notes": "Legacy catch-all RP code — no specific viscosity/formulation"},
        ]
    },
}

COMPETITOR_BRANDS = [
    {
        "brand": "CAM2",
        "type": "Synthetic & Conventional",
        "color": "#DC2626",
        "codes": [
            {"code": "S0W20", "product": "Full Synthetic 0W-20"},
            {"code": "S5W20", "product": "Full Synthetic 5W-20"},
            {"code": "S5W30", "product": "Full Synthetic 5W-30"},
            {"code": "5W20",  "product": "Conventional 5W-20"},
            {"code": "5W30",  "product": "Conventional 5W-30"},
            {"code": "5W40",  "product": "Conventional 5W-40"},
            {"code": "10W30", "product": "Conventional 10W-30"},
            {"code": "10W40", "product": "Conventional 10W-40"},
        ],
        "conversion_note": "S-prefix = Full Synthetic → easiest RP convert (RS same tier). Bare viscosity codes = conventional → upsell opportunity."
    },
    {
        "brand": "Valvoline",
        "type": "Full Synthetic, MaxLife & Conventional",
        "color": "#EA580C",
        "codes": [
            {"code": "VS0W20", "product": "Full Synthetic 0W-20"},
            {"code": "VS5W20", "product": "Full Synthetic 5W-20"},
            {"code": "VS5W30", "product": "Full Synthetic 5W-30"},
            {"code": "VS0W16", "product": "Full Synthetic 0W-16"},
            {"code": "VM5W20", "product": "MaxLife 5W-20"},
            {"code": "VM5W30", "product": "MaxLife 5W-30"},
            {"code": "VB5W20", "product": "Conventional 5W-20"},
            {"code": "VB5W30", "product": "Conventional 5W-30"},
            {"code": "VE5W30", "product": "Conventional 5W-30 (alt)"},
        ],
        "conversion_note": "VS → RS direct swap. VM (MaxLife) → HMX direct swap. VB/VE conventional → upsell to RS or HMX."
    },
    {
        "brand": "Mobil 1",
        "type": "Full Synthetic",
        "color": "#B91C1C",
        "codes": [
            {"code": "M0W20", "product": "Full Synthetic 0W-20"},
            {"code": "M5W20", "product": "Full Synthetic 5W-20"},
            {"code": "M5W30", "product": "Full Synthetic 5W-30"},
        ],
        "conversion_note": "Mobil 1 buyers are already premium synthetic customers — brand loyalty is stronger but RP performance story closes the gap."
    },
    {
        "brand": "Castrol",
        "type": "Full Synthetic (Edge)",
        "color": "#16A34A",
        "codes": [
            {"code": "CS5W20", "product": "Edge Full Synthetic 5W-20"},
            {"code": "CS5W30", "product": "Edge Full Synthetic 5W-30"},
            {"code": "CS0W20", "product": "Edge Full Synthetic 0W-20"},
        ],
        "conversion_note": "CS prefix = Castrol Edge. Same tier as RS — direct competitive swap."
    },
    {
        "brand": "Pennzoil",
        "type": "Platinum, Ultra, & Conventional",
        "color": "#CA8A04",
        "codes": [
            {"code": "PS5W20", "product": "Platinum Full Synthetic 5W-20"},
            {"code": "PS5W30", "product": "Platinum Full Synthetic 5W-30"},
            {"code": "PU0W20", "product": "Ultra Platinum 0W-20"},
            {"code": "PB5W20", "product": "Conventional 5W-20"},
            {"code": "PB5W30", "product": "Conventional 5W-30"},
        ],
        "conversion_note": "PS/PU = premium syn → RS direct. PB conventional → full upsell opportunity."
    },
]

SERVICE_TIERS = [
    {"code": "CK",  "name": "Basic Check", "description": "Quick check service — no oil product"},
    {"code": "S1",  "name": "Service Tier 1", "description": "Entry-level service package"},
    {"code": "S2",  "name": "Service Tier 2", "description": "Standard service package"},
    {"code": "S3",  "name": "Service Tier 3", "description": "Mid-tier service package"},
    {"code": "S4",  "name": "Service Tier 4", "description": "Premium service package"},
    {"code": "S5",  "name": "Service Tier 5", "description": "Full-service package"},
    {"code": "S6",  "name": "Service Tier 6", "description": "Top-tier service package"},
    {"code": "B7",  "name": "Basic Service 7", "description": "Base service tier"},
    {"code": "B8",  "name": "Basic Service 8", "description": "Base service tier"},
    {"code": "B9",  "name": "Basic Service 9", "description": "Base service tier"},
    {"code": "B10", "name": "Full Service (CAM2)", "description": "Full service using CAM2 conventional"},
]

SPEC_FLAGS = [
    {"code": "GF6",    "name": "ILSAC GF-6",  "description": "Latest ILSAC standard — fuel economy + protection"},
    {"code": "DEXOS1", "name": "GM Dexos 1",  "description": "GM factory-fill specification for most GM gas engines"},
    {"code": "AFC",    "name": "AFC",           "description": "Additive or formula certification flag"},
    {"code": "VMATF",  "name": "VMATF",         "description": "Variable motor additive/treatment flag"},
]

VISCOSITY_CROSSWALK = [
    {"current": "*0W20 (any brand)",  "rs": "RS0W20",  "hmx": "HMX0W20", "rsd": "—"},
    {"current": "*5W20 (any brand)",  "rs": "RS5W20",  "hmx": "HMX5W20", "rsd": "—"},
    {"current": "*5W30 (any brand)",  "rs": "RS5W30",  "hmx": "HMX5W30", "rsd": "—"},
    {"current": "*0W16 (any brand)",  "rs": "RS0W16",  "hmx": "—",       "rsd": "—"},
    {"current": "*5W40 (any brand)",  "rs": "RS5W40",  "hmx": "—",       "rsd": "RSD5W40"},
    {"current": "*0W40 (any brand)",  "rs": "RS0W40",  "hmx": "—",       "rsd": "—"},
    {"current": "*15W40 diesel",      "rs": "—",        "hmx": "—",       "rsd": "RSD15W40"},
]

CONVERSION_SEGMENTS = [
    {
        "segment": "Full Synthetic → RS Series",
        "codes": "S*, VS*, M*, CS*, PS*, PU*",
        "difficulty": "Low",
        "color": "#16A34A",
        "rationale": "Already paying premium synthetic prices. RP performance story is the only lift needed.",
        "suggested_sku": "RS (match viscosity)"
    },
    {
        "segment": "MaxLife / High Mileage → HMX",
        "codes": "VM*",
        "difficulty": "Low",
        "color": "#16A34A",
        "rationale": "Already in the high-mileage segment at a comparable price tier. Direct lateral move.",
        "suggested_sku": "HMX (match viscosity)"
    },
    {
        "segment": "Conventional → RS or HMX",
        "codes": "5W30, 5W20, VB*, VE*, PB*",
        "difficulty": "Medium",
        "color": "#D97706",
        "rationale": "Price jump required. Needs an upsell conversation about long-term engine protection value.",
        "suggested_sku": "RS or HMX depending on mileage"
    },
    {
        "segment": "Unknown / No Oil Code",
        "codes": "Ancillary only, or spec flags only",
        "difficulty": "Unknown",
        "color": "#64748B",
        "rationale": "Invoice contained no classifiable oil product code. May be a filter-only visit or data gap.",
        "suggested_sku": "N/A — investigate invoice"
    },
]


ALL_CODES = {}
for series_name, series in RP_PRODUCTS.items():
    for sku in series["skus"]:
        ALL_CODES[sku["code"].upper()] = {
            "brand": "Royal Purple",
            "series": series_name,
            "viscosity": sku["viscosity"],
            "notes": sku["notes"],
            "color": series["color"],
            "category": "rp",
        }
for brand_data in COMPETITOR_BRANDS:
    for sku in brand_data["codes"]:
        ALL_CODES[sku["code"].upper()] = {
            "brand": brand_data["brand"],
            "series": brand_data["type"],
            "viscosity": sku["product"],
            "notes": brand_data["conversion_note"],
            "color": brand_data["color"],
            "category": "competitor",
        }
for st_item in SERVICE_TIERS:
    ALL_CODES[st_item["code"].upper()] = {
        "brand": "Service Tier",
        "series": "Duke of Oil Service Package",
        "viscosity": st_item["name"],
        "notes": st_item["description"],
        "color": "#64748B",
        "category": "service_tier",
    }
for sf in SPEC_FLAGS:
    ALL_CODES[sf["code"].upper()] = {
        "brand": "Spec Flag",
        "series": "Industry Specification",
        "viscosity": sf["name"],
        "notes": sf["description"],
        "color": "#94A3B8",
        "category": "spec_flag",
    }


def render():
    tab1, tab2, tab3, tab4 = st.tabs([
        "RP Product Catalog",
        "Code Lookup",
        "Competitor Brands",
        "Conversion Guide",
    ])

    with tab1:
        _render_rp_catalog()

    with tab2:
        _render_code_lookup()

    with tab3:
        _render_competitor_brands()

    with tab4:
        _render_conversion_guide()


def _badge(text, bg_color, text_color="#FFFFFF", size=11):
    return (
        f'<span style="background:{bg_color};color:{text_color};padding:2px 9px;'
        f'border-radius:10px;font-size:{size}px;font-weight:700;'
        f'white-space:nowrap;display:inline-block;">{text}</span>'
    )


def _render_rp_catalog():
    st.markdown("### Royal Purple Product Catalog")
    st.caption("All known operation codes for Royal Purple products in the Duke of Oil POS system.")
    st.markdown("")

    for series_name, series in RP_PRODUCTS.items():
        badge_color = series["color"]
        badge_label = series["badge"]

        with st.expander(
            f"**{series_name}** — {len(series['skus'])} SKUs",
            expanded=True,
        ):
            col_info, col_skus = st.columns([2, 3])
            with col_info:
                st.markdown(
                    f'{_badge(badge_label, badge_color, size=12)}&nbsp;&nbsp;'
                    f'<span style="color:#4B2D8A;font-weight:600;font-size:14px;">{series_name}</span>',
                    unsafe_allow_html=True,
                )
                st.markdown(f'<p style="color:#475569;font-size:13px;margin-top:6px;">{series["description"]}</p>', unsafe_allow_html=True)
                st.markdown(f'<p style="color:#94A3B8;font-size:12px;"><strong>Best for:</strong> {series["application"]}</p>', unsafe_allow_html=True)

            with col_skus:
                for sku in series["skus"]:
                    cols = st.columns([1, 2, 3])
                    with cols[0]:
                        st.markdown(
                            f'<div style="background:{badge_color};color:white;padding:4px 8px;border-radius:6px;'
                            f'font-size:12px;font-weight:700;text-align:center;">{sku["code"]}</div>',
                            unsafe_allow_html=True,
                        )
                    with cols[1]:
                        st.markdown(
                            f'<div style="font-size:13px;font-weight:600;color:#1E293B;padding-top:4px;">{sku["viscosity"]}</div>',
                            unsafe_allow_html=True,
                        )
                    with cols[2]:
                        st.markdown(
                            f'<div style="font-size:12px;color:#64748B;padding-top:5px;">{sku["notes"]}</div>',
                            unsafe_allow_html=True,
                        )
            st.markdown("")


def _render_code_lookup():
    st.markdown("### Operation Code Lookup")
    st.caption("Enter any operation code from a Duke of Oil export to see its brand classification and recommended RP replacement.")
    st.markdown("")

    search = st.text_input("Code search", placeholder="e.g. RS5W30, VS0W20, HMX0W20, S5W30, GF6, B9...", label_visibility="collapsed")

    if search:
        code_upper = search.strip().upper()
        result = ALL_CODES.get(code_upper)

        if result:
            cat = result["category"]
            if cat == "rp":
                icon = "✅"
                label = "Royal Purple"
                label_color = "#16A34A"
            elif cat == "competitor":
                icon = "⚠️"
                label = "Competitor Oil"
                label_color = "#DC2626"
            elif cat == "service_tier":
                icon = "ℹ️"
                label = "Service Tier"
                label_color = "#64748B"
            else:
                icon = "ℹ️"
                label = "Spec Flag"
                label_color = "#94A3B8"

            st.markdown(
                f'<div style="background:white;border:2px solid {result["color"]};border-radius:10px;padding:16px 20px;">'
                f'<div style="display:flex;align-items:center;gap:12px;margin-bottom:10px;">'
                f'<span style="background:{result["color"]};color:white;padding:5px 14px;border-radius:8px;font-size:18px;font-weight:700;">{code_upper}</span>'
                f'<span style="font-size:14px;font-weight:700;color:{label_color};">{icon} {label}</span>'
                f'</div>'
                f'<table style="font-size:13px;color:#374151;border-collapse:collapse;width:100%">'
                f'<tr><td style="padding:4px 12px 4px 0;color:#94A3B8;font-weight:600;white-space:nowrap;">Brand</td><td style="padding:4px 0;">{result["brand"]}</td></tr>'
                f'<tr><td style="padding:4px 12px 4px 0;color:#94A3B8;font-weight:600;white-space:nowrap;">Series / Type</td><td style="padding:4px 0;">{result["series"]}</td></tr>'
                f'<tr><td style="padding:4px 12px 4px 0;color:#94A3B8;font-weight:600;white-space:nowrap;">Product</td><td style="padding:4px 0;">{result["viscosity"]}</td></tr>'
                f'<tr><td style="padding:4px 12px 4px 0;color:#94A3B8;font-weight:600;white-space:nowrap;">Notes</td><td style="padding:4px 0;">{result["notes"]}</td></tr>'
                f'</table>'
                f'</div>',
                unsafe_allow_html=True,
            )

            if cat == "competitor":
                st.markdown("")
                st.markdown("**Recommended Royal Purple Replacement:**")
                viscosity = result["viscosity"]
                for v_str in ["0W20", "5W20", "5W30", "5W40", "0W40", "0W16", "15W40"]:
                    if v_str in viscosity.replace("-", "").replace(" ", ""):
                        vw = v_str[:2] + "-" + v_str[2:]
                        st.info(
                            f"**Standard:** RS{v_str}  |  **High Mileage (75K+ mi):** HMX{v_str}"
                            if not v_str.startswith("15") else
                            f"**Diesel:** RSD{v_str}"
                        )
                        break

        else:
            _try_prefix_lookup(code_upper)
    else:
        st.markdown("")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(_badge("RS5W30", "#4B2D8A", size=12), unsafe_allow_html=True)
            st.caption("RP High Performance 5W-30")
        with col2:
            st.markdown(_badge("HMX0W20", "#7C3AED", size=12), unsafe_allow_html=True)
            st.caption("RP High Mileage 0W-20")
        with col3:
            st.markdown(_badge("VS5W30", "#EA580C", size=12), unsafe_allow_html=True)
            st.caption("Valvoline Full Syn 5W-30")
        with col4:
            st.markdown(_badge("S5W30", "#DC2626", size=12), unsafe_allow_html=True)
            st.caption("CAM2 Synthetic 5W-30")


def _try_prefix_lookup(code):
    RP_PREFIXES = [
        ("RS",  "Royal Purple", "RS Series — High Performance Synthetic", "#4B2D8A"),
        ("HMX", "Royal Purple", "HMX Series — High Mileage Synthetic",    "#7C3AED"),
        ("RMS", "Royal Purple", "HMX Series — High Mileage Synthetic",    "#7C3AED"),
        ("RSD", "Royal Purple", "Duralec — Diesel Synthetic",              "#1D4ED8"),
        ("RP",  "Royal Purple", "RP Synthetic",                            "#059669"),
    ]
    COMP_PREFIXES = [
        ("S0W",  "CAM2",      "Full Synthetic",   "#DC2626"),
        ("S5W",  "CAM2",      "Full Synthetic",   "#DC2626"),
        ("VS",   "Valvoline", "Full Synthetic",   "#EA580C"),
        ("VM",   "Valvoline", "MaxLife",          "#EA580C"),
        ("VB",   "Valvoline", "Conventional",     "#EA580C"),
        ("VE",   "Valvoline", "Conventional",     "#EA580C"),
        ("M0W",  "Mobil 1",   "Full Synthetic",   "#B91C1C"),
        ("M5W",  "Mobil 1",   "Full Synthetic",   "#B91C1C"),
        ("CS",   "Castrol",   "Edge Synthetic",   "#16A34A"),
        ("PS",   "Pennzoil",  "Platinum Syn",     "#CA8A04"),
        ("PU",   "Pennzoil",  "Ultra Platinum",   "#CA8A04"),
        ("PB",   "Pennzoil",  "Conventional",     "#CA8A04"),
    ]
    for prefix, brand, series, color in RP_PREFIXES:
        if code.startswith(prefix) and any(c.isdigit() for c in code):
            st.markdown(
                f'<div style="background:white;border:2px solid {color};border-radius:10px;padding:16px 20px;">'
                f'<div style="font-weight:700;color:{color};font-size:15px;margin-bottom:8px;">✅ Likely Royal Purple — {series}</div>'
                f'<p style="font-size:13px;color:#475569;">Code <strong>{code}</strong> matches the <strong>{prefix}*</strong> prefix pattern for <strong>{series}</strong>. '
                f'Not in the known code list, but matches RP prefix rules — treat as Royal Purple for revenue calculations.</p>'
                f'</div>',
                unsafe_allow_html=True,
            )
            return
    for prefix, brand, series, color in COMP_PREFIXES:
        if code.startswith(prefix):
            st.markdown(
                f'<div style="background:white;border:2px solid {color};border-radius:10px;padding:16px 20px;">'
                f'<div style="font-weight:700;color:{color};font-size:15px;margin-bottom:8px;">⚠️ Likely Competitor — {brand} {series}</div>'
                f'<p style="font-size:13px;color:#475569;">Code <strong>{code}</strong> matches the <strong>{prefix}*</strong> prefix pattern for <strong>{brand} {series}</strong>. '
                f'Not in the known code table, but matches competitor prefix rules.</p>'
                f'</div>',
                unsafe_allow_html=True,
            )
            return
    st.warning(f'**"{code}"** is not in the known code list and doesn\'t match any known brand prefix. It may be an ancillary item (filter, wiper, air freshener, service charge), a spec flag, or a new/unlisted code. Check the Duke of Oil POS documentation.')


def _render_competitor_brands():
    st.markdown("### Competitor Brand Reference")
    st.caption("All known competitor oil codes in the Duke of Oil POS system, grouped by brand.")
    st.markdown("")

    for brand_data in COMPETITOR_BRANDS:
        color = brand_data["color"]
        with st.expander(
            f"**{brand_data['brand']}** — {brand_data['type']} — {len(brand_data['codes'])} known codes",
        ):
            st.markdown(
                f'<div style="background:{color}11;border-left:4px solid {color};padding:10px 14px;border-radius:0 8px 8px 0;margin-bottom:12px;">'
                f'<span style="font-size:12px;color:{color};font-weight:600;">Conversion Strategy:</span>'
                f'<span style="font-size:13px;color:#374151;"> {brand_data["conversion_note"]}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
            cols = st.columns(3)
            for i, sku in enumerate(brand_data["codes"]):
                with cols[i % 3]:
                    st.markdown(
                        f'<div style="border:1px solid #E2E8F0;border-radius:6px;padding:8px 10px;margin-bottom:6px;">'
                        f'<div style="font-weight:700;font-size:13px;color:{color};">{sku["code"]}</div>'
                        f'<div style="font-size:12px;color:#64748B;">{sku["product"]}</div>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
        st.markdown("")

    st.markdown("---")
    st.markdown("#### Service Tiers & Spec Flags")
    st.caption("These codes appear on invoices alongside oil codes. They do not represent oil products and should not be used for brand classification.")

    col_tier, col_spec = st.columns(2)
    with col_tier:
        st.markdown("**Service Tier Codes**")
        for item in SERVICE_TIERS:
            st.markdown(
                f'<div style="border:1px solid #E2E8F0;border-radius:6px;padding:6px 10px;margin-bottom:4px;">'
                f'<span style="font-weight:700;color:#64748B;font-size:13px;">{item["code"]}</span>'
                f' <span style="font-size:12px;color:#94A3B8;"> — {item["name"]}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
    with col_spec:
        st.markdown("**Spec Flags**")
        for item in SPEC_FLAGS:
            st.markdown(
                f'<div style="border:1px solid #E2E8F0;border-radius:6px;padding:6px 10px;margin-bottom:4px;">'
                f'<span style="font-weight:700;color:#94A3B8;font-size:13px;">{item["code"]}</span>'
                f' <span style="font-size:12px;color:#94A3B8;"> — {item["name"]}: {item["description"]}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )


def _render_conversion_guide():
    st.markdown("### Conversion Guide")
    st.caption("How to identify and target each conversion segment when analyzing a full-code Duke of Oil export.")
    st.markdown("")

    st.markdown("#### Viscosity Crosswalk")
    st.caption("The correct Royal Purple SKU for every viscosity grade a competitor customer might be using.")

    header_cols = st.columns([3, 2, 2, 2])
    with header_cols[0]:
        st.markdown('<div style="font-size:12px;font-weight:700;color:#94A3B8;padding-bottom:4px;border-bottom:2px solid #E2E8F0;">CUSTOMER\'S CURRENT OIL</div>', unsafe_allow_html=True)
    with header_cols[1]:
        st.markdown('<div style="font-size:12px;font-weight:700;color:#4B2D8A;padding-bottom:4px;border-bottom:2px solid #4B2D8A;">→ RS Series</div>', unsafe_allow_html=True)
    with header_cols[2]:
        st.markdown('<div style="font-size:12px;font-weight:700;color:#7C3AED;padding-bottom:4px;border-bottom:2px solid #7C3AED;">→ HMX (High Mileage)</div>', unsafe_allow_html=True)
    with header_cols[3]:
        st.markdown('<div style="font-size:12px;font-weight:700;color:#1D4ED8;padding-bottom:4px;border-bottom:2px solid #1D4ED8;">→ Duralec (Diesel)</div>', unsafe_allow_html=True)

    for row in VISCOSITY_CROSSWALK:
        cols = st.columns([3, 2, 2, 2])
        with cols[0]:
            st.markdown(f'<div style="padding:8px 0;font-size:13px;color:#374151;border-bottom:1px solid #F1F5F9;">{row["current"]}</div>', unsafe_allow_html=True)
        with cols[1]:
            color = "#4B2D8A" if row["rs"] != "—" else "#CBD5E1"
            st.markdown(f'<div style="padding:8px 0;font-size:13px;font-weight:600;color:{color};border-bottom:1px solid #F1F5F9;">{row["rs"]}</div>', unsafe_allow_html=True)
        with cols[2]:
            color = "#7C3AED" if row["hmx"] != "—" else "#CBD5E1"
            st.markdown(f'<div style="padding:8px 0;font-size:13px;font-weight:600;color:{color};border-bottom:1px solid #F1F5F9;">{row["hmx"]}</div>', unsafe_allow_html=True)
        with cols[3]:
            color = "#1D4ED8" if row["rsd"] != "—" else "#CBD5E1"
            st.markdown(f'<div style="padding:8px 0;font-size:13px;font-weight:600;color:{color};border-bottom:1px solid #F1F5F9;">{row["rsd"]}</div>', unsafe_allow_html=True)

    st.markdown("")
    st.markdown("---")
    st.markdown("#### Conversion Segments")
    st.caption("When running the RP classifier against a full-code export, customers fall into these segments.")
    st.markdown("")

    for seg in CONVERSION_SEGMENTS:
        color = seg["color"]
        st.markdown(
            f'<div style="border:1.5px solid {color};border-radius:10px;padding:14px 18px;margin-bottom:12px;">'
            f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">'
            f'<span style="font-size:15px;font-weight:700;color:{color};">{seg["segment"]}</span>'
            f'<span style="background:{color};color:white;padding:2px 12px;border-radius:10px;font-size:12px;font-weight:700;">Difficulty: {seg["difficulty"]}</span>'
            f'</div>'
            f'<div style="font-size:12px;color:#64748B;margin-bottom:6px;"><strong>Codes:</strong> {seg["codes"]}</div>'
            f'<div style="font-size:13px;color:#374151;margin-bottom:6px;">{seg["rationale"]}</div>'
            f'<div style="font-size:12px;color:{color};font-weight:600;">Suggested RP: {seg["suggested_sku"]}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.markdown("")
    st.markdown("---")
    st.markdown("#### How Invoice-Level Classification Works")
    st.info(
        "**Each invoice generates multiple rows** in a full-code export — one per operation code. "
        "All rows for the same Invoice # share the same revenue total.\n\n"
        "**Step 1:** Group all rows by Invoice #\n\n"
        "**Step 2:** Find the oil product code on that invoice (ignore spec flags like GF6/DEXOS1, service tiers like S1–S6, and ancillary items like OF*, AF*, FB)\n\n"
        "**Step 3:** Classify the oil code as Royal Purple or a specific competitor brand using the prefix rules above\n\n"
        "**Step 4:** Assign a conversion segment and calculate RP ticket premium vs. competitor avg ticket"
    )
