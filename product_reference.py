import streamlit as st
import json
import os

CODES_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "codes_db.json")

_DEFAULT_DB = {
    "rp_products": {},
    "competitor_brands": [],
    "service_tiers": [],
    "spec_flags": [],
    "viscosity_crosswalk": [],
    "conversion_segments": [],
}


@st.cache_data(show_spinner=False)
def load_codes_db():
    if os.path.exists(CODES_DB_PATH):
        with open(CODES_DB_PATH) as f:
            return json.load(f)
    return _DEFAULT_DB


def save_codes_db(db: dict):
    with open(CODES_DB_PATH, "w") as f:
        json.dump(db, f, indent=2)
    load_codes_db.clear()


def _build_lookup(db):
    lookup = {}
    for series_name, series in db.get("rp_products", {}).items():
        for sku in series.get("skus", []):
            lookup[sku["code"].upper()] = {
                "brand": "Royal Purple",
                "series": series_name,
                "viscosity": sku["viscosity"],
                "notes": sku.get("notes", ""),
                "color": series.get("color", "#4B2D8A"),
                "category": "rp",
            }
    for brand_data in db.get("competitor_brands", []):
        for sku in brand_data.get("codes", []):
            lookup[sku["code"].upper()] = {
                "brand": brand_data["brand"],
                "series": brand_data.get("type", ""),
                "viscosity": sku.get("product", sku.get("name", sku["code"])),
                "notes": brand_data.get("conversion_note", ""),
                "color": brand_data.get("color", "#DC2626"),
                "category": "competitor",
            }
    for item in db.get("service_tiers", []):
        lookup[item["code"].upper()] = {
            "brand": "Service Tier",
            "series": "Duke of Oil Service Package",
            "viscosity": item["name"],
            "notes": item["description"],
            "color": "#64748B",
            "category": "service_tier",
        }
    for item in db.get("spec_flags", []):
        lookup[item["code"].upper()] = {
            "brand": "Spec Flag",
            "series": "Industry Specification",
            "viscosity": item["name"],
            "notes": item["description"],
            "color": "#94A3B8",
            "category": "spec_flag",
        }
    return lookup


def render():
    db = load_codes_db()
    all_codes = _build_lookup(db)

    tab1, tab2, tab3, tab4 = st.tabs([
        "RP Product Catalog",
        "Code Lookup",
        "Competitor Brands",
        "Conversion Guide",
    ])

    with tab1:
        _render_rp_catalog(db)
    with tab2:
        _render_code_lookup(db, all_codes)
    with tab3:
        _render_competitor_brands(db)
    with tab4:
        _render_conversion_guide(db)


def _badge(text, bg_color, text_color="#FFFFFF", size=11):
    return (
        f'<span style="background:{bg_color};color:{text_color};padding:2px 9px;'
        f'border-radius:10px;font-size:{size}px;font-weight:700;'
        f'white-space:nowrap;display:inline-block;">{text}</span>'
    )


def _render_rp_catalog(db):
    rp_products = db.get("rp_products", {})
    st.markdown("### Royal Purple Product Catalog")
    st.caption("All known operation codes for Royal Purple products in the Duke of Oil POS system.")
    st.markdown("")

    if not rp_products:
        st.info("No RP products defined. Add them in the Admin panel.")
        return

    for series_name, series in rp_products.items():
        badge_color = series.get("color", "#4B2D8A")
        badge_label = series.get("badge", "RP")
        skus = series.get("skus", [])

        with st.expander(f"**{series_name}** — {len(skus)} SKU{'s' if len(skus) != 1 else ''}", expanded=True):
            col_info, col_skus = st.columns([2, 3])
            with col_info:
                st.markdown(
                    f'{_badge(badge_label, badge_color, size=12)}&nbsp;&nbsp;'
                    f'<span style="color:#4B2D8A;font-weight:600;font-size:14px;">{series_name}</span>',
                    unsafe_allow_html=True,
                )
                st.markdown(
                    f'<p style="color:#475569;font-size:13px;margin-top:6px;">{series.get("description","")}</p>',
                    unsafe_allow_html=True,
                )
                st.markdown(
                    f'<p style="color:#94A3B8;font-size:12px;"><strong>Best for:</strong> {series.get("application","")}</p>',
                    unsafe_allow_html=True,
                )
            with col_skus:
                for sku in skus:
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
                            f'<div style="font-size:12px;color:#64748B;padding-top:5px;">{sku.get("notes","")}</div>',
                            unsafe_allow_html=True,
                        )
            st.markdown("")


def _render_code_lookup(db, all_codes):
    st.markdown("### Operation Code Lookup")
    st.caption("Enter any operation code from a Duke of Oil export to see its brand classification and recommended RP replacement.")
    st.markdown("")

    search = st.text_input(
        "Code search",
        placeholder="e.g. RS5W30, VS0W20, HMX0W20, S5W30, GF6, B9...",
        label_visibility="collapsed",
    )

    if search:
        code_upper = search.strip().upper()
        result = all_codes.get(code_upper)

        if result:
            cat = result["category"]
            if cat == "rp":
                icon, label, label_color = "✅", "Royal Purple", "#16A34A"
            elif cat == "competitor":
                icon, label, label_color = "⚠️", "Competitor Oil", "#DC2626"
            elif cat == "service_tier":
                icon, label, label_color = "ℹ️", "Service Tier", "#64748B"
            else:
                icon, label, label_color = "ℹ️", "Spec Flag", "#94A3B8"

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
        rp_products = db.get("rp_products", {})
        competitors = db.get("competitor_brands", [])
        sample_rp = []
        for series in rp_products.values():
            for sku in series.get("skus", [])[:1]:
                sample_rp.append((sku["code"], series.get("color", "#4B2D8A"), sku["viscosity"]))
                if len(sample_rp) >= 2:
                    break
            if len(sample_rp) >= 2:
                break
        sample_comp = []
        for brand in competitors[:2]:
            for sku in brand.get("codes", [])[:1]:
                sample_comp.append((sku["code"], brand.get("color", "#DC2626"), sku.get("product", sku.get("name", sku["code"]))))

        examples = sample_rp + sample_comp
        if examples:
            cols = st.columns(min(len(examples), 4))
            for i, (code, color, label) in enumerate(examples[:4]):
                with cols[i]:
                    st.markdown(_badge(code, color, size=12), unsafe_allow_html=True)
                    st.caption(label)


def _try_prefix_lookup(code):
    RP_PREFIXES = [
        ("RS",  "Royal Purple", "RS Series — High Performance Synthetic", "#4B2D8A"),
        ("HMX", "Royal Purple", "HMX Series — High Mileage Synthetic",    "#7C3AED"),
        ("RMS", "Royal Purple", "HMX Series — High Mileage Synthetic",    "#7C3AED"),
        ("RSD", "Royal Purple", "Duralec — Diesel Synthetic",              "#1D4ED8"),
        ("RP",  "Royal Purple", "RP Synthetic",                            "#059669"),
    ]
    COMP_PREFIXES = [
        ("S0W", "CAM2",      "Full Synthetic",  "#DC2626"),
        ("S5W", "CAM2",      "Full Synthetic",  "#DC2626"),
        ("VS",  "Valvoline", "Full Synthetic",  "#EA580C"),
        ("VM",  "Valvoline", "MaxLife",         "#EA580C"),
        ("VB",  "Valvoline", "Conventional",    "#EA580C"),
        ("VE",  "Valvoline", "Conventional",    "#EA580C"),
        ("M0W", "Mobil 1",   "Full Synthetic",  "#B91C1C"),
        ("M5W", "Mobil 1",   "Full Synthetic",  "#B91C1C"),
        ("CS",  "Castrol",   "Edge Synthetic",  "#16A34A"),
        ("PS",  "Pennzoil",  "Platinum Syn",    "#CA8A04"),
        ("PU",  "Pennzoil",  "Ultra Platinum",  "#CA8A04"),
        ("PB",  "Pennzoil",  "Conventional",    "#CA8A04"),
    ]
    for prefix, brand, series, color in RP_PREFIXES:
        if code.startswith(prefix) and any(c.isdigit() for c in code):
            st.markdown(
                f'<div style="background:white;border:2px solid {color};border-radius:10px;padding:16px 20px;">'
                f'<div style="font-weight:700;color:{color};font-size:15px;margin-bottom:8px;">✅ Likely Royal Purple — {series}</div>'
                f'<p style="font-size:13px;color:#475569;">Code <strong>{code}</strong> matches the <strong>{prefix}*</strong> prefix pattern. '
                f'Not in the known code list — add it in the Admin panel if confirmed.</p>'
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
                f'Not in the known code table — add it in the Admin panel if confirmed.</p>'
                f'</div>',
                unsafe_allow_html=True,
            )
            return
    st.warning(
        f'**"{code}"** is not in the known code list and doesn\'t match any known brand prefix. '
        f'It may be an ancillary item (filter, wiper, air freshener), a spec flag, or a new code. '
        f'Add it in the Admin panel if needed.'
    )


def _render_competitor_brands(db):
    competitor_brands = db.get("competitor_brands", [])
    service_tiers = db.get("service_tiers", [])
    spec_flags = db.get("spec_flags", [])

    st.markdown("### Competitor Brand Reference")
    st.caption("All known competitor oil codes in the Duke of Oil POS system, grouped by brand.")
    st.markdown("")

    if not competitor_brands:
        st.info("No competitor brands defined. Add them in the Admin panel.")
    else:
        for brand_data in competitor_brands:
            color = brand_data.get("color", "#DC2626")
            codes = brand_data.get("codes", [])
            with st.expander(f"**{brand_data['brand']}** — {brand_data.get('type','')} — {len(codes)} known codes"):
                note = brand_data.get("conversion_note", "")
                if note:
                    st.markdown(
                        f'<div style="background:{color}11;border-left:4px solid {color};padding:10px 14px;border-radius:0 8px 8px 0;margin-bottom:12px;">'
                        f'<span style="font-size:12px;color:{color};font-weight:600;">Conversion Strategy:</span>'
                        f'<span style="font-size:13px;color:#374151;"> {note}</span>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                if codes:
                    cols = st.columns(3)
                    for i, sku in enumerate(codes):
                        with cols[i % 3]:
                            st.markdown(
                                f'<div style="border:1px solid #E2E8F0;border-radius:6px;padding:8px 10px;margin-bottom:6px;">'
                                f'<div style="font-weight:700;font-size:13px;color:{color};">{sku["code"]}</div>'
                                f'<div style="font-size:12px;color:#64748B;">{sku.get("product", sku.get("name", sku["code"]))}</div>'
                                f'</div>',
                                unsafe_allow_html=True,
                            )
            st.markdown("")

    st.markdown("---")
    st.markdown("#### Service Tiers & Spec Flags")
    st.caption("These codes appear on invoices alongside oil codes but do not represent oil products.")

    col_tier, col_spec = st.columns(2)
    with col_tier:
        st.markdown("**Service Tier Codes**")
        for item in service_tiers:
            st.markdown(
                f'<div style="border:1px solid #E2E8F0;border-radius:6px;padding:6px 10px;margin-bottom:4px;">'
                f'<span style="font-weight:700;color:#64748B;font-size:13px;">{item["code"]}</span>'
                f' <span style="font-size:12px;color:#94A3B8;"> — {item.get("name","")}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
    with col_spec:
        st.markdown("**Spec Flags**")
        for item in spec_flags:
            st.markdown(
                f'<div style="border:1px solid #E2E8F0;border-radius:6px;padding:6px 10px;margin-bottom:4px;">'
                f'<span style="font-weight:700;color:#94A3B8;font-size:13px;">{item["code"]}</span>'
                f' <span style="font-size:12px;color:#94A3B8;"> — {item.get("name","")}: {item.get("description","")}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )


def _render_conversion_guide(db):
    crosswalk = db.get("viscosity_crosswalk", [])
    segments = db.get("conversion_segments", [])

    st.markdown("### Conversion Guide")
    st.caption("How to identify and target each conversion segment when analyzing a full-code Duke of Oil export.")
    st.markdown("")

    st.markdown("#### Viscosity Crosswalk")
    st.caption("The correct Royal Purple SKU for every viscosity grade a competitor customer might be using.")

    if crosswalk:
        header_cols = st.columns([3, 2, 2, 2])
        labels = [("CUSTOMER'S CURRENT OIL", "#94A3B8", "#E2E8F0"),
                  ("→ RS Series", "#4B2D8A", "#4B2D8A"),
                  ("→ HMX (High Mileage)", "#7C3AED", "#7C3AED"),
                  ("→ Duralec (Diesel)", "#1D4ED8", "#1D4ED8")]
        for col, (text, color, border) in zip(header_cols, labels):
            with col:
                st.markdown(
                    f'<div style="font-size:12px;font-weight:700;color:{color};padding-bottom:4px;border-bottom:2px solid {border};">{text}</div>',
                    unsafe_allow_html=True,
                )
        for row in crosswalk:
            cols = st.columns([3, 2, 2, 2])
            values = [
                (row.get("current", ""), "#374151"),
                (row.get("rs", "—"), "#4B2D8A" if row.get("rs", "—") != "—" else "#CBD5E1"),
                (row.get("hmx", "—"), "#7C3AED" if row.get("hmx", "—") != "—" else "#CBD5E1"),
                (row.get("rsd", "—"), "#1D4ED8" if row.get("rsd", "—") != "—" else "#CBD5E1"),
            ]
            for col, (val, color) in zip(cols, values):
                with col:
                    st.markdown(
                        f'<div style="padding:8px 0;font-size:13px;font-weight:{"600" if color != "#374151" else "400"};color:{color};border-bottom:1px solid #F1F5F9;">{val}</div>',
                        unsafe_allow_html=True,
                    )

    st.markdown("")
    st.markdown("---")
    st.markdown("#### Conversion Segments")
    st.caption("When running the RP classifier against a full-code export, customers fall into these segments.")
    st.markdown("")

    for seg in segments:
        color = seg.get("color", "#64748B")
        st.markdown(
            f'<div style="border:1.5px solid {color};border-radius:10px;padding:14px 18px;margin-bottom:12px;">'
            f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">'
            f'<span style="font-size:15px;font-weight:700;color:{color};">{seg.get("segment","")}</span>'
            f'<span style="background:{color};color:white;padding:2px 12px;border-radius:10px;font-size:12px;font-weight:700;">Difficulty: {seg.get("difficulty","")}</span>'
            f'</div>'
            f'<div style="font-size:12px;color:#64748B;margin-bottom:6px;"><strong>Codes:</strong> {seg.get("codes","")}</div>'
            f'<div style="font-size:13px;color:#374151;margin-bottom:6px;">{seg.get("rationale","")}</div>'
            f'<div style="font-size:12px;color:{color};font-weight:600;">Suggested RP: {seg.get("suggested_sku","")}</div>'
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
