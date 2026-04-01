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
            "series": "Installer Service Package",
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
    rp_products = db.get("rp_products", {})

    # ── Stats bar ──
    total_rp_skus = sum(len(s.get("skus", [])) for s in rp_products.values())
    total_comp_brands = len(db.get("competitor_brands", []))
    total_comp_codes = sum(len(b.get("codes", [])) for b in db.get("competitor_brands", []))

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("RP Product Lines", len(rp_products))
    s2.metric("RP SKUs", total_rp_skus)
    s3.metric("Competitor Brands", total_comp_brands)
    s4.metric("Competitor Codes", total_comp_codes)

    st.markdown("")

    tab_lookup, tab_catalog, tab_competitor, tab_crosswalk = st.tabs([
        "Code Lookup",
        "Royal Purple Catalog",
        "Competitor Reference",
        "Conversion Guide",
    ])

    with tab_lookup:
        _render_code_lookup(db, all_codes, rp_products)
    with tab_catalog:
        _render_rp_catalog(db)
    with tab_competitor:
        _render_competitor_brands(db)
    with tab_crosswalk:
        _render_conversion_guide(db)


# ═══════════════════════════════════════════════════════════════════════
# CODE LOOKUP — the main tool Brian will use day-to-day
# ═══════════════════════════════════════════════════════════════════════

def _render_code_lookup(db, all_codes, rp_products):
    st.markdown(
        '<div style="font-size:13px;color:#C4B5E8;margin-bottom:12px;">'
        'Type any operation code from an installer report to identify the product, brand, and RP replacement.'
        '</div>',
        unsafe_allow_html=True,
    )

    search = st.text_input(
        "Enter an operation code",
        placeholder="e.g. RS5W30, VS0W20, HMX0W20, M5W30, 01320...",
        label_visibility="collapsed",
    )

    if not search:
        _render_quick_reference(rp_products)
        return

    code_upper = search.strip().upper()
    result = all_codes.get(code_upper)

    if not result:
        _try_prefix_lookup(code_upper, rp_products)
        return

    cat = result["category"]
    color = result["color"]

    if cat == "rp":
        _render_rp_result(code_upper, result, rp_products)
    elif cat == "competitor":
        _render_competitor_result(code_upper, result, rp_products)
    else:
        _render_misc_result(code_upper, result, cat)


def _render_quick_reference(rp_products):
    """Show a clean grid of all RP product lines when no search is active."""
    st.markdown("")
    st.markdown(
        '<div style="font-size:10px;font-weight:700;letter-spacing:2px;color:#8888a8;'
        'text-transform:uppercase;margin-bottom:8px;">Royal Purple Product Lines</div>',
        unsafe_allow_html=True,
    )

    for series_name, series in rp_products.items():
        color = series.get("color", "#4B2D8A")
        badge = series.get("badge", "RP")
        skus = series.get("skus", [])
        if not skus:
            continue

        sku_pills = " ".join(
            f'<span style="background:rgba(255,255,255,0.06);border:1px solid #2a2a45;'
            f'padding:4px 10px;border-radius:6px;font-size:12px;color:#e8e8f0;'
            f'font-weight:600;white-space:nowrap;">'
            f'<span style="color:{color};">{s["code"]}</span>'
            f' <span style="color:#8888a8;">{s["viscosity"]}</span></span>'
            for s in skus
        )

        short_name = series_name.split("\u2014")[0].strip() if "\u2014" in series_name else series_name

        st.markdown(
            f'<div style="background:#1a1a2e;border:1px solid #2a2a45;border-radius:10px;'
            f'padding:16px 20px;margin-bottom:10px;">'
            f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">'
            f'<span style="background:{color};color:white;padding:3px 10px;border-radius:6px;'
            f'font-size:12px;font-weight:700;">{badge}</span>'
            f'<span style="font-size:14px;font-weight:700;color:#e8e8f0;">{short_name}</span>'
            f'<span style="font-size:12px;color:#8888a8;margin-left:auto;">{len(skus)} SKUs</span>'
            f'</div>'
            f'<div style="display:flex;flex-wrap:wrap;gap:6px;">{sku_pills}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )


def _render_rp_result(code_upper, result, rp_products):
    """Display a Royal Purple product match."""
    color = result["color"]

    # Find full series data
    series_data = None
    sku_data = None
    for sname, sdata in rp_products.items():
        for sku in sdata.get("skus", []):
            if sku["code"].upper() == code_upper:
                series_data = {**sdata, "_name": sname}
                sku_data = sku
                break
        if sku_data:
            break

    st.markdown(
        f'<div style="background:#1a1a2e;border:2px solid {color};border-radius:12px;padding:24px;">'
        f'<div style="display:flex;align-items:center;gap:14px;margin-bottom:16px;">'
        f'<span style="background:{color};color:white;padding:8px 18px;border-radius:8px;'
        f'font-size:22px;font-weight:800;">{code_upper}</span>'
        f'<div>'
        f'<div style="font-size:16px;font-weight:700;color:#e8e8f0;">Royal Purple</div>'
        f'<div style="font-size:13px;color:#8888a8;">{result["series"]}</div>'
        f'</div>'
        f'<div style="margin-left:auto;background:rgba(5,150,105,0.15);color:#10B981;'
        f'padding:4px 12px;border-radius:6px;font-size:12px;font-weight:700;">RP PRODUCT</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if series_data and sku_data:
        rows = ""
        rows += _detail_row("Product Line", series_data["_name"])
        if sku_data.get("viscosity"):
            rows += _detail_row("Viscosity", sku_data["viscosity"])
        if sku_data.get("notes"):
            rows += _detail_row("Application", sku_data["notes"])
        if series_data.get("description"):
            rows += _detail_row("Description", series_data["description"])
        if series_data.get("application"):
            rows += _detail_row("Best For", series_data["application"])

        st.markdown(
            f'<table style="font-size:13px;border-collapse:collapse;width:100%;">{rows}</table>',
            unsafe_allow_html=True,
        )

        # Show other SKUs in same line
        other_skus = [s for s in series_data.get("skus", []) if s["code"].upper() != code_upper]
        if other_skus:
            pills = " ".join(
                f'<span style="background:rgba(255,255,255,0.06);border:1px solid #2a2a45;'
                f'padding:3px 10px;border-radius:6px;font-size:11px;font-weight:600;color:#e8e8f0;">'
                f'{s["code"]} {s["viscosity"]}</span>'
                for s in other_skus
            )
            st.markdown(
                f'<div style="margin-top:14px;padding-top:14px;border-top:1px solid #2a2a45;">'
                f'<div style="font-size:11px;font-weight:600;color:#8888a8;text-transform:uppercase;'
                f'letter-spacing:1.5px;margin-bottom:6px;">Other SKUs in {series_data.get("badge", "this")} Series</div>'
                f'<div style="display:flex;flex-wrap:wrap;gap:6px;">{pills}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.markdown('</div>', unsafe_allow_html=True)


def _render_competitor_result(code_upper, result, rp_products):
    """Display a competitor product match with RP replacement suggestions."""
    color = result["color"]

    st.markdown(
        f'<div style="background:#1a1a2e;border:2px solid {color};border-radius:12px;padding:24px;">'
        f'<div style="display:flex;align-items:center;gap:14px;margin-bottom:16px;">'
        f'<span style="background:{color};color:white;padding:8px 18px;border-radius:8px;'
        f'font-size:22px;font-weight:800;">{code_upper}</span>'
        f'<div>'
        f'<div style="font-size:16px;font-weight:700;color:{color};">{result["brand"]}</div>'
        f'<div style="font-size:13px;color:#8888a8;">{result["series"]}</div>'
        f'</div>'
        f'<div style="margin-left:auto;background:rgba(220,38,38,0.15);color:#F87171;'
        f'padding:4px 12px;border-radius:6px;font-size:12px;font-weight:700;">COMPETITOR</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    rows = ""
    rows += _detail_row("Brand", result["brand"])
    rows += _detail_row("Type", result["series"])
    rows += _detail_row("Product", result["viscosity"])
    st.markdown(
        f'<table style="font-size:13px;border-collapse:collapse;width:100%;">{rows}</table>',
        unsafe_allow_html=True,
    )

    if result.get("notes"):
        st.markdown(
            f'<div style="margin-top:14px;background:rgba(75,45,138,0.08);border-left:3px solid #4B2D8A;'
            f'padding:10px 14px;border-radius:0 8px 8px 0;">'
            f'<span style="font-size:11px;font-weight:700;color:#C4B5E8;text-transform:uppercase;'
            f'letter-spacing:1px;">Conversion Strategy</span><br>'
            f'<span style="font-size:13px;color:#e8e8f0;">{result["notes"]}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

    # Find RP replacements by viscosity match
    rp_replacements = _find_rp_replacements(result["viscosity"], rp_products)
    if rp_replacements:
        st.markdown(
            f'<div style="margin-top:14px;padding-top:14px;border-top:1px solid #2a2a45;">'
            f'<div style="font-size:11px;font-weight:700;color:#10B981;text-transform:uppercase;'
            f'letter-spacing:1.5px;margin-bottom:8px;">Royal Purple Replacements</div>',
            unsafe_allow_html=True,
        )
        for code, series_name, rp_color, visc in rp_replacements:
            short_name = series_name.split("\u2014")[0].strip() if "\u2014" in series_name else series_name
            st.markdown(
                f'<div style="display:inline-flex;align-items:center;gap:8px;background:rgba(16,185,129,0.08);'
                f'border:1px solid rgba(16,185,129,0.2);border-radius:8px;padding:8px 14px;margin-right:8px;margin-bottom:6px;">'
                f'<span style="background:{rp_color};color:white;padding:2px 8px;border-radius:4px;'
                f'font-size:12px;font-weight:700;">{code}</span>'
                f'<span style="color:#e8e8f0;font-size:13px;font-weight:600;">{short_name}</span>'
                f'<span style="color:#8888a8;font-size:12px;">{visc}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


def _render_misc_result(code_upper, result, cat):
    """Display a service tier or spec flag match."""
    color = result["color"]
    label = "Service Tier" if cat == "service_tier" else "Spec Flag"
    tag_color = "#64748B" if cat == "service_tier" else "#94A3B8"

    st.markdown(
        f'<div style="background:#1a1a2e;border:2px solid {color};border-radius:12px;padding:24px;">'
        f'<div style="display:flex;align-items:center;gap:14px;margin-bottom:16px;">'
        f'<span style="background:{color};color:white;padding:8px 18px;border-radius:8px;'
        f'font-size:22px;font-weight:800;">{code_upper}</span>'
        f'<div>'
        f'<div style="font-size:16px;font-weight:700;color:#e8e8f0;">{label}</div>'
        f'<div style="font-size:13px;color:#8888a8;">Not an oil product</div>'
        f'</div>'
        f'<div style="margin-left:auto;background:rgba(100,116,139,0.15);color:{tag_color};'
        f'padding:4px 12px;border-radius:6px;font-size:12px;font-weight:700;">{label.upper()}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    rows = ""
    rows += _detail_row("Type", result["series"])
    rows += _detail_row("Name", result["viscosity"])
    if result.get("notes"):
        rows += _detail_row("Details", result["notes"])
    st.markdown(
        f'<table style="font-size:13px;border-collapse:collapse;width:100%;">{rows}</table>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div style="margin-top:12px;font-size:12px;color:#8888a8;">'
        'This code appears on invoices alongside oil codes but does not represent an oil product. '
        'It can be safely ignored when classifying oil brand usage.'
        '</div>',
        unsafe_allow_html=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)


def _detail_row(label, value):
    return (
        f'<tr>'
        f'<td style="padding:8px 16px 8px 0;color:#8888a8;font-weight:600;white-space:nowrap;'
        f'vertical-align:top;width:120px;">{label}</td>'
        f'<td style="padding:8px 0;color:#e8e8f0;">{value}</td>'
        f'</tr>'
    )


def _find_rp_replacements(product_text, rp_products):
    """Find RP products matching the viscosity in a competitor product string."""
    viscosity_raw = product_text.replace("-", "").replace(" ", "").upper()
    viscosity_grades = [
        ("0W16", "0W-16"), ("0W20", "0W-20"), ("5W20", "5W-20"), ("5W30", "5W-30"),
        ("5W40", "5W-40"), ("0W40", "0W-40"), ("10W30", "10W-30"), ("10W40", "10W-40"),
        ("15W40", "15W-40"), ("20W50", "20W-50"),
    ]
    replacements = []
    for v_str, v_display in viscosity_grades:
        if v_str in viscosity_raw:
            for sname, sdata in rp_products.items():
                for sku in sdata.get("skus", []):
                    if sku.get("viscosity", "").replace("-", "").replace(" ", "").upper() == v_str:
                        replacements.append((sku["code"], sname, sdata.get("color", "#4B2D8A"), sku["viscosity"]))
            break
    return replacements


def _try_prefix_lookup(code, rp_products):
    """Attempt prefix-based identification for unknown codes."""
    RP_PREFIXES = [
        ("XPR", "Royal Purple", "XPR Series — Extreme Performance Racing", "#B91C1C"),
        ("HPS", "Royal Purple", "HPS Series — High Performance Street", "#7C3AED"),
        ("HMX", "Royal Purple", "HMX Series — High Mileage Synthetic", "#7C3AED"),
        ("RMS", "Royal Purple", "HMX Series — High Mileage Synthetic", "#7C3AED"),
        ("RSD", "Royal Purple", "Duralec — Diesel Synthetic", "#1D4ED8"),
        ("RS", "Royal Purple", "HP API Series — High Performance Synthetic", "#4B2D8A"),
        ("RP", "Royal Purple", "RP Synthetic", "#059669"),
    ]
    COMP_PREFIXES = [
        ("S0W", "CAM2", "Full Synthetic", "#DC2626"),
        ("S5W", "CAM2", "Full Synthetic", "#DC2626"),
        ("VS", "Valvoline", "Full Synthetic", "#EA580C"),
        ("VM", "Valvoline", "MaxLife", "#EA580C"),
        ("VB", "Valvoline", "Conventional", "#EA580C"),
        ("VE", "Valvoline", "Conventional", "#EA580C"),
        ("M0W", "Mobil 1", "Full Synthetic", "#B91C1C"),
        ("M5W", "Mobil 1", "Full Synthetic", "#B91C1C"),
        ("CS", "Castrol", "Edge Synthetic", "#16A34A"),
        ("PS", "Pennzoil", "Platinum Syn", "#CA8A04"),
        ("PU", "Pennzoil", "Ultra Platinum", "#CA8A04"),
        ("PB", "Pennzoil", "Conventional", "#CA8A04"),
    ]

    for prefix, brand, series, color in RP_PREFIXES:
        if code.startswith(prefix) and any(c.isdigit() for c in code):
            st.markdown(
                f'<div style="background:#1a1a2e;border:2px solid {color};border-radius:12px;padding:20px 24px;">'
                f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">'
                f'<span style="background:rgba(5,150,105,0.15);color:#10B981;padding:4px 10px;'
                f'border-radius:6px;font-size:12px;font-weight:700;">LIKELY RP</span>'
                f'<span style="font-weight:700;color:#e8e8f0;font-size:15px;">{series}</span>'
                f'</div>'
                f'<p style="font-size:13px;color:#8888a8;margin:0;">Code <strong style="color:#e8e8f0;">{code}</strong> '
                f'matches the <strong style="color:#e8e8f0;">{prefix}*</strong> prefix pattern. '
                f'Not yet in the database — add it via the Admin page to confirm.</p>'
                f'</div>',
                unsafe_allow_html=True,
            )
            return

    for prefix, brand, series, color in COMP_PREFIXES:
        if code.startswith(prefix):
            st.markdown(
                f'<div style="background:#1a1a2e;border:2px solid {color};border-radius:12px;padding:20px 24px;">'
                f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">'
                f'<span style="background:rgba(220,38,38,0.15);color:#F87171;padding:4px 10px;'
                f'border-radius:6px;font-size:12px;font-weight:700;">LIKELY COMPETITOR</span>'
                f'<span style="font-weight:700;color:#e8e8f0;font-size:15px;">{brand} — {series}</span>'
                f'</div>'
                f'<p style="font-size:13px;color:#8888a8;margin:0;">Code <strong style="color:#e8e8f0;">{code}</strong> '
                f'matches the <strong style="color:#e8e8f0;">{prefix}*</strong> prefix for '
                f'<strong style="color:#e8e8f0;">{brand} {series}</strong>. '
                f'Not yet in the database — add it via the Admin page to confirm.</p>'
                f'</div>',
                unsafe_allow_html=True,
            )

            # Still show RP replacements
            rp_replacements = _find_rp_replacements(code, {k: v for k, v in rp_products.items()})
            if rp_replacements:
                st.markdown(
                    '<div style="margin-top:10px;font-size:11px;font-weight:700;color:#10B981;'
                    'text-transform:uppercase;letter-spacing:1.5px;">Possible RP Replacements</div>',
                    unsafe_allow_html=True,
                )
                for rp_code, sname, rp_color, visc in rp_replacements:
                    short = sname.split("\u2014")[0].strip() if "\u2014" in sname else sname
                    st.markdown(
                        f'<span style="display:inline-block;background:rgba(16,185,129,0.08);'
                        f'border:1px solid rgba(16,185,129,0.2);border-radius:6px;padding:4px 10px;'
                        f'margin-right:6px;font-size:12px;">'
                        f'<span style="color:{rp_color};font-weight:700;">{rp_code}</span> '
                        f'<span style="color:#8888a8;">{short} {visc}</span></span>',
                        unsafe_allow_html=True,
                    )
            return

    st.markdown(
        f'<div style="background:#1a1a2e;border:2px solid #64748B;border-radius:12px;padding:20px 24px;">'
        f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">'
        f'<span style="background:rgba(100,116,139,0.15);color:#94A3B8;padding:4px 10px;'
        f'border-radius:6px;font-size:12px;font-weight:700;">NOT RECOGNIZED</span>'
        f'<span style="font-weight:700;color:#e8e8f0;font-size:15px;">{code}</span>'
        f'</div>'
        f'<p style="font-size:13px;color:#8888a8;margin:0;">'
        f'This code doesn\'t match any known brand prefix. '
        f'It may be an ancillary item (filter, wiper, air freshener), a spec flag, or a new SKU. '
        f'You can add it via the Admin page.</p>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════════════════════════════
# RP CATALOG — clean card-based view of all Royal Purple products
# ═══════════════════════════════════════════════════════════════════════

def _render_rp_catalog(db):
    rp_products = db.get("rp_products", {})

    if not rp_products:
        st.info("No RP products defined. Add them in the Admin page.")
        return

    for series_name, series in rp_products.items():
        badge_color = series.get("color", "#4B2D8A")
        badge_label = series.get("badge", "RP")
        skus = series.get("skus", [])
        description = series.get("description", "")
        application = series.get("application", "")
        short_name = series_name.split("\u2014")[0].strip() if "\u2014" in series_name else series_name

        st.markdown(
            f'<div style="background:#1a1a2e;border:1px solid #2a2a45;border-radius:12px;'
            f'padding:20px 24px;margin-bottom:12px;">'
            f'<div style="display:flex;align-items:center;gap:12px;margin-bottom:12px;">'
            f'<span style="background:{badge_color};color:white;padding:4px 12px;border-radius:6px;'
            f'font-size:13px;font-weight:700;">{badge_label}</span>'
            f'<span style="font-size:16px;font-weight:700;color:#e8e8f0;">{series_name}</span>'
            f'<span style="font-size:12px;color:#8888a8;margin-left:auto;">{len(skus)} SKUs</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

        if description:
            st.markdown(
                f'<p style="font-size:13px;color:#8888a8;line-height:1.6;margin:0 0 8px;">{description}</p>',
                unsafe_allow_html=True,
            )
        if application:
            st.markdown(
                f'<p style="font-size:12px;color:#C4B5E8;margin:0 0 14px;">'
                f'<strong>Best for:</strong> {application}</p>',
                unsafe_allow_html=True,
            )

        # SKU table
        if skus:
            header = (
                '<div style="display:grid;grid-template-columns:100px 80px 1fr;gap:8px;'
                'padding:6px 0;border-bottom:1px solid #2a2a45;margin-bottom:4px;">'
                '<span style="font-size:10px;font-weight:700;color:#8888a8;text-transform:uppercase;letter-spacing:1px;">Code</span>'
                '<span style="font-size:10px;font-weight:700;color:#8888a8;text-transform:uppercase;letter-spacing:1px;">Viscosity</span>'
                '<span style="font-size:10px;font-weight:700;color:#8888a8;text-transform:uppercase;letter-spacing:1px;">Application</span>'
                '</div>'
            )
            rows = ""
            for sku in skus:
                rows += (
                    f'<div style="display:grid;grid-template-columns:100px 80px 1fr;gap:8px;'
                    f'padding:8px 0;border-bottom:1px solid rgba(42,42,69,0.5);">'
                    f'<span style="font-size:13px;font-weight:700;color:{badge_color};">{sku["code"]}</span>'
                    f'<span style="font-size:13px;color:#e8e8f0;">{sku["viscosity"]}</span>'
                    f'<span style="font-size:12px;color:#8888a8;">{sku.get("notes", "")}</span>'
                    f'</div>'
                )
            st.markdown(header + rows, unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════
# COMPETITOR REFERENCE — grouped by brand with conversion notes
# ═══════════════════════════════════════════════════════════════════════

def _render_competitor_brands(db):
    competitor_brands = db.get("competitor_brands", [])
    service_tiers = db.get("service_tiers", [])
    spec_flags = db.get("spec_flags", [])

    if not competitor_brands:
        st.info("No competitor brands defined. Add them in the Admin page.")
    else:
        for brand_data in competitor_brands:
            color = brand_data.get("color", "#DC2626")
            codes = brand_data.get("codes", [])
            note = brand_data.get("conversion_note", "")

            st.markdown(
                f'<div style="background:#1a1a2e;border:1px solid #2a2a45;border-radius:12px;'
                f'padding:20px 24px;margin-bottom:12px;">'
                f'<div style="display:flex;align-items:center;gap:12px;margin-bottom:12px;">'
                f'<span style="background:{color};color:white;padding:4px 12px;border-radius:6px;'
                f'font-size:13px;font-weight:700;">{brand_data["brand"]}</span>'
                f'<span style="font-size:13px;color:#8888a8;">{brand_data.get("type", "")}</span>'
                f'<span style="font-size:12px;color:#8888a8;margin-left:auto;">{len(codes)} codes</span>'
                f'</div>',
                unsafe_allow_html=True,
            )

            if note:
                st.markdown(
                    f'<div style="background:rgba(75,45,138,0.06);border-left:3px solid {color};'
                    f'padding:8px 14px;border-radius:0 8px 8px 0;margin-bottom:12px;">'
                    f'<span style="font-size:11px;font-weight:700;color:{color};text-transform:uppercase;'
                    f'letter-spacing:1px;">Conversion Strategy</span><br>'
                    f'<span style="font-size:13px;color:#e8e8f0;">{note}</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

            if codes:
                pills = " ".join(
                    f'<span style="display:inline-block;background:rgba(255,255,255,0.04);'
                    f'border:1px solid #2a2a45;border-radius:6px;padding:6px 12px;margin-bottom:6px;">'
                    f'<span style="font-weight:700;color:{color};font-size:12px;">{sku["code"]}</span> '
                    f'<span style="color:#8888a8;font-size:12px;">'
                    f'{sku.get("product", sku.get("name", ""))}</span></span>'
                    for sku in codes
                )
                st.markdown(
                    f'<div style="display:flex;flex-wrap:wrap;gap:6px;">{pills}</div>',
                    unsafe_allow_html=True,
                )

            st.markdown('</div>', unsafe_allow_html=True)

    # Service tiers & spec flags
    if service_tiers or spec_flags:
        st.markdown("---")
        st.markdown(
            '<div style="font-size:10px;font-weight:700;letter-spacing:2px;color:#8888a8;'
            'text-transform:uppercase;margin-bottom:8px;">Non-Product Codes (ignore when classifying)</div>',
            unsafe_allow_html=True,
        )

        col_tier, col_spec = st.columns(2)
        with col_tier:
            if service_tiers:
                st.markdown("**Service Tiers**")
                for item in service_tiers:
                    st.markdown(
                        f'<div style="display:flex;align-items:center;gap:8px;padding:4px 0;">'
                        f'<span style="font-weight:700;color:#64748B;font-size:13px;min-width:40px;">{item["code"]}</span>'
                        f'<span style="font-size:12px;color:#8888a8;">{item.get("name", "")}</span>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
        with col_spec:
            if spec_flags:
                st.markdown("**Spec Flags**")
                for item in spec_flags:
                    st.markdown(
                        f'<div style="display:flex;align-items:center;gap:8px;padding:4px 0;">'
                        f'<span style="font-weight:700;color:#94A3B8;font-size:13px;min-width:60px;">{item["code"]}</span>'
                        f'<span style="font-size:12px;color:#8888a8;">{item.get("name", "")}: {item.get("description", "")}</span>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )


# ═══════════════════════════════════════════════════════════════════════
# CONVERSION GUIDE — viscosity crosswalk + segments
# ═══════════════════════════════════════════════════════════════════════

def _render_conversion_guide(db):
    crosswalk = db.get("viscosity_crosswalk", [])
    segments = db.get("conversion_segments", [])

    st.markdown(
        '<div style="font-size:13px;color:#C4B5E8;margin-bottom:16px;">'
        'Use this guide when analyzing an installer export to identify which Royal Purple product '
        'replaces each competitor viscosity grade.'
        '</div>',
        unsafe_allow_html=True,
    )

    if crosswalk:
        st.markdown(
            '<div style="font-size:10px;font-weight:700;letter-spacing:2px;color:#8888a8;'
            'text-transform:uppercase;margin-bottom:8px;">Viscosity Crosswalk</div>',
            unsafe_allow_html=True,
        )

        # Header
        st.markdown(
            '<div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;'
            'padding:8px 0;border-bottom:2px solid #2a2a45;">'
            '<span style="font-size:11px;font-weight:700;color:#8888a8;">CURRENT OIL</span>'
            '<span style="font-size:11px;font-weight:700;color:#4B2D8A;">RS SERIES</span>'
            '<span style="font-size:11px;font-weight:700;color:#7C3AED;">HMX (HIGH MILEAGE)</span>'
            '<span style="font-size:11px;font-weight:700;color:#1D4ED8;">DURALEC (DIESEL)</span>'
            '</div>',
            unsafe_allow_html=True,
        )
        for row in crosswalk:
            rs_val = row.get("rs", "\u2014")
            hmx_val = row.get("hmx", "\u2014")
            rsd_val = row.get("rsd", "\u2014")
            st.markdown(
                f'<div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;'
                f'padding:8px 0;border-bottom:1px solid rgba(42,42,69,0.5);">'
                f'<span style="font-size:13px;color:#e8e8f0;font-weight:600;">{row.get("current", "")}</span>'
                f'<span style="font-size:13px;color:{"#4B2D8A" if rs_val != chr(8212) else "#3a3a55"};font-weight:600;">{rs_val}</span>'
                f'<span style="font-size:13px;color:{"#7C3AED" if hmx_val != chr(8212) else "#3a3a55"};font-weight:600;">{hmx_val}</span>'
                f'<span style="font-size:13px;color:{"#1D4ED8" if rsd_val != chr(8212) else "#3a3a55"};font-weight:600;">{rsd_val}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
        st.markdown("")

    if segments:
        st.markdown("---")
        st.markdown(
            '<div style="font-size:10px;font-weight:700;letter-spacing:2px;color:#8888a8;'
            'text-transform:uppercase;margin-bottom:8px;">Conversion Segments</div>',
            unsafe_allow_html=True,
        )
        st.caption("When classifying a full-code installer export, each customer falls into one of these segments.")
        st.markdown("")

        for seg in segments:
            color = seg.get("color", "#64748B")
            st.markdown(
                f'<div style="background:#1a1a2e;border:1px solid {color};border-radius:10px;'
                f'padding:16px 20px;margin-bottom:10px;">'
                f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">'
                f'<span style="font-size:15px;font-weight:700;color:{color};">{seg.get("segment", "")}</span>'
                f'<span style="background:{color};color:white;padding:2px 10px;border-radius:10px;'
                f'font-size:11px;font-weight:700;">Difficulty: {seg.get("difficulty", "")}</span>'
                f'</div>'
                f'<div style="font-size:12px;color:#8888a8;margin-bottom:6px;"><strong>Codes:</strong> {seg.get("codes", "")}</div>'
                f'<div style="font-size:13px;color:#C4B5E8;margin-bottom:6px;">{seg.get("rationale", "")}</div>'
                f'<div style="font-size:12px;color:{color};font-weight:600;">Suggested RP: {seg.get("suggested_sku", "")}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.markdown("---")
    st.markdown(
        '<div style="font-size:10px;font-weight:700;letter-spacing:2px;color:#8888a8;'
        'text-transform:uppercase;margin-bottom:8px;">How Invoice Classification Works</div>',
        unsafe_allow_html=True,
    )

    steps = [
        ("1", "Group all rows by Invoice #", "Each invoice generates multiple rows in a full-code export — one per operation code. All rows share the same revenue total."),
        ("2", "Find the oil product code", "Ignore spec flags (GF6, DEXOS1), service tiers (S1-S6), and ancillary items (OF*, AF*, FB)."),
        ("3", "Classify the oil code", "Use the prefix rules above to identify Royal Purple vs. specific competitor brands."),
        ("4", "Assign conversion segment", "Calculate the RP ticket premium vs. competitor average ticket."),
    ]

    for num, title, desc in steps:
        st.markdown(
            f'<div style="display:flex;gap:12px;align-items:flex-start;margin-bottom:10px;">'
            f'<span style="background:#4B2D8A;color:white;min-width:28px;height:28px;border-radius:50%;'
            f'display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;">{num}</span>'
            f'<div>'
            f'<div style="font-size:13px;font-weight:700;color:#e8e8f0;">{title}</div>'
            f'<div style="font-size:12px;color:#8888a8;">{desc}</div>'
            f'</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
