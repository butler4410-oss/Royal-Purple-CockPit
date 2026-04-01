import streamlit as st
from product_reference import load_codes_db, save_codes_db

_COLOR_PRESETS = {
    "Purple (RP default)": "#4B2D8A",
    "Deep Purple": "#7C3AED",
    "Blue": "#1D4ED8",
    "Emerald": "#059669",
    "Gold": "#D97706",
    "Red": "#DC2626",
    "Orange": "#EA580C",
    "Dark Red": "#B91C1C",
    "Green": "#16A34A",
    "Yellow": "#CA8A04",
    "Indigo": "#4F46E5",
    "Gray": "#64748B",
}

_PRESET_NAMES = list(_COLOR_PRESETS.keys())
_PRESET_VALUES = list(_COLOR_PRESETS.values())


def _color_index(hex_val):
    """Return the index of a hex color in the preset list, or 0 (first preset) if not found."""
    try:
        return _PRESET_VALUES.index(hex_val)
    except ValueError:
        return 0


def _color_select(label, current_hex, key):
    """Render a color selectbox that shows name + swatch. Returns the selected hex value."""
    idx = _color_index(current_hex)
    chosen = st.selectbox(
        label,
        _PRESET_NAMES,
        index=idx,
        key=key,
        label_visibility="visible",
    )
    return _COLOR_PRESETS[chosen]


def render():
    st.markdown("### Code Database Editor")
    st.caption("Changes auto-save when you edit any field. Use the forms at the bottom of each section to add new entries.")

    st.markdown("")
    tab_rp, tab_comp, tab_misc = st.tabs(["Royal Purple Products", "Competitor Brands", "Service Tiers & Spec Flags"])

    with tab_rp:
        _admin_rp_products()
    with tab_comp:
        _admin_competitor_brands()
    with tab_misc:
        _admin_misc()


# ═══════════════════════════════════════════════════════════════════════
# ROYAL PURPLE PRODUCTS
# ═══════════════════════════════════════════════════════════════════════

def _reorder_series(db, series_names, from_idx, to_idx):
    """Swap two series positions and rebuild the ordered dict."""
    names = list(series_names)
    names[from_idx], names[to_idx] = names[to_idx], names[from_idx]
    db["rp_products"] = {name: db["rp_products"][name] for name in names}
    save_codes_db(db)


def _admin_rp_products():
    db = load_codes_db()
    rp_products = db.get("rp_products", {})

    st.markdown("#### Royal Purple Product Series")
    st.caption("Each series groups related SKUs. Edit any field and it saves automatically.")

    series_names = list(rp_products.keys())

    if not series_names:
        st.info("No RP product series defined yet.")
    else:
        for pos, series_name in enumerate(series_names):
            series = rp_products[series_name]
            skus = series.get("skus", [])
            color = series.get("color", "#4B2D8A")
            badge = series.get("badge", "RP")
            short = series_name.split("\u2014")[0].strip() if "\u2014" in series_name else series_name

            # ── Reorder bar ──
            bar_cols = st.columns([0.5, 0.5, 8])
            with bar_cols[0]:
                if pos > 0:
                    if st.button("\u25B2", key=f"up_{series_name}", help="Move up",
                                  use_container_width=True):
                        _reorder_series(db, series_names, pos, pos - 1)
                        st.rerun()
            with bar_cols[1]:
                if pos < len(series_names) - 1:
                    if st.button("\u25BC", key=f"dn_{series_name}", help="Move down",
                                  use_container_width=True):
                        _reorder_series(db, series_names, pos, pos + 1)
                        st.rerun()
            with bar_cols[2]:
                st.markdown(
                    f'<div style="display:flex;align-items:center;gap:10px;padding:4px 0;">'
                    f'<span style="background:{color};color:white;padding:2px 10px;border-radius:5px;'
                    f'font-size:12px;font-weight:700;">{badge}</span>'
                    f'<span style="font-size:14px;font-weight:600;color:#e8e8f0;">{short}</span>'
                    f'<span style="font-size:12px;color:#8888a8;">{len(skus)} SKUs</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

            with st.expander(f"Edit {short}", expanded=False):
                _edit_rp_series(db, series_name, series)

    st.markdown("---")
    st.markdown("##### Add New Series")
    with st.form("add_rp_series"):
        new_name = st.text_input("Series Name", placeholder="e.g. RS Series — High Performance Synthetic")
        col1, col2 = st.columns(2)
        with col1:
            new_badge = st.text_input("Badge Label", placeholder="e.g. RS", max_chars=6)
        with col2:
            new_color_name = st.selectbox("Color", _PRESET_NAMES, index=0, key="add_series_color")
        new_desc = st.text_area("Description", placeholder="Short description of this product series...")
        new_app = st.text_input("Best For / Application", placeholder="e.g. Modern engines, daily drivers")
        if st.form_submit_button("Add Series", type="primary"):
            if not new_name.strip():
                st.error("Series name is required.")
            elif new_name.strip() in rp_products:
                st.error("A series with that name already exists.")
            else:
                db["rp_products"][new_name.strip()] = {
                    "color": _COLOR_PRESETS[new_color_name],
                    "badge": new_badge.strip().upper() or "RP",
                    "description": new_desc.strip(),
                    "application": new_app.strip(),
                    "skus": [],
                }
                save_codes_db(db)
                st.success(f"Series '{new_name.strip()}' added.")
                st.rerun()


def _edit_rp_series(db, series_name, series):
    skus = series.get("skus", [])
    changed = False

    # ── Series metadata (auto-save) ──
    st.markdown("**Series Details**")
    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        new_badge = st.text_input("Badge", value=series.get("badge", ""), max_chars=6,
                                   key=f"badge_{series_name}")
    with c2:
        new_color = _color_select("Color", series.get("color", "#4B2D8A"),
                                   key=f"color_{series_name}")
    with c3:
        # Live preview
        st.markdown(
            f'<div style="background:{new_color};color:white;padding:8px 12px;border-radius:6px;'
            f'font-size:14px;font-weight:700;text-align:center;margin-top:28px;">'
            f'{new_badge.strip().upper() or series.get("badge", "RP")}</div>',
            unsafe_allow_html=True,
        )

    new_desc = st.text_area("Description", value=series.get("description", ""), height=80,
                             key=f"desc_{series_name}")
    new_app = st.text_input("Best For", value=series.get("application", ""),
                             key=f"app_{series_name}")

    # Detect changes and auto-save
    updated_meta = {
        "badge": new_badge.strip().upper() or "RP",
        "color": new_color,
        "description": new_desc.strip(),
        "application": new_app.strip(),
    }
    current_meta = {
        "badge": series.get("badge", "RP"),
        "color": series.get("color", "#4B2D8A"),
        "description": series.get("description", ""),
        "application": series.get("application", ""),
    }
    if updated_meta != current_meta:
        db["rp_products"][series_name].update(updated_meta)
        save_codes_db(db)
        changed = True

    # ── Delete series ──
    if st.button("Delete Series", key=f"del_series_{series_name}", type="secondary"):
        st.session_state[f"confirm_del_series_{series_name}"] = True

    if st.session_state.get(f"confirm_del_series_{series_name}"):
        st.warning(f"Delete the entire **{series_name}** series and all {len(skus)} SKUs?")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Yes, delete", key=f"confirm_yes_{series_name}", type="primary"):
                del db["rp_products"][series_name]
                save_codes_db(db)
                st.session_state.pop(f"confirm_del_series_{series_name}", None)
                st.rerun()
        with c2:
            if st.button("Cancel", key=f"confirm_no_{series_name}"):
                st.session_state.pop(f"confirm_del_series_{series_name}", None)
                st.rerun()

    # ── Products (auto-save per row) ──
    st.markdown("---")
    st.markdown("**Products**")
    if not skus:
        st.caption("No products in this series yet.")
    else:
        # Column headers
        hv, hn, ha = st.columns([2, 5, 0.6])
        hv.markdown('<span style="font-size:11px;font-weight:700;color:#8888a8;text-transform:uppercase;">Viscosity</span>', unsafe_allow_html=True)
        hn.markdown('<span style="font-size:11px;font-weight:700;color:#8888a8;text-transform:uppercase;">Notes / Application</span>', unsafe_allow_html=True)

        for i, sku in enumerate(skus):
            col_visc, col_notes, col_act = st.columns([2, 5, 0.6])
            with col_visc:
                new_visc = st.text_input("Viscosity", value=sku.get("viscosity", ""),
                                          key=f"sku_visc_{series_name}_{i}", label_visibility="collapsed")
            with col_notes:
                new_notes = st.text_input("Notes", value=sku.get("notes", ""),
                                           key=f"sku_notes_{series_name}_{i}", label_visibility="collapsed")

            # Auto-save changes
            updated_sku = {
                "viscosity": new_visc.strip(),
                "notes": new_notes.strip(),
            }
            if (updated_sku["viscosity"] != sku.get("viscosity", "") or
                    updated_sku["notes"] != sku.get("notes", "")):
                db["rp_products"][series_name]["skus"][i] = updated_sku
                save_codes_db(db)

            with col_act:
                if st.button("Del", key=f"del_sku_{series_name}_{i}", help="Delete product"):
                    db["rp_products"][series_name]["skus"].pop(i)
                    save_codes_db(db)
                    st.rerun()

    # ── Add product (form — needs submit) ──
    st.markdown("")
    with st.form(f"add_sku_{series_name}"):
        st.markdown("**Add Product**")
        c1, c2 = st.columns([2, 5])
        with c1:
            add_visc = st.text_input("Viscosity", placeholder="5W-30")
        with c2:
            add_notes = st.text_input("Notes", placeholder="e.g. Most common viscosity across all platforms")
        if st.form_submit_button("Add Product"):
            if not add_visc.strip():
                st.error("Viscosity is required.")
            else:
                db["rp_products"][series_name]["skus"].append({
                    "viscosity": add_visc.strip(),
                    "notes": add_notes.strip(),
                })
                save_codes_db(db)
                st.success(f"{add_visc.strip()} added.")
                st.rerun()


# ═══════════════════════════════════════════════════════════════════════
# COMPETITOR BRANDS
# ═══════════════════════════════════════════════════════════════════════

def _admin_competitor_brands():
    db = load_codes_db()
    competitor_brands = db.get("competitor_brands", [])

    st.markdown("#### Competitor Brands")
    st.caption("Manage competitor oil brands and their known operation codes. Edits auto-save.")

    for idx, brand_data in enumerate(competitor_brands):
        codes = brand_data.get("codes", [])
        with st.expander(f"**{brand_data['brand']}** — {brand_data.get('type', '')} — {len(codes)} codes"):
            _edit_competitor_brand(db, idx, brand_data)

    st.markdown("---")
    st.markdown("##### Add New Competitor Brand")
    with st.form("add_competitor_brand"):
        c1, c2 = st.columns(2)
        with c1:
            new_brand = st.text_input("Brand Name", placeholder="e.g. Mobil 1")
        with c2:
            new_type = st.text_input("Type Description", placeholder="e.g. Full Synthetic")
        new_color_name = st.selectbox("Color", _PRESET_NAMES, index=_PRESET_NAMES.index("Red"),
                                       key="add_brand_color")
        new_note = st.text_area("Conversion Strategy Note",
                                 placeholder="What's the best approach to convert customers of this brand to Royal Purple?",
                                 height=80)
        if st.form_submit_button("Add Brand", type="primary"):
            if not new_brand.strip():
                st.error("Brand name is required.")
            else:
                db["competitor_brands"].append({
                    "brand": new_brand.strip(),
                    "type": new_type.strip(),
                    "color": _COLOR_PRESETS[new_color_name],
                    "codes": [],
                    "conversion_note": new_note.strip(),
                })
                save_codes_db(db)
                st.success(f"Brand '{new_brand.strip()}' added.")
                st.rerun()


def _edit_competitor_brand(db, idx, brand_data):
    codes = brand_data.get("codes", [])

    # ── Brand metadata (auto-save) ──
    st.markdown("**Brand Details**")
    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        new_brand_name = st.text_input("Brand Name", value=brand_data.get("brand", ""),
                                        key=f"brand_name_{idx}")
    with c2:
        new_type = st.text_input("Type", value=brand_data.get("type", ""),
                                  key=f"brand_type_{idx}")
    with c3:
        new_color = _color_select("Color", brand_data.get("color", "#DC2626"),
                                   key=f"brand_color_{idx}")
    new_note = st.text_area("Conversion Strategy Note", value=brand_data.get("conversion_note", ""),
                             height=70, key=f"brand_note_{idx}")

    # Detect changes and auto-save
    updated = {
        "brand": new_brand_name.strip(),
        "type": new_type.strip(),
        "color": new_color,
        "conversion_note": new_note.strip(),
    }
    current = {
        "brand": brand_data.get("brand", ""),
        "type": brand_data.get("type", ""),
        "color": brand_data.get("color", "#DC2626"),
        "conversion_note": brand_data.get("conversion_note", ""),
    }
    if updated != current:
        db["competitor_brands"][idx].update(updated)
        save_codes_db(db)

    # ── Delete brand ──
    if st.button("Delete Brand", key=f"del_brand_{idx}", type="secondary"):
        st.session_state[f"confirm_del_brand_{idx}"] = True

    if st.session_state.get(f"confirm_del_brand_{idx}"):
        st.warning(f"Delete **{brand_data.get('brand', 'this brand')}** and all {len(codes)} codes?")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Yes, delete", key=f"confirm_del_brand_yes_{idx}", type="primary"):
                db["competitor_brands"].pop(idx)
                save_codes_db(db)
                st.session_state.pop(f"confirm_del_brand_{idx}", None)
                st.rerun()
        with c2:
            if st.button("Cancel", key=f"confirm_del_brand_no_{idx}"):
                st.session_state.pop(f"confirm_del_brand_{idx}", None)
                st.rerun()

    # ── Codes (auto-save per row) ──
    st.markdown("---")
    st.markdown("**Codes**")
    if not codes:
        st.caption("No codes defined for this brand yet.")
    else:
        hc, hp, ha = st.columns([2, 5, 0.6])
        hc.markdown('<span style="font-size:11px;font-weight:700;color:#8888a8;text-transform:uppercase;">Code</span>', unsafe_allow_html=True)
        hp.markdown('<span style="font-size:11px;font-weight:700;color:#8888a8;text-transform:uppercase;">Product Name</span>', unsafe_allow_html=True)

        for i, sku in enumerate(codes):
            col_code, col_product, col_act = st.columns([2, 5, 0.6])
            with col_code:
                new_code = st.text_input("Code", value=sku["code"],
                                          key=f"comp_code_{idx}_{i}", label_visibility="collapsed")
            with col_product:
                new_product = st.text_input("Product Name", value=sku.get("product", ""),
                                             key=f"comp_prod_{idx}_{i}", label_visibility="collapsed")

            # Auto-save
            updated_code = new_code.strip().upper()
            updated_product = new_product.strip()
            if updated_code != sku["code"] or updated_product != sku.get("product", ""):
                db["competitor_brands"][idx]["codes"][i] = {
                    "code": updated_code,
                    "product": updated_product,
                }
                save_codes_db(db)

            with col_act:
                if st.button("Del", key=f"del_comp_{idx}_{i}", help="Delete"):
                    db["competitor_brands"][idx]["codes"].pop(i)
                    save_codes_db(db)
                    st.rerun()

    # ── Add code (form) ──
    st.markdown("")
    with st.form(f"add_comp_code_{idx}"):
        st.markdown("**Add Code**")
        c1, c2 = st.columns([2, 5])
        with c1:
            add_code = st.text_input("Code", placeholder="VS5W30")
        with c2:
            add_product = st.text_input("Product Name", placeholder="e.g. Valvoline Full Synthetic 5W-30")
        if st.form_submit_button("Add Code"):
            if not add_code.strip():
                st.error("Code is required.")
            else:
                existing = [c["code"].upper() for c in db["competitor_brands"][idx]["codes"]]
                if add_code.strip().upper() in existing:
                    st.error(f"Code {add_code.strip().upper()} already exists for this brand.")
                else:
                    db["competitor_brands"][idx]["codes"].append({
                        "code": add_code.strip().upper(),
                        "product": add_product.strip(),
                    })
                    save_codes_db(db)
                    st.success(f"Code {add_code.strip().upper()} added.")
                    st.rerun()


# ═══════════════════════════════════════════════════════════════════════
# SERVICE TIERS & SPEC FLAGS
# ═══════════════════════════════════════════════════════════════════════

def _admin_misc():
    db = load_codes_db()

    st.markdown("#### Service Tier Codes")
    st.caption("Codes like S1-S6, B7-B10 that appear on invoices but do not represent oil products. Edits auto-save.")

    service_tiers = db.get("service_tiers", [])
    if service_tiers:
        hc, hn, hd, ha = st.columns([1, 2, 4, 0.6])
        hc.markdown('<span style="font-size:11px;font-weight:700;color:#8888a8;text-transform:uppercase;">Code</span>', unsafe_allow_html=True)
        hn.markdown('<span style="font-size:11px;font-weight:700;color:#8888a8;text-transform:uppercase;">Name</span>', unsafe_allow_html=True)
        hd.markdown('<span style="font-size:11px;font-weight:700;color:#8888a8;text-transform:uppercase;">Description</span>', unsafe_allow_html=True)

    for i, item in enumerate(service_tiers):
        col_code, col_name, col_desc, col_act = st.columns([1, 2, 4, 0.6])
        with col_code:
            new_code = st.text_input("Code", value=item["code"], key=f"st_code_{i}", label_visibility="collapsed")
        with col_name:
            new_name = st.text_input("Name", value=item.get("name", ""), key=f"st_name_{i}", label_visibility="collapsed")
        with col_desc:
            new_desc = st.text_input("Description", value=item.get("description", ""), key=f"st_desc_{i}", label_visibility="collapsed")

        # Auto-save
        updated_st = {"code": new_code.strip().upper(), "name": new_name.strip(), "description": new_desc.strip()}
        if (updated_st["code"] != item["code"] or
                updated_st["name"] != item.get("name", "") or
                updated_st["description"] != item.get("description", "")):
            db["service_tiers"][i] = updated_st
            save_codes_db(db)

        with col_act:
            if st.button("Del", key=f"del_st_{i}", help="Delete"):
                db["service_tiers"].pop(i)
                save_codes_db(db)
                st.rerun()

    with st.form("add_service_tier"):
        c1, c2, c3 = st.columns([1, 2, 4])
        with c1:
            add_code = st.text_input("Code", placeholder="S7")
        with c2:
            add_name = st.text_input("Name", placeholder="Service Tier 7")
        with c3:
            add_desc = st.text_input("Description", placeholder="Description of this tier")
        if st.form_submit_button("Add Service Tier"):
            if add_code.strip():
                db["service_tiers"].append({"code": add_code.strip().upper(), "name": add_name.strip(), "description": add_desc.strip()})
                save_codes_db(db)
                st.rerun()

    st.markdown("---")
    st.markdown("#### Spec Flags")
    st.caption("Certification flags like GF6, DEXOS1 that appear alongside oil codes but are not oil products themselves.")

    spec_flags = db.get("spec_flags", [])
    if spec_flags:
        hc, hn, hd, ha = st.columns([1, 2, 4, 0.6])
        hc.markdown('<span style="font-size:11px;font-weight:700;color:#8888a8;text-transform:uppercase;">Code</span>', unsafe_allow_html=True)
        hn.markdown('<span style="font-size:11px;font-weight:700;color:#8888a8;text-transform:uppercase;">Name</span>', unsafe_allow_html=True)
        hd.markdown('<span style="font-size:11px;font-weight:700;color:#8888a8;text-transform:uppercase;">Description</span>', unsafe_allow_html=True)

    for i, item in enumerate(spec_flags):
        col_code, col_name, col_desc, col_act = st.columns([1, 2, 4, 0.6])
        with col_code:
            new_code = st.text_input("Code", value=item["code"], key=f"sf_code_{i}", label_visibility="collapsed")
        with col_name:
            new_name = st.text_input("Name", value=item.get("name", ""), key=f"sf_name_{i}", label_visibility="collapsed")
        with col_desc:
            new_desc = st.text_input("Description", value=item.get("description", ""), key=f"sf_desc_{i}", label_visibility="collapsed")

        # Auto-save
        updated_sf = {"code": new_code.strip().upper(), "name": new_name.strip(), "description": new_desc.strip()}
        if (updated_sf["code"] != item["code"] or
                updated_sf["name"] != item.get("name", "") or
                updated_sf["description"] != item.get("description", "")):
            db["spec_flags"][i] = updated_sf
            save_codes_db(db)

        with col_act:
            if st.button("Del", key=f"del_sf_{i}", help="Delete"):
                db["spec_flags"].pop(i)
                save_codes_db(db)
                st.rerun()

    with st.form("add_spec_flag"):
        c1, c2, c3 = st.columns([1, 2, 4])
        with c1:
            add_code = st.text_input("Code", placeholder="GF7")
        with c2:
            add_name = st.text_input("Name", placeholder="ILSAC GF-7")
        with c3:
            add_desc = st.text_input("Description", placeholder="Description")
        if st.form_submit_button("Add Spec Flag"):
            if add_code.strip():
                db["spec_flags"].append({"code": add_code.strip().upper(), "name": add_name.strip(), "description": add_desc.strip()})
                save_codes_db(db)
                st.rerun()
