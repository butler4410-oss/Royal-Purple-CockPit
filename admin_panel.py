import streamlit as st
import os
import copy
from product_reference import load_codes_db, save_codes_db

ADMIN_USERNAME = "admin"

_COLOR_PRESETS = {
    "Purple (RP default)": "#e31837",
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


def _check_password():
    if st.session_state.get("admin_authenticated"):
        return True

    st.markdown(
        '<div style="max-width:400px;margin:60px auto 0;">'
        '<div style="background:#e31837;color:white;padding:20px 24px;border-radius:10px 10px 0 0;">'
        '<h3 style="margin:0;font-size:18px;">Admin Login</h3>'
        '<p style="margin:6px 0 0;font-size:13px;opacity:0.8;">Butler Performance Partnership Hub</p>'
        '</div>',
        unsafe_allow_html=True,
    )

    with st.form("admin_login_form"):
        st.markdown('<div style="background:white;border:1px solid #E2E8F0;border-top:none;padding:20px 24px;border-radius:0 0 10px 10px;">', unsafe_allow_html=True)
        username = st.text_input("Username", placeholder="admin")
        password = st.text_input("Password", type="password", placeholder="••••••••")
        submitted = st.form_submit_button("Sign In", use_container_width=True, type="primary")
        st.markdown("</div>", unsafe_allow_html=True)

    if submitted:
        expected_pw = os.environ.get("ADMIN_PASSWORD", "")
        if username == ADMIN_USERNAME and password == expected_pw and expected_pw:
            st.session_state["admin_authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect username or password.")

    return False


def render():
    if not _check_password():
        return

    col_title, col_logout = st.columns([5, 1])
    with col_title:
        st.markdown("### Code Database Editor")
        st.caption("Add, edit, or remove Butler Performance and competitor operation codes. Changes save immediately to the live database.")
    with col_logout:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Sign Out", type="secondary"):
            st.session_state.pop("admin_authenticated", None)
            st.rerun()

    st.markdown("")
    tab_rp, tab_comp, tab_misc = st.tabs(["Butler Performance Products", "Competitor Brands", "Service Tiers & Spec Flags"])

    with tab_rp:
        _admin_rp_products()
    with tab_comp:
        _admin_competitor_brands()
    with tab_misc:
        _admin_misc()


def _color_picker(label, key, default="#e31837"):
    col_preset, col_hex = st.columns([3, 2])
    with col_preset:
        preset_key = f"{key}_preset"
        preset_names = ["Custom"] + list(_COLOR_PRESETS.keys())
        current_val = st.session_state.get(key, default)
        default_preset = "Custom"
        for name, hex_val in _COLOR_PRESETS.items():
            if hex_val.lower() == current_val.lower():
                default_preset = name
                break
        chosen_preset = st.selectbox(f"{label} Preset", preset_names,
                                      index=preset_names.index(default_preset),
                                      key=preset_key, label_visibility="collapsed")
    with col_hex:
        if chosen_preset != "Custom":
            hex_val = _COLOR_PRESETS[chosen_preset]
        else:
            hex_val = current_val
        val = st.text_input(f"{label} Hex", value=hex_val, key=key, label_visibility="collapsed", placeholder="#e31837")
    return val


def _admin_rp_products():
    db = load_codes_db()
    rp_products = db.get("rp_products", {})

    st.markdown("#### Butler Performance Product Series")
    st.caption("Each series groups related SKUs (e.g. RS Series, HMX, Duralec). SKUs are the individual operation codes.")

    series_names = list(rp_products.keys())

    if not series_names:
        st.info("No RP product series defined yet.")
    else:
        for series_name in series_names:
            series = rp_products[series_name]
            skus = series.get("skus", [])
            with st.expander(f"**{series_name}** — {len(skus)} SKU{'s' if len(skus) != 1 else ''}"):
                _edit_rp_series(db, series_name, series)

    st.markdown("---")
    st.markdown("##### Add New Series")
    with st.form("add_rp_series"):
        new_name = st.text_input("Series Name", placeholder="e.g. RS Series — High Performance Synthetic")
        col1, col2 = st.columns(2)
        with col1:
            new_badge = st.text_input("Badge Label", placeholder="e.g. RS", max_chars=6)
        with col2:
            new_color = st.text_input("Color (hex)", value="#e31837", placeholder="#e31837")
        new_desc = st.text_area("Description", placeholder="Short description of this product series...")
        new_app = st.text_input("Best For / Application", placeholder="e.g. Modern engines, daily drivers")
        if st.form_submit_button("Add Series", type="primary"):
            if not new_name.strip():
                st.error("Series name is required.")
            elif new_name.strip() in rp_products:
                st.error("A series with that name already exists.")
            else:
                db["rp_products"][new_name.strip()] = {
                    "color": new_color.strip() or "#e31837",
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

    col_meta, col_del = st.columns([5, 1])
    with col_meta:
        with st.form(f"edit_series_meta_{series_name}"):
            st.markdown("**Series Details**")
            c1, c2, c3 = st.columns([3, 1, 2])
            with c1:
                new_badge = st.text_input("Badge", value=series.get("badge", ""), max_chars=6)
            with c2:
                new_color = st.text_input("Color", value=series.get("color", "#e31837"))
            with c3:
                st.markdown(
                    f'<div style="background:{series.get("color","#e31837")};color:white;padding:6px 10px;border-radius:6px;font-size:13px;font-weight:700;text-align:center;margin-top:28px;">{series.get("badge","RP")}</div>',
                    unsafe_allow_html=True,
                )
            new_desc = st.text_area("Description", value=series.get("description", ""), height=80)
            new_app = st.text_input("Best For", value=series.get("application", ""))
            if st.form_submit_button("Save Series Details"):
                db["rp_products"][series_name].update({
                    "badge": new_badge.strip().upper() or "RP",
                    "color": new_color.strip() or "#e31837",
                    "description": new_desc.strip(),
                    "application": new_app.strip(),
                })
                save_codes_db(db)
                st.success("Series details saved.")
                st.rerun()

    with col_del:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🗑 Delete Series", key=f"del_series_{series_name}", type="secondary"):
            st.session_state[f"confirm_del_series_{series_name}"] = True

    if st.session_state.get(f"confirm_del_series_{series_name}"):
        st.warning(f"Delete the entire **{series_name}** series and all {len(skus)} SKUs?")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Yes, delete", key=f"confirm_yes_{series_name}", type="primary"):
                del db["rp_products"][series_name]
                save_codes_db(db)
                st.session_state.pop(f"confirm_del_series_{series_name}", None)
                st.success("Series deleted.")
                st.rerun()
        with c2:
            if st.button("Cancel", key=f"confirm_no_{series_name}"):
                st.session_state.pop(f"confirm_del_series_{series_name}", None)
                st.rerun()

    st.markdown("**SKUs**")
    if not skus:
        st.caption("No SKUs in this series yet.")
    else:
        for i, sku in enumerate(skus):
            col_code, col_visc, col_notes, col_act = st.columns([1.5, 1.5, 4, 1])
            with col_code:
                new_code = st.text_input("Code", value=sku["code"], key=f"sku_code_{series_name}_{i}", label_visibility="collapsed")
            with col_visc:
                new_visc = st.text_input("Viscosity", value=sku.get("viscosity", ""), key=f"sku_visc_{series_name}_{i}", label_visibility="collapsed")
            with col_notes:
                new_notes = st.text_input("Notes", value=sku.get("notes", ""), key=f"sku_notes_{series_name}_{i}", label_visibility="collapsed")
            with col_act:
                save_col, del_col = st.columns(2)
                with save_col:
                    if st.button("💾", key=f"save_sku_{series_name}_{i}", help="Save"):
                        db["rp_products"][series_name]["skus"][i] = {
                            "code": new_code.strip().upper(),
                            "viscosity": new_visc.strip(),
                            "notes": new_notes.strip(),
                        }
                        save_codes_db(db)
                        st.success(f"{new_code.strip().upper()} saved.")
                        st.rerun()
                with del_col:
                    if st.button("🗑", key=f"del_sku_{series_name}_{i}", help="Delete"):
                        db["rp_products"][series_name]["skus"].pop(i)
                        save_codes_db(db)
                        st.rerun()

    st.markdown("")
    with st.form(f"add_sku_{series_name}"):
        st.markdown("**Add SKU**")
        c1, c2, c3 = st.columns([1.5, 1.5, 4])
        with c1:
            add_code = st.text_input("Code", placeholder="RS5W30", label_visibility="visible")
        with c2:
            add_visc = st.text_input("Viscosity", placeholder="5W-30", label_visibility="visible")
        with c3:
            add_notes = st.text_input("Notes", placeholder="e.g. Most common viscosity across all platforms", label_visibility="visible")
        if st.form_submit_button("Add SKU"):
            if not add_code.strip():
                st.error("Code is required.")
            else:
                existing_codes = [s["code"].upper() for s in db["rp_products"][series_name]["skus"]]
                if add_code.strip().upper() in existing_codes:
                    st.error(f"Code {add_code.strip().upper()} already exists in this series.")
                else:
                    db["rp_products"][series_name]["skus"].append({
                        "code": add_code.strip().upper(),
                        "viscosity": add_visc.strip(),
                        "notes": add_notes.strip(),
                    })
                    save_codes_db(db)
                    st.success(f"SKU {add_code.strip().upper()} added.")
                    st.rerun()


def _admin_competitor_brands():
    db = load_codes_db()
    competitor_brands = db.get("competitor_brands", [])

    st.markdown("#### Competitor Brands")
    st.caption("Manage competitor oil brands and their known operation codes.")

    for idx, brand_data in enumerate(competitor_brands):
        color = brand_data.get("color", "#DC2626")
        codes = brand_data.get("codes", [])
        with st.expander(f"**{brand_data['brand']}** — {brand_data.get('type','')} — {len(codes)} codes"):
            _edit_competitor_brand(db, idx, brand_data)

    st.markdown("---")
    st.markdown("##### Add New Competitor Brand")
    with st.form("add_competitor_brand"):
        c1, c2 = st.columns(2)
        with c1:
            new_brand = st.text_input("Brand Name", placeholder="e.g. Mobil 1")
        with c2:
            new_type = st.text_input("Type Description", placeholder="e.g. Full Synthetic")
        new_color = st.text_input("Color (hex)", value="#DC2626", placeholder="#DC2626")
        new_note = st.text_area("Conversion Strategy Note", placeholder="What's the best approach to convert customers of this brand to Butler Performance?", height=80)
        if st.form_submit_button("Add Brand", type="primary"):
            if not new_brand.strip():
                st.error("Brand name is required.")
            else:
                db["competitor_brands"].append({
                    "brand": new_brand.strip(),
                    "type": new_type.strip(),
                    "color": new_color.strip() or "#DC2626",
                    "codes": [],
                    "conversion_note": new_note.strip(),
                })
                save_codes_db(db)
                st.success(f"Brand '{new_brand.strip()}' added.")
                st.rerun()


def _edit_competitor_brand(db, idx, brand_data):
    color = brand_data.get("color", "#DC2626")
    codes = brand_data.get("codes", [])

    col_meta, col_del = st.columns([5, 1])
    with col_meta:
        with st.form(f"edit_brand_meta_{idx}"):
            st.markdown("**Brand Details**")
            c1, c2, c3 = st.columns([2, 2, 2])
            with c1:
                new_brand_name = st.text_input("Brand Name", value=brand_data.get("brand", ""))
            with c2:
                new_type = st.text_input("Type", value=brand_data.get("type", ""))
            with c3:
                new_color = st.text_input("Color (hex)", value=brand_data.get("color", "#DC2626"))
            new_note = st.text_area("Conversion Strategy Note", value=brand_data.get("conversion_note", ""), height=70)
            if st.form_submit_button("Save Brand Details"):
                db["competitor_brands"][idx].update({
                    "brand": new_brand_name.strip(),
                    "type": new_type.strip(),
                    "color": new_color.strip() or "#DC2626",
                    "conversion_note": new_note.strip(),
                })
                save_codes_db(db)
                st.success("Brand details saved.")
                st.rerun()

    with col_del:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🗑 Delete Brand", key=f"del_brand_{idx}", type="secondary"):
            st.session_state[f"confirm_del_brand_{idx}"] = True

    if st.session_state.get(f"confirm_del_brand_{idx}"):
        brand_name = brand_data.get("brand", "this brand")
        st.warning(f"Delete **{brand_name}** and all {len(codes)} codes?")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Yes, delete", key=f"confirm_del_brand_yes_{idx}", type="primary"):
                db["competitor_brands"].pop(idx)
                save_codes_db(db)
                st.session_state.pop(f"confirm_del_brand_{idx}", None)
                st.success("Brand deleted.")
                st.rerun()
        with c2:
            if st.button("Cancel", key=f"confirm_del_brand_no_{idx}"):
                st.session_state.pop(f"confirm_del_brand_{idx}", None)
                st.rerun()

    st.markdown("**Codes**")
    if not codes:
        st.caption("No codes defined for this brand yet.")
    else:
        for i, sku in enumerate(codes):
            col_code, col_product, col_act = st.columns([2, 5, 1])
            with col_code:
                new_code = st.text_input("Code", value=sku["code"], key=f"comp_code_{idx}_{i}", label_visibility="collapsed")
            with col_product:
                new_product = st.text_input("Product Name", value=sku.get("product", ""), key=f"comp_prod_{idx}_{i}", label_visibility="collapsed")
            with col_act:
                save_col, del_col = st.columns(2)
                with save_col:
                    if st.button("💾", key=f"save_comp_{idx}_{i}", help="Save"):
                        db["competitor_brands"][idx]["codes"][i] = {
                            "code": new_code.strip().upper(),
                            "product": new_product.strip(),
                        }
                        save_codes_db(db)
                        st.success(f"{new_code.strip().upper()} saved.")
                        st.rerun()
                with del_col:
                    if st.button("🗑", key=f"del_comp_{idx}_{i}", help="Delete"):
                        db["competitor_brands"][idx]["codes"].pop(i)
                        save_codes_db(db)
                        st.rerun()

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


def _admin_misc():
    db = load_codes_db()

    st.markdown("#### Service Tier Codes")
    st.caption("Codes like S1–S6, B7–B10 that appear on invoices but do not represent oil products.")

    service_tiers = db.get("service_tiers", [])
    for i, item in enumerate(service_tiers):
        col_code, col_name, col_desc, col_act = st.columns([1, 2, 4, 1])
        with col_code:
            new_code = st.text_input("Code", value=item["code"], key=f"st_code_{i}", label_visibility="collapsed")
        with col_name:
            new_name = st.text_input("Name", value=item.get("name", ""), key=f"st_name_{i}", label_visibility="collapsed")
        with col_desc:
            new_desc = st.text_input("Description", value=item.get("description", ""), key=f"st_desc_{i}", label_visibility="collapsed")
        with col_act:
            save_col, del_col = st.columns(2)
            with save_col:
                if st.button("💾", key=f"save_st_{i}", help="Save"):
                    db["service_tiers"][i] = {"code": new_code.strip().upper(), "name": new_name.strip(), "description": new_desc.strip()}
                    save_codes_db(db)
                    st.rerun()
            with del_col:
                if st.button("🗑", key=f"del_st_{i}", help="Delete"):
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
    for i, item in enumerate(spec_flags):
        col_code, col_name, col_desc, col_act = st.columns([1, 2, 4, 1])
        with col_code:
            new_code = st.text_input("Code", value=item["code"], key=f"sf_code_{i}", label_visibility="collapsed")
        with col_name:
            new_name = st.text_input("Name", value=item.get("name", ""), key=f"sf_name_{i}", label_visibility="collapsed")
        with col_desc:
            new_desc = st.text_input("Description", value=item.get("description", ""), key=f"sf_desc_{i}", label_visibility="collapsed")
        with col_act:
            save_col, del_col = st.columns(2)
            with save_col:
                if st.button("💾", key=f"save_sf_{i}", help="Save"):
                    db["spec_flags"][i] = {"code": new_code.strip().upper(), "name": new_name.strip(), "description": new_desc.strip()}
                    save_codes_db(db)
                    st.rerun()
            with del_col:
                if st.button("🗑", key=f"del_sf_{i}", help="Delete"):
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
