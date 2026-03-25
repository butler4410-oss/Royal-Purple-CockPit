import streamlit as st
from profit_pdf import generate_profit_pdf

def render():
    st.markdown(
        """<div style="font-size:10px;font-weight:700;letter-spacing:2.5px;color:#9CA3AF;
                    text-transform:uppercase;margin-bottom:12px;">Installer Profit Worksheet</div>""",
        unsafe_allow_html=True,
    )

    col_left, col_right = st.columns([1, 1], gap="large")

    with col_right:
        st.markdown("##### Installer Information")
        installer_name = st.text_input("Customer / Installer Name", value="", key="pc_name")
        r1a, r1b = st.columns(2)
        ocpd = r1a.number_input("Avg Oil Changes Per Day", min_value=1, value=30, step=1, key="pc_ocpd")
        conversion_pct = r1b.number_input("% Converting to Royal Purple", min_value=1, max_value=100, value=10, step=1, key="pc_conv")
        r2a, r2b = st.columns(2)
        gallons_per = r2a.number_input("Gallons Per Oil Change", min_value=0.5, value=1.25, step=0.25, format="%.2f", key="pc_gal")
        days_open = r2b.number_input("Number of Days Open / Year", min_value=1, value=310, step=1, key="pc_days")
        num_locations = st.number_input("Number of Locations", min_value=1, value=1, step=1, key="pc_locs")

        st.markdown("---")
        st.markdown("##### Royal Purple Pricing")
        rp_product = st.text_input("RP Product", value="Royal Purple HP 5W-30", key="pc_rp_prod")
        rp_distributor = st.text_input("Distributor", value="", key="pc_rp_dist")
        rp_selling_price = st.number_input("Suggested RP Selling Price ($)", min_value=0.0, value=0.0, step=1.0, format="%.2f", key="pc_rp_sell")

        rp_pkg = st.selectbox("RP Package Size", ["Bulk", "Drum", "Bag-n-Box", "5 Qt.", "1 Qt.", "1 Gallon"], index=2, key="pc_rp_pkg")
        rp_prices = {}
        st.markdown('<p style="font-size:13px;color:#6B7280;margin-bottom:4px;">RP Distributor Pricing (per gallon conversion):</p>', unsafe_allow_html=True)
        rp_c1, rp_c2, rp_c3 = st.columns(3)
        rp_prices["Bulk"] = rp_c1.number_input("Bulk ($/gal)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="pc_rp_bulk")
        rp_prices["Drum"] = rp_c2.number_input("Drum ($/gal)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="pc_rp_drum")
        rp_prices["Bag-n-Box"] = rp_c3.number_input("BnB ($/gal)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="pc_rp_bnb")
        rp_c4, rp_c5, rp_c6 = st.columns(3)
        rp_prices["5 Qt."] = rp_c4.number_input("5Qt ($/gal)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="pc_rp_5qt")
        rp_prices["1 Qt."] = rp_c5.number_input("1Qt ($/gal)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="pc_rp_1qt")
        rp_prices["1 Gallon"] = rp_c6.number_input("1Gal ($/gal)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="pc_rp_1gal")

        st.markdown("---")
        st.markdown("##### Current Top-Selling Brand")
        comp_brand = st.text_input("Current Brand", value="Mobil 1", key="pc_comp_brand")
        comp_product = st.text_input("Top-Selling Product", value="", key="pc_comp_prod")
        comp_selling_price = st.number_input("Current Selling Price ($)", min_value=0.0, value=0.0, step=1.0, format="%.2f", key="pc_comp_sell")

        comp_pkg = st.selectbox("Competitor Package Size", ["Bulk", "Drum", "Bag-n-Box", "5 Qt.", "1 Qt.", "1 Gallon"], index=0, key="pc_comp_pkg")
        comp_prices = {}
        st.markdown('<p style="font-size:13px;color:#6B7280;margin-bottom:4px;">Competitor Distributor Pricing (per gallon conversion):</p>', unsafe_allow_html=True)
        cp_c1, cp_c2, cp_c3 = st.columns(3)
        comp_prices["Bulk"] = cp_c1.number_input("Bulk ($/gal)", min_value=0.0, value=25.50, step=0.5, format="%.2f", key="pc_cp_bulk")
        comp_prices["Drum"] = cp_c2.number_input("Drum ($/gal)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="pc_cp_drum")
        comp_prices["Bag-n-Box"] = cp_c3.number_input("BnB ($/gal)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="pc_cp_bnb")
        cp_c4, cp_c5, cp_c6 = st.columns(3)
        comp_prices["5 Qt."] = cp_c4.number_input("5Qt ($/gal)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="pc_cp_5qt")
        comp_prices["1 Qt."] = cp_c5.number_input("1Qt ($/gal)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="pc_cp_1qt")
        comp_prices["1 Gallon"] = cp_c6.number_input("1Gal ($/gal)", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="pc_cp_1gal")

    conversion_rate = conversion_pct / 100.0
    total_oil_changes = ocpd * days_open
    rp_converting = total_oil_changes * conversion_rate

    rp_fluid_cost = gallons_per * rp_prices.get(rp_pkg, 0)
    rp_gross_profit = rp_selling_price - rp_fluid_cost

    comp_fluid_cost = gallons_per * comp_prices.get(comp_pkg, 0)
    comp_gross_profit = comp_selling_price - comp_fluid_cost

    incremental_per_service = rp_gross_profit - comp_gross_profit
    annual_per_location = incremental_per_service * rp_converting
    total_annual = annual_per_location * num_locations

    with col_left:
        st.markdown("##### Results Dashboard")
        if installer_name:
            st.markdown(f'<div style="font-size:15px;font-weight:600;color:#4B2D8A;margin-bottom:12px;">{installer_name}</div>', unsafe_allow_html=True)

        st.markdown(
            f"""<div style="background:#F3E8FF;border-radius:10px;padding:16px 20px;margin-bottom:16px;">
                <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;color:#7C3AED;text-transform:uppercase;">Volume Overview</div>
                <div style="display:flex;gap:24px;margin-top:8px;">
                    <div><div style="font-size:24px;font-weight:800;color:#1F2937;">{ocpd}</div><div style="font-size:11px;color:#6B7280;">Oil Changes/Day</div></div>
                    <div><div style="font-size:24px;font-weight:800;color:#1F2937;">{total_oil_changes:,}</div><div style="font-size:11px;color:#6B7280;">Annual Oil Changes</div></div>
                    <div><div style="font-size:24px;font-weight:800;color:#4B2D8A;">{rp_converting:,.0f}</div><div style="font-size:11px;color:#6B7280;">Converting to RP ({conversion_pct}%)</div></div>
                </div>
            </div>""",
            unsafe_allow_html=True,
        )

        rp_col, comp_col = st.columns(2)
        with rp_col:
            st.markdown(
                f"""<div style="background:#ECFDF5;border:2px solid #059669;border-radius:10px;padding:16px;">
                    <div style="font-size:12px;font-weight:700;color:#059669;text-transform:uppercase;letter-spacing:1px;">Royal Purple</div>
                    <div style="font-size:12px;color:#6B7280;margin-top:4px;">{rp_product}</div>
                    <div style="margin-top:12px;">
                        <div style="font-size:11px;color:#6B7280;">Selling Price</div>
                        <div style="font-size:20px;font-weight:700;color:#1F2937;">${rp_selling_price:,.2f}</div>
                    </div>
                    <div style="margin-top:8px;">
                        <div style="font-size:11px;color:#6B7280;">Fluid Cost ({rp_pkg})</div>
                        <div style="font-size:20px;font-weight:700;color:#DC2626;">${rp_fluid_cost:,.2f}</div>
                    </div>
                    <div style="margin-top:8px;padding-top:8px;border-top:1px solid #D1FAE5;">
                        <div style="font-size:11px;color:#6B7280;">Gross Profit / Service</div>
                        <div style="font-size:22px;font-weight:800;color:#059669;">${rp_gross_profit:,.2f}</div>
                    </div>
                </div>""",
                unsafe_allow_html=True,
            )
        with comp_col:
            st.markdown(
                f"""<div style="background:#FEF2F2;border:2px solid #DC2626;border-radius:10px;padding:16px;">
                    <div style="font-size:12px;font-weight:700;color:#DC2626;text-transform:uppercase;letter-spacing:1px;">{comp_brand}</div>
                    <div style="font-size:12px;color:#6B7280;margin-top:4px;">{comp_product or 'Current Top Brand'}</div>
                    <div style="margin-top:12px;">
                        <div style="font-size:11px;color:#6B7280;">Selling Price</div>
                        <div style="font-size:20px;font-weight:700;color:#1F2937;">${comp_selling_price:,.2f}</div>
                    </div>
                    <div style="margin-top:8px;">
                        <div style="font-size:11px;color:#6B7280;">Fluid Cost ({comp_pkg})</div>
                        <div style="font-size:20px;font-weight:700;color:#DC2626;">${comp_fluid_cost:,.2f}</div>
                    </div>
                    <div style="margin-top:8px;padding-top:8px;border-top:1px solid #FEE2E2;">
                        <div style="font-size:11px;color:#6B7280;">Gross Profit / Service</div>
                        <div style="font-size:22px;font-weight:800;color:#DC2626;">${comp_gross_profit:,.2f}</div>
                    </div>
                </div>""",
                unsafe_allow_html=True,
            )

        st.markdown("")

        profit_color = "#059669" if incremental_per_service >= 0 else "#DC2626"
        arrow = "&#9650;" if incremental_per_service >= 0 else "&#9660;"
        st.markdown(
            f"""<div style="background:linear-gradient(135deg,#2D1B5E 0%,#4B2D8A 100%);border-radius:12px;padding:20px 24px;color:white;">
                <div style="font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#C4B5E8;margin-bottom:12px;">Incremental Profitability</div>
                <div style="display:flex;gap:24px;flex-wrap:wrap;">
                    <div style="flex:1;min-width:140px;">
                        <div style="font-size:11px;color:#C4B5E8;">Per Service</div>
                        <div style="font-size:26px;font-weight:800;">{arrow} ${abs(incremental_per_service):,.2f}</div>
                    </div>
                    <div style="flex:1;min-width:140px;">
                        <div style="font-size:11px;color:#C4B5E8;">Annual / Location</div>
                        <div style="font-size:26px;font-weight:800;">${annual_per_location:,.2f}</div>
                    </div>
                </div>
                <div style="margin-top:16px;padding-top:12px;border-top:1px solid rgba(255,255,255,0.2);">
                    <div style="display:flex;justify-content:space-between;align-items:center;">
                        <div>
                            <div style="font-size:11px;color:#C4B5E8;">{num_locations} Location{'s' if num_locations > 1 else ''} — Total Annual Profitability</div>
                            <div style="font-size:32px;font-weight:800;color:#C8A951;">${total_annual:,.2f}</div>
                        </div>
                    </div>
                </div>
            </div>""",
            unsafe_allow_html=True,
        )

        st.markdown("")
        st.markdown(
            f"""<div style="background:#FFFBEB;border:1px solid #F59E0B;border-radius:8px;padding:12px 16px;font-size:12px;color:#92400E;">
                <strong>Key Takeaway:</strong> By converting just {conversion_pct}% of oil changes to Royal Purple,
                {'each location gains' if num_locations > 1 else 'this location gains'}
                <strong>${annual_per_location:,.2f}</strong> in additional annual profit
                {f'across <strong>{num_locations} locations</strong> for a total of <strong>${total_annual:,.2f}</strong>' if num_locations > 1 else ''}.
            </div>""",
            unsafe_allow_html=True,
        )

        st.markdown("")
        pdf_data = {
            "installer_name": installer_name,
            "ocpd": ocpd,
            "conversion_pct": conversion_pct,
            "gallons_per": gallons_per,
            "days_open": days_open,
            "num_locations": num_locations,
            "rp_product": rp_product,
            "rp_distributor": rp_distributor,
            "rp_selling_price": rp_selling_price,
            "rp_pkg": rp_pkg,
            "rp_prices": rp_prices,
            "comp_brand": comp_brand,
            "comp_product": comp_product,
            "comp_selling_price": comp_selling_price,
            "comp_pkg": comp_pkg,
            "comp_prices": comp_prices,
            "total_oil_changes": total_oil_changes,
            "rp_converting": rp_converting,
            "rp_fluid_cost": rp_fluid_cost,
            "rp_gross_profit": rp_gross_profit,
            "comp_fluid_cost": comp_fluid_cost,
            "comp_gross_profit": comp_gross_profit,
            "incremental_per_service": incremental_per_service,
            "annual_per_location": annual_per_location,
            "total_annual": total_annual,
        }
        pdf_bytes = generate_profit_pdf(pdf_data)
        filename = f"Incremental_Profitability_Report_{installer_name.replace(' ', '_') or 'RP'}.pdf"
        st.download_button(
            label="Download Incremental Profitability Report",
            data=pdf_bytes,
            file_name=filename,
            mime="application/pdf",
            use_container_width=True,
        )
