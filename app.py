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

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "RP_Synthetic_Expert_Logo_Black_Text.png")
LOGO_SIDEBAR_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "RPMO_logo_BF_Outline.png")
LOGO_NEVER_SETTLE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "25-RYP-02147 Employee LinkedIn Thumbnails P1-6.jpg")

st.set_page_config(
    page_title="Royal Purple Partnership Hub",
    page_icon="👑",
    layout="wide",
)

with st.sidebar:
    st.markdown("")
    st.markdown("**The Royal Purple Partnership Hub**")
    st.caption("by ThrottlePro")
    st.markdown("---")

    nav = st.radio(
        "Navigation",
        ["Report Generator", "Customer Map", "Product Reference"],
        label_visibility="collapsed",
    )

    st.markdown("---")
    if os.path.exists(LOGO_NEVER_SETTLE):
        st.image(LOGO_NEVER_SETTLE, width="stretch")
    st.caption("More Cars. More Loyalty. Less Stress.")


def page_header(title, subtitle):
    st.markdown(
        f"<h1 style='color:#4B2D8A; margin: 0;'>{title}</h1>"
        f"<p style='color:#94A3B8; margin: 4px 0 0 0;'>{subtitle}</p>",
        unsafe_allow_html=True,
    )


if nav == "Customer Map":
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

        installer_types = ["Promo Only (Not on C4C)", "On Both Lists", "C4C Only", "Rack Installer"]
        installer_total = sum(type_counts.get(t, 0) for t in installer_types)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Locations", len(all_map_data))
        col2.metric("Installer Accounts", installer_total)
        col3.metric("Distributors", type_counts.get("Distributor", 0))
        col4.metric("Powersports", type_counts.get("Powersports/Motorsports", 0))

        col5, col6, col7, col8 = st.columns(4)
        col5.metric("Promo Only", type_counts.get("Promo Only (Not on C4C)", 0))
        col6.metric("On Both Lists", type_counts.get("On Both Lists", 0))
        col7.metric("C4C Only", type_counts.get("C4C Only", 0))
        col8.metric("Rack Installer", type_counts.get("Rack Installer", 0))
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

                    st.success(
                        f"Report generated — {stats['sheets']} sheets: "
                        f"{stats['not_on_c4c']} not on C4C, {stats['c4c_matched']} matched, "
                        f"{stats['states']} US states, {stats['counties']} counties."
                    )

                    st.download_button(
                        label="Download C4C Report",
                        data=report_data,
                        file_name="RP_C4C_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
    else:
        st.info("No customer data available. Upload a CSV file to get started.")

elif nav == "Product Reference":
    page_header("Product Reference", "Royal Purple product categories and descriptions")
    st.markdown("")
    for cat, desc in PRODUCT_DESCRIPTIONS.items():
        with st.container():
            st.markdown(f"#### {cat}")
            st.write(desc)
            st.markdown("---")

elif nav == "Report Generator":
    page_header("Installer Report Generator", "Upload a Royal Purple monthly report to generate a branded PowerPoint presentation.")
    st.markdown("")

    col_upload1, col_upload2 = st.columns(2)
    with col_upload1:
        uploaded_file = st.file_uploader(
            "Upload Royal Purple Excel Report (.xlsx)",
            type=["xlsx"],
            help="The app auto-detects columns, deduplicates multi-product invoices, and computes corrected revenue.",
        )
    with col_upload2:
        map_files = st.file_uploader(
            "Upload Map Images (optional)",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True,
            help="Upload map images to include as slides in the report.",
        )

    if map_files:
        st.caption(f"{len(map_files)} map image(s) uploaded — will be included before store deep dives.")
        with st.expander("Preview Maps"):
            map_cols = st.columns(min(len(map_files), 3))
            for i, mf in enumerate(map_files):
                with map_cols[i % 3]:
                    st.image(mf, caption=mf.name, width="stretch")

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

            dedup_note = f" (deduplicated from {fmt_number(total_raw)} raw lines)" if total_raw > total_inv else ""
            st.success(f"Parsed **{len(stores)}** locations for **{month_year}**{dedup_note}")

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Revenue", fmt_currency(total_rev))
            col2.metric("Unique Invoices", fmt_number(total_inv))
            col3.metric("Avg Rev/Invoice", f"${avg_rev:.2f}")
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

                st.markdown("### Max-Clean Attachment Analysis")
                st.caption(
                    "The RP export only shows Royal Purple products. 'Solo' Max-Clean lines represent "
                    "non-RP oil changes (Castrol, conventional, etc.) where Max-Clean was added as an upsell."
                )

                mc1, mc2, mc3, mc4 = st.columns(4)
                mc1.metric("MC Invoices", fmt_number(network_mc), f"{mc_pct:.1f}% attachment rate")
                mc2.metric("MC Avg Ticket", f"${mc_avg:.2f}")
                mc3.metric("Non-MC Avg Ticket", f"${non_mc_avg:.2f}")
                mc4.metric("MC Ticket Lift", f"+${network_lift:.2f}", f"+{network_lift/non_mc_avg*100:.1f}%" if non_mc_avg else "")

                mc_with_rp = sum(s.get("maxClean", {}).get("withRpOil", 0) for s in stores)
                mc_non_rp = sum(s.get("maxClean", {}).get("withNonRpOil", 0) for s in stores)
                st.markdown("")
                mc_solo = sum(s.get("maxClean", {}).get("soloInData", 0) for s in stores)
                bk1, bk2, bk3 = st.columns(3)
                bk1.metric("MC + RP Oil", fmt_number(mc_with_rp), f"{mc_with_rp/network_mc*100:.1f}%" if network_mc else "")
                bk2.metric("MC + Non-RP Oil", fmt_number(mc_non_rp), f"{mc_non_rp/network_mc*100:.1f}%" if network_mc else "")
                bk3.metric("MC Solo (Non-RP OC)", fmt_number(mc_solo), f"{mc_solo/network_mc*100:.1f}%" if network_mc else "")

            st.markdown("")

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
                        f"selling the RP additive upsell even on conventional oil changes."
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

            st.markdown("")
            if st.button("Generate PowerPoint Report", type="primary"):
                map_temp_paths = []
                try:
                    for mf in (map_files or []):
                        ext = os.path.splitext(mf.name)[1] or ".png"
                        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as mtmp:
                            mtmp.write(mf.getvalue())
                            raw_name = os.path.splitext(mf.name)[0].replace("_", " ").replace("-", " ")
                            map_temp_paths.append({
                                "path": mtmp.name,
                                "title": raw_name,
                            })

                    with st.spinner("Generating branded presentation..."):
                        output_filename = f"Royal_Purple_Partnership_Report_{month_year.replace(' ', '_')}.pptx"
                        output_path = os.path.join(tempfile.gettempdir(), output_filename)
                        generate_report(tmp_path, output_path, map_images=map_temp_paths if map_temp_paths else None)

                        with open(output_path, "rb") as f:
                            pptx_data = f.read()

                        map_note = f" + {len(map_temp_paths)} map image(s)" if map_temp_paths else ""
                        st.success(f"Report generated — {len(stores)} store deep dives{map_note} included.")

                        st.download_button(
                            label="Download PowerPoint Report",
                            data=pptx_data,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="primary",
                        )

                        os.unlink(output_path)
                finally:
                    for mp in map_temp_paths:
                        try:
                            os.unlink(mp["path"])
                        except OSError:
                            pass

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
