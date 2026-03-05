import streamlit as st
import tempfile
import os
import plotly.graph_objects as go
from report_generator import (
    generate_report, parse_excel, fmt_currency, fmt_number,
    PRODUCT_DESCRIPTIONS, get_product_display_name,
)
from distribution_data import STATE_DISTRIBUTORS, DISTRIBUTOR_COLORS, ALL_DISTRIBUTORS

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "RP_Synthetic_Expert_Logo_Black_Text.png")
LOGO_SIDEBAR_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "RPMO_logo_BF_Outline.png")
LOGO_NEVER_SETTLE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "25-RYP-02147 Employee LinkedIn Thumbnails P1-6.jpg")

st.set_page_config(
    page_title="Royal Purple Partnership Hub",
    page_icon="👑",
    layout="wide",
)

with st.sidebar:
    if os.path.exists(LOGO_SIDEBAR_PATH):
        st.image(LOGO_SIDEBAR_PATH, use_container_width=True)
    st.markdown("")
    st.markdown("**The Royal Purple Partnership Hub**")
    st.caption("by ThrottlePro")
    st.markdown("---")

    nav = st.radio(
        "Navigation",
        ["Report Generator", "Distribution Map", "Product Reference"],
        label_visibility="collapsed",
    )

    st.markdown("---")
    if os.path.exists(LOGO_NEVER_SETTLE):
        st.image(LOGO_NEVER_SETTLE, use_container_width=True)
    st.caption("More Cars. More Loyalty. Less Stress.")


def page_header(title, subtitle):
    if os.path.exists(LOGO_PATH):
        col_logo, col_title = st.columns([1, 5])
        with col_logo:
            st.image(LOGO_PATH, width=180)
        with col_title:
            st.markdown(
                f"<h1 style='color:#4B2D8A; margin: 0; padding-top: 20px;'>{title}</h1>"
                f"<p style='color:#94A3B8; margin: 4px 0 0 0;'>{subtitle}</p>",
                unsafe_allow_html=True,
            )
    else:
        st.markdown(f"<h1 style='color:#4B2D8A;'>{title}</h1>", unsafe_allow_html=True)


if nav == "Distribution Map":
    page_header("Distribution Map", "ABE Consumer Distribution Territory Coverage")
    st.markdown("")

    dist_filter = st.multiselect(
        "Filter by Distributor",
        options=ALL_DISTRIBUTORS,
        default=[],
        help="Select distributors to highlight on the map. Leave empty to show all.",
    )

    codes = []
    states_list = []
    colors = []
    hover_texts = []
    for code, info in STATE_DISTRIBUTORS.items():
        dists = info["distributors"]
        if not dists:
            continue
        if dist_filter:
            matching = [d for d in dists if d in dist_filter]
            if not matching:
                continue
            primary = matching[0]
        else:
            primary = dists[0]

        codes.append(code)
        states_list.append(info["state"])
        colors.append(DISTRIBUTOR_COLORS.get(primary, "#808080"))

        dist_list = "<br>".join(f"• {d}" for d in dists)
        hover_texts.append(f"<b>{info['state']}</b><br>{dist_list}")

    fig = go.Figure(data=go.Choropleth(
        locations=codes,
        z=[list(DISTRIBUTOR_COLORS.values()).index(c) if c in DISTRIBUTOR_COLORS.values() else 0 for c in colors],
        locationmode="USA-states",
        colorscale=[[i / max(len(DISTRIBUTOR_COLORS) - 1, 1), c] for i, c in enumerate(DISTRIBUTOR_COLORS.values())],
        showscale=False,
        text=hover_texts,
        hoverinfo="text",
        marker_line_color="white",
        marker_line_width=1.5,
    ))

    fig.update_layout(
        geo=dict(
            scope="usa",
            bgcolor="rgba(0,0,0,0)",
            lakecolor="#F8F5FF",
            landcolor="#E8E0F0",
            showlakes=True,
        ),
        margin=dict(l=0, r=0, t=0, b=0),
        height=500,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )

    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### ABE Legend")
    legend_cols = st.columns(min(len(DISTRIBUTOR_COLORS), 3))
    for i, (dist_name, color) in enumerate(DISTRIBUTOR_COLORS.items()):
        with legend_cols[i % 3]:
            state_count = sum(1 for s in STATE_DISTRIBUTORS.values() if dist_name in s["distributors"])
            st.markdown(
                f"<div style='display:flex;align-items:center;gap:8px;margin-bottom:8px;'>"
                f"<div style='width:20px;height:20px;background:{color};border-radius:3px;flex-shrink:0;'></div>"
                f"<span><b>{dist_name}</b> ({state_count} states)</span>"
                f"</div>",
                unsafe_allow_html=True,
            )

    st.markdown("---")
    st.markdown("### State Details")

    selected_state = st.selectbox(
        "Select a state for details",
        options=[""] + [f"{info['state']} ({code})" for code, info in sorted(STATE_DISTRIBUTORS.items(), key=lambda x: x[1]["state"])],
        format_func=lambda x: "Choose a state..." if x == "" else x,
    )

    if selected_state and selected_state != "":
        state_code = selected_state.split("(")[-1].rstrip(")")
        info = STATE_DISTRIBUTORS.get(state_code, {})
        if info:
            st.markdown(f"#### {info['state']}")
            if info["distributors"]:
                for dist in info["distributors"]:
                    color = DISTRIBUTOR_COLORS.get(dist, "#808080")
                    st.markdown(
                        f"<div style='display:flex;align-items:center;gap:8px;padding:8px 12px;margin-bottom:4px;"
                        f"background:linear-gradient(90deg, {color}22, transparent);border-left:4px solid {color};border-radius:4px;'>"
                        f"<b>{dist}</b></div>",
                        unsafe_allow_html=True,
                    )
            else:
                st.info("No distributor assigned to this state.")

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
            help="Each worksheet should represent one installer location with data columns for products, revenue, invoices, etc. The app auto-detects column layouts.",
        )
    with col_upload2:
        map_files = st.file_uploader(
            "Upload Distribution Maps (optional)",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True,
            help="Upload ABE distribution territory maps to include as slides in the report.",
        )

    if map_files:
        st.caption(f"{len(map_files)} map image(s) uploaded — will be included before store deep dives.")
        with st.expander("Preview Maps"):
            map_cols = st.columns(min(len(map_files), 3))
            for i, mf in enumerate(map_files):
                with map_cols[i % 3]:
                    st.image(mf, caption=mf.name, use_container_width=True)

    if uploaded_file is not None:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            stores, month_year = parse_excel(tmp_path)

            st.success(f"Parsed **{len(stores)}** locations for **{month_year}**")

            total_rev = sum(s["totalRevenue"] for s in stores)
            total_inv = sum(s["invoices"] for s in stores)
            avg_rev = total_rev / total_inv if total_inv else 0
            total_veh = sum(s["vehicles"] for s in stores)

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Revenue", fmt_currency(total_rev))
            col2.metric("Oil Changes", fmt_number(total_inv))
            col3.metric("Avg Rev/Invoice", f"${avg_rev:.2f}")
            col4.metric("Unique Vehicles", fmt_number(total_veh))

            st.markdown("")

            tab_rankings, tab_details = st.tabs(["Store Rankings", "Store Details"])

            with tab_rankings:
                ranking_data = []
                for s in stores:
                    pct = s["totalRevenue"] / total_rev * 100 if total_rev else 0
                    ranking_data.append({
                        "Rank": s["rank"],
                        "Store": s["name"],
                        "Revenue": fmt_currency(s["totalRevenue"]),
                        "Oil Changes": fmt_number(s["invoices"]),
                        "Avg Rev/Inv": f"${s['avgRevPerInvoice']:.2f}",
                        "Share": f"{pct:.1f}%",
                        "Top Product": s["topProduct"],
                    })
                st.dataframe(ranking_data, use_container_width=True, hide_index=True)

            with tab_details:
                for s in stores:
                    with st.expander(f"#{s['rank']} — {s['name']}"):
                        dc1, dc2, dc3 = st.columns(3)
                        dc1.metric("Revenue", fmt_currency(s["totalRevenue"]))
                        dc2.metric("Oil Changes", fmt_number(s["invoices"]))
                        dc3.metric("Avg Rev/Inv", f"${s['avgRevPerInvoice']:.2f}")
                        if s["productBreakdown"]:
                            st.caption("Top Products:")
                            for pb in s["productBreakdown"][:5]:
                                display_name = get_product_display_name(pb["code"])
                                st.text(f"  {display_name} ({pb['category']}) — {fmt_currency(pb['revenue'])}")

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

                        map_note = f" + {len(map_temp_paths)} distribution map(s)" if map_temp_paths else ""
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
        st.info("Upload an Excel file to get started. The app auto-detects column layouts, so reports from any Royal Purple region or distributor should work.")
