import streamlit as st
import tempfile
import os
from report_generator import generate_report, parse_excel, fmt_currency, fmt_number

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "RP_Synthetic_Expert_Logo_Black_Text.png")
LOGO_SIDEBAR_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "RPMO_logo_BF_Outline.png")

st.set_page_config(
    page_title="Royal Purple Report Generator",
    page_icon="👑",
    layout="wide",
)

with st.sidebar:
    if os.path.exists(LOGO_SIDEBAR_PATH):
        st.image(LOGO_SIDEBAR_PATH, use_container_width=True)
    st.markdown("---")
    st.markdown("**Royal Purple**")
    st.markdown("Installer Program Report Generator")
    st.markdown("---")
    st.caption("Upload a monthly Excel report to generate a branded 23-slide PowerPoint deck with executive summary, store rankings, and individual store deep dives.")

if os.path.exists(LOGO_PATH):
    col_logo, col_title = st.columns([1, 4])
    with col_logo:
        st.image(LOGO_PATH, width=200)
    with col_title:
        st.markdown(
            "<h1 style='color:#4B2D8A; margin-top: 10px;'>Installer Report Generator</h1>",
            unsafe_allow_html=True,
        )
else:
    st.markdown(
        "<h1 style='color:#4B2D8A;'>Royal Purple Installer Report Generator</h1>",
        unsafe_allow_html=True,
    )

st.markdown(
    "<p style='color:#94A3B8;'>Upload a Royal Purple monthly Excel report to generate a branded PowerPoint presentation.</p>",
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader(
    "Upload Royal Purple Excel Report (.xlsx)",
    type=["xlsx"],
    help="One sheet per installer location + a Report Summary sheet",
)

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    try:
        stores, month_year = parse_excel(tmp_path)

        st.success(f"Parsed {len(stores)} stores for **{month_year}**")

        total_rev = sum(s["totalRevenue"] for s in stores)
        total_inv = sum(s["invoices"] for s in stores)
        avg_rev = total_rev / total_inv if total_inv else 0
        total_veh = sum(s["vehicles"] for s in stores)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Revenue", fmt_currency(total_rev))
        col2.metric("Oil Changes", fmt_number(total_inv))
        col3.metric("Avg Rev/Invoice", f"${avg_rev:.2f}")
        col4.metric("Unique Vehicles", fmt_number(total_veh))

        st.subheader("Store Rankings")
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

        if st.button("Generate PowerPoint Report", type="primary"):
            with st.spinner("Generating branded presentation..."):
                output_filename = f"Royal_Purple_Installer_Report_{month_year.replace(' ', '_')}.pptx"
                output_path = os.path.join(tempfile.gettempdir(), output_filename)
                generate_report(tmp_path, output_path)

                with open(output_path, "rb") as f:
                    pptx_data = f.read()

                st.success(f"Report generated successfully — {len(stores)} store deep dives included.")

                st.download_button(
                    label="Download PowerPoint Report",
                    data=pptx_data,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                )

                os.unlink(output_path)

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
    finally:
        os.unlink(tmp_path)
else:
    st.info("Upload an Excel file to get started. The file should contain one sheet per installer location and a 'Report Summary' sheet.")
