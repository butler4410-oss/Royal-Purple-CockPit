import io
import os
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF


PURPLE_DARK = HexColor("#2D1B5E")
PURPLE_MID = HexColor("#4B2D8A")
PURPLE_LIGHT = HexColor("#F3E8FF")
GREEN = HexColor("#059669")
GREEN_BG = HexColor("#ECFDF5")
RED = HexColor("#DC2626")
RED_BG = HexColor("#FEF2F2")
GOLD = HexColor("#C8A951")
GRAY = HexColor("#6B7280")
GRAY_LIGHT = HexColor("#F9FAFB")
AMBER_BG = HexColor("#FFFBEB")
AMBER_BORDER = HexColor("#F59E0B")
AMBER_TEXT = HexColor("#92400E")

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "RP_Synthetic_Expert_Logo_Black_Text.png")


def _styles():
    ss = getSampleStyleSheet()
    ss.add(ParagraphStyle("Title_RP", parent=ss["Title"], fontName="Helvetica-Bold",
                          fontSize=18, textColor=PURPLE_DARK, spaceAfter=4, alignment=TA_LEFT))
    ss.add(ParagraphStyle("Subtitle_RP", parent=ss["Normal"], fontName="Helvetica",
                          fontSize=10, textColor=GRAY, spaceAfter=12, alignment=TA_LEFT))
    ss.add(ParagraphStyle("Section", parent=ss["Normal"], fontName="Helvetica-Bold",
                          fontSize=11, textColor=PURPLE_DARK, spaceBefore=14, spaceAfter=6))
    ss.add(ParagraphStyle("Label", parent=ss["Normal"], fontName="Helvetica",
                          fontSize=8, textColor=GRAY))
    ss.add(ParagraphStyle("Value", parent=ss["Normal"], fontName="Helvetica-Bold",
                          fontSize=10, textColor=black))
    ss.add(ParagraphStyle("ValueBig", parent=ss["Normal"], fontName="Helvetica-Bold",
                          fontSize=14, textColor=PURPLE_DARK))
    ss.add(ParagraphStyle("ValueGreen", parent=ss["Normal"], fontName="Helvetica-Bold",
                          fontSize=12, textColor=GREEN))
    ss.add(ParagraphStyle("ValueRed", parent=ss["Normal"], fontName="Helvetica-Bold",
                          fontSize=12, textColor=RED))
    ss.add(ParagraphStyle("ValueGold", parent=ss["Normal"], fontName="Helvetica-Bold",
                          fontSize=16, textColor=GOLD))
    ss.add(ParagraphStyle("Footer", parent=ss["Normal"], fontName="Helvetica",
                          fontSize=7, textColor=GRAY, alignment=TA_CENTER))
    ss.add(ParagraphStyle("Takeaway", parent=ss["Normal"], fontName="Helvetica",
                          fontSize=9, textColor=AMBER_TEXT, leading=13))
    ss.add(ParagraphStyle("WhiteLabel", parent=ss["Normal"], fontName="Helvetica",
                          fontSize=8, textColor=HexColor("#C4B5E8")))
    ss.add(ParagraphStyle("WhiteValue", parent=ss["Normal"], fontName="Helvetica-Bold",
                          fontSize=14, textColor=white))
    ss.add(ParagraphStyle("WhiteValueBig", parent=ss["Normal"], fontName="Helvetica-Bold",
                          fontSize=18, textColor=GOLD))
    return ss


def _header_table(data, ss):
    installer_name = data["installer_name"] or "Incremental Profitability Report"
    date_str = datetime.now().strftime("%B %d, %Y")

    header_data = []
    if os.path.exists(LOGO_PATH):
        logo = Image(LOGO_PATH, width=1.3 * inch, height=0.55 * inch)
        header_data.append([
            logo,
            Paragraph(f"<b>{installer_name}</b>", ss["Title_RP"]),
        ])
    else:
        header_data.append([
            Paragraph("Royal Purple", ss["Title_RP"]),
            Paragraph(f"<b>{installer_name}</b>", ss["Title_RP"]),
        ])

    t = Table(header_data, colWidths=[1.6 * inch, 5.5 * inch])
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))

    elements = [t]
    elements.append(Paragraph(f"Incremental Profitability Report  |  Generated {date_str}", ss["Subtitle_RP"]))

    d = Drawing(7.1 * inch, 2)
    d.add(Rect(0, 0, 7.1 * inch, 2, fillColor=PURPLE_DARK, strokeColor=None))
    elements.append(d)
    elements.append(Spacer(1, 8))
    return elements


def _volume_table(data, ss):
    ocpd = data["ocpd"]
    total_oc = data["total_oil_changes"]
    rp_conv = data["rp_converting"]
    conv_pct = data["conversion_pct"]

    elements = [Paragraph("VOLUME OVERVIEW", ss["Section"])]

    tbl_data = [[
        [Paragraph("Oil Changes / Day", ss["Label"]), Paragraph(f"{ocpd:,}", ss["ValueBig"])],
        [Paragraph("Annual Oil Changes", ss["Label"]), Paragraph(f"{total_oc:,}", ss["ValueBig"])],
        [Paragraph(f"Converting to RP ({conv_pct}%)", ss["Label"]), Paragraph(f"{rp_conv:,.0f}", ss["ValueBig"])],
        [Paragraph("Gallons / Oil Change", ss["Label"]), Paragraph(f"{data['gallons_per']:.2f}", ss["ValueBig"])],
    ]]

    t = Table(tbl_data, colWidths=[1.775 * inch] * 4)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), PURPLE_LIGHT),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("BOX", (0, 0), (-1, -1), 0.5, PURPLE_MID),
        ("ROUNDEDCORNERS", [6, 6, 6, 6]),
    ]))
    elements.append(t)
    return elements


def _comparison_table(data, ss):
    elements = [Paragraph("PROFITABILITY COMPARISON", ss["Section"])]

    rp_label = Paragraph("ROYAL PURPLE", ParagraphStyle("rphead", parent=ss["Label"],
                         fontName="Helvetica-Bold", fontSize=9, textColor=GREEN))
    comp_label = Paragraph(data["comp_brand"].upper(), ParagraphStyle("comphead", parent=ss["Label"],
                           fontName="Helvetica-Bold", fontSize=9, textColor=RED))

    rp_product_p = Paragraph(data["rp_product"], ss["Label"])
    comp_product_p = Paragraph("Current Top-Selling Brand", ss["Label"])

    rows = [
        ["", rp_label, comp_label],
        ["Product", rp_product_p, comp_product_p],
        ["Package Size", Paragraph(data["rp_pkg"], ss["Value"]), Paragraph(data["comp_pkg"], ss["Value"])],
        ["Selling Price", Paragraph(f"${data['rp_selling_price']:,.2f}", ss["Value"]),
         Paragraph(f"${data['comp_selling_price']:,.2f}", ss["Value"])],
        ["Fluid Cost / Service", Paragraph(f"${data['rp_fluid_cost']:,.2f}", ss["ValueRed"]),
         Paragraph(f"${data['comp_fluid_cost']:,.2f}", ss["ValueRed"])],
        ["Gross Profit / Service", Paragraph(f"${data['rp_gross_profit']:,.2f}", ss["ValueGreen"]),
         Paragraph(f"${data['comp_gross_profit']:,.2f}", ss["ValueRed"])],
    ]

    if data.get("rp_distributor"):
        rows.insert(2, ["Distributor", Paragraph(data["rp_distributor"], ss["Value"]), Paragraph("—", ss["Label"])])

    t = Table(rows, colWidths=[1.8 * inch, 2.65 * inch, 2.65 * inch])
    style_cmds = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#E5E7EB")),
        ("BACKGROUND", (0, 0), (0, -1), GRAY_LIGHT),
        ("BACKGROUND", (1, 0), (1, 0), GREEN_BG),
        ("BACKGROUND", (2, 0), (2, 0), RED_BG),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (0, -1), 8),
        ("TEXTCOLOR", (0, 0), (0, -1), GRAY),
        ("LINEBELOW", (0, -1), (-1, -1), 1.5, PURPLE_MID),
    ]
    t.setStyle(TableStyle(style_cmds))
    elements.append(t)
    return elements


def _profitability_block(data, ss):
    elements = [Spacer(1, 8)]

    inc = data["incremental_per_service"]
    arrow = "\u25b2" if inc >= 0 else "\u25bc"
    annual_loc = data["annual_per_location"]
    total = data["total_annual"]
    locs = data["num_locations"]

    rows = [
        [
            [Paragraph("INCREMENTAL PROFITABILITY", ParagraphStyle("s", parent=ss["WhiteLabel"],
             fontName="Helvetica-Bold", fontSize=9, textColor=HexColor("#C4B5E8")))],
            "",
            "",
        ],
        [
            [Paragraph("Per Service", ss["WhiteLabel"]),
             Paragraph(f"{arrow} ${abs(inc):,.2f}", ss["WhiteValue"])],
            [Paragraph("Annual / Location", ss["WhiteLabel"]),
             Paragraph(f"${annual_loc:,.2f}", ss["WhiteValue"])],
            [Paragraph(f"Days Open / Year", ss["WhiteLabel"]),
             Paragraph(f"{data['days_open']}", ss["WhiteValue"])],
        ],
        [
            [Paragraph(f"{locs} Location{'s' if locs > 1 else ''} — Total Annual Profitability", ss["WhiteLabel"]),
             Paragraph(f"${total:,.2f}", ss["WhiteValueBig"])],
            "",
            "",
        ],
    ]

    t = Table(rows, colWidths=[2.37 * inch] * 3)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), PURPLE_DARK),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("SPAN", (0, 0), (2, 0)),
        ("SPAN", (0, 2), (2, 2)),
        ("LINEBELOW", (0, 1), (-1, 1), 0.5, HexColor("#FFFFFF33")),
        ("BOX", (0, 0), (-1, -1), 1, PURPLE_MID),
        ("ROUNDEDCORNERS", [8, 8, 8, 8]),
    ]))
    elements.append(t)
    return elements


def _takeaway(data, ss):
    conv_pct = data["conversion_pct"]
    annual_loc = data["annual_per_location"]
    locs = data["num_locations"]
    total = data["total_annual"]

    multi = locs > 1
    msg = (
        f"<b>Key Takeaway:</b> By converting just {conv_pct}% of oil changes to Royal Purple, "
        f"{'each location gains' if multi else 'this location gains'} "
        f"<b>${annual_loc:,.2f}</b> in additional annual profit"
    )
    if multi:
        msg += f" across <b>{locs} locations</b> for a total of <b>${total:,.2f}</b>"
    msg += "."

    elements = [Spacer(1, 10)]

    tbl = Table([[Paragraph(msg, ss["Takeaway"])]], colWidths=[7.1 * inch])
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), AMBER_BG),
        ("BOX", (0, 0), (-1, -1), 1, AMBER_BORDER),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("ROUNDEDCORNERS", [6, 6, 6, 6]),
    ]))
    elements.append(tbl)
    return elements


def _pricing_detail(data, ss):
    elements = [Spacer(1, 6), Paragraph("DISTRIBUTOR PRICING DETAIL (per gallon)", ss["Section"])]

    pkg_names = ["Bulk", "Drum", "Bag-n-Box", "5 Qt.", "1 Qt.", "1 Gallon"]
    header = ["Package"] + pkg_names
    rp_row = ["Royal Purple"] + [f"${data['rp_prices'].get(p, 0):,.2f}" for p in pkg_names]
    comp_row = [data["comp_brand"]] + [f"${data['comp_prices'].get(p, 0):,.2f}" for p in pkg_names]

    tbl = Table([header, rp_row, comp_row], colWidths=[1.2 * inch] + [0.98 * inch] * 6)
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("TEXTCOLOR", (0, 0), (-1, 0), white),
        ("BACKGROUND", (0, 0), (-1, 0), PURPLE_DARK),
        ("BACKGROUND", (0, 1), (-1, 1), GREEN_BG),
        ("BACKGROUND", (0, 2), (-1, 2), RED_BG),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 1), (0, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#E5E7EB")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
    ]))
    elements.append(tbl)
    return elements


def generate_profit_pdf(data: dict) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        topMargin=0.5 * inch, bottomMargin=0.6 * inch,
        leftMargin=0.7 * inch, rightMargin=0.7 * inch,
    )
    ss = _styles()
    elements = []

    elements.extend(_header_table(data, ss))
    elements.extend(_volume_table(data, ss))
    elements.extend(_comparison_table(data, ss))
    elements.extend(_profitability_block(data, ss))
    elements.extend(_takeaway(data, ss))
    elements.extend(_pricing_detail(data, ss))

    elements.append(Spacer(1, 20))
    elements.append(Paragraph("Royal Purple Partnership Hub by ThrottlePro  |  Confidential", ss["Footer"]))

    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()
