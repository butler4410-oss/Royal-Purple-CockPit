import openpyxl
import json
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
import math

def _set_shape_alpha(shape, alpha_val):
    try:
        from pptx.oxml.ns import qn
        from lxml import etree
        sp_pr = shape._element.spPr
        solid_fill = sp_pr.find(qn('a:solidFill'))
        if solid_fill is not None:
            srgb = solid_fill.find(qn('a:srgbClr'))
            if srgb is not None:
                alpha_el = srgb.find(qn('a:alpha'))
                if alpha_el is None:
                    alpha_el = etree.SubElement(srgb, qn('a:alpha'))
                alpha_el.set('val', str(alpha_val))
    except Exception:
        pass

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
LOGO_WHITE = os.path.join(ASSETS_DIR, "Royal Purple White Logo.png")
LOGO_SYNTHETIC = os.path.join(ASSETS_DIR, "RPMO_logo_BF_Outline.png")
LOGO_EXPERT_YELLOW = os.path.join(ASSETS_DIR, "RP_Synthetic_Expert_Logo_Yellow_Text.png")
LOGO_EXPERT_BLACK = os.path.join(ASSETS_DIR, "RP_Synthetic_Expert_Logo_Black_Text.png")
LOGO_EXPERT_WHITE = os.path.join(ASSETS_DIR, "rp_synthetic_expert_white.png")
BG_NEVER_SETTLE = os.path.join(ASSETS_DIR, "25-RYP-02147 Employee LinkedIn Thumbnails P1-6.jpg")
IMG_BETTER_OIL = os.path.join(ASSETS_DIR, "Better Oil Starts Here.png")


C = {
    "purple": "4B2D8A",
    "purpleMid": "6B44B8",
    "purpleLight": "9B6FD4",
    "gold": "C8973A",
    "goldLight": "E8B85A",
    "white": "FFFFFF",
    "offWhite": "F8F5FF",
    "lightGray": "F2F2F2",
    "midGray": "94A3B8",
    "darkGray": "334155",
    "dark": "1E1035",
    "green": "22C55E",
    "teal": "0D9488",
}

PRODUCT_MAP = {
    "HMX": "High Mileage",
    "RMS": "High Mileage Syn",
    "RS": "Royal Purple High",
    "RP": "Royal Purple Syn",
    "RSD": "Duralec",
    "11722": "Max-Clean",
    "11755": "Royal Purple Premium",
    "18000": "Max-Atomizer",
}

PRODUCT_FULL_NAMES = {
    "11722": "Max-Clean Fuel System Cleaner",
    "18000": "Max-Atomizer Fuel Injector Cleaner",
    "HMX0W20": "HMX 0W-20 High Mileage",
    "HMX5W20": "HMX 5W-20 High Mileage",
    "HMX5W30": "HMX 5W-30 High Mileage",
    "RMS5W20": "HMX Syn 5W-20 High Mileage",
    "RMS5W30": "HMX Syn 5W-30 High Mileage",
    "RS0W16": "RP High Perf 0W-16",
    "RS0W20": "RP High Perf 0W-20",
    "RS0W40": "RP High Perf 0W-40",
    "RS5W20": "RP High Perf 5W-20",
    "RS5W30": "RP High Perf 5W-30",
    "RS5W40": "RP High Perf 5W-40",
    "RP0W16": "RP Synthetic 0W-16",
    "RP0W20": "RP Synthetic 0W-20",
    "RP0W40": "RP Synthetic 0W-40",
    "RP5W20": "RP Synthetic 5W-20",
    "RP5W30": "RP Synthetic 5W-30",
    "RP5W40": "RP Synthetic 5W-40",
    "RSD15W40": "Duralec Super 15W-40",
    "RSD5W40": "Duralec Super 5W-40",
    "11755": "RP Premium Motor Oil",
}

PRODUCT_DESCRIPTIONS = {
    "High Mileage": "Synthetic motor oil for engines with 75,000+ miles. Reduces oil consumption, revitalizes seals, and removes deposits using Synerlec additive technology.",
    "High Mileage Syn": "Full synthetic high mileage formulation with enhanced seal conditioning and superior wear protection for high-mileage engines.",
    "Royal Purple High": "High-performance full synthetic oil with Synerlec technology for superior wear protection, reduced heat, and increased fuel efficiency.",
    "Royal Purple Syn": "Premium full synthetic oil exceeding API/ILSAC standards. Enhanced film strength minimizes metal-to-metal contact in modern engines.",
    "Duralec": "Premium synthetic diesel engine oil (API CK-4) for emission-controlled engines with DPF, EGR, and SCR systems. Extends drain intervals.",
    "Max-Clean": "High-performance fuel system cleaner and stabilizer. Deeply cleans injectors, carburetors, intake valves, and combustion chambers.",
    "Max-Atomizer": "Advanced fuel injector cleaner for optimized spray patterns and improved combustion efficiency.",
    "Royal Purple Premium": "Premium synthetic motor oil with proprietary Synerlec additive technology for maximum engine protection.",
}

def get_product_display_name(code):
    return PRODUCT_FULL_NAMES.get(code, code)

def get_product_category_desc(category):
    return PRODUCT_DESCRIPTIONS.get(category, "")

def rgb(hex_str):
    return RGBColor.from_string(hex_str)

HEADER_PATTERNS = {
    "date": ["invoice date", "date", "service date", "trans date", "transaction date"],
    "product": ["operation code", "op code", "product", "description", "item", "service", "operation"],
    "invoices": ["# of invoices", "invoices", "invoice count", "num invoices", "transactions", "oil changes", "ticket count"],
    "revenue": ["total rev", "revenue", "total sales", "sales amount", "net sales", "gross rev", "gross sales"],
    "avg_rev": ["rev/inv", "avg rev", "average rev", "avg sale", "per invoice", "avg amount", "average sale", "rev per"],
    "vehicles": ["# of vehicles", "vehicles", "vehicle count", "num vehicles", "cars", "unique vehicles"],
    "qty": ["qty", "quantity", "count", "units"],
    "amount": ["amount", "sales", "total"],
}

SKIP_SHEETS = ["report summary", "summary", "totals", "notes", "instructions", "template", "info"]


def _find_column_index(header, field):
    patterns = HEADER_PATTERNS.get(field, [])
    header_lower = [str(h).lower().strip() if h else "" for h in header]
    for pattern in patterns:
        for i, h in enumerate(header_lower):
            if pattern in h:
                return i
    return None


def _safe_float(val, default=0):
    if val is None:
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def _safe_int(val, default=0):
    if val is None:
        return default
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return default


def _detect_date(data_rows, col_map):
    date_idx = col_map.get("date")
    if date_idx is None:
        return None
    for row in data_rows:
        if date_idx >= len(row):
            continue
        date_val = row[date_idx]
        if date_val is None:
            continue
        if hasattr(date_val, 'strftime'):
            return date_val.strftime("%B %Y")
        elif isinstance(date_val, str):
            date_val = date_val.strip()
            parts = date_val.split()
            if len(parts) >= 3:
                return f"{parts[0]} {parts[2].rstrip(',')}"
            elif len(parts) == 2:
                return date_val
    return None


def _get_val(row, idx, default=None):
    if idx is None or idx >= len(row):
        return default
    return row[idx]


def parse_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    stores = []
    month_year = None

    for sheet_name in wb.sheetnames:
        if sheet_name.lower().strip() in SKIP_SHEETS:
            continue
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            continue

        header_row_idx = 0
        header = rows[0]
        for i, row in enumerate(rows[:5]):
            row_strs = [str(c).lower() if c else "" for c in row]
            if any(p in s for s in row_strs for p in ["invoice", "date", "revenue", "product", "operation", "sales"]):
                header = row
                header_row_idx = i
                break

        col_map = {}
        for field in ["date", "product", "invoices", "revenue", "avg_rev", "vehicles"]:
            idx = _find_column_index(header, field)
            if idx is not None:
                col_map[field] = idx

        if "revenue" not in col_map:
            for fallback_field in ["amount"]:
                idx = _find_column_index(header, fallback_field)
                if idx is not None:
                    col_map["revenue"] = idx
                    break

        if "invoices" not in col_map:
            for fallback_field in ["qty"]:
                idx = _find_column_index(header, fallback_field)
                if idx is not None:
                    col_map["invoices"] = idx
                    break

        if "revenue" not in col_map and "invoices" not in col_map:
            continue

        all_data_rows = rows[header_row_idx + 1:]
        first_col = col_map.get("date", col_map.get("product", 0))
        data_rows = [r for r in all_data_rows if len(r) > first_col and r[first_col] is not None]

        last_row = all_data_rows[-1] if all_data_rows else None
        totals_row = None
        if last_row and len(last_row) > first_col and last_row[first_col] is None:
            totals_row = last_row

        if not month_year:
            month_year = _detect_date(data_rows, col_map)

        product_idx = col_map.get("product")
        revenue_idx = col_map.get("revenue")

        product_revenue = {}
        total_rev_calc = 0
        total_inv_calc = 0
        total_veh_calc = 0
        for row in data_rows:
            if product_idx is not None:
                op_desc = str(_get_val(row, product_idx, "")) if _get_val(row, product_idx) else ""
                code = op_desc.split(" - ")[0].strip() if " - " in op_desc else op_desc.strip()
            else:
                code = ""

            rev = _safe_float(_get_val(row, revenue_idx))
            total_rev_calc += rev
            total_inv_calc += _safe_int(_get_val(row, col_map.get("invoices")))
            total_veh_calc += _safe_int(_get_val(row, col_map.get("vehicles")))

            if code:
                product_revenue[code] = product_revenue.get(code, 0) + rev

        if totals_row:
            total_invoices = _safe_int(_get_val(totals_row, col_map.get("invoices"))) or total_inv_calc
            total_revenue = _safe_float(_get_val(totals_row, col_map.get("revenue"))) or total_rev_calc
            avg_rev_inv = _safe_float(_get_val(totals_row, col_map.get("avg_rev")))
            total_vehicles = _safe_int(_get_val(totals_row, col_map.get("vehicles"))) or total_veh_calc
        else:
            total_invoices = total_inv_calc
            total_revenue = total_rev_calc
            avg_rev_inv = 0
            total_vehicles = total_veh_calc

        if (not avg_rev_inv or avg_rev_inv < 1) and total_invoices and total_revenue:
            avg_rev_inv = total_revenue / total_invoices

        if "invoices" not in col_map and total_invoices == 0 and len(data_rows) > 0:
            total_invoices = len(data_rows)
            if total_revenue:
                avg_rev_inv = total_revenue / total_invoices

        sorted_prefixes = sorted(PRODUCT_MAP.keys(), key=len, reverse=True)
        product_breakdown = []
        for code, rev in sorted(product_revenue.items(), key=lambda x: -x[1]):
            cat = "Other"
            for prefix in sorted_prefixes:
                if code.startswith(prefix):
                    cat = PRODUCT_MAP[prefix]
                    break
            product_breakdown.append({
                "code": code,
                "category": cat,
                "revenue": round(rev, 2),
            })

        top_product = product_breakdown[0]["category"] if product_breakdown else "N/A"

        stores.append({
            "name": sheet_name,
            "invoices": int(total_invoices),
            "vehicles": int(total_vehicles),
            "totalRevenue": round(float(total_revenue), 2),
            "avgRevPerInvoice": round(float(avg_rev_inv), 2),
            "topProduct": top_product,
            "productBreakdown": product_breakdown,
        })

    stores.sort(key=lambda s: -s["totalRevenue"])
    for i, s in enumerate(stores):
        s["rank"] = i + 1

    if not month_year:
        from datetime import datetime
        month_year = datetime.now().strftime("%B %Y")

    if not stores:
        raise ValueError("No store data found in the Excel file. Ensure the workbook has at least one data sheet with recognizable column headers (e.g., Revenue, Invoices, Product).")

    return stores, month_year


def add_slide_background(slide, color=C["offWhite"]):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def add_top_bar(slide):
    shape = slide.shapes.add_shape(
        1, Inches(0), Inches(0), Inches(10), Inches(0.55)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(C["purple"])
    shape.line.fill.background()


def add_footer(slide, page_num, total_slides):
    bar = slide.shapes.add_shape(
        1, Inches(0), Inches(5.33), Inches(10), Inches(0.22)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = rgb(C["purple"])
    bar.line.fill.background()

    if os.path.exists(LOGO_WHITE):
        slide.shapes.add_picture(
            LOGO_WHITE, Inches(0.2), Inches(5.35), Inches(0.95), Inches(0.17)
        )

    tf = bar.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.add_run()
    run.text = f"  {page_num} / {total_slides}"
    run.font.size = Pt(7)
    run.font.color.rgb = rgb(C["white"])
    run.font.name = "Calibri"


def add_royal_purple_badge(slide):
    if os.path.exists(LOGO_WHITE):
        slide.shapes.add_picture(
            LOGO_WHITE, Inches(7.8), Inches(0.1), Inches(1.8), Inches(0.35)
        )
    else:
        txBox = slide.shapes.add_textbox(
            Inches(7.8), Inches(0.12), Inches(2.0), Inches(0.3)
        )
        tf = txBox.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        run = p.add_run()
        run.text = "ROYAL PURPLE"
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = rgb(C["gold"])
        run.font.name = "Calibri"


def add_slide_header(slide, title, subtitle=None):
    add_top_bar(slide)
    add_royal_purple_badge(slide)

    accent = slide.shapes.add_shape(
        1, Inches(0.4), Inches(0.7), Inches(0.07), Inches(0.45)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = rgb(C["gold"])
    accent.line.fill.background()

    txBox = slide.shapes.add_textbox(
        Inches(0.6), Inches(0.65), Inches(8.0), Inches(0.5)
    )
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = rgb(C["purple"])
    run.font.name = "Calibri"

    if subtitle:
        txBox2 = slide.shapes.add_textbox(
            Inches(0.6), Inches(1.05), Inches(8.0), Inches(0.3)
        )
        tf2 = txBox2.text_frame
        p2 = tf2.paragraphs[0]
        run2 = p2.add_run()
        run2.text = subtitle
        run2.font.size = Pt(10)
        run2.font.color.rgb = rgb(C["midGray"])
        run2.font.name = "Calibri"


def add_stat_card(slide, x, y, w, h, value, label, sub_label=None):
    card = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    card.fill.solid()
    card.fill.fore_color.rgb = rgb(C["white"])
    card.line.fill.background()

    gold_bar = slide.shapes.add_shape(
        1, Inches(x), Inches(y), Inches(w), Inches(0.06)
    )
    gold_bar.fill.solid()
    gold_bar.fill.fore_color.rgb = rgb(C["gold"])
    gold_bar.line.fill.background()

    val_str = str(value)
    font_size = 20 if len(val_str) > 8 else 26

    val_box = slide.shapes.add_textbox(
        Inches(x + 0.1), Inches(y + 0.15), Inches(w - 0.2), Inches(0.45)
    )
    tf = val_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = val_str
    run.font.size = Pt(font_size)
    run.font.bold = True
    run.font.color.rgb = rgb(C["purple"])
    run.font.name = "Calibri"

    lbl_box = slide.shapes.add_textbox(
        Inches(x + 0.1), Inches(y + 0.6), Inches(w - 0.2), Inches(0.25)
    )
    tf2 = lbl_box.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = label
    run2.font.size = Pt(9)
    run2.font.color.rgb = rgb(C["darkGray"])
    run2.font.name = "Calibri"

    if sub_label:
        sub_box = slide.shapes.add_textbox(
            Inches(x + 0.1), Inches(y + 0.82), Inches(w - 0.2), Inches(0.2)
        )
        tf3 = sub_box.text_frame
        tf3.word_wrap = True
        p3 = tf3.paragraphs[0]
        p3.alignment = PP_ALIGN.CENTER
        run3 = p3.add_run()
        run3.text = sub_label
        run3.font.size = Pt(7)
        run3.font.color.rgb = rgb(C["midGray"])
        run3.font.name = "Calibri"


def fmt_currency(val):
    if val >= 1_000_000:
        return f"${val/1_000_000:.2f}M"
    elif val >= 1_000:
        return f"${val:,.0f}"
    else:
        return f"${val:.2f}"


def fmt_number(val):
    return f"{val:,}"


def build_cover_slide(prs, stores, month_year, total_slides):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide, C["dark"])

    if os.path.exists(BG_NEVER_SETTLE):
        slide.shapes.add_picture(
            BG_NEVER_SETTLE, Inches(0), Inches(0), Inches(10), Inches(5.625)
        )

    overlay = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(5.625))
    overlay.fill.solid()
    overlay.fill.fore_color.rgb = rgb(C["dark"])
    _set_shape_alpha(overlay, 60000)
    overlay.line.fill.background()

    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = rgb(C["gold"])
    bar.line.fill.background()

    logo_path = LOGO_EXPERT_WHITE if os.path.exists(LOGO_EXPERT_WHITE) else LOGO_EXPERT_YELLOW
    if os.path.exists(logo_path):
        slide.shapes.add_picture(
            logo_path, Inches(0.7), Inches(0.4), Inches(2.6), Inches(1.1)
        )

    txBox2 = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(6), Inches(1.0))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    run2 = p2.add_run()
    run2.text = "Partnership Hub Report"
    run2.font.size = Pt(36)
    run2.font.bold = True
    run2.font.color.rgb = rgb(C["white"])
    run2.font.name = "Calibri"

    txBox3 = slide.shapes.add_textbox(Inches(0.8), Inches(2.55), Inches(6), Inches(0.5))
    tf3 = txBox3.text_frame
    p3 = tf3.paragraphs[0]
    run3 = p3.add_run()
    run3.text = f"{month_year} | {len(stores)} Locations"
    run3.font.size = Pt(14)
    run3.font.color.rgb = rgb(C["purpleLight"])
    run3.font.name = "Calibri"

    ns_box = slide.shapes.add_textbox(Inches(0.8), Inches(3.05), Inches(3), Inches(0.35))
    tf_ns = ns_box.text_frame
    p_ns = tf_ns.paragraphs[0]
    r_ns = p_ns.add_run()
    r_ns.text = "NEVER SETTLE"
    r_ns.font.size = Pt(16)
    r_ns.font.bold = True
    r_ns.font.color.rgb = rgb(C["white"])
    r_ns.font.name = "Calibri"

    total_rev = sum(s["totalRevenue"] for s in stores)
    total_inv = sum(s["invoices"] for s in stores)
    avg_rev = total_rev / total_inv if total_inv else 0

    stat_x = 6.8
    stat_y = 1.2
    stat_w = 2.8
    stat_h = 2.2

    stat_bg = slide.shapes.add_shape(1, Inches(stat_x), Inches(stat_y), Inches(stat_w), Inches(stat_h))
    stat_bg.fill.solid()
    stat_bg.fill.fore_color.rgb = rgb(C["purple"])
    stat_bg.line.fill.background()

    gold_top = slide.shapes.add_shape(1, Inches(stat_x), Inches(stat_y), Inches(stat_w), Inches(0.06))
    gold_top.fill.solid()
    gold_top.fill.fore_color.rgb = rgb(C["gold"])
    gold_top.line.fill.background()

    lbl = slide.shapes.add_textbox(Inches(stat_x + 0.15), Inches(stat_y + 0.15), Inches(stat_w - 0.3), Inches(0.25))
    tf_l = lbl.text_frame
    p_l = tf_l.paragraphs[0]
    r_l = p_l.add_run()
    r_l.text = f"{month_year} Summary"
    r_l.font.size = Pt(10)
    r_l.font.bold = True
    r_l.font.color.rgb = rgb(C["goldLight"])
    r_l.font.name = "Calibri"

    items = [
        (fmt_currency(total_rev), "Total Revenue"),
        (fmt_number(total_inv), "Oil Changes"),
        (f"${avg_rev:.2f}", "Avg Rev/Invoice"),
        (stores[0]["name"] if stores else "N/A", "Top Store"),
    ]
    for i, (val, lab) in enumerate(items):
        iy = stat_y + 0.5 + i * 0.42
        vb = slide.shapes.add_textbox(Inches(stat_x + 0.2), Inches(iy), Inches(stat_w - 0.4), Inches(0.22))
        tf_v = vb.text_frame
        p_v = tf_v.paragraphs[0]
        r_v = p_v.add_run()
        r_v.text = val
        r_v.font.size = Pt(14)
        r_v.font.bold = True
        r_v.font.color.rgb = rgb(C["white"])
        r_v.font.name = "Calibri"

        lb = slide.shapes.add_textbox(Inches(stat_x + 0.2), Inches(iy + 0.2), Inches(stat_w - 0.4), Inches(0.15))
        tf_lb = lb.text_frame
        p_lb = tf_lb.paragraphs[0]
        r_lb = p_lb.add_run()
        r_lb.text = lab
        r_lb.font.size = Pt(7)
        r_lb.font.color.rgb = rgb(C["purpleLight"])
        r_lb.font.name = "Calibri"

    footer_txt = slide.shapes.add_textbox(Inches(0.8), Inches(4.8), Inches(8), Inches(0.3))
    tf_f = footer_txt.text_frame
    p_f = tf_f.paragraphs[0]
    r_f = p_f.add_run()
    r_f.text = "ThrottlePro — More Cars. More Loyalty. Less Stress."
    r_f.font.size = Pt(8)
    r_f.font.color.rgb = rgb(C["midGray"])
    r_f.font.name = "Calibri"

    add_footer(slide, 1, total_slides)


def build_toc_slide(prs, total_slides):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide)
    add_slide_header(slide, "Table of Contents", "Report sections overview")
    add_footer(slide, 2, total_slides)

    sections = [
        ("1", "Executive Summary", "KPIs & observations"),
        ("2", "Revenue Overview", "Total revenue breakdown"),
        ("3", "Store Rankings", "Performance table & matrix"),
        ("4", "Product Mix", "Category analysis"),
        ("5", "Store Deep Dives", "Individual store detail"),
        ("6", "Next Steps", "Recommendations"),
    ]
    for i, (num, title, desc) in enumerate(sections):
        col = i % 3
        row = i // 3
        x = 0.5 + col * 3.1
        y = 1.55 + row * 1.7

        card = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(2.8), Inches(1.4))
        card.fill.solid()
        card.fill.fore_color.rgb = rgb(C["white"])
        card.line.fill.background()

        num_box = slide.shapes.add_textbox(Inches(x + 0.15), Inches(y + 0.1), Inches(0.5), Inches(0.5))
        tf = num_box.text_frame
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = num
        r.font.size = Pt(28)
        r.font.bold = True
        r.font.color.rgb = rgb(C["gold"])
        r.font.name = "Calibri"

        t_box = slide.shapes.add_textbox(Inches(x + 0.15), Inches(y + 0.6), Inches(2.5), Inches(0.35))
        tf2 = t_box.text_frame
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run()
        r2.text = title
        r2.font.size = Pt(12)
        r2.font.bold = True
        r2.font.color.rgb = rgb(C["purple"])
        r2.font.name = "Calibri"

        d_box = slide.shapes.add_textbox(Inches(x + 0.15), Inches(y + 0.95), Inches(2.5), Inches(0.3))
        tf3 = d_box.text_frame
        p3 = tf3.paragraphs[0]
        r3 = p3.add_run()
        r3.text = desc
        r3.font.size = Pt(9)
        r3.font.color.rgb = rgb(C["midGray"])
        r3.font.name = "Calibri"


def build_exec_summary_kpis(prs, stores, month_year, total_slides):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide)
    add_slide_header(slide, "Executive Summary", f"Key performance indicators — {month_year}")
    add_footer(slide, 3, total_slides)

    total_rev = sum(s["totalRevenue"] for s in stores)
    total_inv = sum(s["invoices"] for s in stores)
    avg_rev = total_rev / total_inv if total_inv else 0
    total_veh = sum(s["vehicles"] for s in stores)

    cards = [
        (fmt_currency(total_rev), "Total Revenue", f"Across {len(stores)} locations"),
        (fmt_number(total_inv), "Total Oil Changes", f"{month_year}"),
        (f"${avg_rev:.2f}", "Avg Rev / Invoice", "Network average"),
        (fmt_number(total_veh), "Unique Vehicles", f"{month_year}"),
    ]
    card_w = 2.1
    gap = 0.2
    start_x = (10 - (4 * card_w + 3 * gap)) / 2
    for i, (val, lbl, sub) in enumerate(cards):
        x = start_x + i * (card_w + gap)
        add_stat_card(slide, x, 1.55, card_w, 1.1, val, lbl, sub)

    top8 = stores[:8]
    max_rev = top8[0]["totalRevenue"] if top8 else 1

    bar_y = 3.0
    lbl_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.8), Inches(4), Inches(0.3))
    tf = lbl_box.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "Revenue Leaderboard — Top 8 Stores"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = rgb(C["purple"])
    r.font.name = "Calibri"

    for i, store in enumerate(top8):
        y = bar_y + i * 0.28
        bar_w = max(0.3, (store["totalRevenue"] / max_rev) * 6.5)

        name_box = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(2.5), Inches(0.26))
        tf_n = name_box.text_frame
        p_n = tf_n.paragraphs[0]
        p_n.alignment = PP_ALIGN.RIGHT
        r_n = p_n.add_run()
        r_n.text = store["name"]
        r_n.font.size = Pt(8)
        r_n.font.color.rgb = rgb(C["darkGray"])
        r_n.font.name = "Calibri"

        bar_color = C["purple"] if i % 2 == 0 else C["purpleMid"]
        bar = slide.shapes.add_shape(1, Inches(3.1), Inches(y + 0.02), Inches(bar_w), Inches(0.22))
        bar.fill.solid()
        bar.fill.fore_color.rgb = rgb(bar_color)
        bar.line.fill.background()

        val_box = slide.shapes.add_textbox(Inches(3.1 + bar_w + 0.1), Inches(y), Inches(1.2), Inches(0.26))
        tf_v = val_box.text_frame
        p_v = tf_v.paragraphs[0]
        r_v = p_v.add_run()
        r_v.text = fmt_currency(store["totalRevenue"])
        r_v.font.size = Pt(8)
        r_v.font.bold = True
        r_v.font.color.rgb = rgb(C["purple"])
        r_v.font.name = "Calibri"


def build_exec_observations(prs, stores, month_year, total_slides):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide)
    add_slide_header(slide, "Executive Observations", f"Key insights — {month_year}")
    add_footer(slide, 4, total_slides)

    total_rev = sum(s["totalRevenue"] for s in stores)
    total_inv = sum(s["invoices"] for s in stores)
    avg_rev = total_rev / total_inv if total_inv else 0
    top_store = stores[0] if stores else None
    bottom_store = stores[-1] if stores else None

    hm_rev = sum(
        pb["revenue"] for s in stores for pb in s["productBreakdown"]
        if pb["category"] in ("High Mileage", "High Mileage Syn")
    )
    hm_pct = (hm_rev / total_rev * 100) if total_rev else 0

    observations = [
        (
            "Network Revenue",
            f"Total network revenue of {fmt_currency(total_rev)} across {len(stores)} locations with {fmt_number(total_inv)} oil changes at an average of ${avg_rev:.2f} per invoice.",
            C["purple"],
        ),
        (
            "Top Performer",
            f"{top_store['name']} leads with {fmt_currency(top_store['totalRevenue'])} in revenue ({top_store['totalRevenue']/total_rev*100:.1f}% of network total) and {fmt_number(top_store['invoices'])} oil changes." if top_store else "N/A",
            C["gold"],
        ),
        (
            "Product Mix Insight",
            f"High Mileage products account for {hm_pct:.1f}% of total revenue, indicating strong upsell penetration across the network.",
            C["purpleMid"],
        ),
        (
            "Growth Opportunity",
            f"{bottom_store['name']} has the lowest revenue at {fmt_currency(bottom_store['totalRevenue'])} — targeted promotions could lift volume." if bottom_store else "N/A",
            C["teal"],
        ),
    ]

    for i, (title, body, accent_color) in enumerate(observations):
        col = i % 2
        row = i // 2
        x = 0.5 + col * 4.6
        y = 1.55 + row * 1.85
        w = 4.3
        h = 1.65

        card = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
        card.fill.solid()
        card.fill.fore_color.rgb = rgb(C["white"])
        card.line.fill.background()

        accent = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(0.06), Inches(h))
        accent.fill.solid()
        accent.fill.fore_color.rgb = rgb(accent_color)
        accent.line.fill.background()

        t_box = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.12), Inches(w - 0.4), Inches(0.3))
        tf = t_box.text_frame
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = title
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = rgb(accent_color)
        r.font.name = "Calibri"

        b_box = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.45), Inches(w - 0.4), Inches(h - 0.55))
        tf2 = b_box.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run()
        r2.text = body
        r2.font.size = Pt(9)
        r2.font.color.rgb = rgb(C["darkGray"])
        r2.font.name = "Calibri"


def build_revenue_overview(prs, stores, month_year, total_slides):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide)
    add_slide_header(slide, "Revenue Overview", f"Total revenue distribution — {month_year}")
    add_footer(slide, 5, total_slides)

    total_rev = sum(s["totalRevenue"] for s in stores)

    hero = slide.shapes.add_shape(1, Inches(0.5), Inches(1.55), Inches(3.8), Inches(3.5))
    hero.fill.solid()
    hero.fill.fore_color.rgb = rgb(C["purple"])
    hero.line.fill.background()

    h_lbl = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(3.2), Inches(0.3))
    tf = h_lbl.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "TOTAL NETWORK REVENUE"
    r.font.size = Pt(9)
    r.font.bold = True
    r.font.color.rgb = rgb(C["goldLight"])
    r.font.name = "Calibri"

    h_val = slide.shapes.add_textbox(Inches(0.8), Inches(2.1), Inches(3.2), Inches(0.6))
    tf2 = h_val.text_frame
    p2 = tf2.paragraphs[0]
    r2 = p2.add_run()
    r2.text = fmt_currency(total_rev)
    r2.font.size = Pt(32)
    r2.font.bold = True
    r2.font.color.rgb = rgb(C["white"])
    r2.font.name = "Calibri"

    h_sub = slide.shapes.add_textbox(Inches(0.8), Inches(2.7), Inches(3.2), Inches(0.3))
    tf3 = h_sub.text_frame
    p3 = tf3.paragraphs[0]
    r3 = p3.add_run()
    r3.text = f"{len(stores)} Locations | {month_year}"
    r3.font.size = Pt(10)
    r3.font.color.rgb = rgb(C["purpleLight"])
    r3.font.name = "Calibri"

    total_inv = sum(s["invoices"] for s in stores)
    avg_rev = total_rev / total_inv if total_inv else 0
    stats = [
        (fmt_number(total_inv), "Total Oil Changes"),
        (f"${avg_rev:.2f}", "Avg Rev/Invoice"),
        (fmt_number(sum(s["vehicles"] for s in stores)), "Unique Vehicles"),
    ]
    for i, (val, lbl) in enumerate(stats):
        sy = 3.3 + i * 0.55
        vb = slide.shapes.add_textbox(Inches(0.8), Inches(sy), Inches(1.8), Inches(0.25))
        tf_v = vb.text_frame
        p_v = tf_v.paragraphs[0]
        r_v = p_v.add_run()
        r_v.text = val
        r_v.font.size = Pt(14)
        r_v.font.bold = True
        r_v.font.color.rgb = rgb(C["white"])
        r_v.font.name = "Calibri"

        lb = slide.shapes.add_textbox(Inches(2.6), Inches(sy), Inches(1.5), Inches(0.25))
        tf_l = lb.text_frame
        p_l = tf_l.paragraphs[0]
        r_l = p_l.add_run()
        r_l.text = lbl
        r_l.font.size = Pt(8)
        r_l.font.color.rgb = rgb(C["purpleLight"])
        r_l.font.name = "Calibri"

    list_x = 4.6
    list_y = 1.55
    hdr = slide.shapes.add_textbox(Inches(list_x), Inches(list_y), Inches(5), Inches(0.3))
    tf_h = hdr.text_frame
    p_h = tf_h.paragraphs[0]
    r_h = p_h.add_run()
    r_h.text = "Share of Total Revenue by Store"
    r_h.font.size = Pt(10)
    r_h.font.bold = True
    r_h.font.color.rgb = rgb(C["purple"])
    r_h.font.name = "Calibri"

    for i, store in enumerate(stores):
        sy = list_y + 0.35 + i * 0.28
        pct = store["totalRevenue"] / total_rev * 100 if total_rev else 0

        nm = slide.shapes.add_textbox(Inches(list_x), Inches(sy), Inches(2.5), Inches(0.25))
        tf_n = nm.text_frame
        p_n = tf_n.paragraphs[0]
        r_n = p_n.add_run()
        r_n.text = store["name"]
        r_n.font.size = Pt(8)
        r_n.font.color.rgb = rgb(C["darkGray"])
        r_n.font.name = "Calibri"

        bar_w = max(0.15, pct / 100 * 2.5)
        bar = slide.shapes.add_shape(1, Inches(list_x + 2.6), Inches(sy + 0.03), Inches(bar_w), Inches(0.18))
        bar.fill.solid()
        bar.fill.fore_color.rgb = rgb(C["purple"] if i % 2 == 0 else C["purpleMid"])
        bar.line.fill.background()

        pv = slide.shapes.add_textbox(Inches(list_x + 2.6 + bar_w + 0.05), Inches(sy), Inches(0.8), Inches(0.25))
        tf_p = pv.text_frame
        p_p = tf_p.paragraphs[0]
        r_p = p_p.add_run()
        r_p.text = f"{pct:.1f}%"
        r_p.font.size = Pt(7)
        r_p.font.bold = True
        r_p.font.color.rgb = rgb(C["purple"])
        r_p.font.name = "Calibri"


def build_ranking_table(prs, stores, month_year, total_slides, start_page):
    TABLE_TOP = 1.55
    FOOTER_Y = 5.33
    ROW_H = 0.285
    ROWS_PER_PAGE = int(math.floor((FOOTER_Y - TABLE_TOP) / ROW_H)) - 1

    total_rev = sum(s["totalRevenue"] for s in stores)
    chunks = [stores[i:i+ROWS_PER_PAGE] for i in range(0, len(stores), ROWS_PER_PAGE)]
    total_pages = len(chunks)
    pages_built = 0

    for page_idx, chunk in enumerate(chunks):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_slide_background(slide)
        subtitle = f"All locations ranked by total revenue — {month_year}"
        if total_pages > 1:
            subtitle += f"  ({page_idx+1} of {total_pages})"
        add_slide_header(slide, "Store Performance Ranking", subtitle)
        add_footer(slide, start_page + page_idx, total_slides)

        headers = ["Rank", "Store Name", "Revenue", "Oil Changes", "Avg Rev/Inv", "Share %"]
        col_widths = [0.5, 2.8, 1.3, 1.1, 1.1, 0.9]
        col_x = [0.5]
        for w in col_widths[:-1]:
            col_x.append(col_x[-1] + w)

        hy = TABLE_TOP
        for ci, (htext, cw) in enumerate(zip(headers, col_widths)):
            hdr_bg = slide.shapes.add_shape(1, Inches(col_x[ci]), Inches(hy), Inches(cw), Inches(ROW_H))
            hdr_bg.fill.solid()
            hdr_bg.fill.fore_color.rgb = rgb(C["purple"])
            hdr_bg.line.fill.background()

            hdr_tb = slide.shapes.add_textbox(Inches(col_x[ci] + 0.05), Inches(hy + 0.02), Inches(cw - 0.1), Inches(ROW_H - 0.04))
            tf = hdr_tb.text_frame
            tf.word_wrap = False
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if ci < 2 else PP_ALIGN.RIGHT
            r = p.add_run()
            r.text = htext
            r.font.size = Pt(8)
            r.font.bold = True
            r.font.color.rgb = rgb(C["white"])
            r.font.name = "Calibri"

        for ri, store in enumerate(chunk):
            ry = TABLE_TOP + ROW_H * (ri + 1)
            pct = store["totalRevenue"] / total_rev * 100 if total_rev else 0
            row_data = [
                f"#{store['rank']}",
                store["name"],
                fmt_currency(store["totalRevenue"]),
                fmt_number(store["invoices"]),
                f"${store['avgRevPerInvoice']:.2f}",
                f"{pct:.1f}%",
            ]

            bg_color = C["white"] if ri % 2 == 0 else C["lightGray"]
            for ci, (val, cw) in enumerate(zip(row_data, col_widths)):
                cell_bg = slide.shapes.add_shape(1, Inches(col_x[ci]), Inches(ry), Inches(cw), Inches(ROW_H))
                cell_bg.fill.solid()
                cell_bg.fill.fore_color.rgb = rgb(bg_color)
                cell_bg.line.fill.background()

                cell_tb = slide.shapes.add_textbox(Inches(col_x[ci] + 0.05), Inches(ry + 0.02), Inches(cw - 0.1), Inches(ROW_H - 0.04))
                tf = cell_tb.text_frame
                tf.word_wrap = False
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT if ci < 2 else PP_ALIGN.RIGHT
                r = p.add_run()
                r.text = val
                r.font.size = Pt(8)
                r.font.color.rgb = rgb(C["darkGray"])
                r.font.name = "Calibri"
                if ci == 0:
                    r.font.bold = True
                    r.font.color.rgb = rgb(C["gold"])

        pages_built += 1

    return pages_built


def build_performance_matrix(prs, stores, month_year, total_slides, start_page):
    ROW_Y0 = 1.83
    LEGEND_H = 0.30
    ROW_H = 0.258
    ROWS_PER_PAGE = int(math.floor((5.33 - ROW_Y0 - LEGEND_H) / ROW_H))

    total_inv = sum(s["invoices"] for s in stores)
    max_inv = max(s["invoices"] for s in stores) if stores else 1
    max_avg = max(s["avgRevPerInvoice"] for s in stores) if stores else 1

    chunks = [stores[i:i+ROWS_PER_PAGE] for i in range(0, len(stores), ROWS_PER_PAGE)]
    total_pages = len(chunks)
    pages_built = 0

    for page_idx, chunk in enumerate(chunks):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_slide_background(slide)
        subtitle = f"Volume vs. average ticket — {month_year}"
        if total_pages > 1:
            subtitle += f"  ({page_idx+1} of {total_pages})"
        add_slide_header(slide, "Store Performance Matrix", subtitle)
        add_footer(slide, start_page + page_idx, total_slides)

        legend_y = ROW_Y0
        leg1 = slide.shapes.add_shape(1, Inches(5.5), Inches(legend_y), Inches(0.3), Inches(0.15))
        leg1.fill.solid()
        leg1.fill.fore_color.rgb = rgb(C["purple"])
        leg1.line.fill.background()
        l1 = slide.shapes.add_textbox(Inches(5.85), Inches(legend_y - 0.02), Inches(1.2), Inches(0.2))
        tf1 = l1.text_frame
        p1 = tf1.paragraphs[0]
        r1 = p1.add_run()
        r1.text = "Volume (invoices)"
        r1.font.size = Pt(7)
        r1.font.color.rgb = rgb(C["darkGray"])
        r1.font.name = "Calibri"

        leg2 = slide.shapes.add_shape(1, Inches(7.2), Inches(legend_y), Inches(0.3), Inches(0.15))
        leg2.fill.solid()
        leg2.fill.fore_color.rgb = rgb(C["gold"])
        leg2.line.fill.background()
        l2 = slide.shapes.add_textbox(Inches(7.55), Inches(legend_y - 0.02), Inches(1.5), Inches(0.2))
        tf2 = l2.text_frame
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run()
        r2.text = "Avg Rev/Invoice ($)"
        r2.font.size = Pt(7)
        r2.font.color.rgb = rgb(C["darkGray"])
        r2.font.name = "Calibri"

        for ri, store in enumerate(chunk):
            ry = ROW_Y0 + LEGEND_H + ri * ROW_H

            nm = slide.shapes.add_textbox(Inches(0.5), Inches(ry), Inches(2.2), Inches(ROW_H))
            tf_n = nm.text_frame
            p_n = tf_n.paragraphs[0]
            r_n = p_n.add_run()
            r_n.text = store["name"]
            r_n.font.size = Pt(8)
            r_n.font.color.rgb = rgb(C["darkGray"])
            r_n.font.name = "Calibri"

            vol_w = max(0.2, (store["invoices"] / max_inv) * 3.5)
            vol_bar = slide.shapes.add_shape(1, Inches(2.8), Inches(ry + 0.02), Inches(vol_w), Inches(ROW_H * 0.4))
            vol_bar.fill.solid()
            vol_bar.fill.fore_color.rgb = rgb(C["purple"])
            vol_bar.line.fill.background()

            vol_lbl = slide.shapes.add_textbox(Inches(2.8 + vol_w + 0.05), Inches(ry - 0.02), Inches(0.8), Inches(ROW_H * 0.5))
            tf_vl = vol_lbl.text_frame
            p_vl = tf_vl.paragraphs[0]
            r_vl = p_vl.add_run()
            r_vl.text = fmt_number(store["invoices"])
            r_vl.font.size = Pt(6)
            r_vl.font.color.rgb = rgb(C["purple"])
            r_vl.font.name = "Calibri"

            avg_w = max(0.2, (store["avgRevPerInvoice"] / max_avg) * 3.5)
            avg_bar = slide.shapes.add_shape(1, Inches(2.8), Inches(ry + ROW_H * 0.5), Inches(avg_w), Inches(ROW_H * 0.4))
            avg_bar.fill.solid()
            avg_bar.fill.fore_color.rgb = rgb(C["gold"])
            avg_bar.line.fill.background()

            avg_lbl = slide.shapes.add_textbox(Inches(2.8 + avg_w + 0.05), Inches(ry + ROW_H * 0.4), Inches(0.8), Inches(ROW_H * 0.5))
            tf_al = avg_lbl.text_frame
            p_al = tf_al.paragraphs[0]
            r_al = p_al.add_run()
            r_al.text = f"${store['avgRevPerInvoice']:.0f}"
            r_al.font.size = Pt(6)
            r_al.font.color.rgb = rgb(C["gold"])
            r_al.font.name = "Calibri"

        pages_built += 1

    return pages_built


def build_product_mix(prs, stores, month_year, total_slides, page_num):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide)
    add_slide_header(slide, "Product Mix Analysis", f"Revenue by product category — {month_year}")
    add_footer(slide, page_num, total_slides)

    total_rev = sum(s["totalRevenue"] for s in stores)
    cat_rev = {}
    for s in stores:
        for pb in s["productBreakdown"]:
            cat = pb["category"]
            cat_rev[cat] = cat_rev.get(cat, 0) + pb["revenue"]

    rs_rev = sum(v for k, v in cat_rev.items() if k in ("Royal Purple High", "Royal Purple Syn"))
    hmx_rev = sum(v for k, v in cat_rev.items() if k in ("High Mileage", "High Mileage Syn"))
    other_rev = total_rev - rs_rev - hmx_rev
    top3 = [
        ("Royal Purple Synthetic", rs_rev),
        ("High Mileage", hmx_rev),
        ("Other / Specialty", max(other_rev, 0)),
    ]

    colors = [C["purple"], C["gold"], C["purpleMid"]]
    for i, (cat, rev) in enumerate(top3):
        x = 0.5 + i * 3.1
        pct = rev / total_rev * 100 if total_rev else 0

        card = slide.shapes.add_shape(1, Inches(x), Inches(1.55), Inches(2.8), Inches(1.2))
        card.fill.solid()
        card.fill.fore_color.rgb = rgb(C["white"])
        card.line.fill.background()

        accent = slide.shapes.add_shape(1, Inches(x), Inches(1.55), Inches(2.8), Inches(0.06))
        accent.fill.solid()
        accent.fill.fore_color.rgb = rgb(colors[i])
        accent.line.fill.background()

        v_box = slide.shapes.add_textbox(Inches(x + 0.15), Inches(1.7), Inches(2.5), Inches(0.4))
        tf = v_box.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = f"{pct:.1f}%"
        r.font.size = Pt(24)
        r.font.bold = True
        r.font.color.rgb = rgb(colors[i])
        r.font.name = "Calibri"

        n_box = slide.shapes.add_textbox(Inches(x + 0.15), Inches(2.1), Inches(2.5), Inches(0.25))
        tf2 = n_box.text_frame
        p2 = tf2.paragraphs[0]
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = cat
        r2.font.size = Pt(9)
        r2.font.bold = True
        r2.font.color.rgb = rgb(C["darkGray"])
        r2.font.name = "Calibri"

        rv_box = slide.shapes.add_textbox(Inches(x + 0.15), Inches(2.35), Inches(2.5), Inches(0.2))
        tf3 = rv_box.text_frame
        p3 = tf3.paragraphs[0]
        p3.alignment = PP_ALIGN.CENTER
        r3 = p3.add_run()
        r3.text = fmt_currency(rev)
        r3.font.size = Pt(8)
        r3.font.color.rgb = rgb(C["midGray"])
        r3.font.name = "Calibri"

    sku_rev = {}
    for s in stores:
        for pb in s["productBreakdown"]:
            sku_rev[pb["code"]] = sku_rev.get(pb["code"], 0) + pb["revenue"]

    sorted_skus = sorted(sku_rev.items(), key=lambda x: -x[1])[:10]
    max_sku_rev = sorted_skus[0][1] if sorted_skus else 1

    bar_y = 3.1
    lbl_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.9), Inches(4), Inches(0.3))
    tf = lbl_box.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "Revenue by Product SKU (Top 10)"
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = rgb(C["purple"])
    r.font.name = "Calibri"

    for i, (code, rev) in enumerate(sorted_skus):
        y = bar_y + i * 0.22
        bar_w = max(0.2, (rev / max_sku_rev) * 5.0)

        nm = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(2.0), Inches(0.2))
        tf_n = nm.text_frame
        p_n = tf_n.paragraphs[0]
        p_n.alignment = PP_ALIGN.RIGHT
        r_n = p_n.add_run()
        r_n.text = get_product_display_name(code)
        r_n.font.size = Pt(7)
        r_n.font.color.rgb = rgb(C["darkGray"])
        r_n.font.name = "Calibri"

        bar = slide.shapes.add_shape(1, Inches(2.6), Inches(y + 0.02), Inches(bar_w), Inches(0.16))
        bar.fill.solid()
        bar.fill.fore_color.rgb = rgb(C["purple"] if i % 2 == 0 else C["gold"])
        bar.line.fill.background()

        vb = slide.shapes.add_textbox(Inches(2.6 + bar_w + 0.05), Inches(y), Inches(1.2), Inches(0.2))
        tf_v = vb.text_frame
        p_v = tf_v.paragraphs[0]
        r_v = p_v.add_run()
        r_v.text = fmt_currency(rev)
        r_v.font.size = Pt(7)
        r_v.font.bold = True
        r_v.font.color.rgb = rgb(C["purple"])
        r_v.font.name = "Calibri"


def build_section_divider(prs, title, subtitle, total_slides, page_num):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide, C["dark"])

    if os.path.exists(BG_NEVER_SETTLE):
        slide.shapes.add_picture(
            BG_NEVER_SETTLE, Inches(0), Inches(0), Inches(10), Inches(5.625)
        )
        overlay = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(5.625))
        overlay.fill.solid()
        overlay.fill.fore_color.rgb = rgb(C["dark"])
        _set_shape_alpha(overlay, 70000)
        overlay.line.fill.background()

    logo_div = LOGO_EXPERT_WHITE if os.path.exists(LOGO_EXPERT_WHITE) else LOGO_WHITE
    if os.path.exists(logo_div):
        slide.shapes.add_picture(
            logo_div, Inches(3.3), Inches(0.7), Inches(3.4), Inches(1.4)
        )

    bar = slide.shapes.add_shape(1, Inches(2.5), Inches(2.25), Inches(5), Inches(0.04))
    bar.fill.solid()
    bar.fill.fore_color.rgb = rgb(C["gold"])
    bar.line.fill.background()

    t_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(0.7))
    tf = t_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = title
    r.font.size = Pt(30)
    r.font.bold = True
    r.font.color.rgb = rgb(C["white"])
    r.font.name = "Calibri"

    s_box = slide.shapes.add_textbox(Inches(1), Inches(3.2), Inches(8), Inches(0.4))
    tf2 = s_box.text_frame
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = subtitle
    r2.font.size = Pt(12)
    r2.font.color.rgb = rgb(C["purpleLight"])
    r2.font.name = "Calibri"

    ns_box = slide.shapes.add_textbox(Inches(1), Inches(3.65), Inches(8), Inches(0.35))
    tf_ns = ns_box.text_frame
    p_ns = tf_ns.paragraphs[0]
    p_ns.alignment = PP_ALIGN.CENTER
    r_ns = p_ns.add_run()
    r_ns.text = "NEVER SETTLE"
    r_ns.font.size = Pt(14)
    r_ns.font.bold = True
    r_ns.font.color.rgb = rgb(C["goldLight"])
    r_ns.font.name = "Calibri"

    add_footer(slide, page_num, total_slides)


def build_deep_dive(prs, store, stores, month_year, total_slides, page_num):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide)
    add_slide_header(
        slide,
        f"Store Deep Dive: {store['name']}",
        f"Rank #{store['rank']} of {len(stores)} by Revenue | {month_year}"
    )
    add_footer(slide, page_num, total_slides)

    cards = [
        (fmt_currency(store["totalRevenue"]), "Total Revenue", month_year),
        (fmt_number(store["invoices"]), "Oil Changes", "Invoice count"),
        (f"${store['avgRevPerInvoice']:.2f}", "Avg Rev/Invoice", "Per transaction"),
        (f"#{store['rank']}", "Network Rank", f"of {len(stores)} stores"),
    ]
    card_w = 2.05
    gap = 0.15
    start_x = 0.5
    for i, (val, lbl, sub) in enumerate(cards):
        x = start_x + i * (card_w + gap)
        add_stat_card(slide, x, 1.55, card_w, 1.05, val, lbl, sub)

    top_products = store["productBreakdown"][:6]
    if top_products:
        max_prod_rev = top_products[0]["revenue"]
        chart_y = 2.85
        chart_h_total = 2.2
        bar_h = min(0.28, chart_h_total / max(len(top_products), 1))

        lbl = slide.shapes.add_textbox(Inches(0.5), Inches(chart_y - 0.25), Inches(4), Inches(0.25))
        tf = lbl.text_frame
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = "Revenue by Product SKU"
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = rgb(C["purple"])
        r.font.name = "Calibri"

        for pi, prod in enumerate(top_products):
            py = chart_y + pi * (bar_h + 0.05)
            bar_w = max(0.2, (prod["revenue"] / max_prod_rev) * 3.8) if max_prod_rev else 0.2

            nm = slide.shapes.add_textbox(Inches(0.5), Inches(py), Inches(1.8), Inches(bar_h))
            tf_n = nm.text_frame
            tf_n.word_wrap = False
            p_n = tf_n.paragraphs[0]
            p_n.alignment = PP_ALIGN.RIGHT
            r_n = p_n.add_run()
            r_n.text = get_product_display_name(prod["code"])
            r_n.font.size = Pt(7)
            r_n.font.color.rgb = rgb(C["darkGray"])
            r_n.font.name = "Calibri"

            bar_color = C["purple"] if pi % 2 == 0 else C["gold"]
            bar = slide.shapes.add_shape(1, Inches(2.4), Inches(py + 0.02), Inches(bar_w), Inches(bar_h - 0.04))
            bar.fill.solid()
            bar.fill.fore_color.rgb = rgb(bar_color)
            bar.line.fill.background()

            vb = slide.shapes.add_textbox(Inches(2.4 + bar_w + 0.05), Inches(py), Inches(1.2), Inches(bar_h))
            tf_v = vb.text_frame
            p_v = tf_v.paragraphs[0]
            r_v = p_v.add_run()
            r_v.text = fmt_currency(prod["revenue"])
            r_v.font.size = Pt(7)
            r_v.font.bold = True
            r_v.font.color.rgb = rgb(C["purple"])
            r_v.font.name = "Calibri"

    notes_x = 6.5
    notes_y = 2.85
    notes_w = 3.2
    notes_h = 2.2

    notes_bg = slide.shapes.add_shape(1, Inches(notes_x), Inches(notes_y), Inches(notes_w), Inches(notes_h))
    notes_bg.fill.solid()
    notes_bg.fill.fore_color.rgb = rgb(C["purple"])
    notes_bg.line.fill.background()

    n_lbl = slide.shapes.add_textbox(Inches(notes_x + 0.15), Inches(notes_y + 0.1), Inches(notes_w - 0.3), Inches(0.25))
    tf_nl = n_lbl.text_frame
    p_nl = tf_nl.paragraphs[0]
    r_nl = p_nl.add_run()
    r_nl.text = "Location Notes"
    r_nl.font.size = Pt(10)
    r_nl.font.bold = True
    r_nl.font.color.rgb = rgb(C["goldLight"])
    r_nl.font.name = "Calibri"

    total_rev = sum(s["totalRevenue"] for s in stores)
    pct = store["totalRevenue"] / total_rev * 100 if total_rev else 0
    cat_breakdown = {}
    for pb in store["productBreakdown"]:
        cat_breakdown[pb["category"]] = cat_breakdown.get(pb["category"], 0) + pb["revenue"]
    top_cat = max(cat_breakdown.items(), key=lambda x: x[1])[0] if cat_breakdown else "N/A"
    top_cat_pct = (cat_breakdown.get(top_cat, 0) / store["totalRevenue"] * 100) if store["totalRevenue"] else 0

    top_cat_desc = get_product_category_desc(top_cat)
    desc_line = f" {top_cat_desc}" if top_cat_desc else ""

    note_text = (
        f"This location contributes {pct:.1f}% of total network revenue. "
        f"With {fmt_number(store['invoices'])} oil changes and an average ticket of "
        f"${store['avgRevPerInvoice']:.2f}, {store['name']} "
        f"{'leads' if store['rank'] == 1 else 'ranks #' + str(store['rank'])} in the network. "
        f"{top_cat} products represent {top_cat_pct:.0f}% of this location's mix."
        f"{desc_line}"
    )

    n_body = slide.shapes.add_textbox(Inches(notes_x + 0.15), Inches(notes_y + 0.4), Inches(notes_w - 0.3), Inches(notes_h - 0.85))
    tf_nb = n_body.text_frame
    tf_nb.word_wrap = True
    p_nb = tf_nb.paragraphs[0]
    r_nb = p_nb.add_run()
    r_nb.text = note_text
    r_nb.font.size = Pt(8)
    r_nb.font.color.rgb = rgb(C["white"])
    r_nb.font.name = "Calibri"

    n_footer = slide.shapes.add_textbox(Inches(notes_x + 0.15), Inches(notes_y + notes_h - 0.35), Inches(notes_w - 0.3), Inches(0.25))
    tf_nf = n_footer.text_frame
    p_nf = tf_nf.paragraphs[0]
    r_nf = p_nf.add_run()
    r_nf.text = f"Top Product: {store['topProduct']}"
    r_nf.font.size = Pt(8)
    r_nf.font.bold = True
    r_nf.font.color.rgb = rgb(C["goldLight"])
    r_nf.font.name = "Calibri"


def build_next_steps(prs, stores, month_year, total_slides, page_num):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide)
    add_slide_header(slide, "Next Steps & Recommendations", f"Action items for {month_year}")
    add_footer(slide, page_num, total_slides)

    bottom = stores[-1] if stores else None
    top = stores[0] if stores else None
    actions = [
        (
            "Boost Underperformers",
            f"Target {bottom['name']} with promotional offers to increase volume and average ticket size." if bottom else "Identify underperforming locations for targeted promotions.",
            C["purple"],
        ),
        (
            "Expand High Mileage",
            "High Mileage products show strong margins — push upsell training to increase penetration across all locations.",
            C["gold"],
        ),
        (
            "Replicate Top Performer",
            f"Study {top['name']}'s practices and replicate successful strategies across the network." if top else "Analyze top store practices for network-wide adoption.",
            C["purpleMid"],
        ),
        (
            "Monthly Trend Tracking",
            "Establish month-over-month tracking to identify seasonal patterns and measure the impact of promotional campaigns.",
            C["teal"],
        ),
    ]

    for i, (title, body, accent_color) in enumerate(actions):
        col = i % 2
        row = i // 2
        x = 0.5 + col * 4.6
        y = 1.55 + row * 1.85
        w = 4.3
        h = 1.65

        card = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
        card.fill.solid()
        card.fill.fore_color.rgb = rgb(C["white"])
        card.line.fill.background()

        accent = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(0.06), Inches(h))
        accent.fill.solid()
        accent.fill.fore_color.rgb = rgb(accent_color)
        accent.line.fill.background()

        t_box = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.12), Inches(w - 0.4), Inches(0.3))
        tf = t_box.text_frame
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = title
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = rgb(accent_color)
        r.font.name = "Calibri"

        b_box = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.45), Inches(w - 0.4), Inches(h - 0.55))
        tf2 = b_box.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run()
        r2.text = body
        r2.font.size = Pt(9)
        r2.font.color.rgb = rgb(C["darkGray"])
        r2.font.name = "Calibri"


def build_closing_slide(prs, stores, month_year, total_slides):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide, C["dark"])

    if os.path.exists(BG_NEVER_SETTLE):
        slide.shapes.add_picture(
            BG_NEVER_SETTLE, Inches(0), Inches(0), Inches(10), Inches(5.625)
        )
        overlay = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(5.625))
        overlay.fill.solid()
        overlay.fill.fore_color.rgb = rgb(C["dark"])
        _set_shape_alpha(overlay, 60000)
        overlay.line.fill.background()

    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = rgb(C["gold"])
    bar.line.fill.background()

    logo_path = LOGO_EXPERT_WHITE if os.path.exists(LOGO_EXPERT_WHITE) else LOGO_EXPERT_YELLOW
    if os.path.exists(logo_path):
        slide.shapes.add_picture(
            logo_path, Inches(3.5), Inches(0.3), Inches(3.0), Inches(1.25)
        )

    s_box = slide.shapes.add_textbox(Inches(1), Inches(1.6), Inches(8), Inches(0.4))
    tf2 = s_box.text_frame
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = f"Partnership Hub Report | {month_year}"
    r2.font.size = Pt(12)
    r2.font.color.rgb = rgb(C["purpleLight"])
    r2.font.name = "Calibri"

    total_rev = sum(s["totalRevenue"] for s in stores)
    total_inv = sum(s["invoices"] for s in stores)
    avg_rev = total_rev / total_inv if total_inv else 0

    stat_x = 1.5
    stat_y = 2.2
    stat_w = 7.0
    stat_h = 2.0

    stat_bg = slide.shapes.add_shape(1, Inches(stat_x), Inches(stat_y), Inches(stat_w), Inches(stat_h))
    stat_bg.fill.solid()
    stat_bg.fill.fore_color.rgb = rgb(C["purple"])
    stat_bg.line.fill.background()

    gold_top = slide.shapes.add_shape(1, Inches(stat_x), Inches(stat_y), Inches(stat_w), Inches(0.06))
    gold_top.fill.solid()
    gold_top.fill.fore_color.rgb = rgb(C["gold"])
    gold_top.line.fill.background()

    lbl = slide.shapes.add_textbox(Inches(stat_x + 0.3), Inches(stat_y + 0.15), Inches(stat_w - 0.6), Inches(0.25))
    tf_l = lbl.text_frame
    p_l = tf_l.paragraphs[0]
    p_l.alignment = PP_ALIGN.CENTER
    r_l = p_l.add_run()
    r_l.text = f"{month_year} Summary"
    r_l.font.size = Pt(11)
    r_l.font.bold = True
    r_l.font.color.rgb = rgb(C["goldLight"])
    r_l.font.name = "Calibri"

    items = [
        (fmt_currency(total_rev), "Total Revenue"),
        (fmt_number(total_inv), "Oil Changes"),
        (f"${avg_rev:.2f}", "Avg Rev/Invoice"),
        (stores[0]["name"] if stores else "N/A", "Top Store"),
    ]
    col_w = stat_w / 4
    for i, (val, lab) in enumerate(items):
        ix = stat_x + i * col_w + 0.15
        iy = stat_y + 0.55

        vb = slide.shapes.add_textbox(Inches(ix), Inches(iy), Inches(col_w - 0.3), Inches(0.4))
        tf_v = vb.text_frame
        p_v = tf_v.paragraphs[0]
        p_v.alignment = PP_ALIGN.CENTER
        r_v = p_v.add_run()
        r_v.text = val
        font_sz = 16 if len(val) > 10 else 20
        r_v.font.size = Pt(font_sz)
        r_v.font.bold = True
        r_v.font.color.rgb = rgb(C["white"])
        r_v.font.name = "Calibri"

        lb = slide.shapes.add_textbox(Inches(ix), Inches(iy + 0.45), Inches(col_w - 0.3), Inches(0.2))
        tf_lb = lb.text_frame
        p_lb = tf_lb.paragraphs[0]
        p_lb.alignment = PP_ALIGN.CENTER
        r_lb = p_lb.add_run()
        r_lb.text = lab
        r_lb.font.size = Pt(8)
        r_lb.font.color.rgb = rgb(C["purpleLight"])
        r_lb.font.name = "Calibri"

    ns_box = slide.shapes.add_textbox(Inches(1), Inches(4.3), Inches(8), Inches(0.35))
    tf_ns = ns_box.text_frame
    p_ns = tf_ns.paragraphs[0]
    p_ns.alignment = PP_ALIGN.CENTER
    r_ns = p_ns.add_run()
    r_ns.text = "NEVER SETTLE"
    r_ns.font.size = Pt(16)
    r_ns.font.bold = True
    r_ns.font.color.rgb = rgb(C["goldLight"])
    r_ns.font.name = "Calibri"

    contact = slide.shapes.add_textbox(Inches(1), Inches(4.7), Inches(8), Inches(0.5))
    tf_c = contact.text_frame
    p_c = tf_c.paragraphs[0]
    p_c.alignment = PP_ALIGN.CENTER
    r_c = p_c.add_run()
    r_c.text = "ThrottlePro — More Cars. More Loyalty. Less Stress."
    r_c.font.size = Pt(9)
    r_c.font.color.rgb = rgb(C["midGray"])
    r_c.font.name = "Calibri"

    add_footer(slide, total_slides, total_slides)


def build_distribution_map_slide(prs, map_image_path, title, total_slides, page_num):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide, C["offWhite"])

    hdr_bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(1.1))
    hdr_bar.fill.solid()
    hdr_bar.fill.fore_color.rgb = rgb(C["purple"])
    hdr_bar.line.fill.background()

    t_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(7), Inches(0.45))
    tf = t_box.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = rgb(C["white"])
    r.font.name = "Calibri"

    s_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.62), Inches(7), Inches(0.3))
    tf2 = s_box.text_frame
    p2 = tf2.paragraphs[0]
    r2 = p2.add_run()
    r2.text = "ABE Consumer Distribution Territory Map"
    r2.font.size = Pt(11)
    r2.font.color.rgb = rgb(C["goldLight"])
    r2.font.name = "Calibri"

    add_royal_purple_badge(slide)

    map_top = 1.2
    map_bottom = 5.2
    available_h = map_bottom - map_top
    available_w = 9.0
    margin_x = 0.5

    from PIL import Image as PILImage
    try:
        img = PILImage.open(map_image_path)
        img_w, img_h = img.size
        aspect = img_w / img_h
        fit_w = available_w
        fit_h = fit_w / aspect
        if fit_h > available_h:
            fit_h = available_h
            fit_w = fit_h * aspect
        center_x = margin_x + (available_w - fit_w) / 2
        center_y = map_top + (available_h - fit_h) / 2
        slide.shapes.add_picture(
            map_image_path, Inches(center_x), Inches(center_y), Inches(fit_w), Inches(fit_h)
        )
    except Exception:
        slide.shapes.add_picture(
            map_image_path, Inches(margin_x), Inches(map_top), Inches(available_w), Inches(available_h)
        )

    add_footer(slide, page_num, total_slides)


def calculate_total_slides(num_stores, num_maps=0):
    TABLE_TOP = 1.55
    FOOTER_Y = 5.33
    ROW_H = 0.285
    ROWS_PER_PAGE_RANK = int(math.floor((FOOTER_Y - TABLE_TOP) / ROW_H)) - 1

    ROW_Y0 = 1.83
    LEGEND_H = 0.30
    MATRIX_ROW_H = 0.258
    ROWS_PER_PAGE_MATRIX = int(math.floor((5.33 - ROW_Y0 - LEGEND_H) / MATRIX_ROW_H))

    rank_pages = math.ceil(num_stores / ROWS_PER_PAGE_RANK) if num_stores > 0 else 1
    matrix_pages = math.ceil(num_stores / ROWS_PER_PAGE_MATRIX) if num_stores > 0 else 1

    total = (
        1 +  # cover
        1 +  # toc
        1 +  # exec summary kpis
        1 +  # exec observations
        1 +  # revenue overview
        rank_pages +  # ranking table
        matrix_pages +  # performance matrix
        1 +  # product mix
        num_maps +  # distribution maps
        1 +  # section divider
        num_stores +  # deep dives
        1 +  # next steps
        1    # closing
    )
    return total


def generate_report(file_path, output_path=None, map_images=None):
    stores, month_year = parse_excel(file_path)
    maps = map_images or []

    if not output_path:
        month_abbr = month_year.split()[0][:3] if month_year else "Jan"
        year = month_year.split()[-1] if month_year else "2025"
        output_path = f"Royal_Purple_Partnership_Report_{month_abbr}{year}.pptx"

    total_slides = calculate_total_slides(len(stores), len(maps))

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    build_cover_slide(prs, stores, month_year, total_slides)
    build_toc_slide(prs, total_slides)
    build_exec_summary_kpis(prs, stores, month_year, total_slides)
    build_exec_observations(prs, stores, month_year, total_slides)
    build_revenue_overview(prs, stores, month_year, total_slides)

    page = 6
    rank_pages = build_ranking_table(prs, stores, month_year, total_slides, page)
    page += rank_pages

    matrix_pages = build_performance_matrix(prs, stores, month_year, total_slides, page)
    page += matrix_pages

    build_product_mix(prs, stores, month_year, total_slides, page)
    page += 1

    for i, map_info in enumerate(maps):
        map_path = map_info.get("path", map_info) if isinstance(map_info, dict) else map_info
        map_title = map_info.get("title", f"Distribution Map {i + 1}") if isinstance(map_info, dict) else f"Distribution Map {i + 1}"
        build_distribution_map_slide(prs, map_path, map_title, total_slides, page)
        page += 1

    build_section_divider(prs, "Store-Level Deep Dives", f"Individual performance analysis — {month_year}", total_slides, page)
    page += 1

    for store in stores:
        build_deep_dive(prs, store, stores, month_year, total_slides, page)
        page += 1

    build_next_steps(prs, stores, month_year, total_slides, page)
    page += 1

    build_closing_slide(prs, stores, month_year, total_slides)

    prs.save(output_path)
    return output_path, stores, month_year
