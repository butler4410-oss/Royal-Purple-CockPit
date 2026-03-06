import json
import os
import openpyxl as oxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pgeocode
US_STATE_NAMES = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas",
    "CA": "California", "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware",
    "DC": "District of Columbia", "FL": "Florida", "GA": "Georgia", "HI": "Hawaii",
    "ID": "Idaho", "IL": "Illinois", "IN": "Indiana", "IA": "Iowa",
    "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine",
    "MD": "Maryland", "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota",
    "MS": "Mississippi", "MO": "Missouri", "MT": "Montana", "NE": "Nebraska",
    "NV": "Nevada", "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico",
    "NY": "New York", "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio",
    "OK": "Oklahoma", "OR": "Oregon", "PA": "Pennsylvania", "PR": "Puerto Rico",
    "RI": "Rhode Island", "SC": "South Carolina", "SD": "South Dakota",
    "TN": "Tennessee", "TX": "Texas", "UT": "Utah", "VT": "Vermont",
    "VA": "Virginia", "WA": "Washington", "WV": "West Virginia",
    "WI": "Wisconsin", "WY": "Wyoming",
}

CUSTOMERS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "customers.json")
INSTALLER_EXCEL = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "attached_assets", "Installer_Accounts_Not_On_C4C_1772753485907.xlsx"
)
C4C_EXCEL = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "attached_assets", "Royal_Purple_C4C_Installer_List_1772754933619.xlsx"
)
PROMO_EXCEL = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "attached_assets", "Installer_Promotion_Participation_12.13.25_1772754933620.xlsx"
)

HEADER_FILL = PatternFill(start_color="1B1464", end_color="1B1464", fill_type="solid")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
SUBHEADER_FILL = PatternFill(start_color="4B2D8A", end_color="4B2D8A", fill_type="solid")
SUBHEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
DATA_FONT = Font(name="Calibri", size=10)
BOLD_FONT = Font(name="Calibri", bold=True, size=10)
TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="1B1464")
SUBTITLE_FONT = Font(name="Calibri", bold=True, size=11, color="4B2D8A")
AMBER_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
GREEN_FILL = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
LIGHT_GRAY_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin", color="E2E8F0"),
    right=Side(style="thin", color="E2E8F0"),
    top=Side(style="thin", color="E2E8F0"),
    bottom=Side(style="thin", color="E2E8F0"),
)


def _cross_analyze_lists():
    from collections import Counter

    result = {
        "c4c_rows": 0, "c4c_unique": 0, "c4c_dupes": [],
        "promo_rows": 0, "promo_unique": 0, "promo_dupes": [],
        "matched": 0, "only_c4c": 0, "only_promo": 0,
        "matched_names": [], "only_c4c_names": [], "only_promo_names": [],
    }

    c4c_accounts = []
    promo_accounts = []

    if os.path.exists(C4C_EXCEL):
        wb = oxl.load_workbook(C4C_EXCEL)
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            name = str(row[1]).strip() if row[1] else ""
            if not name:
                continue
            c4c_accounts.append({
                "name": name,
                "acct_id": str(row[2]).strip() if row[2] else "",
                "street": str(row[3]).strip() if row[3] else "",
                "city": str(row[4]).strip() if row[4] else "",
                "state": str(row[5]).strip() if row[5] else "",
                "zip": str(row[6]).strip() if row[6] else "",
            })

    if os.path.exists(PROMO_EXCEL):
        wb = oxl.load_workbook(PROMO_EXCEL)
        ws = wb['Summary'] if 'Summary' in wb.sheetnames else wb[wb.sheetnames[0]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            name = str(row[0]).strip() if row[0] else ""
            if not name:
                continue
            promo_accounts.append({
                "name": name,
                "address": str(row[1]).strip() if row[1] else "",
                "city": str(row[3]).strip() if row[3] else "",
                "state": str(row[4]).strip() if row[4] else "",
                "zip": str(row[5]).strip() if row[5] else "",
            })

    result["c4c_rows"] = len(c4c_accounts)
    result["promo_rows"] = len(promo_accounts)

    c4c_names = set(a["name"].upper() for a in c4c_accounts)
    promo_names = set(a["name"].upper() for a in promo_accounts)

    result["c4c_unique"] = len(c4c_names)
    result["promo_unique"] = len(promo_names)

    matched = c4c_names & promo_names
    only_c4c = c4c_names - promo_names
    only_promo = promo_names - c4c_names

    result["matched"] = len(matched)
    result["only_c4c"] = len(only_c4c)
    result["only_promo"] = len(only_promo)

    c4c_name_counts = Counter(a["name"].upper() for a in c4c_accounts)
    for name, cnt in sorted(c4c_name_counts.items(), key=lambda x: (-x[1], x[0])):
        if cnt <= 1:
            continue
        entries = [a for a in c4c_accounts if a["name"].upper() == name]
        locations = "; ".join(f"{e['city']}, {e['state']}" for e in entries)
        result["c4c_dupes"].append({"name": entries[0]["name"], "count": cnt, "locations": locations})

    promo_name_counts = Counter(a["name"].upper() for a in promo_accounts)
    for name, cnt in sorted(promo_name_counts.items(), key=lambda x: (-x[1], x[0])):
        if cnt <= 1:
            continue
        entries = [a for a in promo_accounts if a["name"].upper() == name]
        locations = "; ".join(f"{e['city']}, {e['state']}" for e in entries[:10])
        if len(entries) > 10:
            locations += f" (+{len(entries)-10} more)"
        result["promo_dupes"].append({"name": entries[0]["name"], "count": cnt, "locations": locations})

    for nm in sorted(matched):
        c4c_entry = next((a for a in c4c_accounts if a["name"].upper() == nm), {})
        promo_entry = next((a for a in promo_accounts if a["name"].upper() == nm), {})
        result["matched_names"].append({
            "name": c4c_entry.get("name", nm),
            "c4c_city": f"{c4c_entry.get('city', '')}, {c4c_entry.get('state', '')}",
            "promo_city": f"{promo_entry.get('city', '')}, {promo_entry.get('state', '')}",
        })

    for nm in sorted(only_c4c):
        entry = next((a for a in c4c_accounts if a["name"].upper() == nm), {})
        result["only_c4c_names"].append({
            "name": entry.get("name", nm),
            "city": entry.get("city", ""),
            "state": entry.get("state", ""),
            "acct_id": entry.get("acct_id", ""),
        })

    for nm in sorted(only_promo):
        entry = next((a for a in promo_accounts if a["name"].upper() == nm), {})
        result["only_promo_names"].append({
            "name": entry.get("name", nm),
            "city": entry.get("city", ""),
            "state": entry.get("state", ""),
        })

    return result


def _get_failed_geolocations():
    if not os.path.exists(INSTALLER_EXCEL):
        return []

    nomi = pgeocode.Nominatim('us')
    wb = oxl.load_workbook(INSTALLER_EXCEL)
    failed = []
    seen = set()

    for sheet_name in ['Not on C4C List', 'Matched Accounts']:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        source = "Not on C4C" if "Not" in sheet_name else "C4C Matched"

        for row in ws.iter_rows(min_row=2, values_only=True):
            name = row[0]
            address = row[1]
            city = row[3]
            state = row[4]
            zipcode = row[5]
            phone = row[6] if len(row) > 6 else None
            email = row[7] if len(row) > 7 else None

            if not name:
                continue
            name = str(name).strip()
            if name in ['-', '(RESIDENCE)'] or name.startswith('('):
                continue

            address = str(address).strip() if address else ""
            city = str(city).strip() if city else ""
            state = str(state).strip().upper() if state else ""
            z = str(zipcode).strip() if zipcode else ""
            z_clean = z.split('-')[0][:5]
            phone = str(phone).strip() if phone else ""
            email = str(email).strip() if email else ""
            if phone == "None":
                phone = ""
            if email == "None":
                email = ""

            key = (name.upper(), city.upper(), state)
            if key in seen:
                continue
            seen.add(key)

            reason = ""
            if not city or not state:
                reason = "Missing city/state"
            elif len(state) != 2 or not state.isalpha():
                reason = f"Invalid state code: {state}"
            else:
                if len(z_clean) == 5 and z_clean.isdigit():
                    result = nomi.query_postal_code(z_clean)
                    if result is None or (result.latitude != result.latitude):
                        reason = f"Zip code not found: {z_clean}"
                    else:
                        continue
                else:
                    reason = f"Invalid zip: {z}"

            failed.append({
                "store_name": name,
                "address": address,
                "city": city,
                "state": state,
                "zip": z,
                "phone": phone,
                "email": email,
                "source": source,
                "reason": reason,
            })

    return failed


def _load_customers():
    if os.path.exists(CUSTOMERS_PATH):
        with open(CUSTOMERS_PATH, "r") as f:
            return json.load(f)
    return []


def _apply_header_row(ws, row, num_cols):
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def _apply_data_row(ws, row, num_cols, alt=False):
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = DATA_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
        if alt:
            cell.fill = LIGHT_GRAY_FILL


def _auto_width(ws, num_cols, max_width=40):
    for col in range(1, num_cols + 1):
        max_len = 0
        letter = get_column_letter(col)
        for row in ws.iter_rows(min_col=col, max_col=col, values_only=False):
            for cell in row:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max_len + 3, max_width)


def generate_c4c_report(output_path):
    customers = _load_customers()

    not_c4c = [c for c in customers if c["type"] == "Promo Only (Not on C4C)"]
    c4c_matched = [c for c in customers if c["type"] in ("On Both Lists", "C4C Only")]

    all_states_code = sorted(set(c["state"] for c in customers if c.get("state")))

    state_names = dict(US_STATE_NAMES)

    wb = Workbook()

    # ── Sheet 1: Executive Summary ──
    ws = wb.active
    ws.title = "Executive Summary"
    ws.sheet_properties.tabColor = "1B1464"

    ws.merge_cells("A1:F1")
    ws["A1"] = "Royal Purple — C4C Gap Analysis Report"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:F2")
    ws["A2"] = "Prepared by ThrottlePro"
    ws["A2"].font = Font(name="Calibri", italic=True, size=10, color="64748B")

    row = 4
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"] = "Installer Account Summary"
    ws[f"A{row}"].font = SUBTITLE_FONT
    row += 1

    summary_data = [
        ("Total Installer Accounts on Map", len(not_c4c) + len(c4c_matched)),
        ("Installer Accounts NOT on C4C", len(not_c4c)),
        ("Installer Accounts Matched (on C4C)", len(c4c_matched)),
        ("States Covered", len(all_states_code)),
    ]

    headers = ["Metric", "Count"]
    for ci, h in enumerate(headers, 1):
        ws.cell(row=row, column=ci, value=h)
    _apply_header_row(ws, row, 2)
    row += 1

    for label, val in summary_data:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=val)
        _apply_data_row(ws, row, 2, alt=(row % 2 == 0))
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        row += 1

    row += 2
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"] = "C4C Gap Percentage"
    ws[f"A{row}"].font = SUBTITLE_FONT
    row += 1

    total_installers = len(not_c4c) + len(c4c_matched)
    gap_pct = len(not_c4c) / total_installers * 100 if total_installers else 0
    matched_pct = len(c4c_matched) / total_installers * 100 if total_installers else 0

    gap_data = [
        ("Not on C4C", f"{len(not_c4c)}", f"{gap_pct:.1f}%"),
        ("On C4C (Matched)", f"{len(c4c_matched)}", f"{matched_pct:.1f}%"),
        ("Total", f"{total_installers}", "100%"),
    ]

    for ci, h in enumerate(["Category", "Count", "Percentage"], 1):
        ws.cell(row=row, column=ci, value=h)
    _apply_header_row(ws, row, 3)
    row += 1

    for label, cnt, pct in gap_data:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=cnt)
        ws.cell(row=row, column=3, value=pct)
        _apply_data_row(ws, row, 3, alt=(row % 2 == 0))
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=3).alignment = Alignment(horizontal="center")
        if label == "Total":
            for c in range(1, 4):
                ws.cell(row=row, column=c).font = BOLD_FONT
        row += 1

    _auto_width(ws, 3)
    ws.column_dimensions["A"].width = 40

    # ── Sheet 2: State Breakdown ──
    ws2 = wb.create_sheet("State Breakdown")
    ws2.sheet_properties.tabColor = "4B2D8A"

    ws2.merge_cells("A1:G1")
    ws2["A1"] = "Installer Accounts by State — C4C Gap Analysis"
    ws2["A1"].font = TITLE_FONT
    ws2.row_dimensions[1].height = 28

    row = 3
    state_headers = [
        "State", "State Code", "Not on C4C", "C4C Matched",
        "Total Installers", "Gap %"
    ]
    for ci, h in enumerate(state_headers, 1):
        ws2.cell(row=row, column=ci, value=h)
    _apply_header_row(ws2, row, len(state_headers))
    row += 1

    not_c4c_by_state = {}
    for c in not_c4c:
        not_c4c_by_state[c["state"]] = not_c4c_by_state.get(c["state"], 0) + 1

    c4c_by_state = {}
    for c in c4c_matched:
        c4c_by_state[c["state"]] = c4c_by_state.get(c["state"], 0) + 1

    all_state_codes = sorted(set(list(not_c4c_by_state.keys()) + list(c4c_by_state.keys())))

    total_not_c4c = 0
    total_c4c = 0
    total_all = 0

    for sc in all_state_codes:
        nc = not_c4c_by_state.get(sc, 0)
        mc = c4c_by_state.get(sc, 0)
        total = nc + mc
        gap = nc / total * 100 if total else 0

        total_not_c4c += nc
        total_c4c += mc
        total_all += total

        sname = state_names.get(sc, sc)

        ws2.cell(row=row, column=1, value=sname)
        ws2.cell(row=row, column=2, value=sc)
        ws2.cell(row=row, column=3, value=nc)
        ws2.cell(row=row, column=4, value=mc)
        ws2.cell(row=row, column=5, value=total)
        ws2.cell(row=row, column=6, value=f"{gap:.1f}%")

        _apply_data_row(ws2, row, len(state_headers), alt=(row % 2 == 0))
        for col in [2, 3, 4, 5, 6]:
            ws2.cell(row=row, column=col).alignment = Alignment(horizontal="center")

        if nc > 50:
            ws2.cell(row=row, column=3).fill = AMBER_FILL
        if mc > 0:
            ws2.cell(row=row, column=4).fill = GREEN_FILL

        row += 1

    total_gap = total_not_c4c / total_all * 100 if total_all else 0
    ws2.cell(row=row, column=1, value="TOTAL")
    ws2.cell(row=row, column=3, value=total_not_c4c)
    ws2.cell(row=row, column=4, value=total_c4c)
    ws2.cell(row=row, column=5, value=total_all)
    ws2.cell(row=row, column=6, value=f"{total_gap:.1f}%")
    for col in range(1, len(state_headers) + 1):
        cell = ws2.cell(row=row, column=col)
        cell.font = BOLD_FONT
        cell.border = THIN_BORDER
        if col in [3, 4, 5, 6]:
            cell.alignment = Alignment(horizontal="center")

    _auto_width(ws2, len(state_headers))

    # ── Sheet 3: Not on C4C (Full List) ──
    ws4 = wb.create_sheet("Not on C4C — Full List")
    ws4.sheet_properties.tabColor = "D97706"

    ws4.merge_cells("A1:G1")
    ws4["A1"] = f"Installer Accounts NOT on C4C ({len(not_c4c)} accounts)"
    ws4["A1"].font = TITLE_FONT
    ws4.row_dimensions[1].height = 28

    row = 3
    list_headers = ["Store Name", "Address", "City", "State", "Zip",
                     "Latitude", "Longitude"]
    for ci, h in enumerate(list_headers, 1):
        ws4.cell(row=row, column=ci, value=h)
    _apply_header_row(ws4, row, len(list_headers))
    row += 1

    for c in sorted(not_c4c, key=lambda x: (x["state"], x["store_name"])):
        ws4.cell(row=row, column=1, value=c["store_name"])
        ws4.cell(row=row, column=2, value=c["address"])
        ws4.cell(row=row, column=3, value=c["city"])
        ws4.cell(row=row, column=4, value=c["state"])
        ws4.cell(row=row, column=5, value=c["zip"])
        ws4.cell(row=row, column=6, value=c["latitude"])
        ws4.cell(row=row, column=7, value=c["longitude"])

        _apply_data_row(ws4, row, len(list_headers), alt=(row % 2 == 0))
        ws4.cell(row=row, column=4).alignment = Alignment(horizontal="center")
        row += 1

    _auto_width(ws4, len(list_headers))

    # ── Sheet 5: C4C Matched (Full List) ──
    ws5 = wb.create_sheet("C4C Matched — Full List")
    ws5.sheet_properties.tabColor = "16A34A"

    ws5.merge_cells("A1:G1")
    ws5["A1"] = f"Installer Accounts Matched on C4C ({len(c4c_matched)} accounts)"
    ws5["A1"].font = TITLE_FONT
    ws5.row_dimensions[1].height = 28

    row = 3
    for ci, h in enumerate(list_headers, 1):
        ws5.cell(row=row, column=ci, value=h)
    _apply_header_row(ws5, row, len(list_headers))
    row += 1

    for c in sorted(c4c_matched, key=lambda x: (x["state"], x["store_name"])):
        ws5.cell(row=row, column=1, value=c["store_name"])
        ws5.cell(row=row, column=2, value=c["address"])
        ws5.cell(row=row, column=3, value=c["city"])
        ws5.cell(row=row, column=4, value=c["state"])
        ws5.cell(row=row, column=5, value=c["zip"])
        ws5.cell(row=row, column=6, value=c["latitude"])
        ws5.cell(row=row, column=7, value=c["longitude"])

        _apply_data_row(ws5, row, len(list_headers), alt=(row % 2 == 0))
        ws5.cell(row=row, column=4).alignment = Alignment(horizontal="center")
        row += 1

    _auto_width(ws5, len(list_headers))

    # ── Sheet 6: Top Priority States ──
    ws6 = wb.create_sheet("Top Priority States")
    ws6.sheet_properties.tabColor = "DC143C"

    ws6.merge_cells("A1:F1")
    ws6["A1"] = "States with Highest C4C Gap — Priority for Onboarding"
    ws6["A1"].font = TITLE_FONT
    ws6.row_dimensions[1].height = 28

    row = 3
    priority_headers = ["Rank", "State", "Not on C4C", "C4C Matched",
                         "Total", "Gap %"]
    for ci, h in enumerate(priority_headers, 1):
        ws6.cell(row=row, column=ci, value=h)
    _apply_header_row(ws6, row, len(priority_headers))
    row += 1

    state_priority = []
    for sc in all_state_codes:
        nc = not_c4c_by_state.get(sc, 0)
        mc = c4c_by_state.get(sc, 0)
        total = nc + mc
        if total == 0:
            continue
        gap = nc / total * 100
        sname = state_names.get(sc, sc)
        state_priority.append((sname, sc, nc, mc, total, gap))

    state_priority.sort(key=lambda x: -x[2])

    for rank, (sname, sc, nc, mc, total, gap) in enumerate(state_priority, 1):
        ws6.cell(row=row, column=1, value=rank)
        ws6.cell(row=row, column=2, value=f"{sname} ({sc})")
        ws6.cell(row=row, column=3, value=nc)
        ws6.cell(row=row, column=4, value=mc)
        ws6.cell(row=row, column=5, value=total)
        ws6.cell(row=row, column=6, value=f"{gap:.1f}%")

        _apply_data_row(ws6, row, len(priority_headers), alt=(row % 2 == 0))
        for col in [1, 3, 4, 5, 6]:
            ws6.cell(row=row, column=col).alignment = Alignment(horizontal="center")

        if rank <= 10:
            ws6.cell(row=row, column=3).fill = AMBER_FILL

        row += 1

    _auto_width(ws6, len(priority_headers))

    # ── Sheet 7: Reconciliation Summary ──
    xref = _cross_analyze_lists()

    ws7 = wb.create_sheet("Reconciliation")
    ws7.sheet_properties.tabColor = "2563EB"

    ws7.merge_cells("A1:D1")
    ws7["A1"] = "C4C vs Promo List — Cross-Reference & Duplicates"
    ws7["A1"].font = TITLE_FONT
    ws7.row_dimensions[1].height = 28

    row = 3
    ws7.merge_cells(f"A{row}:B{row}")
    ws7[f"A{row}"] = "Source File Totals"
    ws7[f"A{row}"].font = SUBTITLE_FONT
    row += 1

    for ci, h in enumerate(["Metric", "Count"], 1):
        ws7.cell(row=row, column=ci, value=h)
    _apply_header_row(ws7, row, 2)
    row += 1

    totals = [
        ("C4C Installer List — Total Rows", xref["c4c_rows"]),
        ("C4C Installer List — Unique Names", xref["c4c_unique"]),
        ("Promo Participation List — Total Rows", xref["promo_rows"]),
        ("Promo Participation List — Unique Names", xref["promo_unique"]),
        ("", ""),
        ("Matched (name on BOTH lists)", xref["matched"]),
        ("On C4C ONLY (not on Promo list)", xref["only_c4c"]),
        ("On Promo ONLY (not on C4C)", xref["only_promo"]),
        ("", ""),
        ("C4C Internal Duplicates (same name, multiple rows)", len(xref["c4c_dupes"])),
        ("Promo Internal Duplicates (same name, multiple rows)", len(xref["promo_dupes"])),
    ]

    for label, val in totals:
        ws7.cell(row=row, column=1, value=label)
        ws7.cell(row=row, column=2, value=val if val != "" else "")
        _apply_data_row(ws7, row, 2, alt=(row % 2 == 0))
        ws7.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        if label.startswith("Matched") or label.startswith("On C4C") or label.startswith("On Promo"):
            ws7.cell(row=row, column=1).font = BOLD_FONT
            ws7.cell(row=row, column=2).font = BOLD_FONT
        row += 1

    _auto_width(ws7, 2)
    ws7.column_dimensions["A"].width = 50

    # ── Sheet 8: C4C Duplicates ──
    ws8 = wb.create_sheet("C4C Duplicates")
    ws8.sheet_properties.tabColor = "DC143C"

    ws8.merge_cells("A1:D1")
    ws8["A1"] = f"C4C Installer List — Internal Duplicates ({len(xref['c4c_dupes'])} names with multiple entries)"
    ws8["A1"].font = TITLE_FONT
    ws8.row_dimensions[1].height = 28

    row = 3
    dupe_headers = ["Installer Name", "Occurrences", "Locations"]
    for ci, h in enumerate(dupe_headers, 1):
        ws8.cell(row=row, column=ci, value=h)
    _apply_header_row(ws8, row, len(dupe_headers))
    row += 1

    for d in xref["c4c_dupes"]:
        ws8.cell(row=row, column=1, value=d["name"])
        ws8.cell(row=row, column=2, value=d["count"])
        ws8.cell(row=row, column=3, value=d["locations"])
        _apply_data_row(ws8, row, len(dupe_headers), alt=(row % 2 == 0))
        ws8.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        row += 1

    _auto_width(ws8, len(dupe_headers))
    ws8.column_dimensions["C"].width = 80

    # ── Sheet 9: Promo Duplicates ──
    ws9 = wb.create_sheet("Promo Duplicates")
    ws9.sheet_properties.tabColor = "D97706"

    ws9.merge_cells("A1:D1")
    ws9["A1"] = f"Promo Participation List — Internal Duplicates ({len(xref['promo_dupes'])} names with multiple entries)"
    ws9["A1"].font = TITLE_FONT
    ws9.row_dimensions[1].height = 28

    row = 3
    for ci, h in enumerate(dupe_headers, 1):
        ws9.cell(row=row, column=ci, value=h)
    _apply_header_row(ws9, row, len(dupe_headers))
    row += 1

    for d in xref["promo_dupes"]:
        ws9.cell(row=row, column=1, value=d["name"])
        ws9.cell(row=row, column=2, value=d["count"])
        ws9.cell(row=row, column=3, value=d["locations"])
        _apply_data_row(ws9, row, len(dupe_headers), alt=(row % 2 == 0))
        ws9.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        row += 1

    _auto_width(ws9, len(dupe_headers))
    ws9.column_dimensions["C"].width = 80

    # ── Sheet 10: On C4C Only ──
    ws10 = wb.create_sheet("On C4C Only")
    ws10.sheet_properties.tabColor = "4B2D8A"

    ws10.merge_cells("A1:D1")
    ws10["A1"] = f"Accounts on C4C but NOT on Promo List ({xref['only_c4c']} accounts)"
    ws10["A1"].font = TITLE_FONT
    ws10.row_dimensions[1].height = 28

    row = 3
    c4c_only_headers = ["Installer Name", "City", "State", "Account ID"]
    for ci, h in enumerate(c4c_only_headers, 1):
        ws10.cell(row=row, column=ci, value=h)
    _apply_header_row(ws10, row, len(c4c_only_headers))
    row += 1

    for entry in xref["only_c4c_names"]:
        ws10.cell(row=row, column=1, value=entry["name"])
        ws10.cell(row=row, column=2, value=entry["city"])
        ws10.cell(row=row, column=3, value=entry["state"])
        ws10.cell(row=row, column=4, value=entry["acct_id"])
        _apply_data_row(ws10, row, len(c4c_only_headers), alt=(row % 2 == 0))
        row += 1

    _auto_width(ws10, len(c4c_only_headers))

    # ── Sheet 11: Failed to Geolocate ──
    failed = _get_failed_geolocations()

    ws7 = wb.create_sheet("Failed to Geolocate")
    ws7.sheet_properties.tabColor = "94A3B8"

    ws7.merge_cells("A1:I1")
    ws7["A1"] = f"Installer Accounts — Failed to Geolocate ({len(failed)} accounts)"
    ws7["A1"].font = TITLE_FONT
    ws7.row_dimensions[1].height = 28

    ws7.merge_cells("A2:I2")
    ws7["A2"] = "These accounts could not be placed on the map due to invalid or missing address data."
    ws7["A2"].font = Font(name="Calibri", italic=True, size=10, color="64748B")

    row = 4
    fail_headers = ["Store Name", "Address", "City", "State", "Zip",
                     "Phone", "Email", "Source Sheet", "Reason"]
    for ci, h in enumerate(fail_headers, 1):
        ws7.cell(row=row, column=ci, value=h)
    _apply_header_row(ws7, row, len(fail_headers))
    row += 1

    for f in sorted(failed, key=lambda x: (x["reason"], x["state"], x["store_name"])):
        ws7.cell(row=row, column=1, value=f["store_name"])
        ws7.cell(row=row, column=2, value=f["address"])
        ws7.cell(row=row, column=3, value=f["city"])
        ws7.cell(row=row, column=4, value=f["state"])
        ws7.cell(row=row, column=5, value=f["zip"])
        ws7.cell(row=row, column=6, value=f["phone"])
        ws7.cell(row=row, column=7, value=f["email"])
        ws7.cell(row=row, column=8, value=f["source"])
        ws7.cell(row=row, column=9, value=f["reason"])

        _apply_data_row(ws7, row, len(fail_headers), alt=(row % 2 == 0))
        ws7.cell(row=row, column=4).alignment = Alignment(horizontal="center")
        row += 1

    _auto_width(ws7, len(fail_headers))

    wb.save(output_path)

    return {
        "total_installers": total_installers,
        "not_on_c4c": len(not_c4c),
        "c4c_matched": len(c4c_matched),
        "states": len(all_states_code),
        "failed_geo": len(failed),
        "c4c_dupes": len(xref["c4c_dupes"]),
        "promo_dupes": len(xref["promo_dupes"]),
        "cross_matched": xref["matched"],
        "sheets": 10,
    }
