import json
import os
import openpyxl as oxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pgeocode
from distribution_data import STATE_DISTRIBUTORS, DISTRIBUTOR_COLORS, ALL_DISTRIBUTORS

CUSTOMERS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "customers.json")
INSTALLER_EXCEL = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "attached_assets", "Installer_Accounts_Not_On_C4C_1772753485907.xlsx"
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

    not_c4c = [c for c in customers if c["type"] == "Installer (Not on C4C)"]
    c4c_matched = [c for c in customers if c["type"] == "Installer (C4C Matched)"]
    distributors_list = [c for c in customers if c["type"] == "Distributor"]

    all_states_code = sorted(set(c["state"] for c in customers if c.get("state")))

    state_names = {}
    for code, info in STATE_DISTRIBUTORS.items():
        state_names[code] = info["state"]

    wb = Workbook()

    # ── Sheet 1: Executive Summary ──
    ws = wb.active
    ws.title = "Executive Summary"
    ws.sheet_properties.tabColor = "1B1464"

    ws.merge_cells("A1:F1")
    ws["A1"] = "Royal Purple — C4C Gap Analysis & Distribution Territory Report"
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
        ("ABE Distributors", len(distributors_list)),
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
        "Total Installers", "Gap %", "ABE Distributor(s)"
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

        dist_info = STATE_DISTRIBUTORS.get(sc, {})
        dists = ", ".join(dist_info.get("distributors", [])) if dist_info else ""
        sname = state_names.get(sc, sc)

        ws2.cell(row=row, column=1, value=sname)
        ws2.cell(row=row, column=2, value=sc)
        ws2.cell(row=row, column=3, value=nc)
        ws2.cell(row=row, column=4, value=mc)
        ws2.cell(row=row, column=5, value=total)
        ws2.cell(row=row, column=6, value=f"{gap:.1f}%")
        ws2.cell(row=row, column=7, value=dists)

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

    # ── Sheet 3: Distribution Territory Map ──
    ws3 = wb.create_sheet("Distribution Territories")
    ws3.sheet_properties.tabColor = "228B22"

    ws3.merge_cells("A1:E1")
    ws3["A1"] = "ABE Distributor Territory Assignments"
    ws3["A1"].font = TITLE_FONT
    ws3.row_dimensions[1].height = 28

    row = 3
    dist_headers = ["ABE Distributor", "States Assigned", "State Count",
                     "Installers (Not on C4C)", "Installers (C4C Matched)"]
    for ci, h in enumerate(dist_headers, 1):
        ws3.cell(row=row, column=ci, value=h)
    _apply_header_row(ws3, row, len(dist_headers))
    row += 1

    for dist_name in ALL_DISTRIBUTORS:
        dist_states = []
        for code, info in sorted(STATE_DISTRIBUTORS.items()):
            if dist_name in info.get("distributors", []):
                dist_states.append(code)

        nc_count = sum(not_c4c_by_state.get(s, 0) for s in dist_states)
        mc_count = sum(c4c_by_state.get(s, 0) for s in dist_states)

        ws3.cell(row=row, column=1, value=dist_name)
        ws3.cell(row=row, column=2, value=", ".join(dist_states))
        ws3.cell(row=row, column=3, value=len(dist_states))
        ws3.cell(row=row, column=4, value=nc_count)
        ws3.cell(row=row, column=5, value=mc_count)

        _apply_data_row(ws3, row, len(dist_headers), alt=(row % 2 == 0))
        for col in [3, 4, 5]:
            ws3.cell(row=row, column=col).alignment = Alignment(horizontal="center")
        row += 1

    row += 2
    ws3.merge_cells(f"A{row}:E{row}")
    ws3[f"A{row}"] = "State-Level Territory Detail"
    ws3[f"A{row}"].font = SUBTITLE_FONT
    row += 1

    detail_headers = ["State", "State Code", "ABE Distributor(s)",
                       "Not on C4C", "C4C Matched"]
    for ci, h in enumerate(detail_headers, 1):
        ws3.cell(row=row, column=ci, value=h)
    _apply_header_row(ws3, row, len(detail_headers))
    row += 1

    for code in sorted(STATE_DISTRIBUTORS.keys()):
        info = STATE_DISTRIBUTORS[code]
        dists = ", ".join(info["distributors"]) if info["distributors"] else "—"
        nc = not_c4c_by_state.get(code, 0)
        mc = c4c_by_state.get(code, 0)

        ws3.cell(row=row, column=1, value=info["state"])
        ws3.cell(row=row, column=2, value=code)
        ws3.cell(row=row, column=3, value=dists)
        ws3.cell(row=row, column=4, value=nc)
        ws3.cell(row=row, column=5, value=mc)

        _apply_data_row(ws3, row, len(detail_headers), alt=(row % 2 == 0))
        for col in [2, 4, 5]:
            ws3.cell(row=row, column=col).alignment = Alignment(horizontal="center")
        row += 1

    _auto_width(ws3, len(dist_headers))

    # ── Sheet 4: Not on C4C (Full List) ──
    ws4 = wb.create_sheet("Not on C4C — Full List")
    ws4.sheet_properties.tabColor = "D97706"

    ws4.merge_cells("A1:G1")
    ws4["A1"] = f"Installer Accounts NOT on C4C ({len(not_c4c)} accounts)"
    ws4["A1"].font = TITLE_FONT
    ws4.row_dimensions[1].height = 28

    row = 3
    list_headers = ["Store Name", "Address", "City", "State", "Zip",
                     "ABE Distributor", "Latitude", "Longitude"]
    for ci, h in enumerate(list_headers, 1):
        ws4.cell(row=row, column=ci, value=h)
    _apply_header_row(ws4, row, len(list_headers))
    row += 1

    for c in sorted(not_c4c, key=lambda x: (x["state"], x["store_name"])):
        dist_info = STATE_DISTRIBUTORS.get(c["state"], {})
        dists = ", ".join(dist_info.get("distributors", [])) if dist_info else ""

        ws4.cell(row=row, column=1, value=c["store_name"])
        ws4.cell(row=row, column=2, value=c["address"])
        ws4.cell(row=row, column=3, value=c["city"])
        ws4.cell(row=row, column=4, value=c["state"])
        ws4.cell(row=row, column=5, value=c["zip"])
        ws4.cell(row=row, column=6, value=dists)
        ws4.cell(row=row, column=7, value=c["latitude"])
        ws4.cell(row=row, column=8, value=c["longitude"])

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
        dist_info = STATE_DISTRIBUTORS.get(c["state"], {})
        dists = ", ".join(dist_info.get("distributors", [])) if dist_info else ""

        ws5.cell(row=row, column=1, value=c["store_name"])
        ws5.cell(row=row, column=2, value=c["address"])
        ws5.cell(row=row, column=3, value=c["city"])
        ws5.cell(row=row, column=4, value=c["state"])
        ws5.cell(row=row, column=5, value=c["zip"])
        ws5.cell(row=row, column=6, value=dists)
        ws5.cell(row=row, column=7, value=c["latitude"])
        ws5.cell(row=row, column=8, value=c["longitude"])

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
                         "Total", "Gap %", "ABE Distributor(s)"]
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
        dist_info = STATE_DISTRIBUTORS.get(sc, {})
        dists = ", ".join(dist_info.get("distributors", [])) if dist_info else ""
        state_priority.append((sname, sc, nc, mc, total, gap, dists))

    state_priority.sort(key=lambda x: -x[2])

    for rank, (sname, sc, nc, mc, total, gap, dists) in enumerate(state_priority, 1):
        ws6.cell(row=row, column=1, value=rank)
        ws6.cell(row=row, column=2, value=f"{sname} ({sc})")
        ws6.cell(row=row, column=3, value=nc)
        ws6.cell(row=row, column=4, value=mc)
        ws6.cell(row=row, column=5, value=total)
        ws6.cell(row=row, column=6, value=f"{gap:.1f}%")
        ws6.cell(row=row, column=7, value=dists)

        _apply_data_row(ws6, row, len(priority_headers), alt=(row % 2 == 0))
        for col in [1, 3, 4, 5, 6]:
            ws6.cell(row=row, column=col).alignment = Alignment(horizontal="center")

        if rank <= 10:
            ws6.cell(row=row, column=3).fill = AMBER_FILL

        row += 1

    _auto_width(ws6, len(priority_headers))

    # ── Sheet 7: Failed to Geolocate ──
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
        "distributors": len(distributors_list),
        "states": len(all_states_code),
        "failed_geo": len(failed),
        "sheets": 7,
    }
