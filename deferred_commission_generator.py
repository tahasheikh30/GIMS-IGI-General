"""
Writes the Deferred Commission Expense output workbook. Styled
identically to upr_generator.py / reinsurance_generator.py (same navy/
Arial/thin-border look as the rest of GIMS), and structured the same
way: a "Summary" sheet up front, then a single "Deferred Commission
Working" sheet holding BOTH tables stacked (Table 1 - Commission Paid/
Payable Distribution, then Table 2 - Deferred Commission Expense), both
driven by live Excel SUM formulas so the sheet stays auditable and
recalculates correctly if a cell is ever edited by hand - see
upr_generator.py's own module docstring for the full rationale behind
this layout, carried over unchanged here.
"""
import os
import subprocess
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import deferred_commission_processor as dcp

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial")
BOLD_FONT = Font(name="Arial", bold=True)
TOTAL_FILL = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
DRIVER_FONT = Font(name="Arial", size=9, italic=True, color="6C7A89")
NOTE_FONT = Font(name="Arial", size=9, italic=True, color="6C7A89")
SUBTITLE_FONT = Font(name="Arial", size=10, color="6C7A89")
SECTION_TITLE_FONT = Font(name="Arial", bold=True, size=12, color="1F4E78")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM_FORMAT = "#,##0;(#,##0);-"

SUMMARY_SHEET = "Summary"
WORKING_SHEET = "Deferred Commission Working"

DEPT_COL_WIDTH = 22
DATA_COL_WIDTH = 16
TOTAL_COL_WIDTH = 22


def _col_letter(idx):
    return get_column_letter(idx)


def _write_table_block(ws, start_row, section_label, departments, table_dict, value_formula_fn=None):
    """Identical shape to upr_generator._write_table_block / reinsurance_
    generator._write_table_block - a section label, a header row, one row
    per department, and a bold Total row with SUM formulas."""
    ws.cell(row=start_row, column=2, value=section_label).font = SECTION_TITLE_FONT
    header_row = start_row + 1
    data_start_row = header_row + 1

    columns = dcp.TABLE_COLUMNS + [dcp.TOTAL_LABEL]
    ws.cell(row=header_row, column=2, value="Department").font = HEADER_FONT
    ws.cell(row=header_row, column=2).fill = HEADER_FILL
    ws.cell(row=header_row, column=2).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.cell(row=header_row, column=2).border = BORDER
    for j, col in enumerate(columns, start=3):
        c = ws.cell(row=header_row, column=j, value=col)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[header_row].height = 30

    n_depts = len(departments)
    total_row = data_start_row + n_depts

    for i, dept in enumerate(departments):
        r = data_start_row + i
        name_cell = ws.cell(row=r, column=2, value=dept)
        name_cell.font = BODY_FONT
        name_cell.border = BORDER
        for j, col in enumerate(dcp.TABLE_COLUMNS, start=3):
            if value_formula_fn is not None:
                val = value_formula_fn(dept, r, j, col)
            else:
                val = table_dict[dept][col]
            cell = ws.cell(row=r, column=j, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.number_format = NUM_FORMAT
        first_col_letter = _col_letter(3)
        last_col_letter = _col_letter(2 + len(dcp.TABLE_COLUMNS))
        total_cell = ws.cell(row=r, column=3 + len(dcp.TABLE_COLUMNS),
                              value=f"=SUM({first_col_letter}{r}:{last_col_letter}{r})")
        total_cell.font = BOLD_FONT
        total_cell.border = BORDER
        total_cell.number_format = NUM_FORMAT

    ws.cell(row=total_row, column=2, value=dcp.TOTAL_LABEL).font = BOLD_FONT
    ws.cell(row=total_row, column=2).fill = TOTAL_FILL
    ws.cell(row=total_row, column=2).border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)
    for j in range(3, 3 + len(dcp.TABLE_COLUMNS) + 1):
        col_letter = _col_letter(j)
        cell = ws.cell(row=total_row, column=j, value=f"=SUM({col_letter}{data_start_row}:{col_letter}{total_row - 1})")
        cell.font = BOLD_FONT
        cell.fill = TOTAL_FILL
        cell.number_format = NUM_FORMAT
        cell.border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)

    return header_row, data_start_row, total_row


def _apply_working_sheet_column_widths(ws):
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = DEPT_COL_WIDTH
    for j, _col in enumerate(dcp.TABLE_COLUMNS, start=3):
        ws.column_dimensions[_col_letter(j)].width = DATA_COL_WIDTH
    ws.column_dimensions[_col_letter(3 + len(dcp.TABLE_COLUMNS))].width = TOTAL_COL_WIDTH


def _write_driver_rows(ws, start_row, period_start, label):
    """Identical to upr_generator._write_driver_rows / reinsurance_
    generator._write_driver_rows - the numerator/denominator driver rows
    Table 2's own formulas reference, kept visible and auditable rather
    than baked invisibly into each cell."""
    month_to_position = dcp.up._month_to_period_position(period_start)
    denom = 24 if label == "1/24" else 6

    period_row = start_row
    year_row = start_row + 1

    ws.cell(row=period_row, column=2, value=f"{label} Period").font = DRIVER_FONT
    ws.cell(row=year_row, column=2, value=f"{label} Denominator").font = DRIVER_FONT

    for j, col in enumerate(dcp.TABLE_COLUMNS, start=3):
        if col == dcp.PREVIOUS_LABEL:
            numerator = 0
        elif col == dcp.FUTURE_LABEL:
            numerator = denom
        else:
            n = month_to_position[dcp.MONTH_NAMES.index(col) + 1]
            if label == "1/24":
                numerator = 2 * n - 1
            else:
                numerator = {10: 1, 11: 3, 12: 5}.get(n, 0)
        pcell = ws.cell(row=period_row, column=j, value=numerator)
        pcell.font = DRIVER_FONT
        pcell.alignment = Alignment(horizontal="center")
        ycell = ws.cell(row=year_row, column=j, value=denom)
        ycell.font = DRIVER_FONT
        ycell.alignment = Alignment(horizontal="center")

    return period_row, year_row


def build_deferred_commission_workbook(result, opening_deferred_commission, output_path, run_date=None):
    """
    result: the dict returned by deferred_commission_processor.
        run_deferred_commission_calculation().
    opening_deferred_commission: a plain number - Opening Deferred
        Commission Expense, entered by hand (this isn't in any register -
        it's the prior period's own Closing figure), same "manual for
        now" pattern as UPR/Reinsurance Automation's own Opening balance.
    Returns output_path.
    """
    run_date = run_date or date.today()
    departments = result["departments"]
    commission_table = result["commission_table"]
    deferred_table = result["deferred_table"]
    methods = result["methods"]
    period_start, period_end = result["period_start"], result["period_end"]

    wb = Workbook()
    wb.remove(wb.active)

    # ---- "Deferred Commission Working" - both tables on one sheet ----
    ws_working = wb.create_sheet(WORKING_SHEET)
    ws_working.sheet_view.showGridLines = False
    subtitle = f"Period: {period_start:%d-%b-%Y} to {period_end:%d-%b-%Y}  \u2022  generated {run_date:%d-%b-%Y}"
    ws_working.cell(row=1, column=2, value="Deferred Commission Working").font = Font(name="Arial", bold=True, size=14, color="1F4E78")
    ws_working.cell(row=2, column=2, value=subtitle).font = SUBTITLE_FONT

    comm_header_row, comm_data_row, comm_total_row = _write_table_block(
        ws_working, 4, "Table 1 \u2014 Commission Paid/Payable Distribution", departments, commission_table)

    driver_gap = comm_total_row + 2
    driver_1_24_rows = _write_driver_rows(ws_working, driver_gap, period_start, "1/24")
    driver_1_6_rows = _write_driver_rows(ws_working, driver_gap + 3, period_start, "1/6")
    ws_working.cell(row=driver_gap - 1, column=2,
                     value="Driver rows below feed Table 2's formulas (numerator / denominator per column).").font = NOTE_FONT

    def _deferred_formula(dept, dept_row, col_idx, col_name):
        method = methods[dept]
        if method == dcp.up.METHOD_ZERO:
            return 0
        comm_row = comm_data_row + departments.index(dept)
        col_letter = _col_letter(col_idx)
        comm_cell = f"{col_letter}{comm_row}"
        driver_rows = driver_1_24_rows if method == dcp.up.METHOD_1_24 else driver_1_6_rows
        period_row, year_row = driver_rows
        return f"={comm_cell}*{col_letter}${period_row}/{col_letter}${year_row}"

    table2_start = driver_gap + 3 + 2 + 1
    deferred_header_row, deferred_data_row, deferred_total_row = _write_table_block(
        ws_working, table2_start, "Table 2 \u2014 Deferred Commission Expense", departments, deferred_table,
        value_formula_fn=_deferred_formula)

    _apply_working_sheet_column_widths(ws_working)
    ws_working.freeze_panes = ws_working.cell(row=comm_data_row, column=3).coordinate

    deferred_grand_total_cell = f"'{WORKING_SHEET}'!{_col_letter(3 + len(dcp.TABLE_COLUMNS))}{deferred_total_row}"
    commission_grand_total_cell = f"'{WORKING_SHEET}'!{_col_letter(3 + len(dcp.TABLE_COLUMNS))}{comm_total_row}"

    # ---- "Summary" ----
    ws_summary = wb.create_sheet(SUMMARY_SHEET)
    ws_summary.sheet_view.showGridLines = False
    ws_summary.cell(row=2, column=2, value="Deferred Commission Expense").font = Font(name="Arial", bold=True, size=14, color="1F4E78")
    ws_summary.cell(row=3, column=2, value=subtitle).font = SUBTITLE_FONT

    header_row = 5
    for j, h in enumerate(["Description", "Amount"], start=2):
        c = ws_summary.cell(row=header_row, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="left" if j == 2 else "right", vertical="center", indent=1)

    r = header_row + 1
    ws_summary.cell(row=r, column=2, value="Commission Paid/Payable").font = BODY_FONT
    total_comm_cell = ws_summary.cell(row=r, column=3, value=f"={commission_grand_total_cell}")
    total_comm_cell.number_format = NUM_FORMAT
    total_comm_cell.font = BODY_FONT
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = BORDER
    commission_row = r
    r += 1

    ws_summary.cell(row=r, column=2, value="Add: Opening Deferred Commission Expense").font = BODY_FONT
    opening_cell = ws_summary.cell(row=r, column=3, value=float(opening_deferred_commission or 0))
    opening_cell.number_format = NUM_FORMAT
    opening_cell.font = BODY_FONT
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = BORDER
    opening_row = r
    r += 1

    ws_summary.cell(row=r, column=2, value="Less: Closing Deferred Commission Expense").font = BODY_FONT
    closing_cell = ws_summary.cell(row=r, column=3, value=f"=-{deferred_grand_total_cell}")
    closing_cell.number_format = NUM_FORMAT
    closing_cell.font = BODY_FONT
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = BORDER
    closing_row = r
    r += 1

    ws_summary.cell(row=r, column=2, value="Commission Expense - Net").font = BOLD_FONT
    net_cell = ws_summary.cell(row=r, column=3, value=f"=C{commission_row}+C{opening_row}+C{closing_row}")
    net_cell.font = BOLD_FONT
    net_cell.number_format = NUM_FORMAT
    net_cell.fill = TOTAL_FILL
    ws_summary.cell(row=r, column=2).fill = TOTAL_FILL
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)

    ws_summary.column_dimensions["A"].width = 2
    ws_summary.column_dimensions["B"].width = 42
    ws_summary.column_dimensions["C"].width = 22

    note_row = r + 3
    if result.get("warnings"):
        for w in result["warnings"]:
            ws_summary.cell(row=note_row, column=2, value=f"Note: {w}").font = NOTE_FONT
            note_row += 1
    if result.get("excluded_rows"):
        ws_summary.cell(row=note_row, column=2,
                         value=f"Note: {result['excluded_rows']:,} premium register row(s) were excluded "
                               f"(missing Department, Comm Date, or Gross Premium).").font = NOTE_FONT
        note_row += 1
    ledger_meta = result.get("ledger_meta", {})
    for label, key in (("Commission Expense - Direct", "direct"),
                        ("Commission Expense - Coinsurance", "coinsurance"),
                        ("Service Charges - Coinsurance", "service_charges")):
        m = ledger_meta.get(key)
        if m and m.get("sheets_skipped"):
            ws_summary.cell(row=note_row, column=2,
                             value=f"Note: {label} ledger - {len(m['sheets_skipped'])} sheet(s) skipped "
                                   f"(no recognizable header): "
                                   + ", ".join(f"{fn} [{sn}]" for fn, sn, _reason in m["sheets_skipped"])).font = NOTE_FONT
            note_row += 1

    wb.move_sheet(SUMMARY_SHEET, offset=-1)  # Summary first

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)
    _recalculate(output_path)
    return output_path


def _recalculate(filepath):
    """Best-effort LibreOffice recalculation so formula cells have cached
    values on open - same helper pattern as upr_generator.py /
    reinsurance_generator.py."""
    recalc_script = "/mnt/skills/public/xlsx/scripts/recalc.py"
    if not os.path.exists(recalc_script):
        return
    try:
        subprocess.run([sys.executable, recalc_script, filepath, "60"], capture_output=True, timeout=90)
    except Exception:
        pass


def build_output_filename(period_end, run_date=None):
    run_date = run_date or date.today()
    return f"Deferred_Commission_{period_end:%b%Y}_{run_date:%d%m%Y}.xlsx"


if __name__ == "__main__":
    register_files = ["/mnt/user-data/uploads/Gross_Premium_-_Conventional_Register_2025_v2.xls",
                       "/mnt/user-data/uploads/Gross_Premium_-_Conventional_Register_June_2026_v2.xls"]
    direct_files = ["/mnt/user-data/uploads/Conv_-_Commission_Exp_Direct__2025_.xlsx",
                     "/mnt/user-data/uploads/Conv_-_Commission_Exp_Direct__Jan_2026_-_June_2026_.xlsx"]
    coinsurance_files = ["/mnt/user-data/uploads/Conv_-_Commission_Exp_Coinsurance_Accepted__2025_.xlsx",
                          "/mnt/user-data/uploads/Conv_-_Commission_Exp_Coinsurance_Accepted__Jan_2026_-_June_2026_.xlsx"]
    service_charge_files = ["/mnt/user-data/uploads/Conv_-_Service_Charges_Coinsurance__Dec_2025_.xlsx",
                             "/mnt/user-data/uploads/Conv_-_Service_Charges_Coinsurance__Jan_2026_-_June_2026_.xlsx"]

    result = dcp.run_deferred_commission_calculation(
        register_files, direct_files, coinsurance_files, service_charge_files, 2026, 6, report=print)
    out_path = "/home/claude/work/DEFERRED_COMMISSION_TEST_OUTPUT.xlsx"
    build_deferred_commission_workbook(result, opening_deferred_commission=230000000, output_path=out_path)
    print("Wrote:", out_path)
