"""
Writes the UPR Automation output workbook, styled consistently with the
rest of GIMS (navy header, Arial body, thin borders) and mirroring the
real manually-built working file this module was reverse-engineered
from: BOTH tables live on a single "Unearned Premium Working" sheet -
Table 1 (Gross Premium Distribution) first, then the driver rows, then
Table 2 (Unearned Premium Working) directly below it - exactly the
layout the original file uses (confirmed by inspecting it: rows 27-36
are the Gross Premium table, rows 38-42 the driver rows, rows 44-53 the
weighted working table, all on that same one sheet). Both tables use
live Excel formulas (SUM totals, and Table 2's cells reference Table 1's
own cells multiplied by a driver-row fraction), so the sheet stays
auditable and recalculates correctly if a cell is ever edited by hand.
"""
import os
import subprocess
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import upr_processor as up

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

SUMMARY_SHEET = "Premium Earned Breakup"
WORKING_SHEET = "Unearned Premium Working"

# Column widths - sized to comfortably fit real PKR figures with commas
# (confirmed against the real reference file: a single department's own
# Total can run past 8 billion, and the grand Total-of-Totals past 21
# billion - "21,026,406,428" is 14 characters). Too narrow here is
# exactly what produces Excel's "####" display (the number itself
# calculates fine - Excel just has nowhere to draw it).
DEPT_COL_WIDTH = 22
DATA_COL_WIDTH = 16
TOTAL_COL_WIDTH = 20


def _col_letter(idx):
    return get_column_letter(idx)


def _write_table_block(ws, start_row, section_label, departments, table_dict, value_formula_fn=None):
    """
    Writes one full table (a section label, a header row, one row per
    department, and a bold Total row with SUM formulas) starting at
    `start_row` on an already-open worksheet - used twice on the SAME
    sheet (Table 1 then Table 2, stacked, matching the real reference
    file's own single-sheet layout - see the module docstring) rather
    than writing each table to its own separate sheet.

    value_formula_fn(dept, dept_row, col_idx, col_name) -> a value or
        formula string for that cell. If None, the raw numeric value
        from table_dict is written directly (used for Table 1). Table 2
        passes a function that instead writes a formula referencing
        Table 1's own cells and the driver rows below Table 2.

    Returns (header_row, data_start_row, total_row) - the caller needs
    these to place whatever comes next, and to build a cross-reference
    to this table's own grand-total cell.
    """
    ws.cell(row=start_row, column=2, value=section_label).font = SECTION_TITLE_FONT
    header_row = start_row + 1
    data_start_row = header_row + 1

    columns = up.TABLE_COLUMNS + [up.TOTAL_LABEL]
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
        for j, col in enumerate(up.TABLE_COLUMNS, start=3):
            if value_formula_fn is not None:
                val = value_formula_fn(dept, r, j, col)
            else:
                val = table_dict[dept][col]
            cell = ws.cell(row=r, column=j, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.number_format = NUM_FORMAT
        # row total = SUM across that department's own 14 columns
        first_col_letter = _col_letter(3)
        last_col_letter = _col_letter(2 + len(up.TABLE_COLUMNS))
        total_cell = ws.cell(row=r, column=3 + len(up.TABLE_COLUMNS),
                              value=f"=SUM({first_col_letter}{r}:{last_col_letter}{r})")
        total_cell.font = BOLD_FONT
        total_cell.border = BORDER
        total_cell.number_format = NUM_FORMAT

    # Total row - bold, shaded, SUM formulas down each column
    ws.cell(row=total_row, column=2, value=up.TOTAL_LABEL).font = BOLD_FONT
    ws.cell(row=total_row, column=2).fill = TOTAL_FILL
    ws.cell(row=total_row, column=2).border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)
    for j in range(3, 3 + len(up.TABLE_COLUMNS) + 1):
        col_letter = _col_letter(j)
        cell = ws.cell(row=total_row, column=j, value=f"=SUM({col_letter}{data_start_row}:{col_letter}{total_row - 1})")
        cell.font = BOLD_FONT
        cell.fill = TOTAL_FILL
        cell.number_format = NUM_FORMAT
        cell.border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)

    return header_row, data_start_row, total_row


def _apply_working_sheet_column_widths(ws):
    """Set once per sheet (not per table-block) since Table 1 and Table 2
    now share the exact same columns on the same sheet."""
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = DEPT_COL_WIDTH
    for j, _col in enumerate(up.TABLE_COLUMNS, start=3):
        ws.column_dimensions[_col_letter(j)].width = DATA_COL_WIDTH
    ws.column_dimensions[_col_letter(3 + len(up.TABLE_COLUMNS))].width = TOTAL_COL_WIDTH


def _write_driver_rows(ws, start_row, period_start, label):
    """
    Writes the two small "Period" / "Year" driver rows used by Table 2's
    formulas - exactly the same transparency device the original working
    file uses (see the module docstring): rather than hardcoding a
    fraction into each cell, the fraction is built from a period-position
    numerator (this row) divided by a fixed denominator (the row below
    it), so the driver values themselves are visible and auditable, not
    buried inside each cell's own formula.

    label: "1/24" or "1/6" - which method this driver block is for.
    Returns (period_row, year_row).
    """
    month_to_position = up._month_to_period_position(period_start)
    denom = 24 if label == "1/24" else 6

    period_row = start_row
    year_row = start_row + 1

    ws.cell(row=period_row, column=2, value=f"{label} Period").font = DRIVER_FONT
    ws.cell(row=year_row, column=2, value=f"{label} Denominator").font = DRIVER_FONT

    for j, col in enumerate(up.TABLE_COLUMNS, start=3):
        if col == up.PREVIOUS_LABEL:
            numerator = 0
        elif col == up.FUTURE_LABEL:
            numerator = denom
        else:
            n = month_to_position[up.MONTH_NAMES.index(col) + 1]
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


def build_upr_workbook(result, section1_inputs, output_path, run_date=None):
    """
    result: the dict returned by upr_processor.run_upr_calculation().
    section1_inputs: dict with keys "premium_direct", "premium_coinsurance",
        "premium_reinsurance", "admin_surcharge", "opening_upr" - all
        plain numbers (see upr_ledger_processor.py for the four
        premium-income ones being pulled from a PDF instead of typed by
        hand; Opening UPR stays manual for now).
    Returns output_path.
    """
    run_date = run_date or date.today()
    departments = result["departments"]
    gross_table = result["gross_table"]
    unearned_table = result["unearned_table"]
    methods = result["methods"]
    period_start, period_end = result["period_start"], result["period_end"]

    wb = Workbook()
    wb.remove(wb.active)

    # ---- "Unearned Premium Working" - BOTH tables on one sheet, Table 1
    # then Table 2 stacked directly below it, matching the real reference
    # file's own layout exactly (see the module docstring). ----
    ws_working = wb.create_sheet(WORKING_SHEET)
    ws_working.sheet_view.showGridLines = False
    subtitle = f"Period: {period_start:%d-%b-%Y} to {period_end:%d-%b-%Y}  \u2022  generated {run_date:%d-%b-%Y}"
    ws_working.cell(row=1, column=2, value="Unearned Premium Working").font = Font(name="Arial", bold=True, size=14, color="1F4E78")
    ws_working.cell(row=2, column=2, value=subtitle).font = SUBTITLE_FONT

    # Table 1 - Gross Premium Distribution
    gross_header_row, gross_data_row, gross_total_row = _write_table_block(
        ws_working, 4, "Table 1 \u2014 Gross Premium Distribution", departments, gross_table)

    # Driver rows (1/24 then 1/6), sitting between the two tables - same
    # relative position as the original file (right after Table 1's own
    # Total row, right before Table 2 starts).
    driver_gap = gross_total_row + 2
    driver_1_24_rows = _write_driver_rows(ws_working, driver_gap, period_start, "1/24")
    driver_1_6_rows = _write_driver_rows(ws_working, driver_gap + 3, period_start, "1/6")
    ws_working.cell(row=driver_gap - 1, column=2,
                     value="Driver rows below feed Table 2's formulas (numerator / denominator per column).").font = NOTE_FONT

    def _unearned_formula(dept, dept_row, col_idx, col_name):
        method = methods[dept]
        if method == up.METHOD_ZERO:
            return 0
        gross_row = gross_data_row + departments.index(dept)
        col_letter = _col_letter(col_idx)
        gross_cell = f"{col_letter}{gross_row}"
        driver_rows = driver_1_24_rows if method == up.METHOD_1_24 else driver_1_6_rows
        period_row, year_row = driver_rows
        return f"={gross_cell}*{col_letter}${period_row}/{col_letter}${year_row}"

    # Table 2 - Unearned Premium Working, starting right after the driver rows
    table2_start = driver_gap + 3 + 2 + 1  # driver block (2 rows) + gap + blank spacer
    unearned_header_row, unearned_data_row, unearned_total_row = _write_table_block(
        ws_working, table2_start, "Table 2 \u2014 Unearned Premium Working", departments, unearned_table,
        value_formula_fn=_unearned_formula)

    _apply_working_sheet_column_widths(ws_working)
    ws_working.freeze_panes = ws_working.cell(row=gross_data_row, column=3).coordinate

    unearned_grand_total_cell = f"'{WORKING_SHEET}'!{_col_letter(3 + len(up.TABLE_COLUMNS))}{unearned_total_row}"

    # ---- "Premium Earned Breakup" (Section 1 summary) ----
    ws_summary = wb.create_sheet(SUMMARY_SHEET)
    ws_summary.sheet_view.showGridLines = False
    ws_summary.cell(row=2, column=2, value="Premium Earned Breakup").font = Font(name="Arial", bold=True, size=14, color="1F4E78")
    ws_summary.cell(row=3, column=2, value=subtitle).font = SUBTITLE_FONT

    header_row = 5
    for j, h in enumerate(["Description", "Amount"], start=2):
        c = ws_summary.cell(row=header_row, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="left" if j == 2 else "right", vertical="center", indent=1)

    rows = [
        ("Premium Direct", section1_inputs.get("premium_direct", 0)),
        ("Premium Income Coinsurance", section1_inputs.get("premium_coinsurance", 0)),
        ("Premium Income Reinsurance", section1_inputs.get("premium_reinsurance", 0)),
        ("Administrative Surcharge", section1_inputs.get("admin_surcharge", 0)),
    ]
    r = header_row + 1
    first_input_row = r
    for label, val in rows:
        ws_summary.cell(row=r, column=2, value=label).font = BODY_FONT
        amt_cell = ws_summary.cell(row=r, column=3, value=float(val))
        amt_cell.number_format = NUM_FORMAT
        amt_cell.font = BODY_FONT
        for col in (2, 3):
            ws_summary.cell(row=r, column=col).border = BORDER
        r += 1
    last_input_row = r - 1

    gross_premium_row = r
    ws_summary.cell(row=r, column=2, value="Gross Premium").font = BOLD_FONT
    gp_cell = ws_summary.cell(row=r, column=3, value=f"=SUM(C{first_input_row}:C{last_input_row})")
    gp_cell.font = BOLD_FONT
    gp_cell.number_format = NUM_FORMAT
    gp_cell.fill = TOTAL_FILL
    ws_summary.cell(row=r, column=2).fill = TOTAL_FILL
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)
    r += 1

    ws_summary.cell(row=r, column=2, value="Add: Opening Unearned Premium Revenue").font = BODY_FONT
    opening_cell = ws_summary.cell(row=r, column=3, value=float(section1_inputs.get("opening_upr", 0)))
    opening_cell.number_format = NUM_FORMAT
    opening_cell.font = BODY_FONT
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = BORDER
    opening_row = r
    r += 1

    ws_summary.cell(row=r, column=2, value="Less: Closing Unearned Premium Revenue").font = BODY_FONT
    closing_cell = ws_summary.cell(row=r, column=3, value=f"=-{unearned_grand_total_cell}")
    closing_cell.number_format = NUM_FORMAT
    closing_cell.font = BODY_FONT
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = BORDER
    closing_row = r
    r += 1

    ws_summary.cell(row=r, column=2, value="Premium Earned").font = BOLD_FONT
    earned_cell = ws_summary.cell(row=r, column=3, value=f"=C{gross_premium_row}+C{opening_row}+C{closing_row}")
    earned_cell.font = BOLD_FONT
    earned_cell.number_format = NUM_FORMAT
    earned_cell.fill = TOTAL_FILL
    ws_summary.cell(row=r, column=2).fill = TOTAL_FILL
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)

    ws_summary.column_dimensions["A"].width = 2
    ws_summary.column_dimensions["B"].width = 40
    ws_summary.column_dimensions["C"].width = 24

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
    sheets_used = ", ".join(f"{fn} [{sn}]" for fn, sn in result.get("sheets_used", []))
    if sheets_used:
        ws_summary.cell(row=note_row, column=2, value=f"Source sheet(s): {sheets_used}").font = NOTE_FONT

    wb.move_sheet(SUMMARY_SHEET, offset=-1)  # Summary first, ahead of the working sheet

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)
    _recalculate(output_path)
    return output_path


def _recalculate(filepath):
    """Best-effort LibreOffice recalculation so formula cells have cached
    values on open - same helper pattern as loss_ratio_generator.py."""
    recalc_script = "/mnt/skills/public/xlsx/scripts/recalc.py"
    if not os.path.exists(recalc_script):
        return
    try:
        subprocess.run([sys.executable, recalc_script, filepath, "60"], capture_output=True, timeout=90)
    except Exception:
        pass


def build_output_filename(period_end, run_date=None):
    run_date = run_date or date.today()
    return f"UPR_Working_{period_end:%b%Y}_{run_date:%d%m%Y}.xlsx"


if __name__ == "__main__":
    files = sys.argv[1:] or ["/mnt/user-data/uploads/Gross_Premium_and_Unearned_Premium_Working_-_Conventional_June_2026.xlsx"]

    result = up.run_upr_calculation(files, 2026, 6, report=print)
    section1 = {
        "premium_direct": 5842287567,
        "premium_coinsurance": 1035145247,
        "premium_reinsurance": 8926539,
        "admin_surcharge": 77318060,
        "opening_upr": 4804087358,
    }
    out_path = "/home/claude/UPR_TEST_OUTPUT.xlsx"
    build_upr_workbook(result, section1, out_path)
    print("Wrote:", out_path)
