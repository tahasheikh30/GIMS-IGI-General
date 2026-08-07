"""
Writes the UPR Automation output workbook: three sheets, styled
consistently with the rest of GIMS (navy header, Arial body, thin
borders) and mirroring the real manually-built working file this module
was reverse-engineered from - "Gross Premium Distribution" and "Unearned
Premium Working" both use live Excel formulas (SUM totals, and Table 2's
cells reference Table 1's cells multiplied by a driver-row fraction),
exactly like the original, so the sheet stays auditable and recalculates
correctly if a cell is ever edited by hand.
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
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM_FORMAT = "#,##0;(#,##0);-"

SUMMARY_SHEET = "Premium Earned Breakup"
GROSS_SHEET = "Gross Premium Distribution"
UNEARNED_SHEET = "Unearned Premium Working"

# Row layout on the two table sheets - department rows start at DATA_START_ROW
HEADER_ROW = 3
DATA_START_ROW = 4


def _col_letter(idx):
    return get_column_letter(idx)


def _write_table_sheet(ws, title, subtitle, departments, table_dict, value_formula_fn=None):
    """
    Shared layout for both Table 1 and Table 2: a title row, a header row
    (Previous, Jan..Dec, Future Policies, Total), one row per department,
    a bold Total row, with SUM formulas for every row/column total.

    value_formula_fn(dept, dept_row, col_idx, col_name) -> a value or
        formula string for that cell. If None, the raw numeric value from
        table_dict is written directly (used for Table 1). Table 2 passes
        a function that instead writes a formula referencing Table 1 and
        the driver rows below the grid.
    Returns the row index of the Total row (so the caller can reference
    the grand-total cell afterwards).
    """
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=2, value=title).font = Font(name="Arial", bold=True, size=14, color="1F4E78")
    if subtitle:
        ws.cell(row=2, column=2, value=subtitle).font = SUBTITLE_FONT

    columns = up.TABLE_COLUMNS + [up.TOTAL_LABEL]
    ws.cell(row=HEADER_ROW, column=2, value="Department").font = HEADER_FONT
    ws.cell(row=HEADER_ROW, column=2).fill = HEADER_FILL
    ws.cell(row=HEADER_ROW, column=2).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.cell(row=HEADER_ROW, column=2).border = BORDER
    for j, col in enumerate(columns, start=3):
        c = ws.cell(row=HEADER_ROW, column=j, value=col)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[HEADER_ROW].height = 30

    n_depts = len(departments)
    total_row = DATA_START_ROW + n_depts

    for i, dept in enumerate(departments):
        r = DATA_START_ROW + i
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
        cell = ws.cell(row=total_row, column=j, value=f"=SUM({col_letter}{DATA_START_ROW}:{col_letter}{total_row - 1})")
        cell.font = BOLD_FONT
        cell.fill = TOTAL_FILL
        cell.number_format = NUM_FORMAT
        cell.border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 20
    for j in range(3, 3 + len(columns)):
        ws.column_dimensions[_col_letter(j)].width = 13
    ws.freeze_panes = ws.cell(row=DATA_START_ROW, column=3).coordinate

    return total_row


def _write_driver_rows(ws, start_row, period_start, label):
    """
    Writes the two small "Period" / "Year" driver rows used by Table 2's
    formulas - exactly the same transparency device the original working
    file uses (see claims of module docstring): rather than hardcoding a
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
        plain numbers, entered manually by the user for now (per the
        initial-development spec; a future version can pull Gross
        Premium and Opening UPR automatically from the previous period's
        own output file instead).
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

    # ---- Sheet 2: Gross Premium Distribution (Table 1) ----
    ws_gross = wb.create_sheet(GROSS_SHEET)
    subtitle = f"Period: {period_start:%d-%b-%Y} to {period_end:%d-%b-%Y}  \u2022  generated {run_date:%d-%b-%Y}"
    gross_total_row = _write_table_sheet(ws_gross, "Gross Premium Distribution", subtitle, departments, gross_table)
    gross_grand_total_cell = f"{_col_letter(3 + len(up.TABLE_COLUMNS))}{gross_total_row}"

    # ---- Sheet 3: Unearned Premium Working (Table 2) ----
    ws_unearned = wb.create_sheet(UNEARNED_SHEET)

    def _unearned_formula(dept, dept_row, col_idx, col_name):
        method = methods[dept]
        gross_cell = f"'{GROSS_SHEET}'!{_col_letter(col_idx)}{dept_row}"
        if method == up.METHOD_ZERO:
            return 0
        driver_rows = driver_1_24_rows if method == up.METHOD_1_24 else driver_1_6_rows
        period_row, year_row = driver_rows
        col_letter = _col_letter(col_idx)
        return f"={gross_cell}*{col_letter}${period_row}/{col_letter}${year_row}"

    unearned_total_row = _write_table_sheet(ws_unearned, "Unearned Premium Working", subtitle, departments,
                                              unearned_table, value_formula_fn=None)
    # _write_table_sheet already wrote placeholder numeric values above;
    # now overwrite the department cells with the live formulas, since
    # the driver rows (written below the grid) need to already exist at
    # fixed row numbers before the formulas can reference them. Simpler
    # to lay the driver rows out first, then rewrite the grid - done here
    # rather than restructuring _write_table_sheet's signature further.
    driver_gap = unearned_total_row + 2
    driver_1_24_rows = _write_driver_rows(ws_unearned, driver_gap, period_start, "1/24")
    driver_1_6_rows = _write_driver_rows(ws_unearned, driver_gap + 3, period_start, "1/6")

    for i, dept in enumerate(departments):
        r = DATA_START_ROW + i
        method = methods[dept]
        for j, col in enumerate(up.TABLE_COLUMNS, start=3):
            cell = ws_unearned.cell(row=r, column=j)
            if method == up.METHOD_ZERO:
                cell.value = 0
            else:
                cell.value = _unearned_formula(dept, r, j, col)

    ws_unearned.cell(row=driver_gap - 1, column=2,
                      value="Driver rows below feed the formulas above (numerator / denominator per column).").font = NOTE_FONT

    unearned_grand_total_cell = f"'{UNEARNED_SHEET}'!{_col_letter(3 + len(up.TABLE_COLUMNS))}{unearned_total_row}"

    # ---- Sheet 1: Premium Earned Breakup (Section 1 summary) ----
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
    sheets_used = ", ".join(f"{fn} [{sn}]" for fn, sn in result.get("sheets_used", []))
    if sheets_used:
        ws_summary.cell(row=note_row, column=2, value=f"Source sheet(s): {sheets_used}").font = NOTE_FONT

    wb.move_sheet(SUMMARY_SHEET, offset=-2)  # Summary first

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
    import sys
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
