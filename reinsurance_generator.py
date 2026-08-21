"""
Writes the Reinsurance Automation output workbook. Styled identically to
upr_generator.py (same navy/Arial/thin-border look as the rest of GIMS),
and structured the same way: a "Summary" sheet up front, then a single
"Prepaid RI Ceded Working" sheet holding BOTH tables stacked (Table 1 -
Reinsurance Premium Ceded Distribution, then Table 2 - Prepaid
Reinsurance Premium Ceded), both driven by live Excel SUM formulas so
the sheet stays auditable and recalculates correctly if a cell is ever
edited by hand - see upr_generator.py's own module docstring for the
full rationale behind this layout, which is carried over unchanged here
per the module's own spec ("the rest of the working is same as UPR
module").

The Summary sheet also carries a small Treaty/Local/Foreign breakdown
table per department (component_totals, from reinsurance_processor.py) -
this is the same "how much of this is treaty vs facultative" figure the
real manually-maintained working file shows, ahead of the Premium
Ceded / Opening / Closing net-figure block.
"""
import os
import subprocess
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import reinsurance_processor as rp

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
WORKING_SHEET = "Prepaid RI Ceded Working"

DEPT_COL_WIDTH = 22
DATA_COL_WIDTH = 16
TOTAL_COL_WIDTH = 20


def _col_letter(idx):
    return get_column_letter(idx)


def _write_table_block(ws, start_row, section_label, departments, table_dict, value_formula_fn=None):
    """Identical shape to upr_generator._write_table_block - a section
    label, a header row, one row per department, and a bold Total row
    with SUM formulas - see that module's own docstring for the full
    rationale (this is deliberately the same layout, reused by copy
    rather than cross-import since the two modules' output workbooks are
    independent artifacts that should each remain self-contained)."""
    ws.cell(row=start_row, column=2, value=section_label).font = SECTION_TITLE_FONT
    header_row = start_row + 1
    data_start_row = header_row + 1

    columns = rp.TABLE_COLUMNS + [rp.TOTAL_LABEL]
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
        for j, col in enumerate(rp.TABLE_COLUMNS, start=3):
            if value_formula_fn is not None:
                val = value_formula_fn(dept, r, j, col)
            else:
                val = table_dict[dept][col]
            cell = ws.cell(row=r, column=j, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.number_format = NUM_FORMAT
        first_col_letter = _col_letter(3)
        last_col_letter = _col_letter(2 + len(rp.TABLE_COLUMNS))
        total_cell = ws.cell(row=r, column=3 + len(rp.TABLE_COLUMNS),
                              value=f"=SUM({first_col_letter}{r}:{last_col_letter}{r})")
        total_cell.font = BOLD_FONT
        total_cell.border = BORDER
        total_cell.number_format = NUM_FORMAT

    ws.cell(row=total_row, column=2, value=rp.TOTAL_LABEL).font = BOLD_FONT
    ws.cell(row=total_row, column=2).fill = TOTAL_FILL
    ws.cell(row=total_row, column=2).border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)
    for j in range(3, 3 + len(rp.TABLE_COLUMNS) + 1):
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
    for j, _col in enumerate(rp.TABLE_COLUMNS, start=3):
        ws.column_dimensions[_col_letter(j)].width = DATA_COL_WIDTH
    ws.column_dimensions[_col_letter(3 + len(rp.TABLE_COLUMNS))].width = TOTAL_COL_WIDTH


def _write_driver_rows(ws, start_row, period_start, label):
    """Identical to upr_generator._write_driver_rows - the numerator/
    denominator driver rows that Table 2's own formulas reference, kept
    visible and auditable rather than baked invisibly into each cell."""
    month_to_position = rp.up._month_to_period_position(period_start)
    denom = 24 if label == "1/24" else 6

    period_row = start_row
    year_row = start_row + 1

    ws.cell(row=period_row, column=2, value=f"{label} Period").font = DRIVER_FONT
    ws.cell(row=year_row, column=2, value=f"{label} Denominator").font = DRIVER_FONT

    for j, col in enumerate(rp.TABLE_COLUMNS, start=3):
        if col == rp.PREVIOUS_LABEL:
            numerator = 0
        elif col == rp.FUTURE_LABEL:
            numerator = denom
        else:
            n = month_to_position[rp.MONTH_NAMES.index(col) + 1]
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


def _write_component_breakdown(ws, top_row, departments, component_totals):
    """Small Department x (Treaty / FAC / Foreign FAC / Total) table -
    the Summary sheet's own "how much of this is treaty vs facultative"
    figure, computed across every row actually loaded (not period-
    scoped) - see reinsurance_processor.build_component_totals.

    Column labels here ("Treaty" / "FAC" / "Foreign FAC") deliberately
    match the wording on the real, existing manually-built Summary sheet
    this module's output is meant to replace - NOT the register's own
    raw column headers ("Treaty Premium" / "Local Premium" / "Foreign
    Premium"), which is what those three columns are called everywhere
    else (the register itself, Table 1/Table 2 on the working sheet, and
    reinsurance_processor.py's own internal field names). "Local
    Premium" in the register is the same thing the existing workbook
    calls "FAC" (facultative reinsurance placed locally), and "Foreign
    Premium" is what it calls "Foreign FAC" (facultative reinsurance
    placed abroad) - only the LABEL differs here, not the underlying
    figure."""
    headers = ["Department", "Treaty", "FAC", "Foreign FAC", "Total"]
    for j, h in enumerate(headers, start=2):
        c = ws.cell(row=top_row, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="left" if j == 2 else "center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[top_row].height = 26

    r = top_row + 1
    first_data_row = r
    for dept in departments:
        t = component_totals.get(dept, {})
        ws.cell(row=r, column=2, value=dept).font = BODY_FONT
        ws.cell(row=r, column=2).border = BORDER
        for j, key in enumerate(("Treaty Premium", "Local Premium", "Foreign Premium", "Total Reinsurance Ceded"), start=3):
            cell = ws.cell(row=r, column=j, value=float(t.get(key, 0.0)))
            cell.font = BODY_FONT
            cell.number_format = NUM_FORMAT
            cell.border = BORDER
        r += 1
    last_data_row = r - 1

    ws.cell(row=r, column=2, value="Total").font = BOLD_FONT
    ws.cell(row=r, column=2).fill = TOTAL_FILL
    ws.cell(row=r, column=2).border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)
    for j, col_letter in enumerate((_col_letter(3), _col_letter(4), _col_letter(5), _col_letter(6)), start=3):
        cell = ws.cell(row=r, column=j, value=f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})")
        cell.font = BOLD_FONT
        cell.fill = TOTAL_FILL
        cell.number_format = NUM_FORMAT
        cell.border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)

    return r  # the Total row - caller uses this as the "Premium Ceded" reference point if desired


def build_reinsurance_workbook(result, opening_prepaid, output_path, run_date=None):
    """
    result: the dict returned by reinsurance_processor.run_reinsurance_calculation().
    opening_prepaid: a plain number - Opening Prepaid Reinsurance Premium
        Ceded, entered by hand (this isn't in any register - it's the
        prior period's own Closing figure), same "manual for now" pattern
        as UPR Automation's own Opening UPR.
    Returns output_path.
    """
    run_date = run_date or date.today()
    departments = result["departments"]
    ceded_table = result["ceded_table"]
    prepaid_table = result["prepaid_table"]
    methods = result["methods"]
    component_totals = result["component_totals"]
    period_start, period_end = result["period_start"], result["period_end"]

    wb = Workbook()
    wb.remove(wb.active)

    # ---- "Prepaid RI Ceded Working" - both tables on one sheet, exactly
    # mirroring upr_generator's own single-sheet Table1-then-Table2 layout. ----
    ws_working = wb.create_sheet(WORKING_SHEET)
    ws_working.sheet_view.showGridLines = False
    subtitle = f"Period: {period_start:%d-%b-%Y} to {period_end:%d-%b-%Y}  \u2022  generated {run_date:%d-%b-%Y}"
    ws_working.cell(row=1, column=2, value="Prepaid RI Premium Ceded Working").font = Font(name="Arial", bold=True, size=14, color="1F4E78")
    ws_working.cell(row=2, column=2, value=subtitle).font = SUBTITLE_FONT

    ceded_header_row, ceded_data_row, ceded_total_row = _write_table_block(
        ws_working, 4, "Table 1 \u2014 Reinsurance Premium Ceded Distribution", departments, ceded_table)

    driver_gap = ceded_total_row + 2
    driver_1_24_rows = _write_driver_rows(ws_working, driver_gap, period_start, "1/24")
    driver_1_6_rows = _write_driver_rows(ws_working, driver_gap + 3, period_start, "1/6")
    ws_working.cell(row=driver_gap - 1, column=2,
                     value="Driver rows below feed Table 2's formulas (numerator / denominator per column).").font = NOTE_FONT

    def _prepaid_formula(dept, dept_row, col_idx, col_name):
        method = methods[dept]
        if method == rp.up.METHOD_ZERO:
            return 0
        ceded_row = ceded_data_row + departments.index(dept)
        col_letter = _col_letter(col_idx)
        ceded_cell = f"{col_letter}{ceded_row}"
        driver_rows = driver_1_24_rows if method == rp.up.METHOD_1_24 else driver_1_6_rows
        period_row, year_row = driver_rows
        return f"={ceded_cell}*{col_letter}${period_row}/{col_letter}${year_row}"

    table2_start = driver_gap + 3 + 2 + 1
    prepaid_header_row, prepaid_data_row, prepaid_total_row = _write_table_block(
        ws_working, table2_start, "Table 2 \u2014 Prepaid Reinsurance Premium Ceded", departments, prepaid_table,
        value_formula_fn=_prepaid_formula)

    _apply_working_sheet_column_widths(ws_working)
    ws_working.freeze_panes = ws_working.cell(row=ceded_data_row, column=3).coordinate

    prepaid_grand_total_cell = f"'{WORKING_SHEET}'!{_col_letter(3 + len(rp.TABLE_COLUMNS))}{prepaid_total_row}"
    # "Premium Ceded" is business INCEPTING during the selected period
    # only - i.e. the SUM of just Table 1's own January..December
    # columns on its Total row, deliberately excluding the Previous
    # column (older business from before the period, still amortizing)
    # and the Future Policies column (business already booked for after
    # it) - summing the WHOLE row (Table 1's own "Total" column, which is
    # every bucket combined) would double up the period figure with
    # however much old/future business happens to sit in the register,
    # a different, inconsistent amount per department. Built from
    # TABLE_COLUMNS' own layout rather than hardcoded column letters, so
    # this stays correct if that column order is ever changed.
    first_month_col = 3 + rp.TABLE_COLUMNS.index(rp.MONTH_NAMES[0])
    last_month_col = 3 + rp.TABLE_COLUMNS.index(rp.MONTH_NAMES[-1])
    premium_ceded_in_period_cell = (
        f"SUM('{WORKING_SHEET}'!{_col_letter(first_month_col)}{ceded_total_row}:"
        f"{_col_letter(last_month_col)}{ceded_total_row})"
    )

    # ---- "Summary" ----
    ws_summary = wb.create_sheet(SUMMARY_SHEET)
    ws_summary.sheet_view.showGridLines = False
    ws_summary.cell(row=2, column=2, value="Prepaid Reinsurance Premium Ceded").font = Font(name="Arial", bold=True, size=14, color="1F4E78")
    ws_summary.cell(row=3, column=2, value=subtitle).font = SUBTITLE_FONT

    ws_summary.cell(row=5, column=2, value="Reinsurance Ceded by Department & Type").font = SECTION_TITLE_FONT
    breakdown_total_row = _write_component_breakdown(ws_summary, 6, departments, component_totals)

    r = breakdown_total_row + 3
    ws_summary.cell(row=r, column=2, value="Premium Ceded (selected period)").font = BODY_FONT
    premium_ceded_cell = ws_summary.cell(row=r, column=3, value=f"={premium_ceded_in_period_cell}")
    premium_ceded_cell.number_format = NUM_FORMAT
    premium_ceded_cell.font = BODY_FONT
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = BORDER
    premium_ceded_row = r
    r += 1

    ws_summary.cell(row=r, column=2, value="Add: Opening Prepaid Reinsurance Ceded").font = BODY_FONT
    opening_cell = ws_summary.cell(row=r, column=3, value=float(opening_prepaid or 0))
    opening_cell.number_format = NUM_FORMAT
    opening_cell.font = BODY_FONT
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = BORDER
    opening_row = r
    r += 1

    ws_summary.cell(row=r, column=2, value="Less: Closing Prepaid Reinsurance Ceded").font = BODY_FONT
    closing_cell = ws_summary.cell(row=r, column=3, value=f"=-{prepaid_grand_total_cell}")
    closing_cell.number_format = NUM_FORMAT
    closing_cell.font = BODY_FONT
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = BORDER
    closing_row = r
    r += 1

    ws_summary.cell(row=r, column=2, value="Reinsurance Premium Ceded - Net").font = BOLD_FONT
    net_cell = ws_summary.cell(row=r, column=3, value=f"=C{premium_ceded_row}+C{opening_row}+C{closing_row}")
    net_cell.font = BOLD_FONT
    net_cell.number_format = NUM_FORMAT
    net_cell.fill = TOTAL_FILL
    ws_summary.cell(row=r, column=2).fill = TOTAL_FILL
    for col in (2, 3):
        ws_summary.cell(row=r, column=col).border = Border(top=Side(style="thin", color="1F4E78"), left=THIN, right=THIN, bottom=THIN)

    ws_summary.column_dimensions["A"].width = 2
    ws_summary.column_dimensions["B"].width = 42
    ws_summary.column_dimensions["C"].width = 22
    ws_summary.column_dimensions["D"].width = 20
    ws_summary.column_dimensions["E"].width = 20
    ws_summary.column_dimensions["F"].width = 22

    note_row = r + 3
    if result.get("warnings"):
        for w in result["warnings"]:
            ws_summary.cell(row=note_row, column=2, value=f"Note: {w}").font = NOTE_FONT
            note_row += 1
    if result.get("excluded_rows"):
        ws_summary.cell(row=note_row, column=2,
                         value=f"Note: {result['excluded_rows']:,} register row(s) were excluded "
                               f"(missing or implausible Comm Date).").font = NOTE_FONT
        note_row += 1
    files_note = ", ".join(f"{fn} -> {dept}" for fn, dept in result.get("departments_by_file", {}).items())
    if files_note:
        ws_summary.cell(row=note_row, column=2, value=f"Source file(s): {files_note}").font = NOTE_FONT

    wb.move_sheet(SUMMARY_SHEET, offset=-1)  # Summary first

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)
    _recalculate(output_path)
    return output_path


def _recalculate(filepath):
    """Best-effort LibreOffice recalculation so formula cells have cached
    values on open - same helper pattern as upr_generator.py."""
    recalc_script = "/mnt/skills/public/xlsx/scripts/recalc.py"
    if not os.path.exists(recalc_script):
        return
    try:
        subprocess.run([sys.executable, recalc_script, filepath, "60"], capture_output=True, timeout=90)
    except Exception:
        pass


def build_output_filename(period_end, run_date=None):
    run_date = run_date or date.today()
    return f"Prepaid_RI_Ceded_{period_end:%b%Y}_{run_date:%d%m%Y}.xlsx"


if __name__ == "__main__":
    files = sys.argv[1:] or ["Eng__2025_v2.xls", "Eng__June_2026_v2.xls"]

    result = rp.run_reinsurance_calculation(files, 2026, 6, report=print)
    out_path = "/home/claude/work/REINSURANCE_TEST_OUTPUT.xlsx"
    build_reinsurance_workbook(result, opening_prepaid=250000000, output_path=out_path)
    print("Wrote:", out_path)
