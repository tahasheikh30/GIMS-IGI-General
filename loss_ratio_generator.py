"""
Builds the Loss Ratio Summary Excel output, replicating the real IGI
"Group Summary" template exactly (verified cell-by-cell against a real
sample file):

  Row 2: merged group-name banner, bold, bright green fill (B2:F2)
  Row 3: header CLIENT | GP | CLAIM | COMMISSION | LR, bold, light grey fill
  Rows 4+: one row per company - GP/CLAIM/COMMISSION as plain numeric
    values (accounting format), LR as a live formula:
        =SUM(D{row}:E{row})/C{row}%
    (Excel's "%" operator divides by 100, so this evaluates to
    (CLAIM+COMMISSION)/GP*100 - matching the real template's own formula
    convention exactly rather than hardcoding a Python-computed number.)
  Group mode only: a bold TOTAL row with SUM() formulas for GP/CLAIM/
    COMMISSION and the same LR formula pattern, with the LR cell
    highlighted yellow - matching the real template's total-row styling.

As an enhancement beyond the original template, individual company rows
also get a light red/amber fill on the LR value when it crosses the
configurable threshold (red) or a fixed 70% (amber) - this doesn't
conflict with anything in the real template, which leaves those rows
unstyled.
"""
import os
import subprocess
import sys
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

GREEN_FILL = PatternFill("solid", start_color="FF92D050", end_color="FF92D050")
HEADER_FILL = PatternFill("solid", start_color="FFD9D9D9", end_color="FFD9D9D9")  # light grey, matches theme(0)-15%
YELLOW_FILL = PatternFill("solid", start_color="FFFFFF00", end_color="FFFFFF00")
RED_FILL = PatternFill("solid", start_color="FFF8D7DA", end_color="FFF8D7DA")
RED_FONT = Font(name="Calibri", size=10, color="FF842029")
AMBER_FILL = PatternFill("solid", start_color="FFFFF3CD", end_color="FFFFF3CD")
AMBER_FONT = Font(name="Calibri", size=10, color="FF8A6100")
NA_FILL = PatternFill("solid", start_color="FFE9ECEF", end_color="FFE9ECEF")
NA_FONT = Font(name="Calibri", size=10, italic=True, color="FF6C757D")

FONT_BOLD = Font(name="Calibri", size=10, bold=True)
FONT_NORMAL = Font(name="Calibri", size=10)
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

ACCOUNTING_FORMAT = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'
LR_FORMAT = "0.00"
AMBER_THRESHOLD = 70.0  # fixed secondary tier, e.g. flag anything >= 70%

HEADERS = ["CLIENT", "GP", "CLAIM", "COMMISSION", "LR"]


def _write_text_safe(cell):
    """
    Guard against Excel/DDE formula injection. openpyxl auto-classifies
    ANY string cell value starting with "=" (also "+", "-", "@" in some
    Excel configurations) as a live formula. Company Name here comes
    straight from an uploaded Excel or PDF source file
    (loss_ratio_processor.py) - untrusted input - so a value like
    '=cmd|"/c calc.exe"!A0' would otherwise be written out as a live,
    auto-executing formula in the report handed to Finance/the CFO.
    Forcing the cell's data_type back to plain text after the fact
    neutralizes this with zero visible change to legitimate values -
    same pattern already used by every other GIMS generator (see
    report_generator.py / receivable_generator.py /
    premium_register_generator.py / claims_register_generator.py).
    """
    if cell.data_type == "f":
        cell.data_type = "s"
    return cell


def build_loss_ratio_workbook(rows, mode, output_path, group_name="LOSS RATIO SUMMARY", red_threshold_pct=100.0):
    """
    rows: list of company row dicts from loss_ratio_processor (Company Name,
        Gross Premium Written, Incurred Loss, Commission Accrual, loss_ratio,
        warning)
    mode: "individual" or "group" - group mode adds a bold TOTAL row
    output_path: full filepath to save to (from a Save As dialog)
    group_name: text shown in the green banner row (e.g. "POPULAR GROUP")
    red_threshold_pct: Loss Ratio % above which a company row is
        additionally highlighted red (values >= 70% are always amber)
    """
    if not rows:
        raise ValueError("No rows to write - nothing was successfully processed.")

    wb = Workbook()
    ws = wb.active
    ws.title = "Group Summary" if mode == "group" else "Loss Ratio Summary"

    ws.merge_cells("B2:F2")
    banner = ws["B2"]
    banner.value = group_name
    _write_text_safe(banner)  # group_name is free-typed by the user - guard against formula injection
    banner.font = FONT_BOLD
    banner.fill = GREEN_FILL
    banner.alignment = Alignment(horizontal="center", vertical="top")

    header_row = 3
    for j, h in enumerate(HEADERS, start=2):  # starts at column B, matching the real template
        c = ws.cell(row=header_row, column=j, value=h)
        c.font = FONT_BOLD
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="top")
        c.border = BORDER

    r = header_row + 1
    first_data_row = r
    for row in rows:
        _write_company_row(ws, r, row, red_threshold_pct)
        r += 1
    last_data_row = r - 1

    if mode == "group":
        _write_total_row(ws, r, first_data_row, last_data_row)
        r += 1

    last_row = r - 1

    ws.column_dimensions["A"].width = 4.3
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 14.3
    ws.column_dimensions["D"].width = 13.5
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 9

    footer_row = last_row + 2
    ws.cell(row=footer_row, column=2,
             value=f"Generated on: {datetime.now().strftime('%d-%b-%Y %H:%M')}").font = Font(name="Calibri", size=8, italic=True, color="FF6C757D")
    ws.cell(row=footer_row + 1, column=2,
             value=f"Loss Ratio = (Claim + Commission) / GP  |  Red > {red_threshold_pct:.0f}%  |  Amber \u2265 {AMBER_THRESHOLD:.0f}%").font = Font(name="Calibri", size=8, italic=True, color="FF6C757D")

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)

    _recalculate(output_path)
    return output_path


def _write_company_row(ws, r, row, red_threshold_pct):
    gp = row["Gross Premium Written"]
    claim = row["Incurred Loss"]
    comm = row["Commission Accrual"]

    name_cell = ws.cell(row=r, column=2, value=row["Company Name"])
    _write_text_safe(name_cell)  # Company Name is untrusted - comes straight from the uploaded source file
    name_cell.font = FONT_NORMAL
    ws.cell(row=r, column=3, value=gp)
    ws.cell(row=r, column=4, value=claim)
    ws.cell(row=r, column=5, value=comm)

    for col in (3, 4, 5):
        c = ws.cell(row=r, column=col)
        c.font = FONT_NORMAL
        c.number_format = ACCOUNTING_FORMAT
        c.border = BORDER
    ws.cell(row=r, column=2).border = BORDER

    lr_cell = ws.cell(row=r, column=6)
    if gp == 0:
        lr_cell.value = "N/A"
        lr_cell.font = NA_FONT
        lr_cell.fill = NA_FILL
        lr_cell.alignment = Alignment(horizontal="center")
    else:
        lr_cell.value = f"=SUM(D{r}:E{r})/C{r}%"
        lr_cell.number_format = LR_FORMAT
        lr_cell.alignment = Alignment(horizontal="center")
        lr_value_pct = row["loss_ratio"] * 100 if row["loss_ratio"] is not None else 0
        if lr_value_pct > red_threshold_pct:
            lr_cell.fill = RED_FILL
            lr_cell.font = RED_FONT
        elif lr_value_pct >= AMBER_THRESHOLD:
            lr_cell.fill = AMBER_FILL
            lr_cell.font = AMBER_FONT
        else:
            lr_cell.font = FONT_NORMAL
    lr_cell.border = BORDER


def _write_total_row(ws, r, first_data_row, last_data_row):
    ws.cell(row=r, column=2, value="TOTAL").font = FONT_BOLD

    for col, letter in ((3, "C"), (4, "D"), (5, "E")):
        c = ws.cell(row=r, column=col, value=f"=SUM({letter}{first_data_row}:{letter}{last_data_row})")
        c.font = FONT_BOLD
        c.number_format = ACCOUNTING_FORMAT

    lr_cell = ws.cell(row=r, column=6, value=f"=SUM(D{r}:E{r})/C{r}%")
    lr_cell.number_format = LR_FORMAT
    lr_cell.font = FONT_BOLD
    lr_cell.fill = YELLOW_FILL
    lr_cell.alignment = Alignment(horizontal="center")


def _recalculate(filepath):
    """Run the project's LibreOffice-based recalculation so formula cells
    have cached values when opened (mandatory per the xlsx skill whenever
    a file contains formulas). Non-fatal if the recalc script isn't
    available in this environment - Excel will simply compute the
    formulas itself the moment the user opens the file."""
    recalc_script = "/mnt/skills/public/xlsx/scripts/recalc.py"
    if not os.path.exists(recalc_script):
        return
    try:
        subprocess.run([sys.executable, recalc_script, filepath, "30"],
                        capture_output=True, timeout=40)
    except Exception:
        pass  # best-effort - formulas still evaluate fine when opened in real Excel


if __name__ == "__main__":
    import loss_ratio_processor as lrp
    import tempfile
    import pandas as pd

    tmp_dir = tempfile.mkdtemp()
    files = []
    data = [
        ("POPULAR ASEPTIC PACKAGING", 8572128, 1222808, 984695),
        ("POPULAR CEMENT INDUSTRIES", 142229, 100925, 17289),
        ("POPULAR JUICE INDUSTRIES", 3538618, 12204510, 417252),  # >100% -> red
    ]
    for name, gp, claim, comm in data:
        fp = os.path.join(tmp_dir, f"{name}.xlsx")
        pd.DataFrame({
            "Company Name": [name],
            "Gross Premium Written": [gp],
            "Incurred Loss": [claim],
            "Commission Accrual": [comm],
        }).to_excel(fp, index=False)
        files.append(fp)

    rows, errors = lrp.process_files(files)
    out = build_loss_ratio_workbook(rows, mode="group", output_path=os.path.join(tmp_dir, "LossRatio_Test.xlsx"), group_name="POPULAR GROUP")
    print("Wrote:", out)

    import openpyxl
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb.active
    for r in range(1, ws.max_row + 1):
        print(r, [ws.cell(row=r, column=c).value for c in range(2, 7)])
