"""
Writes the Withholding Tax (Income Tax) working output: the General
Ledger with every original column preserved exactly as-is, plus the
columns withholding_tax_processor.py adds (NTN/CNIC match results,
Status/Address/Location, and one Filer/Non-Filer column per ATL sheet).

Any row withholding_tax_processor flagged with a "Manual Review Reason"
(no confident NTN/CNIC could be established) is highlighted across its
ENTIRE row - the same "never force a bad match, highlight it instead"
convention every other matching module in GIMS uses (Policy Check,
Claims Check, Claims Register) - so it's immediately obvious which
payees still need a person to look at them before the withholding tax
figures are relied on, without silently guessing.

Single flat sheet, not a multi-sheet dashboard - this is a working
paper for review/audit, not a set of per-company statements, so the
simpler shape used by premium_register_generator.py (one sheet, matched
rows styled plainly, problem rows highlighted red) is the closer fit
than the Summary-plus-charts pattern used by the coinsurer/company-split
modules.
"""
import os
import re
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from perf_utils import RowProgress, StyleIndexCache
except ImportError:  # perf_utils.py not yet added to this checkout, degrade gracefully
    class RowProgress:
        def __init__(self, *a, **k):
            pass

        def maybe(self, *a, **k):
            pass

        def final(self, *a, **k):
            pass

    class StyleIndexCache:
        """Plain, always-correct fallback if perf_utils isn't available -
        no caching, just sets each style attribute directly every time
        (the original, slower-but-safe behavior)."""
        def apply(self, cell, font=None, border=None, fill=None, alignment=None, number_format=None):
            if font is not None:
                cell.font = font
            if border is not None:
                cell.border = border
            if fill is not None:
                cell.fill = fill
            if alignment is not None:
                cell.alignment = alignment
            if number_format is not None:
                cell.number_format = number_format

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial")

REVIEW_FILL = PatternFill("solid", start_color="F8D7DA", end_color="F8D7DA")     # light red - manual review needed
REVIEW_FONT = Font(name="Arial", color="842029")
FILER_FILL = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")       # light green - a confirmed "Filer" cell
NOT_DETERMINED_FILL = PatternFill("solid", start_color="FDF3EA", end_color="FDF3EA")  # light amber - "Not Determined" ATL cell
NEW_COL_FILL = PatternFill("solid", start_color="EAF2FA", end_color="EAF2FA")     # light blue - other new (matched) columns

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

SHEET_NAME = "Withholding Tax Working"

# Every column this module adds, other than the ATL Filer/Non-Filer
# columns themselves (those come from result["atl_sheet_order"], since
# their names/count depend on the uploaded ATL file) - used to decide
# which columns get the light-blue "new data" tint on an otherwise-clean
# row.
NEW_BASE_COLUMNS = [
    "NTN (Locator Match)", "CNIC (Locator Match)", "Locator Match Status",
    "NTN (Final)", "CNIC (Final)", "Manual Review Reason",
    "Status", "Address", "Location",
]


def _write_text_safe(cell):
    """Same junk-placeholder-text guard used by every other GIMS
    generator - a source value like '=../-' gets auto-classified by
    openpyxl as a live formula, which then makes Excel flag the whole
    workbook as needing repair on open."""
    if cell.data_type == "f":
        cell.data_type = "s"
    return cell


def _safe_filename_part(name):
    name = re.sub(r"[^\w]+", "_", str(name).strip())
    return re.sub(r"_+", "_", name).strip("_")[:80] or "output"


def build_output_filename(gl_source_filename, run_date=None):
    run_date = run_date or date.today()
    base = os.path.splitext(os.path.basename(gl_source_filename))[0]
    return f"{_safe_filename_part(base)}_WHT_Working_{run_date.strftime('%d%m%Y')}.xlsx"


def _prepare_for_excel(df, date_cols):
    """Vectorized, one-pass NaN/date cleanup - see the matching helper in
    report_generator.py for the full rationale."""
    out = {}
    for col in df.columns:
        s = df[col]
        if col in date_cols:
            out[col] = s.dt.date.astype(object).where(s.notna(), None) if pd.api.types.is_datetime64_any_dtype(s) else s
        else:
            s = s.astype(object)
            out[col] = s.where(pd.notna(s), None)
    return pd.DataFrame(out, columns=df.columns)


def build_withholding_tax_workbook(result, output_path, gl_source_filename=None, report=None):
    """
    result: the dict returned by
        withholding_tax_processor.run_withholding_tax_processing().
    output_path: full filepath to write to.
    report: optional callable(str) for progress on a large GL - if it's
        a perf_utils.Reporter, row-level progress also drives the numeric
        progress bar. Safe to leave as None.

    Every row whose "Manual Review Reason" is set gets its entire row
    highlighted light red, so it's immediately visible which payees
    still need a person to confirm an NTN/CNIC before the Status and
    Filer/Non-Filer figures next to them are relied on. On an otherwise-
    clean row, the columns this module added are tinted light blue for
    easy visual separation from the original GL data, with a "Filer"
    cell tinted green and a "Not Determined" ATL cell tinted amber (an
    ATL result that isn't a confident Filer/Non-Filer - see
    withholding_tax_processor.match_atl for why that's tracked
    separately from Non-Filer).
    Returns output_path.
    """
    df = result["merged_df"]
    atl_sheet_order = result["atl_sheet_order"]
    new_cols = set(NEW_BASE_COLUMNS) | set(atl_sheet_order)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    columns = list(df.columns)
    date_cols = {c for c in columns if pd.api.types.is_datetime64_any_dtype(df[c])}

    for j, col in enumerate(columns, start=1):
        c = ws.cell(row=1, column=j, value=str(col))
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 20

    n_total = len(df)
    review_col_idx = columns.index("Manual Review Reason")

    clean_df = _prepare_for_excel(df, date_cols)
    progress = RowProgress(report, total=n_total, label="rows written") if report else None
    style_cache = StyleIndexCache()

    for i, row in enumerate(clean_df.itertuples(index=False, name=None), start=2):
        needs_review = row[review_col_idx] is not None
        for j, (col, val) in enumerate(zip(columns, row), start=1):
            cell = ws.cell(row=i, column=j, value=val)
            _write_text_safe(cell)
            if val is None:
                continue
            if needs_review:
                style_cache.apply(cell, font=REVIEW_FONT, border=BORDER, fill=REVIEW_FILL)
            else:
                fill = None
                if col in atl_sheet_order:
                    if val == "Filer":
                        fill = FILER_FILL
                    elif val == "Not Determined":
                        fill = NOT_DETERMINED_FILL
                    else:
                        fill = NEW_COL_FILL
                elif col in new_cols:
                    fill = NEW_COL_FILL
                style_cache.apply(cell, font=BODY_FONT, border=BORDER, fill=fill)
        if progress:
            progress.maybe(i - 1)
    if progress:
        progress.final()

    for j, col in enumerate(columns, start=1):
        width = 20 if col in new_cols else max(12, min(30, len(str(col)) + 4))
        ws.column_dimensions[get_column_letter(j)].width = width

    n_review = int(df["Manual Review Reason"].notna().sum())
    n_matched_locator = int((df["Locator Match Status"] == "Matched").sum())
    pct = f"{n_matched_locator / n_total:.1%}" if n_total else "N/A"
    note_row = n_total + 3
    ws.cell(row=note_row, column=1,
            value=f"{n_matched_locator} of {n_total} ({pct}) payee(s) matched to an NTN/CNIC via the NTN "
                  f"Locator. {n_review} row(s) highlighted below need manual review - no confident "
                  f"NTN/CNIC could be determined for that payee.")\
        .font = Font(name="Arial", size=9, italic=True, color="6C7A89")

    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    import sys
    import withholding_tax_processor as wtp

    gl_fp = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/GL_-_Combined-26-07-2026.xlsx"
    atl_fp = sys.argv[2] if len(sys.argv) > 2 else "/mnt/user-data/uploads/ATL_IT.xlsx"
    output_dir = sys.argv[3] if len(sys.argv) > 3 else "/mnt/user-data/outputs"

    result = wtp.run_withholding_tax_processing(
        gl_fp, "WORKING WHT (2)", gl_fp, None, gl_fp, None, atl_fp, report=print)

    filename = build_output_filename(gl_fp)
    fp = build_withholding_tax_workbook(result, os.path.join(output_dir, filename), gl_source_filename=gl_fp)
    print("Wrote:", fp)
