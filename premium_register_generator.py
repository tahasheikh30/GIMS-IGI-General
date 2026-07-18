"""
Writes the Premium Register matching output: every original column from
the leader's file, plus the matched IGI Policy Number and match
diagnostics, with unmatched rows highlighted so they're immediately
visible as needing manual entry into the Premium Register.

Styling mirrors the rest of GIMS (report_generator.py / receivable_
generator.py): navy header row, thin borders, Arial body font.

PERFORMANCE NOTE: a leader file can itself run into the hundreds of
thousands of rows, so this writer follows the same pattern as
report_generator.py / receivable_generator.py: NaN/date cleanup is done
once, vectorized, over whole columns rather than per-cell, iteration
uses itertuples() instead of iterrows(), and perf_utils.style_budget_ok()
gates the purely cosmetic per-cell styling on very large outputs.
Critically, the red "Not Matched" row highlighting is NOT gated by the
style budget even on huge files - that highlight is the whole point of
this report (it's what tells someone which rows need manual entry), not
a cosmetic flourish, so it's always applied regardless of size. Only the
green "matched, new column" tint and the plain body font on ordinary
matched cells are skipped past the threshold.
"""
import os
import re
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from perf_utils import style_budget_ok, RowProgress
except ImportError:  # perf_utils.py not yet added to this checkout, degrade gracefully
    def style_budget_ok(n_rows):
        return True

    class RowProgress:
        def __init__(self, *a, **k):
            pass

        def maybe(self, *a, **k):
            pass

        def final(self, *a, **k):
            pass

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial")
UNMATCHED_FILL = PatternFill("solid", start_color="F8D7DA", end_color="F8D7DA")  # light red
UNMATCHED_FONT = Font(name="Arial", color="842029")
MATCHED_NEW_COL_FILL = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")  # light green, matched rows only, new columns
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NEW_COLUMNS = ["IGI Policy Number", "Matched Leader Ref (Register)", "Match Score", "Match Status"]


def _safe_filename_part(name):
    name = re.sub(r"[^\w]+", "_", str(name).strip())
    return re.sub(r"_+", "_", name).strip("_")[:80] or "output"


def build_output_filename(leader_source_filename, run_date=None):
    run_date = run_date or date.today()
    base = os.path.splitext(os.path.basename(leader_source_filename))[0]
    return f"{_safe_filename_part(base)}_PolicyMatched_{run_date.strftime('%d%m%Y')}.xlsx"


def _prepare_for_excel(df, date_cols):
    """Vectorized, one-pass NaN/date cleanup - see the matching helper in
    report_generator.py for the full rationale."""
    out = {}
    for col in df.columns:
        s = df[col]
        if col in date_cols:
            out[col] = s.dt.date.astype(object).where(s.notna(), None)
        else:
            s = s.astype(object)
            out[col] = s.where(pd.notna(s), None)
    return pd.DataFrame(out, columns=df.columns)


def build_match_output(result_df, output_path, leader_source_filename=None, report=None):
    """
    result_df: the dataframe returned by
        premium_register_processor.match_leader_file() - every original
        leader-file column plus NEW_COLUMNS.
    output_path: full filepath to write to.
    report: optional callable(str) for progress on large results - if
        it's a perf_utils.Reporter, row-level write progress also drives
        the numeric progress bar. Safe to leave as None.
    Every row whose "Match Status" is "Not Matched" gets its entire row
    highlighted light red, so it's visually obvious at a glance which
    rows still need to be entered into the Premium Register by hand.
    Returns output_path.
    """
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Policy Number Matches"

    columns = list(result_df.columns)
    date_cols = {c for c in columns if pd.api.types.is_datetime64_any_dtype(result_df[c])}

    for j, col in enumerate(columns, start=1):
        c = ws.cell(row=1, column=j, value=str(col))
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 20

    n_matched = int((result_df["Match Status"] == "Matched").sum())
    n_total = len(result_df)
    apply_style = style_budget_ok(n_total)
    status_col_idx = columns.index("Match Status")

    clean_df = _prepare_for_excel(result_df, date_cols)
    progress = RowProgress(report, total=n_total, label="rows written") if report else None

    for i, row in enumerate(clean_df.itertuples(index=False, name=None), start=2):
        is_unmatched = row[status_col_idx] != "Matched"
        for j, (col, val) in enumerate(zip(columns, row), start=1):
            cell = ws.cell(row=i, column=j, value=val)
            if val is None:
                continue
            # The unmatched-row highlight is functional information (it's
            # what flags a row for manual entry), never skipped regardless
            # of file size. Everything else is cosmetic and respects the
            # style budget.
            if is_unmatched:
                cell.border = BORDER
                cell.font = UNMATCHED_FONT
                cell.fill = UNMATCHED_FILL
            elif apply_style:
                cell.border = BORDER
                cell.font = BODY_FONT
                if col in NEW_COLUMNS:
                    cell.fill = MATCHED_NEW_COL_FILL
            if col in date_cols:
                cell.number_format = "DD-MMM-YYYY"
            elif col == "Match Score":
                cell.number_format = "0.000"
        if progress:
            progress.maybe(i - 1)

    if progress:
        progress.final()
    if not apply_style and report:
        report(f"  Note: {n_total:,} rows exceeds the styling budget - per-cell font/fill on MATCHED rows "
               f"was skipped for speed (unmatched-row highlighting, values, and number formats are unaffected).")

    for j, col in enumerate(columns, start=1):
        width = 26 if col in NEW_COLUMNS else max(12, min(30, len(str(col)) + 4))
        ws.column_dimensions[get_column_letter(j)].width = width

    # Small summary note below the table
    note_row = len(result_df) + 3
    ws.cell(row=note_row, column=1,
             value=f"Matched {n_matched} of {n_total} ({n_matched / n_total:.1%}) - "
                   f"unmatched rows are highlighted and need manual entry into the Premium Register.")\
        .font = Font(name="Arial", size=9, italic=True, color="6C7A89")

    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    import sys
    import premium_register_processor as prp

    leader_fp = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/Premium_EFU_Jinnah_Division__statement.xlsx"
    register_fp = sys.argv[2] if len(sys.argv) > 2 else "/mnt/user-data/uploads/Premium_Register_KU4.xls"
    output_dir = sys.argv[3] if len(sys.argv) > 3 else "/mnt/user-data/outputs"

    leader_df, nums_col = prp.load_leader_file(leader_fp)
    register_df, leader_ref_col, igi_policy_col = prp.load_register_file(register_fp)
    result = prp.match_leader_file(leader_df, nums_col, register_df, leader_ref_col, igi_policy_col)

    filename = build_output_filename(leader_fp)
    fp = build_match_output(result, os.path.join(output_dir, filename), leader_source_filename=leader_fp)
    matched = int((result["Match Status"] == "Matched").sum())
    print(f"Matched {matched}/{len(result)} ({matched/len(result):.1%})")
    print("Wrote:", fp)
