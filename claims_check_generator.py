"""
Writes the Claims Check matching output: every original column from the
leader's own claims sheet, plus the matched IGI Policy Number (Doc Ref #)
and match diagnostics, with unmatched rows highlighted so they're
immediately visible as needing manual review - i.e. a claim the leader
billed us for that we can't find any record of in our own Premium
Register.

Styling mirrors premium_register_generator.py (the Policy Check module
this is the claims-side counterpart to) so output looks consistent with
the rest of GIMS.

PERFORMANCE NOTE: same treatment as premium_register_generator.py - NaN/
date cleanup is vectorized over whole columns, iteration uses
itertuples() instead of iterrows(), and perf_utils.StyleIndexCache makes
per-cell styling cheap regardless of row count. The red "Not Matched" row
highlight is never skipped regardless of file size - it's the whole
point of this report, not decoration.
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
UNMATCHED_FILL = PatternFill("solid", start_color="F8D7DA", end_color="F8D7DA")  # light red
UNMATCHED_FONT = Font(name="Arial", color="842029")
MATCHED_NEW_COL_FILL = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")  # light green, matched rows only, new columns
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NEW_COLUMNS = ["IGI Policy Number", "Matched Leader Ref (Register)", "Match Score", "Match Status"]


def _write_text_safe(cell):
    """Same junk-placeholder-text guard used by every other GIMS
    generator (premium_register_generator.py / receivable_generator.py /
    claims_register_generator.py) - a source value like '=../-' gets
    auto-classified by openpyxl as a live formula, which then makes Excel
    flag the whole workbook as needing repair on open."""
    if cell.data_type == "f":
        cell.data_type = "s"
    return cell


def _safe_filename_part(name):
    name = re.sub(r"[^\w]+", "_", str(name).strip())
    return re.sub(r"_+", "_", name).strip("_")[:80] or "output"


def build_output_filename(leader_source_filename, run_date=None):
    run_date = run_date or date.today()
    base = os.path.splitext(os.path.basename(leader_source_filename))[0]
    return f"{_safe_filename_part(base)}_ClaimsChecked_{run_date.strftime('%d%m%Y')}.xlsx"


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
        claims_check_processor.match_leader_claims_file() - every
        original leader-claims-file column plus NEW_COLUMNS.
    output_path: full filepath to write to.
    report: optional callable(str) for progress on large results - if
        it's a perf_utils.Reporter, row-level write progress also drives
        the numeric progress bar. Safe to leave as None.
    Every row whose "Match Status" is "Not Matched" gets its entire row
    highlighted light red, so it's visually obvious at a glance which
    claims couldn't be found in our own Premium Register and need manual
    review. Returns output_path.
    """
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Claims Checked"

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
    status_col_idx = columns.index("Match Status")

    clean_df = _prepare_for_excel(result_df, date_cols)
    progress = RowProgress(report, total=n_total, label="rows written") if report else None
    style_cache = StyleIndexCache()

    for i, row in enumerate(clean_df.itertuples(index=False, name=None), start=2):
        is_unmatched = row[status_col_idx] != "Matched"
        for j, (col, val) in enumerate(zip(columns, row), start=1):
            cell = ws.cell(row=i, column=j, value=val)
            _write_text_safe(cell)
            if val is None:
                continue
            number_format = None
            if col in date_cols:
                number_format = "DD-MMM-YYYY"
            elif col == "Match Score":
                number_format = "0.000"
            if is_unmatched:
                style_cache.apply(cell, font=UNMATCHED_FONT, border=BORDER, fill=UNMATCHED_FILL, number_format=number_format)
            else:
                fill = MATCHED_NEW_COL_FILL if col in NEW_COLUMNS else None
                style_cache.apply(cell, font=BODY_FONT, border=BORDER, fill=fill, number_format=number_format)
        if progress:
            progress.maybe(i - 1)

    if progress:
        progress.final()

    for j, col in enumerate(columns, start=1):
        width = 26 if col in NEW_COLUMNS else max(12, min(30, len(str(col)) + 4))
        ws.column_dimensions[get_column_letter(j)].width = width

    note_row = len(result_df) + 3
    ws.cell(row=note_row, column=1,
             value=f"Matched {n_matched} of {n_total} ({n_matched / n_total:.1%}) - "
                   f"unmatched claims are highlighted and need manual review (the policy they reference "
                   f"could not be found in our own Premium Register).")\
        .font = Font(name="Arial", size=9, italic=True, color="6C7A89")

    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    import sys
    import claims_check_processor as ccp

    leader_fp = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/Indus_Settlement_F.xlsx"
    register_fp = sys.argv[2] if len(sys.argv) > 2 else "/mnt/user-data/uploads/Indus_Motor_Premium.xls"
    output_dir = sys.argv[3] if len(sys.argv) > 3 else "/mnt/user-data/outputs"

    leader_df, pmd_col = ccp.load_leader_claims_file(leader_fp)
    register_df, doc_col, base_doc_col = ccp.load_register_file(register_fp)
    result = ccp.match_leader_claims_file(leader_df, pmd_col, register_df, doc_col, base_doc_col)

    filename = build_output_filename(leader_fp)
    fp = build_match_output(result, os.path.join(output_dir, filename), leader_source_filename=leader_fp)
    matched = int((result["Match Status"] == "Matched").sum())
    print(f"Matched {matched}/{len(result)} ({matched/len(result):.1%})")
    print("Wrote:", fp)
