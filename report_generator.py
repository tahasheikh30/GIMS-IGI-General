"""
Generates one Excel file per qualifying client, showing the
Department x Year outstanding-balance breakdown, styled to match
a standard pivot-table look.

PERFORMANCE NOTE: _write_raw_data_into() is the path most likely to see
a huge dataframe in practice (a client's own raw transaction rows, which
on a full production export can run into the hundreds of thousands of
rows). It's been rewritten to, first, convert the NaN/date handling once,
vectorized, over whole columns instead of per-cell in a Python loop,
second, iterate with itertuples() instead of iterrows() (meaningfully
faster for large frames), and third, skip per-cell font/border past
perf_utils.STYLE_ROW_THRESHOLD rows. Values and number formats are
always written correctly either way, only the cosmetic per-cell
font/border is skipped on very large sheets, since applying an openpyxl
style object to millions of individual cells is the dominant cost once
you're past roughly 50,000 rows, well before the pandas iteration itself
becomes the bottleneck.
"""
import os
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from perf_utils import style_budget_ok, ThrottledReporter
except ImportError:  # perf_utils.py not yet added to this checkout, degrade gracefully
    def style_budget_ok(n_rows):
        return True

    class ThrottledReporter:
        def __init__(self, *a, **k):
            pass

        def maybe(self, *a, **k):
            pass

        def final(self, *a, **k):
            pass

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
TOTAL_FONT = Font(name="Arial", bold=True)
BODY_FONT = Font(name="Arial")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM_FORMAT = "#,##0;(#,##0);-"


def _write_text_safe(cell):
    """
    Guard against a confirmed real-world data-quality issue: some source
    exports carry junk placeholder text in a text column (e.g. "=../-",
    "====-=-") instead of a real value. openpyxl auto-detects ANY string
    value starting with "=" as a live Excel formula - and since junk text
    like that isn't valid formula syntax, Excel then flags the whole
    workbook as containing "unreadable content" and silently strips it on
    open (the "we found a problem with some content" repair prompt). This
    sheet writes every original source column verbatim (raw transaction
    rows), so it carries the same risk as receivable_generator.py's own
    data sheets - force the cell back to plain text whenever openpyxl
    auto-classified it as a formula, preserving the original text exactly
    (no stray leading apostrophe) with zero functional loss.
    """
    if cell.data_type == "f":
        cell.data_type = "s"
    return cell


def safe_filename(name):
    name = re.sub(r'[\\/*?:"<>|]', "_", str(name)).strip()
    return name[:150]  # keep filenames reasonable length


def _write_raw_data_sheet(wb, raw_df):
    """Add a 'Raw Data' sheet containing the client's original transaction rows."""
    ws = wb.create_sheet("Raw Data")
    _write_raw_data_into(ws, raw_df)


def _prepare_for_excel(raw_df, date_cols):
    """
    One-time, vectorized cleanup of the whole dataframe before writing,
    replacing what used to be a per-cell isna()/hasattr() chain executed
    once for every single cell. All of this runs as pandas C-level
    column operations, so it stays fast even at a million rows:
      - datetime columns: NaT -> None, else stripped down to a plain
        date object (no "00:00:00" time component)
      - every other column: NaN -> None so openpyxl never sees a NaN
        float, and any numpy scalar dtype is converted to a plain Python
        object via astype(object), which openpyxl can always write
    Returns a NEW dataframe of plain Python objects, safe to iterate with
    itertuples() and hand straight to ws.cell(value=...).
    """
    out = {}
    for col in raw_df.columns:
        s = raw_df[col]
        if col in date_cols:
            out[col] = s.dt.date.astype(object).where(s.notna(), None)
        else:
            s = s.astype(object)
            out[col] = s.where(pd.notna(s), None)
    return pd.DataFrame(out, columns=raw_df.columns)


def _write_raw_data_into(ws, raw_df, report=None):
    """Write raw_df into an already-created worksheet, starting at A1 with
    no other content on the sheet - shared by both the fallback 'Raw Data'
    sheet and the clean single-sheet workbook used as native pivot source.

    report: optional callable(str) for progress lines on very large
    frames (see perf_utils.run_in_background) - safe to leave as None for
    normal, small-file use."""
    columns = list(raw_df.columns)
    n_rows = len(raw_df)
    apply_style = style_budget_ok(n_rows)

    # Columns that are genuinely datetime-typed get a clean date-only format
    # (no "00:00:00" time component cluttering every cell).
    date_cols = {col for col in columns if pd.api.types.is_datetime64_any_dtype(raw_df[col])}

    for j, col in enumerate(columns, start=1):
        c = ws.cell(row=1, column=j, value=str(col))
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")

    clean_df = _prepare_for_excel(raw_df, date_cols)
    reporter = ThrottledReporter(report, min_rows=25000) if report else None

    for i, row in enumerate(clean_df.itertuples(index=False, name=None), start=2):
        for j, (col, val) in enumerate(zip(columns, row), start=1):
            cell = ws.cell(row=i, column=j, value=val)
            _write_text_safe(cell)
            # Only style a cell that actually has a value - see the matching
            # comment in receivable_generator.py's _write_sheet(). Setting
            # .font on a still-valueless cell forces openpyxl to write it out
            # as an invalid empty numeric cell (<c t="n"> with no <v>),
            # which Excel then "repairs" (silently drops) on every open.
            # Real transaction data routinely has blank cells, so this
            # matters in production even though the original small test
            # data never happened to trigger it.
            if val is not None and apply_style:
                cell.font = BODY_FONT
                if col in date_cols:
                    cell.number_format = "DD-MMM-YYYY"
            elif val is not None and col in date_cols:
                # Number format is not cosmetic (a raw date serial without
                # it is unreadable), so it's kept even past the style
                # budget - only the per-cell Font object is skipped.
                cell.number_format = "DD-MMM-YYYY"
        if reporter:
            reporter.maybe(i - 1, total=n_rows, label="rows written")

    if reporter:
        reporter.final(n_rows, label="rows written")
    if not apply_style and report:
        report(f"  Note: {n_rows:,} rows exceeds the styling budget - per-cell font formatting "
               f"was skipped for speed (values and number formats are unaffected).")

    for j, col in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(14, min(30, len(str(col)) + 4))

    ws.freeze_panes = "A2"


def write_raw_only_workbook(client_name, raw_df, output_dir, report=None):
    """
    Write just the client's raw transaction rows to a clean single-sheet
    workbook (header at row 1, nothing else on the sheet) - used as the
    source data for a genuine Excel PivotTable built afterwards via
    pivot_native_generator.add_native_pivot_table(). Keeping this sheet
    free of any extra header rows/text matters: Excel's PivotTable source
    range detection works most reliably against a plain rectangular table
    starting at A1.

    report: optional callable(str) for progress on large clients - see
    perf_utils.run_in_background.
    """
    os.makedirs(output_dir, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Raw Data"
    _write_raw_data_into(ws, raw_df, report=report)
    filename = f"{safe_filename(client_name)}_OutstandingBalance_Report.xlsx"
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)
    return filepath


def write_client_report(client_name, pivot_df, output_dir, extra_info=None, raw_df=None, report=None):
    """
    Write a single client's Department x Year pivot to its own Excel file.
    extra_info: optional dict of {label: value} shown above the table
        (e.g. {"Due Date": ..., "Policy Ref": ...})
    raw_df: optional dataframe of that client's original transaction rows
        (from Sheet1). If provided, written as a second sheet "Raw Data"
        in the same workbook.
    """
    os.makedirs(output_dir, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Outstanding Balance"

    row = 1
    ws.cell(row=row, column=1, value="Client:").font = Font(name="Arial", bold=True)
    ws.cell(row=row, column=2, value=client_name).font = Font(name="Arial", bold=True)
    row += 1

    if extra_info:
        for label, value in extra_info.items():
            ws.cell(row=row, column=1, value=f"{label}:").font = BODY_FONT
            ws.cell(row=row, column=2, value=value).font = BODY_FONT
            row += 1
    row += 1

    header_row = row
    ws.cell(row=header_row, column=1, value="Department").font = HEADER_FONT
    ws.cell(row=header_row, column=1).fill = HEADER_FILL
    for j, col in enumerate(pivot_df.columns, start=2):
        c = ws.cell(row=header_row, column=j, value=str(col))
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")

    for i, (dept, r) in enumerate(pivot_df.iterrows(), start=header_row + 1):
        is_total_row = (dept == "Grand Total")
        cell = ws.cell(row=i, column=1, value=dept)
        cell.font = TOTAL_FONT if is_total_row else BODY_FONT
        if is_total_row:
            cell.fill = TOTAL_FILL
        for j, col in enumerate(pivot_df.columns, start=2):
            val = r[col]
            c = ws.cell(row=i, column=j, value=float(val))
            c.number_format = NUM_FORMAT
            c.font = TOTAL_FONT if is_total_row else BODY_FONT
            if is_total_row:
                c.fill = TOTAL_FILL

    last_row = header_row + len(pivot_df)
    last_col = 1 + len(pivot_df.columns)
    for r in range(header_row, last_row + 1):
        for c in range(1, last_col + 1):
            ws.cell(row=r, column=c).border = BORDER

    ws.column_dimensions["A"].width = 22
    for j in range(2, last_col + 1):
        ws.column_dimensions[get_column_letter(j)].width = 16

    if raw_df is not None and len(raw_df) > 0:
        ws_raw = wb.create_sheet("Raw Data")
        _write_raw_data_into(ws_raw, raw_df, report=report)

    filename = f"{safe_filename(client_name)}_OutstandingBalance_Report.xlsx"
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)
    return filepath


if __name__ == "__main__":
    import data_processor as dp
    raw = dp.load_raw_data("/mnt/user-data/uploads/Karachi_Conventional_Outstanding.xlsx", sheet_name="Sheet1")
    clean, issues, _dropped = dp.validate_data(raw)
    df, totals = dp.build_pivot(clean)
    q = dp.qualifying_clients(df, totals, threshold=2_000_000)
    out_dir = "/home/claude/balance_app/sample_output"
    generated = []
    for client in list(q.index)[:3]:
        pivot = dp.client_pivot_table(df, client)
        raw_rows = dp.get_client_raw_rows(df, client)
        fp = write_client_report(client, pivot, out_dir, raw_df=raw_rows)
        generated.append(fp)
    print("\n".join(generated))
