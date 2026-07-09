"""
Generates one Excel file per qualifying client, showing the
Department x Year outstanding-balance breakdown, styled to match
a standard pivot-table look.
"""
import os
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
TOTAL_FONT = Font(name="Arial", bold=True)
BODY_FONT = Font(name="Arial")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM_FORMAT = "#,##0;(#,##0);-"


def safe_filename(name):
    name = re.sub(r'[\\/*?:"<>|]', "_", str(name)).strip()
    return name[:150]  # keep filenames reasonable length


def _write_raw_data_sheet(wb, raw_df):
    """Add a 'Raw Data' sheet containing the client's original transaction rows."""
    ws = wb.create_sheet("Raw Data")
    _write_raw_data_into(ws, raw_df)


def _write_raw_data_into(ws, raw_df):
    """Write raw_df into an already-created worksheet, starting at A1 with
    no other content on the sheet - shared by both the fallback 'Raw Data'
    sheet and the clean single-sheet workbook used as native pivot source."""
    columns = list(raw_df.columns)

    # Columns that are genuinely datetime-typed get a clean date-only format
    # (no "00:00:00" time component cluttering every cell).
    date_cols = {col for col in columns if pd.api.types.is_datetime64_any_dtype(raw_df[col])}

    for j, col in enumerate(columns, start=1):
        c = ws.cell(row=1, column=j, value=str(col))
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")

    for i, (_, row) in enumerate(raw_df.iterrows(), start=2):
        for j, col in enumerate(columns, start=1):
            val = row[col]
            is_date_col = col in date_cols
            # openpyxl can't write pandas Timestamp/NaT or numpy types directly in all cases
            if pd.isna(val):
                val = None
            elif is_date_col and hasattr(val, "date"):
                val = val.date()  # strip the time component entirely
            elif hasattr(val, "to_pydatetime"):
                val = val.to_pydatetime()
            elif hasattr(val, "item"):  # numpy scalar -> native python
                val = val.item()
            cell = ws.cell(row=i, column=j, value=val)
            cell.font = BODY_FONT
            if is_date_col and val is not None:
                cell.number_format = "DD-MMM-YYYY"

    for j, col in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(14, min(30, len(str(col)) + 4))

    ws.freeze_panes = "A2"


def write_raw_only_workbook(client_name, raw_df, output_dir):
    """
    Write just the client's raw transaction rows to a clean single-sheet
    workbook (header at row 1, nothing else on the sheet) - used as the
    source data for a genuine Excel PivotTable built afterwards via
    pivot_native_generator.add_native_pivot_table(). Keeping this sheet
    free of any extra header rows/text matters: Excel's PivotTable source
    range detection works most reliably against a plain rectangular table
    starting at A1.
    """
    os.makedirs(output_dir, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Raw Data"
    _write_raw_data_into(ws, raw_df)
    filename = f"{safe_filename(client_name)}_OutstandingBalance_Report.xlsx"
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)
    return filepath


def write_client_report(client_name, pivot_df, output_dir, extra_info=None, raw_df=None):
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
        _write_raw_data_sheet(wb, raw_df)

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
