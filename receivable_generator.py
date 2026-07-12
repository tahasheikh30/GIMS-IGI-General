"""
Writes one Excel workbook per Client (+ Company) group, containing exactly
two sheets: "Premium Receivable" and "Claims Receivable". Styling mirrors
report_generator.py (Outstanding Premium module) so output looks
consistent with the rest of GIMS.

Works directly off the ORIGINAL source column names (Class, Leader Ref.
No, Code, Client Name, Company, Premium Receivable, Claim Receivable,
etc.) as returned by receivable_processor.build_groups() - nothing is
renamed, so the output sheets read exactly like the source file's own
columns, just filtered to that Client+Company group and split by category.
"""
import os
import re
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial")
EMPTY_FONT = Font(name="Arial", italic=True, color="808080")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM_FORMAT = "#,##0;(#,##0);-"
DATE_FORMAT = "DD-MMM-YYYY"

PREMIUM_SHEET = "Premium Receivable"
CLAIMS_SHEET = "Claims Receivable"


def safe_filename_part(name):
    """Sanitize a single name component for filesystem/attachment safety:
    strips special characters and collapses whitespace to underscores, per
    spec ('no spaces or special characters')."""
    name = str(name).strip()
    name = re.sub(r"[^\w]+", "_", name)  # anything not alnum/underscore -> _
    name = re.sub(r"_+", "_", name).strip("_")
    return name[:100] if name else "UNKNOWN"


def build_output_filename(client_name, company_name, run_date=None):
    run_date = run_date or date.today()
    date_str = run_date.strftime("%d%m%Y")
    return f"{safe_filename_part(client_name)}_{safe_filename_part(company_name)}_Receivable_{date_str}.xlsx"


def _write_sheet(ws, df, amount_col, drop_cols=None):
    """
    amount_col: name of this sheet's own amount column (e.g. "Premium
        Receivable" or "Claim Receivable") - gets the accounting number
        format; every other column is written as-is.
    drop_cols: columns to omit from this sheet (typically the OTHER
        category's amount column, so the Premium Receivable sheet doesn't
        also show a redundant all-zero Claim Receivable column).
    """
    drop_cols = [c for c in (drop_cols or []) if c in df.columns]
    columns = [c for c in df.columns if c not in drop_cols]

    for j, col in enumerate(columns, start=1):
        c = ws.cell(row=1, column=j, value=str(col))
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")
    for j in range(1, len(columns) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 18
    ws.freeze_panes = "A2"

    if len(df) == 0:
        ws.cell(row=2, column=1, value="No records in this category.").font = EMPTY_FONT
        return

    date_cols = {col for col in columns if col in df.columns and pd.api.types.is_datetime64_any_dtype(df[col])}

    for i, (_, row) in enumerate(df.iterrows(), start=2):
        for j, col in enumerate(columns, start=1):
            val = row[col]
            is_date_col = col in date_cols
            if pd.isna(val):
                val = None
            elif is_date_col and hasattr(val, "date"):
                val = val.date()
            elif hasattr(val, "to_pydatetime"):
                val = val.to_pydatetime()
            elif hasattr(val, "item"):
                val = val.item()
            cell = ws.cell(row=i, column=j, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            if col == amount_col and val is not None:
                cell.number_format = NUM_FORMAT
            elif is_date_col and val is not None:
                cell.number_format = DATE_FORMAT

    for j, col in enumerate(columns, start=1):
        width = max(14, min(30, len(str(col)) + 4))
        ws.column_dimensions[get_column_letter(j)].width = width


def build_receivable_workbook(client_name, company_name, premium_rows, claims_rows, output_dir,
                               premium_amount_col, claims_amount_col, run_date=None):
    """
    Write a single Client(+Company) workbook with exactly two sheets:
    "Premium Receivable" and "Claims Receivable". Either sheet is written
    headers-only if that category has no rows for this group.

    premium_amount_col / claims_amount_col: the ORIGINAL source column
        names for each amount (e.g. "Premium Receivable", "Claim
        Receivable"), from receivable_processor's colmap - so the Premium
        sheet gets number-formatted on its own amount column and drops
        the (redundant, all-zero) Claims amount column, and vice versa.
        In "single Amount column" mode these can both be the same name;
        nothing is dropped in that case.
    Returns the full filepath written.
    """
    os.makedirs(output_dir, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)

    ws_premium = wb.create_sheet(PREMIUM_SHEET)
    premium_drop = [claims_amount_col] if claims_amount_col != premium_amount_col else []
    _write_sheet(ws_premium, premium_rows, amount_col=premium_amount_col, drop_cols=premium_drop)

    ws_claims = wb.create_sheet(CLAIMS_SHEET)
    claims_drop = [premium_amount_col] if claims_amount_col != premium_amount_col else []
    _write_sheet(ws_claims, claims_rows, amount_col=claims_amount_col, drop_cols=claims_drop)

    filename = build_output_filename(client_name, company_name, run_date=run_date)
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)
    return filepath


if __name__ == "__main__":
    premium_df = pd.DataFrame({
        "Class": ["Fire", "Marine"],
        "Leader Ref. No": ["LR-1001", "LR-1002"],
        "Code": ["C01", "C01"],
        "Client Name": ["ACME CORP", "ACME CORP"],
        "Company": ["ADAMJEE INSURANCE COMPANY LTD", "ADAMJEE INSURANCE COMPANY LTD"],
        "Premium Receivable": [150000, 92000],
        "Claim Receivable": [0, 0],
    })
    claims_df = pd.DataFrame(columns=list(premium_df.columns))
    fp = build_receivable_workbook(
        "ACME CORP", "ADAMJEE INSURANCE COMPANY LTD", premium_df, claims_df,
        "/tmp/receivable_test_output", premium_amount_col="Premium Receivable", claims_amount_col="Claim Receivable",
    )
    print("Wrote:", fp)
