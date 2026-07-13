"""
Writes one Excel workbook per insurance company (all of that company's
clients combined into a single file, per spec), containing exactly three
sheets: "Premium Receivable", "Claims Receivable", and "Summary" (a
client-by-client chart + totals). Styling mirrors report_generator.py
(Outstanding Premium module) so output looks consistent with the rest of
GIMS.

Works directly off the ORIGINAL source column names (Class, Leader Ref.
No, Code, Client Name, Company, Premium Receivable, Claim Receivable,
etc.) as returned by receivable_processor.build_company_groups() -
nothing is renamed, so the two data sheets read exactly like the source
file's own columns, just filtered to that company's bucket and split by
category. Companies whose name contains "OFC" are pooled by
build_company_groups() into a single shared "OFC" bucket rather than each
getting their own file - this module doesn't need to know about that
distinction itself, it just writes whatever bucket_label + rows it's given.
"""
import os
import re
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial")
EMPTY_FONT = Font(name="Arial", italic=True, color="808080")
TITLE_FONT = Font(name="Arial", bold=True, size=14, color="1F4E78")
TOTAL_LABEL_FONT = Font(name="Arial", size=10, color="5B6B79")
TOTAL_VALUE_FONT = Font(name="Arial", bold=True, size=12)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM_FORMAT = "#,##0;(#,##0);-"
DATE_FORMAT = "DD-MMM-YYYY"

PREMIUM_SHEET = "Premium Receivable"
CLAIMS_SHEET = "Claims Receivable"
SUMMARY_SHEET = "Summary"


def safe_filename_part(name):
    """Sanitize a single name component for filesystem/attachment safety:
    strips special characters and collapses whitespace to underscores, per
    spec ('no spaces or special characters')."""
    name = str(name).strip()
    name = re.sub(r"[^\w]+", "_", name)  # anything not alnum/underscore -> _
    name = re.sub(r"_+", "_", name).strip("_")
    return name[:100] if name else "UNKNOWN"


def build_output_filename(bucket_label, run_date=None):
    """bucket_label: the insurance company name, or "OFC" for the pooled
    OFC bucket - one filename component now (not client+company), since
    each output file covers a whole company's worth of clients."""
    run_date = run_date or date.today()
    date_str = run_date.strftime("%d%m%Y")
    return f"{safe_filename_part(bucket_label)}_Receivable_{date_str}.xlsx"


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


def _write_summary_sheet(ws, bucket_label, client_breakdown_df, premium_total, claims_total, num_clients):
    """
    Third sheet: a short totals block, a small Premium-vs-Claims pie
    (share of the two totals), and a per-client bar chart built from
    real embedded openpyxl Excel chart objects (interactive/refreshable
    in Excel itself - not a static picture) fed by a plain data table
    also written on this sheet, so the numbers behind every chart are
    visible and auditable, not hidden inside the chart object.
    """
    ws.sheet_view.showGridLines = False

    ws.merge_cells("B2:F2")
    banner = ws["B2"]
    banner.value = f"Receivable Summary \u2014 {bucket_label}"
    banner.font = TITLE_FONT
    banner.alignment = Alignment(horizontal="left", vertical="center")

    ws["B3"] = f"{num_clients} client{'s' if num_clients != 1 else ''} \u2022 generated {date.today().strftime('%d-%b-%Y')}"
    ws["B3"].font = TOTAL_LABEL_FONT

    # ---- small totals table (also the pie chart's data source) ----
    totals_header_row = 5
    ws.cell(row=totals_header_row, column=2, value="Category").font = HEADER_FONT
    ws.cell(row=totals_header_row, column=2).fill = HEADER_FILL
    ws.cell(row=totals_header_row, column=3, value="Total").font = HEADER_FONT
    ws.cell(row=totals_header_row, column=3).fill = HEADER_FILL

    ws.cell(row=totals_header_row + 1, column=2, value="Premium Receivable").font = BODY_FONT
    ws.cell(row=totals_header_row + 1, column=3, value=float(premium_total)).number_format = NUM_FORMAT
    ws.cell(row=totals_header_row + 1, column=3).font = TOTAL_VALUE_FONT

    ws.cell(row=totals_header_row + 2, column=2, value="Claims Receivable").font = BODY_FONT
    ws.cell(row=totals_header_row + 2, column=3, value=float(claims_total)).number_format = NUM_FORMAT
    ws.cell(row=totals_header_row + 2, column=3).font = TOTAL_VALUE_FONT

    for r in (totals_header_row, totals_header_row + 1, totals_header_row + 2):
        for c in (2, 3):
            ws.cell(row=r, column=c).border = BORDER

    if premium_total or claims_total:
        pie = PieChart()
        pie.title = "Premium vs Claims Share"
        pie_data = Reference(ws, min_col=3, min_row=totals_header_row, max_row=totals_header_row + 2)
        pie_cats = Reference(ws, min_col=2, min_row=totals_header_row + 1, max_row=totals_header_row + 2)
        pie.add_data(pie_data, titles_from_data=True)
        pie.set_categories(pie_cats)
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        pie.height = 8
        pie.width = 12
        ws.add_chart(pie, "E5")

    # ---- per-client breakdown table + bar chart ----
    table_header_row = 15
    ws.cell(row=table_header_row, column=2, value="Client Name").font = HEADER_FONT
    ws.cell(row=table_header_row, column=2).fill = HEADER_FILL
    ws.cell(row=table_header_row, column=3, value="Premium Receivable").font = HEADER_FONT
    ws.cell(row=table_header_row, column=3).fill = HEADER_FILL
    ws.cell(row=table_header_row, column=4, value="Claims Receivable").font = HEADER_FONT
    ws.cell(row=table_header_row, column=4).fill = HEADER_FILL

    if len(client_breakdown_df) == 0:
        ws.cell(row=table_header_row + 1, column=2, value="No client data available.").font = EMPTY_FONT
    else:
        for i, row in enumerate(client_breakdown_df.itertuples(index=False), start=table_header_row + 1):
            ws.cell(row=i, column=2, value=str(row[0])).font = BODY_FONT
            ws.cell(row=i, column=3, value=float(row[1])).number_format = NUM_FORMAT
            ws.cell(row=i, column=4, value=float(row[2])).number_format = NUM_FORMAT
            for c in (2, 3, 4):
                ws.cell(row=i, column=c).border = BORDER

        last_row = table_header_row + len(client_breakdown_df)
        chart = BarChart()
        chart.type = "col"
        chart.title = "Premium vs Claims Receivable by Client"
        chart.y_axis.title = "Amount (PKR)"
        chart.x_axis.title = "Client"
        chart.style = 10
        data = Reference(ws, min_col=3, max_col=4, min_row=table_header_row, max_row=last_row)
        cats = Reference(ws, min_col=2, min_row=table_header_row + 1, max_row=last_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.width = 26
        chart.height = 13
        ws.add_chart(chart, f"F{table_header_row}")

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20


def build_company_workbook(bucket_label, premium_rows, claims_rows, client_breakdown_df, output_dir,
                            premium_amount_col, claims_amount_col, run_date=None):
    """
    Write a single insurance-company workbook covering every client for
    that company, with exactly three sheets:
      1. "Premium Receivable" - every premium row for this company/bucket
      2. "Claims Receivable"  - every claims row for this company/bucket
      3. "Summary"            - totals, a Premium-vs-Claims pie, and a
                                 per-client bar chart

    bucket_label: the exact Company name, or "OFC" for the pooled bucket
        of every company whose name contains "OFC" (see
        receivable_processor.build_company_groups). Either data sheet is
        written headers-only if that category has no rows.

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

    ws_summary = wb.create_sheet(SUMMARY_SHEET)
    premium_total = pd.to_numeric(premium_rows[premium_amount_col], errors="coerce").fillna(0).sum() if len(premium_rows) else 0.0
    claims_total = pd.to_numeric(claims_rows[claims_amount_col], errors="coerce").fillna(0).sum() if len(claims_rows) else 0.0
    num_clients = int(client_breakdown_df["Client Name"].nunique()) if len(client_breakdown_df) else 0
    _write_summary_sheet(ws_summary, bucket_label, client_breakdown_df, premium_total, claims_total, num_clients)

    # Summary first, then the two data sheets - the visual overview
    # should be the first thing opened, with the line-item detail behind it.
    wb.move_sheet(SUMMARY_SHEET, offset=-2)

    filename = build_output_filename(bucket_label, run_date=run_date)
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)
    return filepath


if __name__ == "__main__":
    premium_df = pd.DataFrame({
        "Class": ["Fire", "Marine", "Fire"],
        "Leader Ref. No": ["LR-1001", "LR-1002", "LR-1003"],
        "Code": ["C01", "C01", "C02"],
        "Client Name": ["ACME CORP", "ACME CORP", "BETA MILLS"],
        "Company": ["ADAMJEE INSURANCE COMPANY LTD", "ADAMJEE INSURANCE COMPANY LTD", "ADAMJEE INSURANCE COMPANY LTD"],
        "Premium Receivable": [150000, 92000, 45000],
        "Claim Receivable": [0, 0, 0],
    })
    claims_df = pd.DataFrame(columns=list(premium_df.columns))
    breakdown = pd.DataFrame({
        "Client Name": ["ACME CORP", "BETA MILLS"],
        "premium_total": [242000, 45000],
        "claims_total": [0, 0],
    })
    fp = build_company_workbook(
        "ADAMJEE INSURANCE COMPANY LTD", premium_df, claims_df, breakdown,
        "/tmp/receivable_test_output", premium_amount_col="Premium Receivable", claims_amount_col="Claim Receivable",
    )
    print("Wrote:", fp)
