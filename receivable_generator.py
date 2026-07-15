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
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.title import Title
from openpyxl.chart.text import RichText, Text
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RegularTextRun, Font as DrawFont

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

# ---- Summary "dashboard" styling ----
NAVY = "1F4E78"
NAVY_DARK = "163A5C"
ACCENT_BLUE = "2E86AB"
CLAIMS_ORANGE = "D97B29"
SLATE_TEXT = "1C2733"
MUTED_TEXT = "6C7A89"
CARD_BORDER = "D9E2EC"
ROW_ALT_FILL = PatternFill("solid", start_color="F4F7FA", end_color="F4F7FA")
ROW_WHITE_FILL = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")

BANNER_FILL = PatternFill("solid", start_color=NAVY, end_color=NAVY)
BANNER_ACCENT_FILL = PatternFill("solid", start_color=ACCENT_BLUE, end_color=ACCENT_BLUE)
KPI_PREMIUM_FILL = PatternFill("solid", start_color="EAF2FA", end_color="EAF2FA")
KPI_CLAIMS_FILL = PatternFill("solid", start_color="FDF3EA", end_color="FDF3EA")
KPI_TOTAL_FILL = PatternFill("solid", start_color=NAVY, end_color=NAVY)
KPI_PREMIUM_ACCENT = PatternFill("solid", start_color=ACCENT_BLUE, end_color=ACCENT_BLUE)
KPI_CLAIMS_ACCENT = PatternFill("solid", start_color="D97B29", end_color="D97B29")
KPI_BORDER = Border(left=Side(style="thin", color=CARD_BORDER), right=Side(style="thin", color=CARD_BORDER),
                     top=Side(style="thin", color=CARD_BORDER), bottom=Side(style="thin", color=CARD_BORDER))

KPI_LABEL_FONT = Font(name="Arial", size=9, bold=True, color=MUTED_TEXT)
KPI_VALUE_FONT = Font(name="Arial", size=18, bold=True, color=SLATE_TEXT)
KPI_VALUE_FONT_WHITE = Font(name="Arial", size=18, bold=True, color="FFFFFF")
KPI_LABEL_FONT_WHITE = Font(name="Arial", size=9, bold=True, color="D6E4F0")
KPI_SUB_FONT = Font(name="Arial", size=8, italic=True, color=MUTED_TEXT)
TABLE_TITLE_FONT = Font(name="Arial", bold=True, size=11, color=NAVY)
TOTAL_FONT = Font(name="Arial", bold=True)

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


KPI_NUM_FORMAT = '"PKR "#,##0;("PKR "#,##0);"-"'


def _chart_title(text, color=NAVY, size=1400, bold=True):
    """Build a styled openpyxl chart Title in the dashboard's own navy
    Arial branding. The plain string API (chart.title = "text") always
    renders in Excel's generic default chart-title style (a small grey
    Calibri caption), so a title has to be built manually as rich text to
    actually match the rest of the sheet."""
    cp = CharacterProperties(sz=size, b=bold, solidFill=color, latin=DrawFont(typeface="Arial"))
    run = RegularTextRun(t=text, rPr=cp)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[run])
    return Title(tx=Text(rich=RichText(p=[para])))


def _chart_text(size=900, color=MUTED_TEXT, bold=False):
    """Reusable small Arial text-properties block for axis labels,
    legends, and data labels, so every element on a chart reads in the
    same clean type as the rest of the dashboard instead of Excel's
    generic default chart font."""
    cp = CharacterProperties(sz=size, b=bold, solidFill=color, latin=DrawFont(typeface="Arial"))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp))])


def _style_chart_frame(chart):
    """Give a chart the same card look as the KPI tiles: a white plot
    area on a thin light border, so it reads as a built-in piece of the
    dashboard rather than a bare chart floating loose on the sheet."""
    chart.graphical_properties = GraphicalProperties(
        solidFill="FFFFFF",
        ln=LineProperties(solidFill=CARD_BORDER, w=9525),
    )


def _apply_outer_border(ws, min_row, max_row, min_col, max_col, color=CARD_BORDER):
    """Draw a rectangular outer border around a (possibly merged) cell
    block. Openpyxl gives merged ranges no border of their own, so every
    boundary cell needs its own Border with only the outward-facing
    side(s) set."""
    thin = Side(style="thin", color=color)
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(
                left=thin if c == min_col else None,
                right=thin if c == max_col else None,
                top=thin if r == min_row else None,
                bottom=thin if r == max_row else None,
            )


def _write_kpi_card(ws, top_row, left_col, col_span, label, value_formula, fill, accent_hex,
                     value_font=KPI_VALUE_FONT, label_font=KPI_LABEL_FONT, sub_text=None):
    """A single dashboard KPI card: a thin accent strip, a small caps
    label, and a large value - occupying `col_span` columns starting at
    `left_col`, 5 rows tall starting at `top_row`. Returns the cell
    holding the value (so callers can reference it in later formulas)."""
    right_col = left_col + col_span - 1
    accent_fill = PatternFill("solid", start_color=accent_hex, end_color=accent_hex)

    ws.merge_cells(start_row=top_row, start_column=left_col, end_row=top_row, end_column=right_col)
    ws.cell(row=top_row, column=left_col).fill = accent_fill
    ws.row_dimensions[top_row].height = 5

    ws.merge_cells(start_row=top_row + 1, start_column=left_col, end_row=top_row + 1, end_column=right_col)
    lbl_cell = ws.cell(row=top_row + 1, column=left_col, value=label.upper())
    lbl_cell.font = label_font
    lbl_cell.fill = fill
    lbl_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws.merge_cells(start_row=top_row + 2, start_column=left_col, end_row=top_row + 3, end_column=right_col)
    val_cell = ws.cell(row=top_row + 2, column=left_col, value=value_formula)
    val_cell.font = value_font
    val_cell.fill = fill
    val_cell.number_format = KPI_NUM_FORMAT
    val_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    bottom_row = top_row + 4
    ws.merge_cells(start_row=bottom_row, start_column=left_col, end_row=bottom_row, end_column=right_col)
    sub_cell = ws.cell(row=bottom_row, column=left_col, value=sub_text or "")
    sub_cell.font = KPI_SUB_FONT
    sub_cell.fill = fill
    sub_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    for r in (top_row + 1, top_row + 2, top_row + 3, bottom_row):
        for c in range(left_col, right_col + 1):
            ws.cell(row=r, column=c).fill = fill
    _apply_outer_border(ws, top_row, bottom_row, left_col, right_col)

    return val_cell


def _write_summary_sheet(ws, bucket_label, client_breakdown_df, premium_total, claims_total, num_clients):
    """
    Third sheet, redesigned as a compact one-page dashboard:
      - a navy title banner naming the company/file and generation date
      - three KPI cards (Premium Receivable / Claims Receivable / Total
        Receivable) whose values are live SUM formulas over the
        per-client table on this same sheet, plus a Premium-vs-Claims
        pie chart
      - the full per-client breakdown table with zebra striping and a
        bold formula-driven Total row, plus a per-client bar chart
    All numbers stay auditable next to the charts that visualize them,
    and every total recalculates automatically if a value is edited.
    """
    ws.sheet_view.showGridLines = False

    n_clients = len(client_breakdown_df)

    # Row plan, top to bottom, with a fixed gap after the KPI/pie-chart
    # band so nothing placed below it can ever collide with the pie chart
    # regardless of how many client rows the table ends up with:
    #   rows 2-5   banner
    #   rows 8-12  KPI cards (+ pie chart floating beside them, columns K+)
    #   row 23     "Client Breakdown" title
    #   row 25     table header
    #   ...        table data + Total row
    #   +3 rows    bar chart, full width, always BELOW the table so it
    #              never shares a row range with the pie chart above
    card_row = 8
    table_header_row = 25
    first_data_row = table_header_row + 1
    last_data_row = table_header_row + n_clients if n_clients else table_header_row
    total_row = last_data_row + 1

    # ---- Title banner (3 stacked navy rows read as one continuous block) ----
    LAST_BANNER_COL = 12  # column L
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=LAST_BANNER_COL)
    ws.cell(row=2, column=2, value="PREMIUM & CLAIMS RECEIVABLE STATEMENT").font = KPI_LABEL_FONT_WHITE
    ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=LAST_BANNER_COL)
    ws.cell(row=3, column=2, value=bucket_label).font = Font(name="Arial", bold=True, size=18, color="FFFFFF")
    ws.merge_cells(start_row=4, start_column=2, end_row=4, end_column=LAST_BANNER_COL)
    subtitle = f"{num_clients} client{'s' if num_clients != 1 else ''}   \u2022   generated {date.today().strftime('%d-%b-%Y')}"
    ws.cell(row=4, column=2, value=subtitle).font = Font(name="Arial", size=10, color="D6E4F0")
    for r in (2, 3, 4):
        for c in range(2, LAST_BANNER_COL + 1):
            ws.cell(row=r, column=c).fill = BANNER_FILL
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 26
    ws.row_dimensions[4].height = 18
    # slim accent strip under the banner
    ws.merge_cells(start_row=5, start_column=2, end_row=5, end_column=LAST_BANNER_COL)
    ws.cell(row=5, column=2).fill = BANNER_ACCENT_FILL
    ws.row_dimensions[5].height = 4

    # ---- KPI cards (values reference the per-client table below via formula) ----
    premium_range = f"C{first_data_row}:C{last_data_row}" if n_clients else None
    claims_range = f"D{first_data_row}:D{last_data_row}" if n_clients else None

    premium_cell = _write_kpi_card(
        ws, card_row, 2, 2, "Premium Receivable",
        f"=SUM({premium_range})" if premium_range else 0,
        KPI_PREMIUM_FILL, ACCENT_BLUE, sub_text="Owed to IGI by the co-insurer",
    )
    claims_cell = _write_kpi_card(
        ws, card_row, 4, 2, "Claims Receivable",
        f"=SUM({claims_range})" if claims_range else 0,
        KPI_CLAIMS_FILL, CLAIMS_ORANGE, sub_text="Reimbursable to IGI for claims paid",
    )
    _write_kpi_card(
        ws, card_row, 7, 2, "Total Receivable",
        f"={premium_cell.coordinate}+{claims_cell.coordinate}",
        KPI_TOTAL_FILL, "5FA8D3", value_font=KPI_VALUE_FONT_WHITE, label_font=KPI_LABEL_FONT_WHITE,
        sub_text=f"Across {num_clients} client{'s' if num_clients != 1 else ''}",
    )

    # ---- small labeled data table feeding the pie chart (kept visible & auditable) ----
    pie_src_row = card_row + 1
    ws.cell(row=pie_src_row, column=10, value="Category").font = Font(name="Arial", size=8, bold=True, color=MUTED_TEXT)
    ws.cell(row=pie_src_row, column=11, value="Total").font = Font(name="Arial", size=8, bold=True, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 1, column=10, value="Premium Receivable").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 1, column=11, value=f"={premium_cell.coordinate}").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 1, column=11).number_format = NUM_FORMAT
    ws.cell(row=pie_src_row + 2, column=10, value="Claims Receivable").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 2, column=11, value=f"={claims_cell.coordinate}").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 2, column=11).number_format = NUM_FORMAT

    pie = PieChart()
    pie.title = _chart_title("Premium vs Claims Share", size=1300)
    pie_data = Reference(ws, min_col=11, min_row=pie_src_row, max_row=pie_src_row + 2)
    pie_cats = Reference(ws, min_col=10, min_row=pie_src_row + 1, max_row=pie_src_row + 2)
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)

    # Slice colors matched to the KPI cards above (blue = Premium, orange
    # = Claims) with a thin white separator, instead of Excel's default
    # generic theme palette - ties the pie directly back to the cards.
    premium_slice = DataPoint(idx=0)
    premium_slice.graphicalProperties = GraphicalProperties(solidFill=ACCENT_BLUE, ln=LineProperties(solidFill="FFFFFF", w=19050))
    claims_slice = DataPoint(idx=1)
    claims_slice.graphicalProperties = GraphicalProperties(solidFill=CLAIMS_ORANGE, ln=LineProperties(solidFill="FFFFFF", w=19050))
    pie.series[0].data_points = [premium_slice, claims_slice]

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showCatName = False
    pie.dataLabels.showVal = False
    pie.dataLabels.numFmt = "0.0%"
    pie.dataLabels.dLblPos = "outEnd"
    pie.dataLabels.txPr = _chart_text(size=1000, color=SLATE_TEXT, bold=True)

    pie.legend.position = "b"
    pie.legend.overlay = False
    pie.legend.txPr = _chart_text(size=900, color=SLATE_TEXT)

    _style_chart_frame(pie)
    pie.height = 8.6
    pie.width = 12.6
    ws.add_chart(pie, "M8")

    # ---- Client Breakdown section title ----
    ws.cell(row=table_header_row - 2, column=2, value="Client Breakdown").font = TABLE_TITLE_FONT

    # ---- per-client breakdown table + bar chart ----
    for j, h in enumerate(["Client Name", "Premium Receivable", "Claims Receivable"], start=2):
        c = ws.cell(row=table_header_row, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="left" if j == 2 else "right", vertical="center", indent=1)
    ws.row_dimensions[table_header_row].height = 20

    if n_clients == 0:
        ws.cell(row=first_data_row, column=2, value="No client data available.").font = EMPTY_FONT
    else:
        for i, row in enumerate(client_breakdown_df.itertuples(index=False), start=first_data_row):
            stripe = ROW_WHITE_FILL if (i - first_data_row) % 2 == 0 else ROW_ALT_FILL
            name_cell = ws.cell(row=i, column=2, value=str(row[0]))
            name_cell.font = BODY_FONT
            prem_cell = ws.cell(row=i, column=3, value=float(row[1]))
            prem_cell.number_format = NUM_FORMAT
            claim_cell = ws.cell(row=i, column=4, value=float(row[2]))
            claim_cell.number_format = NUM_FORMAT
            for cell in (name_cell, prem_cell, claim_cell):
                cell.border = BORDER
                cell.fill = stripe

        # Bold formula-driven Total row, so it always reflects the table above
        total_label = ws.cell(row=total_row, column=2, value="TOTAL")
        total_premium = ws.cell(row=total_row, column=3, value=f"=SUM(C{first_data_row}:C{last_data_row})")
        total_claims = ws.cell(row=total_row, column=4, value=f"=SUM(D{first_data_row}:D{last_data_row})")
        total_premium.number_format = NUM_FORMAT
        total_claims.number_format = NUM_FORMAT
        for cell in (total_label, total_premium, total_claims):
            cell.font = TOTAL_FONT
            cell.fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
            cell.border = Border(top=Side(style="thin", color=NAVY), left=THIN, right=THIN, bottom=THIN)

        # Horizontal bars, not vertical columns: client/company names are
        # long text labels that would otherwise get rotated or truncated
        # under a vertical column chart. Capped to the top clients (the
        # table above is already sorted largest Premium Receivable first)
        # so the chart stays legible instead of a wall of illegible bars
        # when a company file covers dozens of clients - the full list
        # still lives in the table itself, just above the chart.
        MAX_CHART_CLIENTS = 12
        chart_rows_shown = min(n_clients, MAX_CHART_CLIENTS)
        chart_last_row = first_data_row + chart_rows_shown - 1

        chart_title_text = "Premium vs Claims Receivable by Client"
        if n_clients > MAX_CHART_CLIENTS:
            chart_title_text += f"  (Top {MAX_CHART_CLIENTS} of {n_clients})"

        chart = BarChart()
        chart.type = "bar"
        chart.title = _chart_title(chart_title_text, size=1200)
        chart.y_axis.title = _chart_title("Amount (PKR)", color=MUTED_TEXT, size=900, bold=False)
        chart.gapWidth = 60
        chart.overlap = -10

        data = Reference(ws, min_col=3, max_col=4, min_row=table_header_row, max_row=chart_last_row)
        cats = Reference(ws, min_col=2, min_row=first_data_row, max_row=chart_last_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        # Series colors matched to the KPI cards / pie slices above, with
        # no bar border, instead of Excel's default generic theme colors.
        chart.series[0].graphicalProperties = GraphicalProperties(solidFill=ACCENT_BLUE, ln=LineProperties(noFill=True))
        chart.series[1].graphicalProperties = GraphicalProperties(solidFill=CLAIMS_ORANGE, ln=LineProperties(noFill=True))

        chart.legend.position = "t"
        chart.legend.overlay = False
        chart.legend.txPr = _chart_text(size=900, color=SLATE_TEXT)

        chart.x_axis.txPr = _chart_text(size=900, color=SLATE_TEXT)    # client name labels
        chart.y_axis.txPr = _chart_text(size=850, color=MUTED_TEXT)    # amount axis labels
        chart.y_axis.number_format = "#,##0"
        chart.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=CARD_BORDER, w=9525)))
        chart.x_axis.majorGridlines = None
        chart.x_axis.delete = False
        chart.y_axis.delete = False

        # Excel's own default for a horizontal bar chart plots the FIRST
        # category at the bottom - reversing this makes the largest
        # client (row 1, since the table is sorted descending) appear at
        # the top, matching the table above it instead of reading upside-down.
        chart.x_axis.scaling.orientation = "maxMin"

        _style_chart_frame(chart)
        chart.width = 22
        chart.height = max(8, min(15, 2.2 + 0.85 * chart_rows_shown))
        # Anchored two rows below the Total row (i.e. below the whole
        # table), not beside it - this guarantees it never shares a row
        # range with the pie chart above, regardless of how tall either
        # chart ends up rendering or how many client rows the table has.
        ws.add_chart(chart, f"B{total_row + 2}")

    # ---- Column widths ----
    # Cards sit at Premium=B-C, Claims=D-E, Total=G-H (F is the only true
    # gap column, kept narrow). Column D doubles as the client table's own
    # "Claims Receivable" column, so it needs real width either way - this
    # replaces the old layout where column E was hard-coded to width 3 as
    # a leftover spacer, which happened to be the exact column the Claims
    # Receivable card started in, collapsing that card and truncating the
    # table header next to it. Every column below gets enough room for
    # "PKR 19,850,294" at the KPI font size and for the "Premium
    # Receivable" / "Claims Receivable" table headers.
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 19
    ws.column_dimensions["D"].width = 19
    ws.column_dimensions["E"].width = 19
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 19
    ws.column_dimensions["H"].width = 19
    ws.column_dimensions["I"].width = 3
    ws.column_dimensions["J"].width = 20
    ws.column_dimensions["K"].width = 15


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
