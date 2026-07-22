"""
Writes one Excel workbook per coinsurer (all of that coinsurer's matched
claims combined into a single file, per spec - a claim shared by several
coinsurers appears in each of their own files, each showing only THEIR
own share), containing exactly two sheets: "Summary" (a dashboard - KPI
cards, a pie of Paid vs Outstanding share, and a per-insured chart) and
"Claims Data" (every matched claim row for that coinsurer, original claim
columns preserved plus the fan-out columns from claims_register_processor).

Also writes a single combined "Unmatched" workbook (one sheet, red-
highlighted like premium_register_generator's own unmatched-row
treatment) for claims that couldn't be matched to a policy, or whose
matched policy had no coinsurer on it at all - collected for manual
review rather than silently dropped.

Styling mirrors the rest of GIMS (receivable_generator.py's Summary
dashboard pattern, premium_register_generator.py's unmatched-row
highlighting) so output looks consistent with every other module.
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
BODY_FONT = Font(name="Arial")
EMPTY_FONT = Font(name="Arial", italic=True, color="808080")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM_FORMAT = "#,##0;(#,##0);-"
DATE_FORMAT = "DD-MMM-YYYY"

UNMATCHED_FILL = PatternFill("solid", start_color="F8D7DA", end_color="F8D7DA")  # light red
UNMATCHED_FONT = Font(name="Arial", color="842029")

# ---- Dashboard styling (same palette as receivable_generator.py) ----
NAVY = "1F4E78"
ACCENT_BLUE = "2E86AB"
CLAIMS_ORANGE = "D97B29"
SLATE_TEXT = "1C2733"
MUTED_TEXT = "6C7A89"
CARD_BORDER = "D9E2EC"
ROW_ALT_FILL = PatternFill("solid", start_color="F4F7FA", end_color="F4F7FA")
ROW_WHITE_FILL = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
BANNER_FILL = PatternFill("solid", start_color=NAVY, end_color=NAVY)
BANNER_ACCENT_FILL = PatternFill("solid", start_color=ACCENT_BLUE, end_color=ACCENT_BLUE)
KPI_PAID_FILL = PatternFill("solid", start_color="EAF2FA", end_color="EAF2FA")
KPI_OS_FILL = PatternFill("solid", start_color="FDF3EA", end_color="FDF3EA")
KPI_TOTAL_FILL = PatternFill("solid", start_color=NAVY, end_color=NAVY)
KPI_LABEL_FONT = Font(name="Arial", size=9, bold=True, color=MUTED_TEXT)
KPI_VALUE_FONT = Font(name="Arial", size=18, bold=True, color=SLATE_TEXT)
KPI_VALUE_FONT_WHITE = Font(name="Arial", size=18, bold=True, color="FFFFFF")
KPI_LABEL_FONT_WHITE = Font(name="Arial", size=9, bold=True, color="D6E4F0")
KPI_SUB_FONT = Font(name="Arial", size=8, italic=True, color=MUTED_TEXT)
KPI_NUM_FORMAT = '"PKR "#,##0;("PKR "#,##0);"-"'
TABLE_TITLE_FONT = Font(name="Arial", bold=True, size=11, color=NAVY)
TOTAL_FONT = Font(name="Arial", bold=True)

SUMMARY_SHEET = "Summary"
DATA_SHEET = "Claims Data"
UNMATCHED_SHEET = "Unmatched Claims"

NEW_COLUMNS = [
    "Coinsurer Name", "Coinsurer Share %", "Coinsurer Share of Claim Paid",
    "Coinsurer Share of Claim O/S", "Matched Policy Doc No.", "Match Score",
]


def _write_text_safe(cell):
    """Same junk-placeholder-text guard as every other generator in this
    project (report_generator.py, receivable_generator.py,
    premium_register_generator.py) - a source value like "=../-" gets
    auto-classified by openpyxl as a live formula, which then makes Excel
    flag the whole workbook as needing repair on open."""
    if cell.data_type == "f":
        cell.data_type = "s"
    return cell


def safe_filename_part(name):
    name = str(name).strip()
    name = re.sub(r"[^\w]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name[:100] if name else "UNKNOWN"


def build_output_filename(bucket_label, run_date=None):
    run_date = run_date or date.today()
    return f"{safe_filename_part(bucket_label)}_ClaimsRecovery_{run_date.strftime('%d%m%Y')}.xlsx"


def _prepare_for_excel(df, date_cols):
    out = {}
    for col in df.columns:
        s = df[col]
        if col in date_cols:
            out[col] = s.dt.date.astype(object).where(s.notna(), None) if pd.api.types.is_datetime64_any_dtype(s) else s
        else:
            s = s.astype(object)
            out[col] = s.where(pd.notna(s), None)
    return pd.DataFrame(out, columns=df.columns)


def _write_data_sheet(ws, df, highlight_new_cols=None, report=None):
    """Writes every column of df starting at A1. highlight_new_cols (if
    given) get a light-green tint, matching premium_register_generator's
    'matched, new column' treatment, so the fan-out columns this module
    added are visually distinguishable from the claim's own original data."""
    highlight_new_cols = set(highlight_new_cols or [])
    columns = list(df.columns)
    n_rows = len(df)

    for j, col in enumerate(columns, start=1):
        c = ws.cell(row=1, column=j, value=str(col))
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 20

    if n_rows == 0:
        ws.cell(row=2, column=1, value="No claims in this category.").font = EMPTY_FONT
        for j, col in enumerate(columns, start=1):
            ws.column_dimensions[get_column_letter(j)].width = max(14, min(28, len(str(col)) + 4))
        return

    apply_style = style_budget_ok(n_rows)
    date_cols = {c for c in columns if pd.api.types.is_datetime64_any_dtype(df[c])}
    clean_df = _prepare_for_excel(df, date_cols)
    reporter = ThrottledReporter(report, min_rows=25000) if report else None

    for i, row in enumerate(clean_df.itertuples(index=False, name=None), start=2):
        for j, (col, val) in enumerate(zip(columns, row), start=1):
            cell = ws.cell(row=i, column=j, value=val)
            _write_text_safe(cell)
            if val is None:
                continue
            if col in ("Coinsurer Share of Claim Paid", "Coinsurer Share of Claim O/S", "Claim Paid", "Claim O/S",
                       "Act. Loss", "Suminsured/Covered"):
                cell.number_format = NUM_FORMAT
            elif col == "Coinsurer Share %":
                cell.number_format = "0.00"
            elif col == "Match Score":
                cell.number_format = "0.000"
            elif col in date_cols:
                cell.number_format = DATE_FORMAT
            if apply_style:
                cell.border = BORDER
                cell.font = BODY_FONT
                if col in highlight_new_cols:
                    cell.fill = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
        if reporter:
            reporter.maybe(i - 1, total=n_rows, label="rows written")
    if reporter:
        reporter.final(n_rows, label="rows written")

    for j, col in enumerate(columns, start=1):
        width = 22 if col in highlight_new_cols else max(12, min(28, len(str(col)) + 4))
        ws.column_dimensions[get_column_letter(j)].width = width


def _write_unmatched_sheet(ws, df, report=None):
    """Every unmatched claim row, entirely highlighted red (matches
    premium_register_generator's own unmatched-row convention) since the
    whole sheet is 'needs manual review', plus the Reason column
    explaining why."""
    columns = list(df.columns)
    n_rows = len(df)

    for j, col in enumerate(columns, start=1):
        c = ws.cell(row=1, column=j, value=str(col))
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 20

    if n_rows == 0:
        ws.cell(row=2, column=1, value="No unmatched claims - everything matched to a coinsurer.").font = EMPTY_FONT
        return

    date_cols = {c for c in columns if pd.api.types.is_datetime64_any_dtype(df[c])}
    clean_df = _prepare_for_excel(df, date_cols)
    reporter = ThrottledReporter(report, min_rows=25000) if report else None

    for i, row in enumerate(clean_df.itertuples(index=False, name=None), start=2):
        for j, (col, val) in enumerate(zip(columns, row), start=1):
            cell = ws.cell(row=i, column=j, value=val)
            _write_text_safe(cell)
            if val is None:
                continue
            cell.border = BORDER
            cell.font = UNMATCHED_FONT
            cell.fill = UNMATCHED_FILL
            if col in date_cols:
                cell.number_format = DATE_FORMAT
        if reporter:
            reporter.maybe(i - 1, total=n_rows, label="rows written")
    if reporter:
        reporter.final(n_rows, label="rows written")

    for j, col in enumerate(columns, start=1):
        width = 26 if col == "Reason" else max(12, min(28, len(str(col)) + 4))
        ws.column_dimensions[get_column_letter(j)].width = width

    note_row = n_rows + 3
    ws.cell(row=note_row, column=1,
             value=f"{n_rows} claim(s) need manual review - either no matching policy was found in the Premium "
                   f"Register, or the matched policy has no coinsurer on it.").font = Font(name="Arial", size=9, italic=True, color="6C7A89")


def _chart_title(text, color=NAVY, size=1300, bold=True):
    cp = CharacterProperties(sz=size, b=bold, solidFill=color, latin=DrawFont(typeface="Arial"))
    run = RegularTextRun(t=text, rPr=cp)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[run])
    return Title(tx=Text(rich=RichText(p=[para])))


def _chart_text(size=900, color=MUTED_TEXT, bold=False):
    cp = CharacterProperties(sz=size, b=bold, solidFill=color, latin=DrawFont(typeface="Arial"))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp))])


def _style_chart_frame(chart):
    chart.graphical_properties = GraphicalProperties(
        solidFill="FFFFFF", ln=LineProperties(solidFill=CARD_BORDER, w=9525))


def _blank_styled_cell(ws, row, column):
    """See receivable_generator.py's identical helper for the full
    rationale - only ever call this on a merge's own top-left anchor
    cell."""
    cell = ws.cell(row=row, column=column)
    if cell.value is None:
        cell.value = " "
    return cell


def _apply_outer_border(ws, anchor_rows, top_row, bottom_row, col, color=CARD_BORDER):
    thin = Side(style="thin", color=color)
    for r in anchor_rows:
        cell = _blank_styled_cell(ws, r, col)
        existing = cell.border
        cell.border = Border(
            left=thin, right=thin,
            top=thin if r == top_row else existing.top,
            bottom=thin if r == bottom_row else existing.bottom,
        )


def _write_kpi_card(ws, top_row, left_col, col_span, label, value_formula, fill, accent_hex,
                     value_font=KPI_VALUE_FONT, label_font=KPI_LABEL_FONT, sub_text=None):
    right_col = left_col + col_span - 1
    accent_fill = PatternFill("solid", start_color=accent_hex, end_color=accent_hex)

    ws.merge_cells(start_row=top_row, start_column=left_col, end_row=top_row, end_column=right_col)
    _blank_styled_cell(ws, top_row, left_col).fill = accent_fill
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
    sub_cell = ws.cell(row=bottom_row, column=left_col, value=sub_text or " ")
    sub_cell.font = KPI_SUB_FONT
    sub_cell.fill = fill
    sub_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    _apply_outer_border(ws, [top_row, top_row + 1, top_row + 2, bottom_row], top_row, bottom_row, left_col)
    return val_cell


def _write_summary_sheet(ws, bucket_label, breakdown_df, total_paid, total_os, num_claims):
    """Compact one-page dashboard: navy title banner, three KPI cards
    (Claims Count / Share of Claim Paid / Share of Claim O/S) whose
    values are live SUM formulas over the per-insured table on this same
    sheet, a Paid-vs-Outstanding pie, and a per-insured bar chart -
    mirrors receivable_generator.py's own Summary sheet layout."""
    ws.sheet_view.showGridLines = False
    n_insureds = len(breakdown_df)

    card_row = 8
    table_header_row = 25
    first_data_row = table_header_row + 1
    last_data_row = table_header_row + n_insureds if n_insureds else table_header_row
    total_row = last_data_row + 1

    LAST_BANNER_COL = 12
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=LAST_BANNER_COL)
    ws.cell(row=2, column=2, value="CLAIMS RECOVERY STATEMENT").font = KPI_LABEL_FONT_WHITE
    ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=LAST_BANNER_COL)
    ws.cell(row=3, column=2, value=bucket_label).font = Font(name="Arial", bold=True, size=18, color="FFFFFF")
    ws.merge_cells(start_row=4, start_column=2, end_row=4, end_column=LAST_BANNER_COL)
    subtitle = f"{num_claims} claim{'s' if num_claims != 1 else ''}   \u2022   generated {date.today().strftime('%d-%b-%Y')}"
    ws.cell(row=4, column=2, value=subtitle).font = Font(name="Arial", size=10, color="D6E4F0")
    for r in (2, 3, 4):
        cell = ws.cell(row=r, column=2)
        cell.fill = BANNER_FILL
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 26
    ws.row_dimensions[4].height = 18
    ws.merge_cells(start_row=5, start_column=2, end_row=5, end_column=LAST_BANNER_COL)
    _blank_styled_cell(ws, 5, 2).fill = BANNER_ACCENT_FILL
    ws.row_dimensions[5].height = 4

    paid_range = f"C{first_data_row}:C{last_data_row}" if n_insureds else None
    os_range = f"D{first_data_row}:D{last_data_row}" if n_insureds else None

    paid_cell = _write_kpi_card(
        ws, card_row, 2, 2, "Share of Claim Paid",
        f"=SUM({paid_range})" if paid_range else 0,
        KPI_PAID_FILL, ACCENT_BLUE, sub_text="This coinsurer's cut of claims already paid",
    )
    os_cell = _write_kpi_card(
        ws, card_row, 4, 2, "Share of Claim O/S",
        f"=SUM({os_range})" if os_range else 0,
        KPI_OS_FILL, CLAIMS_ORANGE, sub_text="This coinsurer's cut of claims still outstanding",
    )
    _write_kpi_card(
        ws, card_row, 7, 2, "Total Recoverable",
        f"={paid_cell.coordinate}+{os_cell.coordinate}",
        KPI_TOTAL_FILL, "5FA8D3", value_font=KPI_VALUE_FONT_WHITE, label_font=KPI_LABEL_FONT_WHITE,
        sub_text=f"Across {num_claims} claim{'s' if num_claims != 1 else ''}",
    )

    pie_src_row = card_row + 1
    ws.cell(row=pie_src_row, column=10, value="Category").font = Font(name="Arial", size=8, bold=True, color=MUTED_TEXT)
    ws.cell(row=pie_src_row, column=11, value="Total").font = Font(name="Arial", size=8, bold=True, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 1, column=10, value="Claim Paid Share").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 1, column=11, value=f"={paid_cell.coordinate}").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 1, column=11).number_format = NUM_FORMAT
    ws.cell(row=pie_src_row + 2, column=10, value="Claim O/S Share").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 2, column=11, value=f"={os_cell.coordinate}").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 2, column=11).number_format = NUM_FORMAT

    pie = PieChart()
    pie.title = _chart_title("Paid vs Outstanding Share", size=1300)
    pie_data = Reference(ws, min_col=11, min_row=pie_src_row, max_row=pie_src_row + 2)
    pie_cats = Reference(ws, min_col=10, min_row=pie_src_row + 1, max_row=pie_src_row + 2)
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)

    paid_slice = DataPoint(idx=0)
    paid_slice.graphicalProperties = GraphicalProperties(solidFill=ACCENT_BLUE, ln=LineProperties(solidFill="FFFFFF", w=19050))
    os_slice = DataPoint(idx=1)
    os_slice.graphicalProperties = GraphicalProperties(solidFill=CLAIMS_ORANGE, ln=LineProperties(solidFill="FFFFFF", w=19050))
    pie.series[0].data_points = [paid_slice, os_slice]

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
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

    ws.cell(row=table_header_row - 2, column=2, value="Claims Recovery by Insured").font = TABLE_TITLE_FONT

    for j, h in enumerate(["Insured Name", "Claim Paid Share", "Claim O/S Share", "Claims"], start=2):
        c = ws.cell(row=table_header_row, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="left" if j == 2 else "right", vertical="center", indent=1)
    ws.row_dimensions[table_header_row].height = 20

    if n_insureds == 0:
        ws.cell(row=first_data_row, column=2, value="No claim data available.").font = EMPTY_FONT
    else:
        for i, row in enumerate(breakdown_df.itertuples(index=False), start=first_data_row):
            stripe = ROW_WHITE_FILL if (i - first_data_row) % 2 == 0 else ROW_ALT_FILL
            name_cell = ws.cell(row=i, column=2, value=str(row[0]))
            name_cell.font = BODY_FONT
            name_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            paid_c = ws.cell(row=i, column=3, value=float(row[1]))
            paid_c.number_format = NUM_FORMAT
            paid_c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            os_c = ws.cell(row=i, column=4, value=float(row[2]))
            os_c.number_format = NUM_FORMAT
            os_c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            n_c = ws.cell(row=i, column=5, value=int(row[3]))
            n_c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            for cell in (name_cell, paid_c, os_c, n_c):
                cell.border = BORDER
                cell.fill = stripe
            ws.row_dimensions[i].height = 26

        total_label = ws.cell(row=total_row, column=2, value="TOTAL")
        total_paid_c = ws.cell(row=total_row, column=3, value=f"=SUM(C{first_data_row}:C{last_data_row})")
        total_os_c = ws.cell(row=total_row, column=4, value=f"=SUM(D{first_data_row}:D{last_data_row})")
        total_paid_c.number_format = NUM_FORMAT
        total_os_c.number_format = NUM_FORMAT
        for cell in (total_label, total_paid_c, total_os_c):
            cell.font = TOTAL_FONT
            cell.fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
            cell.border = Border(top=Side(style="thin", color=NAVY), left=THIN, right=THIN, bottom=THIN)

        MAX_CHART_ROWS = 12
        chart_rows_shown = min(n_insureds, MAX_CHART_ROWS)
        chart_last_row = first_data_row + chart_rows_shown - 1
        chart_title_text = "Claim Paid vs O/S Share by Insured"
        if n_insureds > MAX_CHART_ROWS:
            chart_title_text += f"  (Top {MAX_CHART_ROWS} of {n_insureds})"

        chart = BarChart()
        chart.type = "col"
        chart.title = _chart_title(chart_title_text, size=1200)
        chart.gapWidth = 40
        chart.overlap = -10

        data = Reference(ws, min_col=3, max_col=4, min_row=table_header_row, max_row=chart_last_row)
        cats = Reference(ws, min_col=2, min_row=first_data_row, max_row=chart_last_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.series[0].graphicalProperties = GraphicalProperties(solidFill=ACCENT_BLUE, ln=LineProperties(noFill=True))
        chart.series[1].graphicalProperties = GraphicalProperties(solidFill=CLAIMS_ORANGE, ln=LineProperties(noFill=True))

        chart.legend.position = "b"
        chart.legend.overlay = False
        chart.legend.txPr = _chart_text(size=900, color=SLATE_TEXT)

        name_label_props = _chart_text(size=850, color=SLATE_TEXT)
        name_label_props.bodyPr.rot = -2700000
        name_label_props.bodyPr.vert = "horz"
        chart.x_axis.txPr = name_label_props
        chart.y_axis.txPr = _chart_text(size=850, color=MUTED_TEXT)
        chart.y_axis.number_format = "#,##0"
        chart.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=CARD_BORDER, w=9525)))
        chart.x_axis.majorGridlines = None
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.x_axis.lblAlgn = "r"
        chart.x_axis.lblOffset = 100
        chart.x_axis.axPos = "b"
        chart.y_axis.axPos = "l"
        _style_chart_frame(chart)
        chart.width = max(20, min(34, 1.7 * chart_rows_shown + 6))
        chart.height = 13
        ws.add_chart(chart, f"B{total_row + 2}")

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 19
    ws.column_dimensions["H"].width = 19
    ws.column_dimensions["I"].width = 3
    ws.column_dimensions["J"].width = 20
    ws.column_dimensions["K"].width = 15


def build_coinsurer_workbook(bucket_label, claims_df, breakdown_df, output_dir, run_date=None, report=None):
    """
    Write a single coinsurer's workbook:
      1. "Summary" - KPI cards, Paid-vs-O/S pie, per-insured bar chart
      2. "Claims Data" - every matched claim row for this coinsurer

    breakdown_df: per-insured Insured Name / share_paid_total /
        share_os_total / num_claims (from
        claims_register_processor.build_coinsurer_buckets).
    Returns the full filepath written.
    """
    os.makedirs(output_dir, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)

    ws_data = wb.create_sheet(DATA_SHEET)
    _write_data_sheet(ws_data, claims_df, highlight_new_cols=NEW_COLUMNS, report=report)

    ws_summary = wb.create_sheet(SUMMARY_SHEET)
    total_paid = pd.to_numeric(claims_df["Coinsurer Share of Claim Paid"], errors="coerce").fillna(0).sum() if "Coinsurer Share of Claim Paid" in claims_df.columns else 0.0
    total_os = pd.to_numeric(claims_df["Coinsurer Share of Claim O/S"], errors="coerce").fillna(0).sum() if "Coinsurer Share of Claim O/S" in claims_df.columns else 0.0
    num_claims = len(claims_df)
    _write_summary_sheet(ws_summary, bucket_label, breakdown_df, total_paid, total_os, num_claims)

    wb.move_sheet(SUMMARY_SHEET, offset=-1)  # Summary first

    filename = build_output_filename(bucket_label, run_date=run_date)
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)
    return filepath


def build_unmatched_workbook(unmatched_df, output_dir, run_date=None, report=None):
    """Single-sheet workbook of every claim needing manual review - not
    per-company, since there's no coinsurer identified to route it to
    yet. Returns the filepath written, or None if unmatched_df is empty
    (nothing to write)."""
    if len(unmatched_df) == 0:
        return None
    os.makedirs(output_dir, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = UNMATCHED_SHEET
    _write_unmatched_sheet(ws, unmatched_df, report=report)
    run_date = run_date or date.today()
    filename = f"Unmatched_Claims_{run_date.strftime('%d%m%Y')}.xlsx"
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)
    return filepath


if __name__ == "__main__":
    premium_df = pd.DataFrame({
        "Claim/Benefit Ref. #": ["CL-001", "CL-002"],
        "Policy/PMD #": ["2025/06/HOFFNIP00746", "2025/06/HOFFNIP00747"],
        "Insured/Participant": ["ACME CORP", "BETA MILLS"],
        "Claim Paid": [500000, 0],
        "Claim O/S": [0, 250000],
        "Coinsurer Name": ["EFU GENERAL INSURANCE LIMITED", "EFU GENERAL INSURANCE LIMITED"],
        "Coinsurer Share %": [50, 50],
        "Coinsurer Share of Claim Paid": [250000, 0],
        "Coinsurer Share of Claim O/S": [0, 125000],
        "Matched Policy Doc No.": ["2026/01/HOFFNIE00050", "2026/01/HOFFNIE00051"],
        "Match Score": [0.9, 0.85],
    })
    breakdown = pd.DataFrame({
        "Insured Name": ["ACME CORP", "BETA MILLS"],
        "share_paid_total": [250000, 0],
        "share_os_total": [0, 125000],
        "num_claims": [1, 1],
    })
    fp = build_coinsurer_workbook("EFU GENERAL INSURANCE LIMITED", premium_df, breakdown, "/tmp/claims_register_test_output")
    print("Wrote:", fp)
