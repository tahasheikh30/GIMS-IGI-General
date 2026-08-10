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

import company_alias_resolver as car

try:
    from perf_utils import style_budget_ok, ThrottledReporter, StyleIndexCache
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
    # A fixed, requested naming convention applies to a known set of
    # companies (see company_alias_resolver.CANONICAL_COMPANY_FILENAMES) -
    # e.g. always "UBL_Insurance_Company_LTD" / "United_Insurance_Company_LTD"
    # regardless of exactly how that company's raw/merged bucket label
    # happens to be worded. Any company not on that fixed list keeps the
    # original sanitized-label naming, unchanged.
    canonical = car.canonical_filename_for_label(bucket_label)
    stub = canonical if canonical else safe_filename_part(bucket_label)
    return f"{stub}_ClaimsRecovery_{run_date.strftime('%d%m%Y')}.xlsx"


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

    date_cols = {c for c in columns if pd.api.types.is_datetime64_any_dtype(df[c])}
    clean_df = _prepare_for_excel(df, date_cols)
    reporter = ThrottledReporter(report, min_rows=25000) if report else None
    # StyleIndexCache (see its docstring in perf_utils.py) makes per-cell
    # font/border/fill/number_format application cheap regardless of row
    # count, by registering each distinct style combination once and
    # reusing its internal openpyxl style index directly for every later
    # cell that needs the identical combination - this replaces the old
    # style_budget_ok() row-count gate that used to skip font/border
    # entirely above 50,000 rows. Confirmed via cProfile against a real
    # coinsurer bucket that this per-cell style hashing (not the pandas
    # iteration, not the save() itself) was the dominant cost of
    # generating an output file.
    style_cache = StyleIndexCache()
    HIGHLIGHT_FILL = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")

    for i, row in enumerate(clean_df.itertuples(index=False, name=None), start=2):
        for j, (col, val) in enumerate(zip(columns, row), start=1):
            cell = ws.cell(row=i, column=j, value=val)
            _write_text_safe(cell)
            if val is None:
                continue
            if col in ("Coinsurer Share of Claim Paid", "Coinsurer Share of Claim O/S", "Claim Paid", "Claim O/S",
                       "Act. Loss", "Suminsured/Covered"):
                number_format = NUM_FORMAT
            elif col == "Coinsurer Share %":
                number_format = "0.00"
            elif col == "Match Score":
                number_format = "0.000"
            elif col in date_cols:
                number_format = DATE_FORMAT
            else:
                number_format = None
            fill = HIGHLIGHT_FILL if col in highlight_new_cols else None
            style_cache.apply(cell, font=BODY_FONT, border=BORDER, fill=fill, number_format=number_format)
        if reporter:
            reporter.maybe(i - 1, total=n_rows, label="rows written")
    if reporter:
        reporter.final(n_rows, label="rows written")

    for j, col in enumerate(columns, start=1):
        width = 22 if col in highlight_new_cols else max(12, min(28, len(str(col)) + 4))
        ws.column_dimensions[get_column_letter(j)].width = width


def _write_unmatched_sheet(ws, df, report=None, note=None, empty_message=None):
    """Every unmatched claim row, entirely highlighted red (matches
    premium_register_generator's own unmatched-row convention) since the
    whole sheet is 'needs manual review', plus the Reason column
    explaining why.

    note / empty_message: optional overrides for the summary footnote and
    the "nothing here" placeholder, so this same writer can be reused for
    a differently-flavored manual-review sheet without duplicating the
    styling logic - default text (unmatched-claims wording) is used if
    omitted."""
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
        ws.cell(row=2, column=1,
                 value=empty_message or "No unmatched claims - everything matched to a coinsurer.").font = EMPTY_FONT
        return

    date_cols = {c for c in columns if pd.api.types.is_datetime64_any_dtype(df[c])}
    clean_df = _prepare_for_excel(df, date_cols)
    reporter = ThrottledReporter(report, min_rows=25000) if report else None
    style_cache = StyleIndexCache()

    for i, row in enumerate(clean_df.itertuples(index=False, name=None), start=2):
        for j, (col, val) in enumerate(zip(columns, row), start=1):
            cell = ws.cell(row=i, column=j, value=val)
            _write_text_safe(cell)
            if val is None:
                continue
            number_format = DATE_FORMAT if col in date_cols else None
            style_cache.apply(cell, font=UNMATCHED_FONT, border=BORDER, fill=UNMATCHED_FILL, number_format=number_format)
        if reporter:
            reporter.maybe(i - 1, total=n_rows, label="rows written")
    if reporter:
        reporter.final(n_rows, label="rows written")

    for j, col in enumerate(columns, start=1):
        width = 26 if col == "Reason" else max(12, min(28, len(str(col)) + 4))
        ws.column_dimensions[get_column_letter(j)].width = width

    note_row = n_rows + 3
    default_note = (f"{n_rows} claim(s) need manual review - either no matching policy was found in the Premium "
                    f"Register, or the matched policy has no coinsurer on it.")
    ws.cell(row=note_row, column=1, value=note or default_note).font = Font(name="Arial", size=9, italic=True, color="6C7A89")


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

    # Row plan, top to bottom. The per-insured bar chart now sits ABOVE
    # the table (previously it was below), so it reads as part of the
    # dashboard overview at a glance rather than being buried under a
    # potentially long table. bar_chart_row is placed well clear of the
    # pie chart's own row range (the pie sits at "M8" and, at its own
    # fixed height, never runs past roughly row 24) so the two charts
    # never overlap regardless of how many insured rows follow, and
    # table_header_row leaves enough vertical room below bar_chart_row for
    # the bar chart's own fixed height before the table starts:
    #   rows 2-5   banner
    #   rows 8-12  KPI cards (+ pie chart floating beside them, columns K+)
    #   row 26     per-insured bar chart, full width
    #   row 53     "Claims Recovery by Insured" title
    #   row 55     table header
    #   ...        table data + Total row
    card_row = 8
    bar_chart_row = 26
    table_header_row = 55
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

    # ---- per-insured bar chart, placed ABOVE the table (see the row-plan
    # note at the top of this function) - built here, before the table
    # itself is written, using row numbers (table_header_row/first_data_row)
    # that are already fixed regardless of write order; openpyxl only
    # needs valid cell coordinates for a chart's data/category references,
    # not for those cells to already contain data at the time the chart
    # object is created. ----
    if n_insureds > 0:
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
        # Anchored at bar_chart_row, well clear of the pie chart above it
        # and well above the table below it (see the row-plan note at the
        # top of this function).
        ws.add_chart(chart, f"B{bar_chart_row}")

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
        style_cache = StyleIndexCache()
        ALIGN_NAME = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", indent=1)
        for i, row in enumerate(breakdown_df.itertuples(index=False), start=first_data_row):
            stripe = ROW_WHITE_FILL if (i - first_data_row) % 2 == 0 else ROW_ALT_FILL
            name_cell = ws.cell(row=i, column=2, value=str(row[0]))
            style_cache.apply(name_cell, font=BODY_FONT, border=BORDER, fill=stripe, alignment=ALIGN_NAME)
            paid_c = ws.cell(row=i, column=3, value=float(row[1]))
            style_cache.apply(paid_c, border=BORDER, fill=stripe, alignment=ALIGN_RIGHT, number_format=NUM_FORMAT)
            os_c = ws.cell(row=i, column=4, value=float(row[2]))
            style_cache.apply(os_c, border=BORDER, fill=stripe, alignment=ALIGN_RIGHT, number_format=NUM_FORMAT)
            n_c = ws.cell(row=i, column=5, value=int(row[3]))
            style_cache.apply(n_c, border=BORDER, fill=stripe, alignment=ALIGN_RIGHT)
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


def build_coinsurer_workbooks_parallel(groups, breakdowns, output_dir, run_date=None, max_workers=None, report=None,
                                        unmatched_df=None):
    """
    Generate every coinsurer's workbook using a process pool instead of
    one at a time - see receivable_generator.build_company_workbooks_parallel
    for the full rationale (same technique, same fallback conditions).

    This module's own bucket sizes are typically far more skewed than the
    Receivable module's: confirmed against a real 27,976-claim run, one
    coinsurer bucket alone (58,423 matched claim-coinsurer rows, out of
    274,000 total across 98 coinsurers) took roughly a minute to write on
    its own, purely from openpyxl's own per-cell object construction cost
    at that scale (millions of individual cells) - a cost that's now
    genuinely at the floor of what openpyxl's normal (non-streaming) mode
    can do, since this module's Summary sheet needs real merged cells and
    embedded charts, which openpyxl's much faster "write_only" streaming
    mode does not support mixed into the same workbook as those features
    (confirmed directly: attempting to mix a write-only worksheet into a
    normal, chart-containing workbook silently drops every row written to
    it - not attempted here for exactly that reason). Running every
    bucket concurrently instead of one after another means the WHOLE
    BATCH's wall-clock time is bound by whichever single bucket is
    largest, not the sum of all of them - on a real multi-core machine,
    generating 98 files where 91 are small and 7 are huge should take
    roughly as long as the biggest handful of those, not 98 files' worth
    of sequential time.

    unmatched_df: optional - if given (and non-empty), the "Unmatched
        Claims" workbook is submitted to the SAME process pool as one
        more independent task, so it runs concurrently with every
        coinsurer's workbook instead of being generated afterward as a
        separate, strictly-sequential step (which would otherwise leave
        every other CPU core idle while it alone was being written).

    Returns (generated_filepaths, filename_lookup, errors) - same shape
    the app.py "Generate Excel Reports" worker already builds around a
    plain loop, so this is a drop-in replacement for that loop. The
    unmatched-claims filepath (if any) is included in generated_filepaths
    but not filename_lookup (it isn't looked up by bucket_label the way
    coinsurer files are).
    """
    import os as _os
    bucket_labels = list(groups.keys())
    include_unmatched = unmatched_df is not None and len(unmatched_df) > 0
    total_tasks = len(bucket_labels) + (1 if include_unmatched else 0)
    n_cpus = _os.cpu_count() or 1
    MIN_BUCKETS_FOR_PARALLEL = 4  # even 2-3 buckets can be worth it here, given how skewed bucket sizes are
    use_parallel = n_cpus >= 2 and total_tasks >= MIN_BUCKETS_FOR_PARALLEL

    generated, filename_lookup, errors = [], {}, []

    if use_parallel:
        try:
            import concurrent.futures
            workers = max(1, min(max_workers or n_cpus, n_cpus, total_tasks))
            if report:
                report(f"Generating {total_tasks} workbook(s) across {workers} worker process(es)...")
            tasks = [
                ("coinsurer", bucket_label, groups[bucket_label], breakdowns.get(bucket_label), output_dir, run_date)
                for bucket_label in bucket_labels
            ]
            if include_unmatched:
                tasks.append(("unmatched", None, unmatched_df, None, output_dir, run_date))
            done_count = 0
            with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
                for label, fp, err in executor.map(_build_one_generic_workbook_worker, tasks):
                    done_count += 1
                    if err:
                        errors.append(f"{label}: {err}")
                    elif fp:
                        generated.append(fp)
                        if label != "Unmatched Claims":
                            filename_lookup[label] = os.path.basename(fp)
                    if report:
                        report(f"[{done_count}/{total_tasks}] {label}")
                        if hasattr(report, "progress"):
                            report.progress(done_count, total_tasks,
                                             text=f"{done_count} / {total_tasks} workbook(s) generated")
            return generated, filename_lookup, errors
        except Exception as e:
            if report:
                report(f"Parallel generation unavailable ({e}) - falling back to sequential generation.")
            generated, filename_lookup, errors = [], {}, []

    for i, bucket_label in enumerate(bucket_labels, start=1):
        claims_sub_df = groups[bucket_label]
        breakdown = breakdowns.get(bucket_label)
        if report:
            report(f"[{i}/{total_tasks}] {bucket_label}")
        try:
            fp = build_coinsurer_workbook(bucket_label, claims_sub_df, breakdown, output_dir, run_date=run_date)
            generated.append(fp)
            filename_lookup[bucket_label] = os.path.basename(fp)
        except Exception as e:
            errors.append(f"{bucket_label}: {e}")
    if include_unmatched:
        if report:
            report(f"[{total_tasks}/{total_tasks}] Unmatched Claims")
        try:
            fp = build_unmatched_workbook(unmatched_df, output_dir, run_date=run_date)
            if fp:
                generated.append(fp)
        except Exception as e:
            errors.append(f"Unmatched Claims: {e}")
    return generated, filename_lookup, errors


def _build_one_generic_workbook_worker(args):
    """Module-level (picklable) worker used by build_coinsurer_workbooks_parallel
    for BOTH task kinds (a coinsurer's own workbook, or the shared
    Unmatched Claims workbook) so both can be submitted to the same
    process pool and run concurrently rather than the unmatched workbook
    being a separate, strictly-sequential step afterward."""
    kind, bucket_label, df, breakdown_df, output_dir, run_date = args
    try:
        if kind == "coinsurer":
            fp = build_coinsurer_workbook(bucket_label, df, breakdown_df, output_dir, run_date=run_date)
            return bucket_label, fp, None
        else:
            fp = build_unmatched_workbook(df, output_dir, run_date=run_date)
            return "Unmatched Claims", fp, None
    except Exception as e:
        return (bucket_label or "Unmatched Claims"), None, str(e)


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
