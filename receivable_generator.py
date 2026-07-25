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
TITLE_FONT = Font(name="Arial", bold=True, size=14, color="1F4E78")
TOTAL_LABEL_FONT = Font(name="Arial", size=10, color="5B6B79")
TOTAL_VALUE_FONT = Font(name="Arial", bold=True, size=12)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM_FORMAT = "#,##0;(#,##0);-"
DATE_FORMAT = "DD-MMM-YYYY"


def _write_text_safe(cell):
    """
    Guard against a real, confirmed data-quality issue in the source
    export: a handful of rows have junk placeholder text in columns like
    "Leader Ref. No" (e.g. "=../-", "====-=-") instead of a real
    reference number. openpyxl auto-detects ANY string value starting
    with "=" as a live Excel formula - and since that junk text isn't
    valid formula syntax, Excel then flags the whole workbook as
    containing "unreadable content" and silently strips it on open
    (the repair prompt). None of these pass-through data cells are ever
    meant to hold a real formula, so force the cell back to plain text
    whenever openpyxl auto-classified it as one - this preserves the
    original text exactly (no stray leading apostrophe) with zero
    functional loss.
    """
    if cell.data_type == "f":
        cell.data_type = "s"
    return cell

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


def _prepare_for_excel(df, date_cols):
    """
    Vectorized, one-pass cleanup of the whole dataframe before writing
    (see the matching helper in report_generator.py for the full
    rationale) - replaces a per-cell isna()/hasattr() chain that used to
    run once per cell with a handful of whole-column pandas operations,
    which stays fast even at hundreds of thousands of rows.
    """
    out = {}
    for col in df.columns:
        s = df[col]
        if col in date_cols:
            out[col] = s.dt.date.astype(object).where(s.notna(), None)
        else:
            s = s.astype(object)
            out[col] = s.where(pd.notna(s), None)
    return pd.DataFrame(out, columns=df.columns)


def _write_sheet(ws, df, amount_col, drop_cols=None, report=None):
    """
    amount_col: name of this sheet's own amount column (e.g. "Premium
        Receivable" or "Claim Receivable") - gets the accounting number
        format; every other column is written as-is.
    drop_cols: columns to omit from this sheet (typically the OTHER
        category's amount column, so the Premium Receivable sheet doesn't
        also show a redundant all-zero Claim Receivable column).
    report: optional callable(str) for progress on large files (see
        perf_utils.run_in_background) - safe to leave as None.
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

    n_rows = len(df)
    date_cols = {col for col in columns if col in df.columns and pd.api.types.is_datetime64_any_dtype(df[col])}
    clean_df = _prepare_for_excel(df[columns], date_cols)
    reporter = ThrottledReporter(report, min_rows=25000) if report else None
    # StyleIndexCache makes per-cell font/border/number_format application
    # cheap (see its own docstring in perf_utils.py) - the entire sheet
    # only ever needs 3 distinct style combinations (amount columns, date
    # columns, everything else), each registered once and reused for
    # every remaining cell, instead of openpyxl re-hashing the same
    # Font+Border+number_format combination on every single cell. This
    # replaces the old style_budget_ok() row-count gate that used to skip
    # font/border entirely above 50,000 rows - with the cache, applying
    # full styling costs almost nothing extra even on a very large sheet,
    # so there's no longer a reason to skip it.
    style_cache = StyleIndexCache()

    for i, row in enumerate(clean_df.itertuples(index=False, name=None), start=2):
        for j, (col, val) in enumerate(zip(columns, row), start=1):
            cell = ws.cell(row=i, column=j, value=val)
            _write_text_safe(cell)
            # IMPORTANT: only style a cell that actually has a value. In
            # openpyxl, calling ws.cell(value=None) alone writes nothing
            # to the sheet - the cell simply doesn't exist in the saved
            # file, which is correct and harmless. But if ANY other
            # attribute (font, border, fill, ...) is then set on that
            # same still-valueless cell, openpyxl is forced to
            # materialize it - and it gets written out as an empty
            # numeric cell (<c t="n"> with no <v>), which is invalid
            # per the OOXML schema. Real production data is full of
            # genuinely blank cells (Base No., Base date, etc.), so this
            # matters even though small hand-built test data never
            # triggers it. Skipping styling on blank cells keeps every
            # written cell valid; the only visible cost is that a blank
            # cell in the middle of a data row won't carry the thin grid
            # border its neighbors have.
            if val is None:
                continue
            if col == amount_col:
                number_format = NUM_FORMAT
            elif col in date_cols:
                number_format = DATE_FORMAT
            else:
                number_format = None
            style_cache.apply(cell, font=BODY_FONT, border=BORDER, number_format=number_format)
        if reporter:
            reporter.maybe(i - 1, total=n_rows, label="rows written")

    if reporter:
        reporter.final(n_rows, label="rows written")

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


def _apply_outer_border(ws, anchor_rows, top_row, bottom_row, col, color=CARD_BORDER):
    """Draw an outer border around a block made of several stacked
    single-column merges (e.g. a KPI card: an accent-strip row, a label
    row, a 2-row value merge, and a subtitle row, each merged across the
    same `col`:right_col span). `anchor_rows` lists the row of each
    merge's own top-left anchor cell (e.g. the value merge spanning
    top_row+2:top_row+3 has only ONE anchor, at top_row+2 - top_row+3
    itself has no cell of its own to style).

    Deliberately only ever touches the anchor cell of each merge - see
    _blank_styled_cell's docstring for why touching any other cell in a
    merged range is unsafe."""
    thin = Side(style="thin", color=color)
    for r in anchor_rows:
        cell = _blank_styled_cell(ws, r, col)
        existing = cell.border
        cell.border = Border(
            left=thin, right=thin,
            top=thin if r == top_row else existing.top,
            bottom=thin if r == bottom_row else existing.bottom,
        )


def _blank_styled_cell(ws, row, column):
    """
    Return the cell at (row, column), giving it a single invisible space
    as its value if it doesn't already have a real one.

    IMPORTANT: only ever call this on a merge's own top-left ANCHOR cell
    (or a cell that isn't part of any merge at all) - never on one of the
    other cells absorbed into a merged range. Once a range is merged,
    openpyxl turns every cell in it other than the anchor into a
    read-only MergedCell proxy whose .value can never be set (assigning
    to it raises AttributeError), so there is no way to give those cells
    real content. Fortunately there's also no need to: in Excel, a
    merged block's fill/border/alignment are entirely driven by its
    anchor cell - the merged block renders as a single visual cell, so
    styling only the anchor is both sufficient and correct.

    The underlying problem this works around: openpyxl only actually
    WRITES a cell that nobody has ever assigned a real value to if some
    other attribute (fill, border, font, alignment...) is then set on it
    too - and when it does, it serializes that cell as an empty numeric
    cell (<c t="n"> with no <v>), which is invalid per the OOXML spec.
    That's exactly what makes Excel show its "we found a problem with
    some content" repair prompt every time the file is opened - it
    silently strips those cells back out on load. A single space
    character is visually indistinguishable from genuinely blank against
    any fill color, so using it as a placeholder on the (always real,
    always settable) anchor cell keeps it validly serializable without
    changing how the sheet looks.
    """
    cell = ws.cell(row=row, column=column)
    if cell.value is None:
        cell.value = " "
    return cell


def _write_kpi_card(ws, top_row, left_col, col_span, label, value_formula, fill, accent_hex,
                     value_font=KPI_VALUE_FONT, label_font=KPI_LABEL_FONT, sub_text=None):
    """A single dashboard KPI card: a thin accent strip, a small caps
    label, and a large value - occupying `col_span` columns starting at
    `left_col`, 5 rows tall starting at `top_row`. Returns the cell
    holding the value (so callers can reference it in later formulas).

    Each row-tier is its own single-row (or, for the value, 2-row)
    merge spanning left_col:right_col - so within any given tier, only
    left_col ever holds a real, stylable cell; right_col is always
    absorbed into that same merge. Every fill/alignment/border call
    below is therefore applied to left_col only - see
    _blank_styled_cell's docstring for why this is both necessary
    (touching a merged-away cell either crashes or corrupts the file)
    and sufficient (Excel renders a merged block's styling from its
    anchor alone)."""
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
        cell = ws.cell(row=r, column=2)
        cell.fill = BANNER_FILL
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 26
    ws.row_dimensions[4].height = 18
    # slim accent strip under the banner
    ws.merge_cells(start_row=5, start_column=2, end_row=5, end_column=LAST_BANNER_COL)
    _blank_styled_cell(ws, 5, 2).fill = BANNER_ACCENT_FILL
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
        style_cache = StyleIndexCache()
        ALIGN_NAME = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ALIGN_AMOUNT = Alignment(horizontal="right", vertical="center", indent=1)
        for i, row in enumerate(client_breakdown_df.itertuples(index=False), start=first_data_row):
            stripe = ROW_WHITE_FILL if (i - first_data_row) % 2 == 0 else ROW_ALT_FILL
            name_cell = ws.cell(row=i, column=2, value=str(row[0]))
            # Wrap rather than clip - even at the wider column B, a handful
            # of client names run long enough to still overflow a single
            # line, so this lets those wrap onto a second line instead of
            # being cut off.
            style_cache.apply(name_cell, font=BODY_FONT, border=BORDER, fill=stripe, alignment=ALIGN_NAME)
            prem_cell = ws.cell(row=i, column=3, value=float(row[1]))
            style_cache.apply(prem_cell, border=BORDER, fill=stripe, alignment=ALIGN_AMOUNT, number_format=NUM_FORMAT)
            claim_cell = ws.cell(row=i, column=4, value=float(row[2]))
            style_cache.apply(claim_cell, border=BORDER, fill=stripe, alignment=ALIGN_AMOUNT, number_format=NUM_FORMAT)
            ws.row_dimensions[i].height = 26

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

        # Switched from a horizontal "bar" chart to a vertical "column"
        # chart. The horizontal version needed its category axis order
        # reversed (largest client on top, to match the table), and that
        # reversal is what dragged the value-axis numbers up to the top of
        # the plot area in real Excel, overlapping the title, no matter how
        # explicitly axPos/tickLblPos/crosses were pinned back down - a
        # known rendering quirk with reversed category axes on bar-
        # direction charts. A column chart needs no such reversal at all:
        # the source rows are already sorted largest-first, so the first
        # (leftmost) column is naturally the largest, reading left-to-right
        # like a normal Pareto-style chart, and the value axis has always
        # sat at the bottom by default for a column chart - nothing to
        # fight with Excel over.
        MAX_CHART_CLIENTS = 12
        chart_rows_shown = min(n_clients, MAX_CHART_CLIENTS)
        chart_last_row = first_data_row + chart_rows_shown - 1

        chart_title_text = "Premium vs Claims Receivable by Client"
        if n_clients > MAX_CHART_CLIENTS:
            chart_title_text += f"  (Top {MAX_CHART_CLIENTS} of {n_clients})"

        chart = BarChart()
        chart.type = "col"
        chart.title = _chart_title(chart_title_text, size=1200)
        chart.gapWidth = 40
        chart.overlap = -10

        data = Reference(ws, min_col=3, max_col=4, min_row=table_header_row, max_row=chart_last_row)
        cats = Reference(ws, min_col=2, min_row=first_data_row, max_row=chart_last_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        # Series colors matched to the KPI cards / pie slices above, with
        # no bar border, instead of Excel's default generic theme colors.
        chart.series[0].graphicalProperties = GraphicalProperties(solidFill=ACCENT_BLUE, ln=LineProperties(noFill=True))
        chart.series[1].graphicalProperties = GraphicalProperties(solidFill=CLAIMS_ORANGE, ln=LineProperties(noFill=True))

        chart.legend.position = "b"
        chart.legend.overlay = False
        chart.legend.txPr = _chart_text(size=900, color=SLATE_TEXT)

        # Client names are long, so the category labels are rotated -45
        # degrees along the bottom rather than left horizontal (which
        # would either overlap each other or force the labels to auto-
        # rotate unpredictably) - "rot" is in 60,000ths of a degree.
        name_label_props = _chart_text(size=850, color=SLATE_TEXT)
        name_label_props.bodyPr.rot = -2700000
        name_label_props.bodyPr.vert = "horz"
        chart.x_axis.txPr = name_label_props
        chart.y_axis.txPr = _chart_text(size=850, color=MUTED_TEXT)    # amount axis labels
        chart.y_axis.number_format = "#,##0"
        chart.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=CARD_BORDER, w=9525)))
        chart.x_axis.majorGridlines = None
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.x_axis.lblAlgn = "r"
        chart.x_axis.lblOffset = 100
        # openpyxl always defaults both axes' axPos to "l" regardless of
        # chart type - correct for a bar chart's category axis, but wrong
        # here: for a column chart the category axis (client names) needs
        # to sit at the bottom and the value axis (amounts) at the left.
        chart.x_axis.axPos = "b"
        chart.y_axis.axPos = "l"

        _style_chart_frame(chart)
        # Width scales with how many client columns are shown (not a fixed
        # value) so the rotated name labels underneath have breathing room
        # and don't crowd into each other as more clients are added.
        chart.width = max(20, min(34, 1.7 * chart_rows_shown + 6))
        chart.height = 13
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
    # Widened from the original 30/19/19/19 - at those widths, both the
    # "Premium Receivable"/"Claims Receivable" table headers (right-
    # aligned, ~19-20 characters) and longer client names ("SINDH
    # NOORIABAD POWER COMPANY PHASE II (PVT.) LIMITED", etc.) were being
    # visually clipped since the text overflowed into a neighboring cell
    # that already had its own content. Client Name also gets wrap_text
    # (below) so a long name wraps onto a second line within the wider
    # column instead of being cut off outright.
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 23
    ws.column_dimensions["D"].width = 23
    ws.column_dimensions["E"].width = 19
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 19
    ws.column_dimensions["H"].width = 19
    ws.column_dimensions["I"].width = 3
    ws.column_dimensions["J"].width = 20
    ws.column_dimensions["K"].width = 15


def build_company_workbook(bucket_label, premium_rows, claims_rows, client_breakdown_df, output_dir,
                            premium_amount_col, claims_amount_col, run_date=None, report=None):
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
    _write_sheet(ws_premium, premium_rows, amount_col=premium_amount_col, drop_cols=premium_drop, report=report)

    ws_claims = wb.create_sheet(CLAIMS_SHEET)
    claims_drop = [premium_amount_col] if claims_amount_col != premium_amount_col else []
    _write_sheet(ws_claims, claims_rows, amount_col=claims_amount_col, drop_cols=claims_drop, report=report)

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


def _build_one_workbook_worker(args):
    """
    Module-level (picklable) worker for build_company_workbooks_parallel -
    a ProcessPoolExecutor worker function has to be a plain module-level
    function, not a closure/lambda/method, since worker processes import
    this module fresh and call it by reference rather than by value.
    Returns (bucket_label, filepath, error_str_or_None) rather than
    raising, so one bucket's failure (e.g. a genuinely malformed group)
    doesn't take down the whole batch or need special multiprocessing
    exception plumbing - the caller decides what to do with a per-bucket
    error the same way the existing sequential loop in app.py already does.
    """
    bucket_label, premium_rows, claims_rows, breakdown, output_dir, premium_amount_col, claims_amount_col, run_date = args
    try:
        fp = build_company_workbook(bucket_label, premium_rows, claims_rows, breakdown, output_dir,
                                     premium_amount_col=premium_amount_col, claims_amount_col=claims_amount_col,
                                     run_date=run_date)
        return bucket_label, fp, None
    except Exception as e:
        return bucket_label, None, str(e)


def build_company_workbooks_parallel(groups, breakdowns, output_dir, premium_amount_col, claims_amount_col,
                                      run_date=None, max_workers=None, report=None):
    """
    Generate every bucket's workbook using a process pool instead of one
    at a time - each company's workbook is entirely independent of every
    other (its own DataFrame slice in, its own .xlsx file out), so this
    is an "embarrassingly parallel" batch: openpyxl's own per-file cost
    (building the cell/style/chart XML and writing it to disk - see
    perf_utils.StyleIndexCache's docstring for why that cost is CPU-bound,
    not I/O-bound) is spread across every CPU core instead of running
    entirely on one, which is what actually matters for a real machine
    generating dozens to hundreds of company files in one "Generate Excel
    Reports" click.

    Falls back to plain sequential generation (calling
    build_company_workbook in a loop, exactly as before this function
    existed) whenever parallelism wouldn't help or isn't safely available:
      - fewer than 2 CPU cores detected (nothing to parallelize across)
      - fewer than MIN_BUCKETS_FOR_PARALLEL buckets (process start-up /
        pickling overhead would outweigh the benefit on a small run)
      - ProcessPoolExecutor fails to start at all for any reason (e.g. a
        locked-down environment that disallows spawning subprocesses) -
        caught broadly and degrades to sequential rather than failing
        the whole report generation over an infrastructure problem.

    report: optional callable(str) - progress is reported per completed
        workbook (bucket_label), same granularity as the sequential loop
        it replaces, so the calling UI's progress bar behaves the same
        way regardless of which path actually ran.

    Returns (generated_filepaths, filename_lookup, errors) - same shape
    the app.py "Generate Excel Reports" worker already builds by hand
    around a plain loop, so this is a drop-in replacement for that loop.
    """
    import os as _os
    bucket_labels = list(groups.keys())
    n_cpus = _os.cpu_count() or 1
    MIN_BUCKETS_FOR_PARALLEL = 8  # below this, process start-up overhead isn't worth it
    use_parallel = n_cpus >= 2 and len(bucket_labels) >= MIN_BUCKETS_FOR_PARALLEL

    generated, filename_lookup, errors = [], {}, []

    if use_parallel:
        try:
            import concurrent.futures
            workers = max(1, min(max_workers or n_cpus, n_cpus, len(bucket_labels)))
            if report:
                report(f"Generating {len(bucket_labels)} workbook(s) across {workers} worker process(es)...")
            tasks = [
                (bucket_label, groups[bucket_label]["premium_rows"], groups[bucket_label]["claims_rows"],
                 breakdowns.get(bucket_label), output_dir, premium_amount_col, claims_amount_col, run_date)
                for bucket_label in bucket_labels
            ]
            done_count = 0
            with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
                for bucket_label, fp, err in executor.map(_build_one_workbook_worker, tasks):
                    done_count += 1
                    if err:
                        errors.append(f"{bucket_label}: {err}")
                    else:
                        generated.append(fp)
                        filename_lookup[bucket_label] = os.path.basename(fp)
                    if report:
                        report(f"[{done_count}/{len(bucket_labels)}] {bucket_label}")
                        if hasattr(report, "progress"):
                            report.progress(done_count, len(bucket_labels),
                                             text=f"{done_count} / {len(bucket_labels)} workbook(s) generated")
            return generated, filename_lookup, errors
        except Exception as e:
            # Infrastructure-level failure (pool couldn't start at all) -
            # fall through to the sequential path below rather than
            # losing the whole report generation over it.
            if report:
                report(f"Parallel generation unavailable ({e}) - falling back to sequential generation.")
            generated, filename_lookup, errors = [], {}, []

    for i, bucket_label in enumerate(bucket_labels, start=1):
        g = groups[bucket_label]
        breakdown = breakdowns.get(bucket_label)
        if report:
            report(f"[{i}/{len(bucket_labels)}] {bucket_label}")
        try:
            fp = build_company_workbook(bucket_label, g["premium_rows"], g["claims_rows"], breakdown, output_dir,
                                         premium_amount_col=premium_amount_col, claims_amount_col=claims_amount_col,
                                         run_date=run_date)
            generated.append(fp)
            filename_lookup[bucket_label] = os.path.basename(fp)
        except Exception as e:
            errors.append(f"{bucket_label}: {e}")
    return generated, filename_lookup, errors


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
