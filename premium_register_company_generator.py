"""
Writes one Excel workbook per coinsurer (every "Leader = No" policy row
belonging to that company, pooled the same way every other GIMS module
pools an "OFC" bucket), containing exactly two sheets: "Summary" (KPI
cards, a Gross-vs-Net pie, and a per-insured bar chart) and "Premium
Data" (every original Premium Register column for that company's own
rows, untouched).

Styling mirrors receivable_generator.py / claims_register_generator.py
so output looks consistent with the rest of GIMS.
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
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RegularTextRun, Font as DrawFont

import company_alias_resolver as car

try:
    from perf_utils import ThrottledReporter, StyleIndexCache
except ImportError:  # perf_utils.py not yet added to this checkout, degrade gracefully
    class ThrottledReporter:
        def __init__(self, *a, **k):
            pass

        def maybe(self, *a, **k):
            pass

        def final(self, *a, **k):
            pass

    class StyleIndexCache:
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

NAVY = "1F4E78"
ACCENT_BLUE = "2E86AB"
PREMIUM_ORANGE = "D97B29"
SLATE_TEXT = "1C2733"
MUTED_TEXT = "6C7A89"
CARD_BORDER = "D9E2EC"
ROW_ALT_FILL = PatternFill("solid", start_color="F4F7FA", end_color="F4F7FA")
ROW_WHITE_FILL = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
BANNER_FILL = PatternFill("solid", start_color=NAVY, end_color=NAVY)
BANNER_ACCENT_FILL = PatternFill("solid", start_color=ACCENT_BLUE, end_color=ACCENT_BLUE)
KPI_GROSS_FILL = PatternFill("solid", start_color="EAF2FA", end_color="EAF2FA")
KPI_NET_FILL = PatternFill("solid", start_color="FDF3EA", end_color="FDF3EA")
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
DATA_SHEET = "Premium Data"


def _write_text_safe(cell):
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
    canonical = car.canonical_filename_for_label(bucket_label)
    stub = canonical if canonical else safe_filename_part(bucket_label)
    return f"{stub}_PremiumRegister_{run_date.strftime('%d%m%Y')}.xlsx"


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


def _write_data_sheet(ws, df, report=None):
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
        ws.cell(row=2, column=1, value="No policies for this company.").font = EMPTY_FONT
        for j, col in enumerate(columns, start=1):
            ws.column_dimensions[get_column_letter(j)].width = max(14, min(28, len(str(col)) + 4))
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
            if col in ("Gross Premium", "Net Premium", "Sum Insured", "Admin", "C.E.D", "F.I.F", "Stamp Duty"):
                number_format = NUM_FORMAT
            elif col == "Share":
                number_format = "0.00"
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
        ws.column_dimensions[get_column_letter(j)].width = max(12, min(28, len(str(col)) + 4))


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


def _write_summary_sheet(ws, bucket_label, breakdown_df, total_gross, total_net, num_policies):
    ws.sheet_view.showGridLines = False
    n_insureds = len(breakdown_df)

    card_row = 8
    bar_chart_row = 26
    table_header_row = 55
    first_data_row = table_header_row + 1
    last_data_row = table_header_row + n_insureds if n_insureds else table_header_row
    total_row = last_data_row + 1

    LAST_BANNER_COL = 12
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=LAST_BANNER_COL)
    ws.cell(row=2, column=2, value="PREMIUM REGISTER STATEMENT").font = KPI_LABEL_FONT_WHITE
    ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=LAST_BANNER_COL)
    ws.cell(row=3, column=2, value=bucket_label).font = Font(name="Arial", bold=True, size=18, color="FFFFFF")
    ws.merge_cells(start_row=4, start_column=2, end_row=4, end_column=LAST_BANNER_COL)
    subtitle = f"{num_policies} polic{'y' if num_policies == 1 else 'ies'}   \u2022   generated {date.today().strftime('%d-%b-%Y')}"
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

    gross_range = f"C{first_data_row}:C{last_data_row}" if n_insureds else None
    net_range = f"D{first_data_row}:D{last_data_row}" if n_insureds else None

    gross_cell = _write_kpi_card(
        ws, card_row, 2, 2, "Gross Premium",
        f"=SUM({gross_range})" if gross_range else 0,
        KPI_GROSS_FILL, ACCENT_BLUE, sub_text="This company's total gross premium share",
    )
    net_cell = _write_kpi_card(
        ws, card_row, 4, 2, "Net Premium",
        f"=SUM({net_range})" if net_range else 0,
        KPI_NET_FILL, PREMIUM_ORANGE, sub_text="This company's total net premium share",
    )
    _write_kpi_card(
        ws, card_row, 7, 2, "Policies",
        num_policies,
        KPI_TOTAL_FILL, "5FA8D3", value_font=KPI_VALUE_FONT_WHITE, label_font=KPI_LABEL_FONT_WHITE,
        sub_text="Total policy rows in this file",
    )
    ws.cell(row=card_row + 2, column=7).number_format = "#,##0"

    pie_src_row = card_row + 1
    ws.cell(row=pie_src_row, column=10, value="Category").font = Font(name="Arial", size=8, bold=True, color=MUTED_TEXT)
    ws.cell(row=pie_src_row, column=11, value="Total").font = Font(name="Arial", size=8, bold=True, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 1, column=10, value="Gross Premium").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 1, column=11, value=f"={gross_cell.coordinate}").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 1, column=11).number_format = NUM_FORMAT
    ws.cell(row=pie_src_row + 2, column=10, value="Net Premium").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 2, column=11, value=f"={net_cell.coordinate}").font = Font(name="Arial", size=8, color=MUTED_TEXT)
    ws.cell(row=pie_src_row + 2, column=11).number_format = NUM_FORMAT

    pie = PieChart()
    pie.title = _chart_title("Gross vs Net Premium", size=1300)
    pie_data = Reference(ws, min_col=11, min_row=pie_src_row, max_row=pie_src_row + 2)
    pie_cats = Reference(ws, min_col=10, min_row=pie_src_row + 1, max_row=pie_src_row + 2)
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)

    gross_slice = DataPoint(idx=0)
    gross_slice.graphicalProperties = GraphicalProperties(solidFill=ACCENT_BLUE, ln=LineProperties(solidFill="FFFFFF", w=19050))
    net_slice = DataPoint(idx=1)
    net_slice.graphicalProperties = GraphicalProperties(solidFill=PREMIUM_ORANGE, ln=LineProperties(solidFill="FFFFFF", w=19050))
    pie.series[0].data_points = [gross_slice, net_slice]

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showCatName = False
    pie.dataLabels.showVal = False
    pie.dataLabels.showSerName = False
    pie.dataLabels.showLegendKey = False
    pie.dataLabels.numFmt = "0.0%"
    pie.dataLabels.dLblPos = "outEnd"
    pie.dataLabels.txPr = _chart_text(size=1000, color=SLATE_TEXT, bold=True)
    pie.legend.position = "b"
    pie.legend.overlay = False
    pie.legend.txPr = _chart_text(size=900, color=SLATE_TEXT)

    pie.layout = Layout(
        manualLayout=ManualLayout(
            layoutTarget="inner",
            xMode="edge", yMode="edge",
            x=0.06, y=0.24, w=0.88, h=0.62,
        )
    )

    _style_chart_frame(pie)
    pie.height = 8.6
    pie.width = 12.6
    ws.add_chart(pie, "M8")

    if n_insureds > 0:
        MAX_CHART_ROWS = 12
        chart_rows_shown = min(n_insureds, MAX_CHART_ROWS)
        chart_last_row = first_data_row + chart_rows_shown - 1
        chart_title_text = "Gross vs Net Premium by Insured"
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
        chart.series[1].graphicalProperties = GraphicalProperties(solidFill=PREMIUM_ORANGE, ln=LineProperties(noFill=True))

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
        ws.add_chart(chart, f"B{bar_chart_row}")

    ws.cell(row=table_header_row - 2, column=2, value="Premium by Insured").font = TABLE_TITLE_FONT

    for j, h in enumerate(["Insured Name", "Gross Premium", "Net Premium", "Policies"], start=2):
        c = ws.cell(row=table_header_row, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="left" if j == 2 else "right", vertical="center", indent=1)
    ws.row_dimensions[table_header_row].height = 20

    if n_insureds == 0:
        ws.cell(row=first_data_row, column=2, value="No policy data available.").font = EMPTY_FONT
    else:
        style_cache = StyleIndexCache()
        ALIGN_NAME = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", indent=1)
        for i, row in enumerate(breakdown_df.itertuples(index=False), start=first_data_row):
            stripe = ROW_WHITE_FILL if (i - first_data_row) % 2 == 0 else ROW_ALT_FILL
            name_cell = ws.cell(row=i, column=2, value=str(row[0]))
            style_cache.apply(name_cell, font=BODY_FONT, border=BORDER, fill=stripe, alignment=ALIGN_NAME)
            gross_c = ws.cell(row=i, column=3, value=float(row[1]))
            style_cache.apply(gross_c, border=BORDER, fill=stripe, alignment=ALIGN_RIGHT, number_format=NUM_FORMAT)
            net_c = ws.cell(row=i, column=4, value=float(row[2]))
            style_cache.apply(net_c, border=BORDER, fill=stripe, alignment=ALIGN_RIGHT, number_format=NUM_FORMAT)
            n_c = ws.cell(row=i, column=5, value=int(row[3]))
            style_cache.apply(n_c, border=BORDER, fill=stripe, alignment=ALIGN_RIGHT)
            ws.row_dimensions[i].height = 26

        total_label = ws.cell(row=total_row, column=2, value="TOTAL")
        total_gross_c = ws.cell(row=total_row, column=3, value=f"=SUM(C{first_data_row}:C{last_data_row})")
        total_net_c = ws.cell(row=total_row, column=4, value=f"=SUM(D{first_data_row}:D{last_data_row})")
        total_gross_c.number_format = NUM_FORMAT
        total_net_c.number_format = NUM_FORMAT
        for cell in (total_label, total_gross_c, total_net_c):
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


def build_company_workbook(bucket_label, rows_df, breakdown_df, output_dir, run_date=None, report=None):
    """
    Write a single company's workbook:
      1. "Summary" - KPI cards, Gross-vs-Net pie, per-insured bar chart
      2. "Premium Data" - every original Premium Register row for that
         company

    Returns the full filepath written.
    """
    os.makedirs(output_dir, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)

    ws_data = wb.create_sheet(DATA_SHEET)
    _write_data_sheet(ws_data, rows_df, report=report)

    ws_summary = wb.create_sheet(SUMMARY_SHEET)
    gross_col = next((c for c in rows_df.columns if c == "Gross Premium"), None)
    net_col = next((c for c in rows_df.columns if c == "Net Premium"), None)
    total_gross = pd.to_numeric(rows_df[gross_col], errors="coerce").fillna(0).sum() if gross_col else 0.0
    total_net = pd.to_numeric(rows_df[net_col], errors="coerce").fillna(0).sum() if net_col else 0.0
    num_policies = len(rows_df)
    _write_summary_sheet(ws_summary, bucket_label, breakdown_df if breakdown_df is not None else pd.DataFrame(),
                          total_gross, total_net, num_policies)

    wb.move_sheet(SUMMARY_SHEET, offset=-1)

    filename = build_output_filename(bucket_label, run_date=run_date)
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)
    return filepath


def _build_one_workbook_worker(args):
    bucket_label, rows_df, breakdown_df, output_dir, run_date = args
    try:
        fp = build_company_workbook(bucket_label, rows_df, breakdown_df, output_dir, run_date=run_date)
        return bucket_label, fp, None
    except Exception as e:
        return bucket_label, None, str(e)


def build_company_workbooks_parallel(groups, breakdowns, output_dir, run_date=None, max_workers=None, report=None):
    """
    Generate every company's workbook using a process pool instead of
    one at a time - identical pattern (and identical fallback
    conditions) to receivable_generator.build_company_workbooks_parallel /
    claims_register_generator.build_coinsurer_workbooks_parallel.

    Returns (generated_filepaths, filename_lookup, errors).
    """
    import os as _os
    bucket_labels = list(groups.keys())
    n_cpus = _os.cpu_count() or 1
    MIN_BUCKETS_FOR_PARALLEL = 8
    use_parallel = n_cpus >= 2 and len(bucket_labels) >= MIN_BUCKETS_FOR_PARALLEL

    generated, filename_lookup, errors = [], {}, []

    if use_parallel:
        try:
            import concurrent.futures
            workers = max(1, min(max_workers or n_cpus, n_cpus, len(bucket_labels)))
            if report:
                report(f"Generating {len(bucket_labels)} workbook(s) across {workers} worker process(es)...")
            tasks = [
                (bucket_label, groups[bucket_label], breakdowns.get(bucket_label), output_dir, run_date)
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
            if report:
                report(f"Parallel generation unavailable ({e}) - falling back to sequential generation.")
            generated, filename_lookup, errors = [], {}, []

    for i, bucket_label in enumerate(bucket_labels, start=1):
        rows_df = groups[bucket_label]
        breakdown = breakdowns.get(bucket_label)
        if report:
            report(f"[{i}/{len(bucket_labels)}] {bucket_label}")
        try:
            fp = build_company_workbook(bucket_label, rows_df, breakdown, output_dir, run_date=run_date)
            generated.append(fp)
            filename_lookup[bucket_label] = os.path.basename(fp)
        except Exception as e:
            errors.append(f"{bucket_label}: {e}")
    return generated, filename_lookup, errors
