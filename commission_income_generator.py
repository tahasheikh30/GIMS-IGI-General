"""
Writes the Commission Income output workbook - modelled on the manually
built "Commission Income & Unearned" reference workbook this module was
reverse-engineered from, but driven by live Excel formulas end to end so
the working stays auditable and a changed % flows all the way through:

  Summary                      Treaty / FAC / Total commission income per
                               department for each register year, the
                               opening / closing RI Unearned Commission
                               and the earned figure, plus a tie-out check.
  RI Unearned Comm. Working    Treaty and FAC commission per department x
                               Period (Previous / Jan..Dec / Future
                               Policies) for both years, their total, then
                               the 1/24 (1/6 Marine/Travel, zero BIMA)
                               amortization -> RI Unearned Commission.
  Check                        Per-department, per-Period treaty / FAC / total
                               commission summed straight from the detail
                               sheets and tied back to the Summary (laid
                               out like the reference workbook's Check sheet).
  Treaty Rates                 One row per department x Treaty
                               Description: total Treaty Premium, the
                               commission % (the ONLY place a % lives -
                               edit it here and everything recalculates)
                               and the resulting commission.
  RI <Dept> <Year>             Policy-level detail for every premium
                               register: Treaty Premium, one commission
                               column per Treaty Description
                               (= premium x that description's %), total
                               and the Period.
  RI Commission FAC <Year>     Policy-level FAC detail with its Period.

Detail sheets are tens of thousands of rows (Motor/Marine run to 70,000+),
so every sheet is written in openpyxl's write-only streaming mode - a
normal-mode workbook of this size would need well over a gigabyte of RAM.
"""
import os
import subprocess
import sys
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import commission_income_processor as cip

NAVY = "1F4E78"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOTAL_BORDER = Border(top=Side(style="thin", color=NAVY), left=THIN, right=THIN, bottom=THIN)
NUM = "#,##0;(#,##0);-"
NUM2 = "#,##0.00;(#,##0.00);-"
PCT = "0.00%"
DATE_FMT = "dd-mmm-yyyy"

# name -> kwargs for WriteOnlyCell styling
STYLES = {
    "title": dict(font=Font(name="Arial", bold=True, size=14, color=NAVY)),
    "subtitle": dict(font=Font(name="Arial", size=10, color="6C7A89")),
    "section": dict(font=Font(name="Arial", bold=True, size=12, color=NAVY)),
    "note": dict(font=Font(name="Arial", size=9, italic=True, color="6C7A89")),
    "driver": dict(font=Font(name="Arial", size=9, italic=True, color="6C7A89"), alignment=Alignment(horizontal="center")),
    "driver_label": dict(font=Font(name="Arial", size=9, italic=True, color="6C7A89")),
    "header": dict(font=Font(name="Arial", bold=True, color="FFFFFF"), fill=PatternFill("solid", start_color=NAVY, end_color=NAVY),
                   alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=BORDER),
    "header_left": dict(font=Font(name="Arial", bold=True, color="FFFFFF"), fill=PatternFill("solid", start_color=NAVY, end_color=NAVY),
                        alignment=Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True), border=BORDER),
    "body": dict(font=Font(name="Arial"), border=BORDER),
    "body_num": dict(font=Font(name="Arial"), border=BORDER, number_format=NUM),
    "body_num2": dict(font=Font(name="Arial"), border=BORDER, number_format=NUM2),
    "bold_num": dict(font=Font(name="Arial", bold=True), border=BORDER, number_format=NUM),
    "input_pct": dict(font=Font(name="Arial", color="0000FF"), border=BORDER, number_format=PCT,
                      fill=PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")),
    "input_num": dict(font=Font(name="Arial", color="0000FF"), border=BORDER, number_format=NUM,
                      fill=PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")),
    "total_label": dict(font=Font(name="Arial", bold=True), fill=PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2"), border=TOTAL_BORDER),
    "total_num": dict(font=Font(name="Arial", bold=True), fill=PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2"),
                      border=TOTAL_BORDER, number_format=NUM),
    "total_pct": dict(font=Font(name="Arial", bold=True), fill=PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2"),
                      border=TOTAL_BORDER, number_format=PCT),
    "pct_link": dict(font=Font(name="Arial", size=9, italic=True, color="008000"), number_format=PCT, alignment=Alignment(horizontal="right")),
    "label_small": dict(font=Font(name="Arial", size=9, italic=True, color="6C7A89"), alignment=Alignment(horizontal="right")),
    "col_total": dict(font=Font(name="Arial", bold=True, size=10), number_format=NUM),
}

WORKING_SHEET = "RI Unearned Comm. Working"
CHECK_SHEET = "Check"
SUMMARY_FIRST_ROW = 7   # first department row on the Summary sheet
SUMMARY_SHEET = "Summary"
RATES_SHEET = "Treaty Rates"

FIRST_DATA_ROW = 7
HEADER_ROW = 6
TOTALS_ROW = 5


def _L(idx):
    return get_column_letter(idx)


def _q(sheet):
    return "'" + sheet.replace("'", "''") + "'"


def _styled(ws, value, style=None, number_format=None):
    cell = WriteOnlyCell(ws, value=value)
    if style:
        spec = STYLES[style]
        for attr, val in spec.items():
            setattr(cell, attr, val)
    if number_format:
        cell.number_format = number_format
    return cell


class _Buf:
    """Small-sheet buffer: set cells by (row, col), flush in row order
    (write-only worksheets can only be appended to sequentially)."""

    def __init__(self):
        self.cells = {}
        self.widths = {}
        self.heights = {}
        self.freeze = None

    def set(self, r, c, value, style=None):
        self.cells[(r, c)] = (value, style)

    def flush(self, ws):
        for c, w in self.widths.items():
            ws.column_dimensions[_L(c)].width = w
        if self.freeze:
            ws.freeze_panes = self.freeze
        ws.sheet_view.showGridLines = False
        if not self.cells:
            return
        max_r = max(r for r, _ in self.cells)
        max_c = max(c for _, c in self.cells)
        for r in range(1, max_r + 1):
            row = []
            for c in range(1, max_c + 1):
                item = self.cells.get((r, c))
                if item is None:
                    row.append(None)
                else:
                    row.append(_styled(ws, item[0], item[1]))
            ws.append(row)


# ---------------------------------------------------------------------------
# Layout helpers shared by every sheet (formulas must agree on coordinates)
# ---------------------------------------------------------------------------

def _register_keys(df, result_summary, dept):
    """Treaty Description keys that get a commission column on this
    register's sheet: every key present in the register, except a blank
    description carrying no premium at all."""
    keys = []
    for key in sorted(df["Desc Key"].unique(), key=lambda k: (k == "", k)):
        if key == "":
            prem = df.loc[df["Desc Key"] == "", "Treaty Premium"].abs().sum()
            if prem < 0.005:
                continue
        keys.append(key)
    return keys


def _detail_layout(keys):
    layout = {"serial": 1, "policy": 2, "insured": 3, "desc": 4, "date": 5, "prem": 6}
    layout["desc_cols"] = {k: 7 + i for i, k in enumerate(keys)}
    layout["total"] = 7 + len(keys)
    layout["period"] = 8 + len(keys)
    return layout


def _sheet_name(dept, year):
    return f"RI {dept} {year}"


def _fac_sheet_name(year):
    return f"RI Commission FAC {year}"


# FAC sheet columns (fixed): A Serial, B Policy, C Insured, D Closing#,
# E Department, F Comm Date, G FAC Premium, H FAC Commission, I Period,
# J Date Used.
FAC_COLS = {"dept": "E", "commission": "H", "period": "I", "premium": "G"}


# ---------------------------------------------------------------------------
# Detail sheets (streamed)
# ---------------------------------------------------------------------------

def _write_register_sheet(wb, result, dept, slot, df, layout, keys, rate_row_by_key):
    year = result["year_labels"][slot]
    ws = wb.create_sheet(_sheet_name(dept, year))
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = f"A{FIRST_DATA_ROW}"
    widths = {1: 8, 2: 24, 3: 38, 4: 34, 5: 13, 6: 16}
    for c, w in widths.items():
        ws.column_dimensions[_L(c)].width = w
    for c in list(layout["desc_cols"].values()) + [layout["total"]]:
        ws.column_dimensions[_L(c)].width = 22
    ws.column_dimensions[_L(layout["period"])].width = 16

    n = len(df)
    last = FIRST_DATA_ROW + n - 1 if n else FIRST_DATA_ROW
    ncols = layout["period"]
    rate_col_letter = "F" if slot == cip.SLOT_CURRENT else "G"

    def blank_row():
        return [None] * ncols

    # Row 1-2 titles
    r1 = blank_row(); r1[0] = _styled(ws, "IGI General Insurance Company Limited", "title"); ws.append(r1)
    r2 = blank_row()
    r2[0] = _styled(ws, f"RI {dept} Premium Register - {year} register  \u2022  period {result['period_start']:%d-%b-%Y} to {result['period_end']:%d-%b-%Y}", "subtitle")
    ws.append(r2)
    ws.append(blank_row())
    # Row 4: commission % per description (linked to the Treaty Rates sheet)
    r4 = blank_row()
    r4[layout["prem"] - 1] = _styled(ws, "Commission % \u2192", "label_small")
    for key, col in layout["desc_cols"].items():
        rr = rate_row_by_key[(dept, key)]
        r4[col - 1] = _styled(ws, f"={_q(RATES_SHEET)}!${rate_col_letter}${rr}", "pct_link")
    ws.append(r4)
    # Row 5: column totals
    r5 = blank_row()
    r5[0] = _styled(ws, "Total", "col_total")
    for col in [layout["prem"]] + list(layout["desc_cols"].values()) + [layout["total"]]:
        L = _L(col)
        r5[col - 1] = _styled(ws, f"=SUM({L}{FIRST_DATA_ROW}:{L}{last})", "col_total")
    ws.append(r5)
    # Row 6: header
    r6 = blank_row()
    heads = {layout["serial"]: "Serial No", layout["policy"]: "Policy Number", layout["insured"]: "Insured",
             layout["desc"]: "Treaty Description", layout["date"]: "Comm. Date", layout["prem"]: "Treaty Premium",
             layout["total"]: "Total Treaty Commission", layout["period"]: "Period"}
    descs_text = df.drop_duplicates("Desc Key").set_index("Desc Key")["Treaty Description"].to_dict()
    for key, col in layout["desc_cols"].items():
        heads[col] = (descs_text.get(key) or cip.BLANK_DESC_LABEL) + " Commission"
    for col, text in heads.items():
        r6[col - 1] = _styled(ws, text, "header")
    ws.append(r6)

    # Data rows
    desc_cols = layout["desc_cols"]
    first_desc_L = _L(7) if desc_cols else None
    last_desc_L = _L(6 + len(desc_cols)) if desc_cols else None
    policies = df["Policy"].tolist()
    insureds = df["Insured"].tolist()
    descs = df["Treaty Description"].tolist()
    keys_col = df["Desc Key"].tolist()
    dates = df["Comm Date"].dt.to_pydatetime().tolist()
    prems = df["Treaty Premium"].tolist()
    periods = df["Period"].tolist()
    prem_fmt = NUM
    date_cell_fmt = DATE_FMT
    for i in range(n):
        r = FIRST_DATA_ROW + i
        row = [None] * ncols
        row[0] = i + 1
        row[1] = policies[i]
        ins = insureds[i]
        row[2] = None if pd.isna(ins) else ins
        row[3] = descs[i] or None
        dc = WriteOnlyCell(ws, value=dates[i]); dc.number_format = date_cell_fmt
        row[4] = dc
        pc = WriteOnlyCell(ws, value=prems[i]); pc.number_format = prem_fmt
        row[5] = pc
        col = desc_cols.get(keys_col[i])
        if col is not None:
            fc = WriteOnlyCell(ws, value=f"=F{r}*{_L(col)}$4"); fc.number_format = prem_fmt
            row[col - 1] = fc
        if desc_cols:
            tc = WriteOnlyCell(ws, value=f"=SUM({first_desc_L}{r}:{last_desc_L}{r})"); tc.number_format = prem_fmt
            row[layout["total"] - 1] = tc
        row[layout["period"] - 1] = periods[i]
        ws.append(row)
    return {"sheet": ws.title, "last": last, "n": n}


def _write_fac_sheet(wb, result, slot, df):
    year = result["year_labels"][slot]
    ws = wb.create_sheet(_fac_sheet_name(year))
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = f"A{FIRST_DATA_ROW}"
    for c, w in {1: 8, 2: 22, 3: 38, 4: 20, 5: 16, 6: 13, 7: 16, 8: 20, 9: 16, 10: 26}.items():
        ws.column_dimensions[_L(c)].width = w
    n = len(df)
    last = FIRST_DATA_ROW + n - 1 if n else FIRST_DATA_ROW

    def blank():
        return [None] * 10

    r = blank(); r[0] = _styled(ws, "IGI General Insurance Company Limited", "title"); ws.append(r)
    r = blank()
    r[0] = _styled(ws, f"RI FAC Commission Register - {year} register  \u2022  period {result['period_start']:%d-%b-%Y} to {result['period_end']:%d-%b-%Y}", "subtitle")
    ws.append(r)
    r = blank(); r[0] = _styled(ws, "Marine lines carry no commencement date in the FAC register - their Acceptance Date is used (see 'Date Used').", "note"); ws.append(r)
    ws.append(blank())
    r = blank()
    r[0] = _styled(ws, "Total", "col_total")
    for col in (7, 8):
        L = _L(col)
        r[col - 1] = _styled(ws, f"=SUM({L}{FIRST_DATA_ROW}:{L}{last})", "col_total")
    ws.append(r)
    r = blank()
    for i, text in enumerate(["Serial No", "Policy", "Insured", "Closing#", "Department", "Comm. Date",
                               "FAC Premium", "FAC Commission Income Amount", "Period", "Date Used"]):
        r[i] = _styled(ws, text, "header")
    ws.append(r)

    policies = df["Policy"].tolist()
    insureds = df["Insured"].tolist()
    closings = df["Closing"].tolist()
    depts = df["Department"].tolist()
    dates = df["Comm Date"].dt.to_pydatetime().tolist()
    prems = df["FAC Premium"].tolist()
    comms = df["FAC Commission"].tolist()
    periods = df["Period"].tolist()
    sources = df["Date Source"].tolist()
    for i in range(n):
        row = blank()
        row[0] = i + 1
        row[1] = policies[i]
        ins = insureds[i]
        row[2] = None if pd.isna(ins) else ins
        cl = closings[i]
        row[3] = None if pd.isna(cl) else cl
        row[4] = depts[i]
        dc = WriteOnlyCell(ws, value=dates[i]); dc.number_format = DATE_FMT
        row[5] = dc
        pc = WriteOnlyCell(ws, value=prems[i]); pc.number_format = NUM
        row[6] = pc
        cc = WriteOnlyCell(ws, value=comms[i]); cc.number_format = NUM
        row[7] = cc
        row[8] = periods[i]
        row[9] = sources[i]
        ws.append(row)
    return {"sheet": ws.title, "last": last, "n": n}


# ---------------------------------------------------------------------------
# Small sheets
# ---------------------------------------------------------------------------

def _build_rates_buf(result, rate_rows):
    """rate_rows: [(summary_row_dict, excel_row_number), ...]"""
    b = _Buf()
    cur_y, prev_y = result["year_labels"][cip.SLOT_CURRENT], result["year_labels"][cip.SLOT_PREVIOUS]
    b.set(1, 1, "Treaty Commission Rates", "title")
    b.set(2, 1, f"Total Treaty Premium per Treaty Description, the commission % applied to it, and the resulting commission.  "
                 f"Period {result['period_start']:%d-%b-%Y} to {result['period_end']:%d-%b-%Y}", "subtitle")
    b.set(3, 1, "The yellow % cells are the only inputs - change one and every detail sheet, the working and the summary recalculate.", "note")
    heads = ["Department", "Treaty Description", f"Treaty Premium {cur_y}", f"Treaty Premium {prev_y}", "Total Treaty Premium",
             f"Commission % {cur_y}", f"Commission % {prev_y}", f"Commission {cur_y}", f"Commission {prev_y}", "Total Commission"]
    for j, h in enumerate(heads, start=1):
        b.set(5, j, h, "header_left" if j <= 2 else "header")
    first = 6
    for row, r in rate_rows:
        b.set(r, 1, row["department"], "body")
        b.set(r, 2, row["description"], "body")
        b.set(r, 3, row["current_premium"], "body_num")
        b.set(r, 4, row["previous_premium"], "body_num")
        b.set(r, 5, f"=C{r}+D{r}", "body_num")
        b.set(r, 6, (row["current_pct"] or 0.0) / 100.0, "input_pct")
        b.set(r, 7, (row["previous_pct"] or 0.0) / 100.0, "input_pct")
        b.set(r, 8, f"=C{r}*F{r}", "body_num")
        b.set(r, 9, f"=D{r}*G{r}", "body_num")
        b.set(r, 10, f"=H{r}+I{r}", "body_num")
    last = rate_rows[-1][1] if rate_rows else first
    tr = last + 1
    b.set(tr, 1, "Total", "total_label"); b.set(tr, 2, None, "total_label")
    for c in (3, 4, 5, 8, 9, 10):
        L = _L(c)
        b.set(tr, c, f"=SUM({L}{first}:{L}{last})", "total_num")
    b.set(tr, 6, None, "total_label"); b.set(tr, 7, None, "total_label")
    for c, w in {1: 18, 2: 42, 3: 20, 4: 20, 5: 20, 6: 16, 7: 16, 8: 20, 9: 20, 10: 20}.items():
        b.widths[c] = w
    b.freeze = "C6"
    return b


def _build_working_buf(result, sheet_info):
    b = _Buf()
    departments = result["departments"]
    methods = result["methods"]
    period_start = result["period_start"]
    cur_y, prev_y = result["year_labels"][cip.SLOT_CURRENT], result["year_labels"][cip.SLOT_PREVIOUS]
    cols = cip.TABLE_COLUMNS
    ncols = len(cols)
    total_col = 3 + ncols

    b.set(1, 2, "RI Unearned Commission Working", "title")
    b.set(2, 2, f"Period {period_start:%d-%b-%Y} to {result['period_end']:%d-%b-%Y}  \u2022  generated {date.today():%d-%b-%Y}", "subtitle")

    anchors = {}

    def block(start, label, formula_fn, key):
        b.set(start, 2, label, "section")
        hdr = start + 1
        b.set(hdr, 2, "Department", "header_left")
        for j, col in enumerate(cols, start=3):
            b.set(hdr, j, col, "header")
        b.set(hdr, total_col, cip.TOTAL_LABEL, "header")
        first = hdr + 1
        for i, dept in enumerate(departments):
            r = first + i
            b.set(r, 2, dept, "body")
            for j, col in enumerate(cols, start=3):
                b.set(r, j, formula_fn(dept, r, j, col, hdr), "body_num")
            b.set(r, total_col, f"=SUM({_L(3)}{r}:{_L(2 + ncols)}{r})", "bold_num")
        tr = first + len(departments)
        b.set(tr, 2, cip.TOTAL_LABEL, "total_label")
        for j in range(3, total_col + 1):
            L = _L(j)
            b.set(tr, j, f"=SUM({L}{first}:{L}{tr - 1})", "total_num")
        anchors[key] = {"hdr": hdr, "first": first, "total": tr}
        return tr + 2

    def treaty_formula(slot):
        def fn(dept, r, j, col, hdr):
            info = sheet_info.get((dept, slot))
            if info is None:
                return 0
            sh, layout, last = _q(info["sheet"]), info["layout"], info["last"]
            tot, per = _L(layout["total"]), _L(layout["period"])
            return (f"=SUMIFS({sh}!${tot}${FIRST_DATA_ROW}:${tot}${last},"
                    f"{sh}!${per}${FIRST_DATA_ROW}:${per}${last},{_L(j)}${hdr})")
        return fn

    def fac_formula(slot):
        def fn(dept, r, j, col, hdr):
            info = sheet_info.get(("FAC", slot))
            if info is None:
                return 0
            sh, last = _q(info["sheet"]), info["last"]
            return (f"=SUMIFS({sh}!$H${FIRST_DATA_ROW}:$H${last},{sh}!$E${FIRST_DATA_ROW}:$E${last},$B{r},"
                    f"{sh}!$I${FIRST_DATA_ROW}:$I${last},{_L(j)}${hdr})")
        return fn

    row = 4
    row = block(row, f"Table 1 \u2014 Treaty Commission, {cur_y} registers", treaty_formula(cip.SLOT_CURRENT), "t_cur")
    row = block(row, f"Table 2 \u2014 FAC Commission, {cur_y} FAC register", fac_formula(cip.SLOT_CURRENT), "f_cur")
    row = block(row, f"Table 3 \u2014 Treaty Commission, {prev_y} registers", treaty_formula(cip.SLOT_PREVIOUS), "t_prev")
    row = block(row, f"Table 4 \u2014 FAC Commission, {prev_y} FAC register", fac_formula(cip.SLOT_PREVIOUS), "f_prev")

    def total_formula(dept, r, j, col, hdr):
        i = r - (hdr + 1)
        L = _L(j)
        parts = [f"{L}{anchors[k]['first'] + i}" for k in ("t_cur", "f_cur", "t_prev", "f_prev")]
        return "=" + "+".join(parts)
    row = block(row, "Table 5 \u2014 Total RI Commission (Treaty + FAC, both years)", total_formula, "total")

    # Driver rows
    b.set(row, 2, "Driver rows below feed Table 6's formulas (numerator / denominator per column).", "note")
    row += 1
    month_pos = cip.up._month_to_period_position(period_start)
    driver_rows = {}
    for label, denom in (("1/24", 24), ("1/6", 6)):
        pr, yr = row, row + 1
        b.set(pr, 2, f"{label} Period", "driver_label")
        b.set(yr, 2, f"{label} Denominator", "driver_label")
        for j, col in enumerate(cols, start=3):
            if col == cip.PREVIOUS_LABEL:
                num = 0
            elif col == cip.FUTURE_LABEL:
                num = denom
            else:
                n = month_pos[cip.MONTH_NAMES.index(col) + 1]
                num = (2 * n - 1) if label == "1/24" else {10: 1, 11: 3, 12: 5}.get(n, 0)
            b.set(pr, j, num, "driver")
            b.set(yr, j, denom, "driver")
        driver_rows[label] = (pr, yr)
        row += 3

    def unearned_formula(dept, r, j, col, hdr):
        method = methods[dept]
        if method == cip.up.METHOD_ZERO:
            return 0
        i = r - (hdr + 1)
        L = _L(j)
        pr, yr = driver_rows["1/24" if method == cip.up.METHOD_1_24 else "1/6"]
        return f"={L}{anchors['total']['first'] + i}*{L}${pr}/{L}${yr}"
    row = block(row, "Table 6 \u2014 RI Unearned Commission", unearned_formula, "unearned")

    b.widths[1] = 2
    b.widths[2] = 26
    for j in range(3, total_col):
        b.widths[j] = 16
    b.widths[total_col] = 20
    b.freeze = "C4"
    return b, anchors, total_col


def _build_summary_buf(result, sheet_info, anchors, total_col, rates_info, opening, check_cells):
    b = _Buf()
    departments = result["departments"]
    cur_y, prev_y = result["year_labels"][cip.SLOT_CURRENT], result["year_labels"][cip.SLOT_PREVIOUS]
    ps, pe = result["period_start"], result["period_end"]

    b.set(1, 2, "IGI General Insurance Company Limited", "title")
    b.set(2, 2, "RI Commission Income & Unearned Commission", "section")
    b.set(3, 2, f"For the period {ps:%d-%b-%Y} to {pe:%d-%b-%Y}  \u2022  generated {date.today():%d-%b-%Y}", "subtitle")

    b.set(5, 3, f"{cur_y} registers", "section")
    b.set(5, 7, f"{prev_y} registers", "section")
    for j, h in enumerate(["Department", "Treaty", "FAC", "Total"], start=2):
        b.set(6, j, h, "header_left" if j == 2 else "header")
    for j, h in enumerate(["Department", "Treaty", "FAC", "Total"], start=6):
        b.set(6, j, h, "header_left" if j == 6 else "header")

    def treaty_ref(dept, slot):
        info = sheet_info.get((dept, slot))
        if info is None:
            return 0
        return f"={_q(info['sheet'])}!{_L(info['layout']['total'])}{TOTALS_ROW}"

    def fac_ref(dept, slot, r_for_name, name_col):
        info = sheet_info.get(("FAC", slot))
        if info is None:
            return 0
        sh, last = _q(info["sheet"]), info["last"]
        return f"=SUMIFS({sh}!$H${FIRST_DATA_ROW}:$H${last},{sh}!$E${FIRST_DATA_ROW}:$E${last},${name_col}{r_for_name})"

    first = SUMMARY_FIRST_ROW
    for i, dept in enumerate(departments):
        r = first + i
        b.set(r, 2, dept, "body")
        b.set(r, 3, treaty_ref(dept, cip.SLOT_CURRENT), "body_num")
        b.set(r, 4, fac_ref(dept, cip.SLOT_CURRENT, r, "B"), "body_num")
        b.set(r, 5, f"=C{r}+D{r}", "bold_num")
        b.set(r, 6, dept, "body")
        b.set(r, 7, treaty_ref(dept, cip.SLOT_PREVIOUS), "body_num")
        b.set(r, 8, fac_ref(dept, cip.SLOT_PREVIOUS, r, "F"), "body_num")
        b.set(r, 9, f"=G{r}+H{r}", "bold_num")
    tr = first + len(departments)
    b.set(tr, 2, cip.TOTAL_LABEL, "total_label"); b.set(tr, 6, cip.TOTAL_LABEL, "total_label")
    for c in (3, 4, 5, 7, 8, 9):
        L = _L(c)
        b.set(tr, c, f"=SUM({L}{first}:{L}{tr - 1})", "total_num")

    r = tr + 2
    b.set(r, 2, "RI Commission Income", "body"); b.set(r, 3, f"=E{tr}", "body_num")
    income_row = r
    r += 1
    b.set(r, 2, "Add: RI Unearned Commission Opening", "body"); b.set(r, 3, float(opening or 0.0), "input_num")
    opening_row = r
    r += 1
    unearned_total = f"{_q(WORKING_SHEET)}!{_L(total_col)}{anchors['unearned']['total']}"
    b.set(r, 2, "Less: RI Unearned Commission Closing", "body"); b.set(r, 3, f"=-{unearned_total}", "body_num")
    closing_row = r
    r += 1
    b.set(r, 2, "RI Commission Earned", "total_label")
    b.set(r, 3, f"=C{income_row}+C{opening_row}+C{closing_row}", "total_num")
    r += 2

    # Tie-out: Treaty Rates sheet commission vs the detail sheets
    rates_total = f"{_q(RATES_SHEET)}!J{rates_info['total_row']}"
    detail_total = f"C{tr}+G{tr}"
    b.set(r, 2, "Check: Treaty Rates commission less register detail", "body")
    b.set(r, 3, f"={rates_total}-({detail_total})", "body_num2")
    b.set(r, 4, "should be zero", "note")
    r += 2

    # Every difference on the Check sheet, as one number (absolute values so
    # opposite differences cannot cancel).
    if check_cells:
        total_diff = "=" + "+".join(f"ABS({_q(CHECK_SHEET)}!{c})" for c in check_cells)
    else:
        total_diff = 0
    b.set(r, 2, "Check: all differences on the 'Check' sheet", "body")
    b.set(r, 3, total_diff, "body_num2")
    b.set(r, 4, "should be zero", "note")
    r += 2

    b.set(r, 2, "Notes", "section"); r += 1
    notes = [
        f"Period: registers are bucketed against {ps:%d-%b-%Y} - {pe:%d-%b-%Y}; a Comm Date before the period is 'Previous', after it 'Future Policies'.",
        "Treaty commission = Treaty Premium x the Treaty Description's % (see 'Treaty Rates'). FAC commission is the Comm amt already on the FAC register.",
        "Marine FAC lines use the Acceptance Date (the FAC register has no commencement date for Marine).",
        "RI Unearned Commission uses the 1/24 method (1/6 for Marine and Travel, nil for BIMA) - same as UPR Automation.",
    ]
    for w in result.get("warnings", []):
        notes.append(f"Warning: {w}")
    for n in notes:
        b.set(r, 2, n, "note"); r += 1

    for c, w in {1: 2, 2: 40, 3: 20, 4: 20, 5: 20, 6: 22, 7: 20, 8: 20, 9: 20}.items():
        b.widths[c] = w
    return b



def _build_check_buf(result, sheet_info):
    """Check sheet, laid out like the reference workbook's: per
    department, one line per Period with Treaty / FAC / Total commission
    (each summed straight from the detail sheets, independently of the
    Working sheet), a Total line, the matching Summary figure and a
    Check (difference) line for each of the three columns. Done for the
    current-year registers, then the previous-year registers. Returns
    (buffer, [addresses of every difference cell])."""
    b = _Buf()
    cols = cip.TABLE_COLUMNS
    departments = result["departments"]
    cur_y, prev_y = result["year_labels"][cip.SLOT_CURRENT], result["year_labels"][cip.SLOT_PREVIOUS]
    # Standard register departments always; any other department only if it carries commission.
    shown = []
    for i, dept in enumerate(departments):
        has_value = any(abs(result["income"][sl][dept]["total"]) > 0.005 for sl in cip.SLOTS)
        if dept in cip.DEPARTMENTS or has_value:
            shown.append((i, dept))

    b.set(1, 2, "Check", "title")
    b.set(2, 2, "Each department's commission by Period, summed straight from the detail sheets and tied back to the Summary. "
                 "Every Check line should be zero.", "subtitle")
    check_cells = []
    r = 4
    # slot -> (year, Summary columns for treaty / FAC / total)
    for slot, year, summ_cols in ((cip.SLOT_CURRENT, cur_y, ("C", "D", "E")),
                                  (cip.SLOT_PREVIOUS, prev_y, ("G", "H", "I"))):
        b.set(r, 2, f"{year} registers", "section")
        r += 1
        for j, h in enumerate(["Period", "Commission Income Treaty", "Commission Income FAC", "Total Commission"], start=2):
            b.set(r, j, h, "header_left" if j == 2 else "header")
        r += 1
        for i, dept in shown:
            b.set(r, 2, dept, "section")
            r += 1
            first = r
            tinfo = sheet_info.get((dept, slot))
            finfo = sheet_info.get(("FAC", slot))
            for col in cols:
                b.set(r, 2, col, "body")
                if tinfo is None:
                    b.set(r, 3, 0, "body_num")
                else:
                    sh, lay, last = _q(tinfo["sheet"]), tinfo["layout"], tinfo["last"]
                    tot, per = _L(lay["total"]), _L(lay["period"])
                    b.set(r, 3, f"=SUMIFS({sh}!${tot}${FIRST_DATA_ROW}:${tot}${last},"
                                f"{sh}!${per}${FIRST_DATA_ROW}:${per}${last},$B{r})", "body_num")
                if finfo is None:
                    b.set(r, 4, 0, "body_num")
                else:
                    sh, last = _q(finfo["sheet"]), finfo["last"]
                    b.set(r, 4, f"=SUMIFS({sh}!$H${FIRST_DATA_ROW}:$H${last},{sh}!$E${FIRST_DATA_ROW}:$E${last},\"{dept}\","
                                f"{sh}!$I${FIRST_DATA_ROW}:$I${last},$B{r})", "body_num")
                b.set(r, 5, f"=C{r}+D{r}", "bold_num")
                r += 1
            last_row = r - 1
            b.set(r, 2, cip.TOTAL_LABEL, "total_label")
            for c in (3, 4, 5):
                L = _L(c)
                b.set(r, c, f"=SUM({L}{first}:{L}{last_row})", "total_num")
            total_row = r
            r += 1
            b.set(r, 2, "Per Summary", "body")
            for c, sc in zip((3, 4, 5), summ_cols):
                b.set(r, c, f"={_q(SUMMARY_SHEET)}!{sc}{SUMMARY_FIRST_ROW + i}", "body_num")
            summary_row = r
            r += 1
            b.set(r, 2, "Check", "bold_num")
            for c in (3, 4, 5):
                L = _L(c)
                b.set(r, c, f"={L}{total_row}-{L}{summary_row}", "body_num2")
                check_cells.append(f"{L}{r}")
            r += 2
        r += 1
    for c, w in {1: 2, 2: 26, 3: 24, 4: 24, 5: 24}.items():
        b.widths[c] = w
    return b, check_cells

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def build_commission_income_workbook(result, opening_unearned_commission, output_path, report=None):
    """result: dict from commission_income_processor.run_commission_income_calculation.
    opening_unearned_commission: prior period's closing RI Unearned
    Commission, entered by hand (same 'manual for now' pattern as UPR /
    Deferred Wakala Fee). Returns output_path."""
    log = report or (lambda m: None)
    wb = Workbook(write_only=True)

    departments = result["departments"]
    summary_rows = result["treaty_summary"]

    # Treaty Rates row numbers - fixed up front so every detail sheet can
    # point its % cells at them.
    rate_rows, rate_row_by_key = [], {}
    r = 6
    for row in summary_rows:
        if row["key"] == "" and abs(row["current_premium"]) + abs(row["previous_premium"]) < 0.005:
            continue
        rate_rows.append((row, r))
        rate_row_by_key[(row["department"], row["key"])] = r
        r += 1

    # Layout (sheet names, column positions, last data row) of every detail
    # sheet - the small sheets' formulas point at these coordinates.
    sheet_info = {}
    for (dept, slot), df in result["detail"].items():
        keys = _register_keys(df, summary_rows, dept)
        layout = _detail_layout(keys)
        last = FIRST_DATA_ROW + len(df) - 1 if len(df) else FIRST_DATA_ROW
        sheet_info[(dept, slot)] = {"sheet": _sheet_name(dept, result["year_labels"][slot]),
                                     "layout": layout, "keys": keys, "last": last}
    for slot, df in result["fac_detail"].items():
        last = FIRST_DATA_ROW + len(df) - 1 if len(df) else FIRST_DATA_ROW
        sheet_info[("FAC", slot)] = {"sheet": _fac_sheet_name(result["year_labels"][slot]), "last": last}

    # Sheet order in a write-only workbook = creation order. Summary and
    # Working must come first in the file but only depend on layout
    # (coordinates), not on the detail data - so every small sheet is
    # BUILT first into a buffer, then the sheets are created and
    # flushed in the order they should appear.
    working_buf, anchors, total_col = _build_working_buf(result, sheet_info)
    rates_total_row = (rate_rows[-1][1] if rate_rows else 6) + 1
    rates_buf = _build_rates_buf(result, rate_rows)
    check_buf, check_cells = _build_check_buf(result, sheet_info)
    summary_buf = _build_summary_buf(result, sheet_info, anchors, total_col, {"total_row": rates_total_row},
                                      opening_unearned_commission, check_cells)
    log("Writing Summary, Working, Check and Treaty Rates sheets...")
    summary_buf.flush(wb.create_sheet(SUMMARY_SHEET))
    working_buf.flush(wb.create_sheet(WORKING_SHEET))
    check_buf.flush(wb.create_sheet(CHECK_SHEET))
    rates_buf.flush(wb.create_sheet(RATES_SHEET))

    for slot in (cip.SLOT_CURRENT, cip.SLOT_PREVIOUS):
        for dept in departments:
            df = result["detail"].get((dept, slot))
            if df is None:
                continue
            info = sheet_info[(dept, slot)]
            log(f"Writing {info['sheet']} ({len(df):,} rows)...")
            _write_register_sheet(wb, result, dept, slot, df, info["layout"], info["keys"], rate_row_by_key)
        fac_df = result["fac_detail"].get(slot)
        if fac_df is not None:
            log(f"Writing {sheet_info[('FAC', slot)]['sheet']} ({len(fac_df):,} rows)...")
            _write_fac_sheet(wb, result, slot, fac_df)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)
    _recalculate(output_path)
    return output_path


def _recalculate(filepath):
    """Best-effort LibreOffice recalculation so formula cells carry cached
    values on open - same helper pattern as the other generators."""
    recalc_script = "/mnt/skills/public/xlsx/scripts/recalc.py"
    if not os.path.exists(recalc_script):
        return
    try:
        subprocess.run([sys.executable, recalc_script, filepath, "180"], capture_output=True, timeout=240)
    except Exception:
        pass


def build_output_filename(period_end, run_date=None):
    run_date = run_date or date.today()
    return f"Commission_Income_{period_end:%b%Y}_{run_date:%d%m%Y}.xlsx"
