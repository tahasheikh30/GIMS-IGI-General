"""
Writes the Trading Reports workbook: two sheets for either working.

  Summary        the totals - per bank for a STOCK working, per stock for a
                 BANK working - with a Total line.
  Transactions   every trade behind the summary, sorted by trade date, with
                 the bank that performed it (stock working) or the account
                 header (bank working), and a Total line.

A workbook for SEVERAL stocks (kind "multi") keeps the same two sheets: the
Summary shows each stock's banks with a total line per stock and a grand
total of the amounts (share quantities of different stocks are never added
together); the Transactions sheet lists every trade, grouped by stock.

All figures are written as plain VALUES, not formulas, on purpose: these
files are e-mailed to the banks, and an Excel attachment opened in Protected
View (or previewed on a phone) does not calculate formulas, so formula cells
would show empty or zero until the reader clicked Enable Editing.
"""
import os
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import trading_reports_processor as trp

NAVY = "1F4E78"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOTAL_BORDER = Border(top=Side(style="thin", color=NAVY), left=THIN, right=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", start_color=NAVY, end_color=NAVY)
TOTAL_FILL = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
FLAG_FILL = PatternFill("solid", start_color="FCE4D6", end_color="FCE4D6")

QTY = "#,##0;(#,##0);-"
AMT = "#,##0.00;(#,##0.00);-"
RATE = "#,##0.0000;(#,##0.0000);-"
DATE_FMT = "dd-mmm-yyyy"


def _header(ws, row, labels, start_col=1):
    for j, text in enumerate(labels, start=start_col):
        c = ws.cell(row=row, column=j, value=text)
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32


def _title_block(ws, working, subtitle):
    ws["A1"] = "IGI General Insurance Limited"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color=NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = Font(name="Arial", bold=True, size=12, color=NAVY)


def _put(ws, row, col, value, fmt=None, bold=False, fill=None, border=BORDER, align=None, color=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Arial", bold=bold, color=color)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = border
    if align:
        c.alignment = Alignment(horizontal=align)
    return c


def _isnull(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT or (not isinstance(v, str) and pd.isna(v))


def _summary_columns(side):
    """[(header, key, fmt)] for the summary's numeric columns."""
    buy = [("Buy Qty", "buy_qty", QTY), ("Buy Amount", "buy_amount", AMT), ("Avg Buy Rate", "buy_rate", RATE)]
    sell = [("Sell Qty", "sell_qty", QTY), ("Sell Amount", "sell_amount", AMT), ("Avg Sell Rate", "sell_rate", RATE)]
    net_qty = [("Net Qty (Bought - Sold)", "net_qty", QTY)]
    tail = [("Commission", "comm", AMT), ("Sales Tax", "tax", AMT), ("Net Amount", "net", AMT), ("Trades", "trades", "0")]
    if side == trp.SIDE_BUY:
        return buy + tail
    if side == trp.SIDE_SELL:
        return sell + tail
    return buy + sell + net_qty + tail


def _write_summary_multi(wb, working):
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    side, meta = working["side"], working["meta"]
    _title_block(ws, working, "Trading Summary - " + trp.working_title(working))
    info = [("Stocks", ", ".join(f"{w['symbol']} - {w['company']}" for w in working["stocks"])),
            (meta["date_label"], trp._period_text(meta)), ("Transactions", side),
            ("No. of transactions", working["total"]["trades"]), ("Prepared on", f"{date.today():%d-%b-%Y}")]
    if working["skipped"]:
        info.append(("Not included", ", ".join(f"{sym} (no matching trades)" for sym, _why in working["skipped"])))
    r = 3
    for label, value in info:
        _put(ws, r, 1, label, bold=True, border=None, color="6C7A89")
        _put(ws, r, 2, value, border=None, align="left")
        r += 1

    num_cols = _summary_columns(side)
    header_row = r + 1
    _header(ws, header_row, ["Symbol", "Stock Name", "Bank", "Account"] + [h for h, _k, _f in num_cols])
    row = header_row + 1
    amount_keys = {"buy_amount", "sell_amount", "comm", "tax", "net", "trades"}
    for stock in working["stocks"]:
        for line in stock["rows"]:
            if not line["trades"]:
                continue     # banks that did not trade this stock are left out of the multi-stock summary
            _put(ws, row, 1, stock["symbol"])
            _put(ws, row, 2, stock["company"])
            _put(ws, row, 3, line["bank"])
            _put(ws, row, 4, line["code"])
            for j, (_h, key, fmt) in enumerate(num_cols, start=5):
                v = line[key]
                _put(ws, row, j, None if _isnull(v) else v, fmt=fmt)
            row += 1
        t = stock["total"]
        _put(ws, row, 1, f"Total {stock['symbol']}", bold=True, fill=TOTAL_FILL, border=TOTAL_BORDER)
        for c in (2, 3, 4):
            _put(ws, row, c, None, fill=TOTAL_FILL, border=TOTAL_BORDER)
        for j, (_h, key, fmt) in enumerate(num_cols, start=5):
            v = t[key]
            _put(ws, row, j, None if _isnull(v) else v, fmt=fmt, bold=True, fill=TOTAL_FILL, border=TOTAL_BORDER)
        row += 1
    grand = working["total"]
    navy = dict(bold=True, fill=HEADER_FILL, border=TOTAL_BORDER, color="FFFFFF")
    _put(ws, row, 1, "Grand total (amounts)", **navy)
    for c in (2, 3, 4):
        _put(ws, row, c, None, fill=HEADER_FILL, border=TOTAL_BORDER)
    for j, (_h, key, fmt) in enumerate(num_cols, start=5):
        v = grand[key] if key in amount_keys else None      # quantities of different stocks are not added together
        _put(ws, row, j, None if _isnull(v) else v, fmt=fmt, **navy)
    row += 2

    notes = ["Net Amount follows the broker report: positive is cash paid on buys, negative is cash received on sells.",
             "Share quantities and average rates are per stock; only amounts are added across stocks in the grand total."]
    if meta["flagged"]:
        notes.append(f"{meta['flagged']} transaction(s) are flagged in the Flag column of the Transactions sheet (figures are as reported).")
    for n in notes:
        _put(ws, row, 1, n, border=None).font = Font(name="Arial", italic=True, size=9, color="6C7A89")
        row += 1
    for j in range(1, 5 + len(num_cols)):
        ws.column_dimensions[get_column_letter(j)].width = {1: 22, 2: 34, 3: 16, 4: 10}.get(j, 17)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _write_summary(wb, working):
    if working["kind"] == "multi":
        return _write_summary_multi(wb, working)
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    kind, side, meta = working["kind"], working["side"], working["meta"]
    _title_block(ws, working, "Trading Summary - " + trp.working_title(working))

    info = [(meta["date_label"], trp._period_text(meta)), ("Transactions", side),
            ("No. of transactions", working["total"]["trades"]), ("Prepared on", f"{date.today():%d-%b-%Y}")]
    if kind == "stock":
        info.insert(0, ("Stock", trp.working_title(working)))
    else:
        info.insert(0, ("Account", trp.working_title(working)))
    r = 3
    for label, value in info:
        _put(ws, r, 1, label, bold=True, border=None, color="6C7A89")
        _put(ws, r, 2, value, border=None, align="left")
        r += 1

    first_cols = ["Bank", "Account"] if kind == "stock" else ["Symbol", "Stock Name"]
    num_cols = _summary_columns(side)
    header_row = r + 1
    _header(ws, header_row, first_cols + [h for h, _k, _f in num_cols])

    row = header_row + 1
    for line in working["rows"]:
        if kind == "stock":
            _put(ws, row, 1, line["bank"])
            _put(ws, row, 2, line["code"])
        else:
            _put(ws, row, 1, line["symbol"])
            _put(ws, row, 2, line["company"])
        for j, (_h, key, fmt) in enumerate(num_cols, start=3):
            v = line[key]
            _put(ws, row, j, None if _isnull(v) else v, fmt=fmt)
        row += 1

    total = working["total"]
    _put(ws, row, 1, "Total", bold=True, fill=TOTAL_FILL, border=TOTAL_BORDER)
    _put(ws, row, 2, None, fill=TOTAL_FILL, border=TOTAL_BORDER)
    for j, (_h, key, fmt) in enumerate(num_cols, start=3):
        v = total[key]
        _put(ws, row, j, None if _isnull(v) else v, fmt=fmt, bold=True, fill=TOTAL_FILL, border=TOTAL_BORDER)
    row += 2

    notes = ["Net Amount follows the broker report: positive is cash paid on buys, negative is cash received on sells."]
    if meta["flagged"]:
        notes.append(f"{meta['flagged']} transaction(s) are flagged in the Flag column of the Transactions sheet "
                     f"(figures are as reported).")
    for n in notes:
        _put(ws, row, 1, n, border=None, color="6C7A89").font = Font(name="Arial", italic=True, size=9, color="6C7A89")
        row += 1

    widths = {1: 22, 2: 34 if kind == "bank" else 12}
    for j in range(1, 3 + len(num_cols)):
        ws.column_dimensions[get_column_letter(j)].width = widths.get(j, 17)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _write_transactions(wb, working):
    ws = wb.create_sheet("Transactions")
    ws.sheet_view.showGridLines = False
    kind, meta = working["kind"], working["meta"]
    tx = working["transactions"]
    _title_block(ws, working, "Transactions - " + trp.working_title(working))
    ws["A3"] = (f"{meta['date_label']} {trp._period_text(meta)}  \u2022  {working['side']}  \u2022  {len(tx):,} transaction(s)  \u2022  "
                + ("grouped by stock, then trade date" if kind == "multi" else "sorted by trade date"))
    ws["A3"].font = Font(name="Arial", size=9, italic=True, color="6C7A89")

    # (header, DataFrame column, format, width)
    common = [("Buy Qty", "Buy Qty", QTY, 12), ("Sell Qty", "Sell Qty", QTY, 12), ("Rate", "Rate", RATE, 14),
              ("Buy Amount", "Buy Amount", AMT, 16), ("Sell Amount", "Sell Amount", AMT, 16),
              ("Comm Rate", "Comm Rate", RATE, 12), ("Comm Amount", "Comm Amount", AMT, 14),
              ("Sales Tax", "Sales Tax", AMT, 12), ("Net Amount", "Net Amount", AMT, 16)]
    if kind == "multi":
        lead = [("Symbol", "Symbol", None, 10), ("Stock Name", "Stock Name", None, 34), ("Bank", "Bank", None, 16),
                ("Account", "Account Code", None, 10), ("Trade Date", "Trade Date", DATE_FMT, 13),
                ("Settlement Date", "Settlement Date", DATE_FMT, 13)]
    elif kind == "stock":
        lead = [("Bank", "Bank", None, 16), ("Account", "Account Code", None, 10), ("Symbol", "Symbol", None, 10),
                ("Stock Name", "Stock Name", None, 34), ("Trade Date", "Trade Date", DATE_FMT, 13),
                ("Settlement Date", "Settlement Date", DATE_FMT, 13)]
    else:
        lead = [("Trade Date", "Trade Date", DATE_FMT, 13), ("Settlement Date", "Settlement Date", DATE_FMT, 13),
                ("Symbol", "Symbol", None, 10), ("Stock Name", "Stock Name", None, 34)]
    columns = lead + common
    show_flag = bool((tx["Flag"] != "").any())
    if show_flag:
        columns = columns + [("Flag", "Flag", None, 30)]

    header_row = 5
    _header(ws, header_row, [h for h, _c, _f, _w in columns])
    for j, (_h, _c, _f, w) in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    r = header_row + 1
    for rec in tx.to_dict("records"):
        flagged = bool(rec.get("Flag"))
        for j, (_h, col, fmt, _w) in enumerate(columns, start=1):
            v = rec.get(col)
            if _isnull(v):
                v = None
            elif isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            _put(ws, r, j, v, fmt=fmt, fill=FLAG_FILL if flagged else None)
        r += 1
    last = r - 1

    # Total line
    total_cols = {"Buy Qty", "Sell Qty", "Buy Amount", "Sell Amount", "Comm Amount", "Sales Tax", "Net Amount"}
    if kind == "multi":
        total_cols -= {"Buy Qty", "Sell Qty"}   # quantities of different stocks are not added together
    label_col = 2 if kind == "multi" else 4     # the stock-name column carries the word "Total"
    for j, (_h, col, fmt, _w) in enumerate(columns, start=1):
        if col in total_cols:
            _put(ws, r, j, float(tx[col].sum()), fmt=fmt, bold=True, fill=TOTAL_FILL, border=TOTAL_BORDER)
        else:
            _put(ws, r, j, "Total" if j == label_col else None, bold=True, fill=TOTAL_FILL, border=TOTAL_BORDER)

    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(columns))}{last}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{header_row}:{header_row}"


def write_trading_workbook(working, output_path):
    """Write the two-sheet workbook for a stock or bank working."""
    wb = Workbook()
    _write_summary(wb, working)
    _write_transactions(wb, working)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)
    return output_path
