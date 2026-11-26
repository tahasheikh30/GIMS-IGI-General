"""
Trading Reports (Finance > Trading > Trading Reports).

Reads the broker trading-account reports - one report per trading account,
all with the same layout - for the company's accounts managed by different
fund managers, and produces two kinds of working:

  1. STOCK working  - one chosen stock across ALL the accounts: how many
     shares each bank bought (and sold), the amounts, average rates,
     commission, sales tax and net amount, plus every transaction of that
     stock with the bank that performed it.
  2. BANK working   - one chosen bank/account: its transactions summarised
     by stock plus the full transaction list, to send to that bank.

REPORT LAYOUT (one row per trade, headers as in the Finex trading account
reports): CLIENT NAME, COMPANY (stock name), SYMBOL, TD (trade date), SD
(settlement date), BUY QTY, SELL QTY, RATE, BUY AMOUNT, SELL AMOUNT, COMM
RATE, COMM AMOUNT, SST (sales tax), NET AMOUNT. Every row is either a buy
(BUY QTY / BUY AMOUNT filled) or a sell (SELL QTY / SELL AMOUNT filled).
Net Amount is the cash payable on the trade: positive for a buy (amount +
commission + tax), negative for a sell (-(amount - commission - tax)).
Any extra unnamed columns some of the reports carry are ignored.

WHICH BANK IS WHICH: the account code sits in each row's CLIENT NAME, for
example "IGI GENERAL INSURANCE LIMITED (22302)". ACCOUNTS maps those codes to
the fund manager. (If a row has no code in its client name, the code is taken
from the sheet name, then from the file name.) A code not listed in ACCOUNTS
is still processed, shown as "Account <code>", and reported as a warning.

CHECKS (reported, never silently "fixed" - the report's own figures are
always what is shown):
  * Net Amount that does not reconcile to amount +/- commission +/- tax
  * a negative rate or negative amount on a trade
  * a row with no buy/sell quantity, or with both
The same account appearing in two different files is an error, because
adding the same report twice would double-count every trade.

DATE FILTER: a month's report can be narrowed to a period or a single day on
the trade date (or settlement date) before either working is built -
apply_date_filter(). MULTIPLE STOCKS: build_multi_stock_working() combines
several stocks into one workbook; the screen can alternatively write one
workbook per stock.
"""
import os
import re

import pandas as pd

# Account code (as it appears in the report's client name) -> fund manager,
# in the order the accounts are listed throughout the module.
ACCOUNTS = [
    ("22994", "UBL Funds"),
    ("22402", "Alfalah Funds"),
    ("22403", "MCB Funds"),
    ("22302", "NBP Funds"),
]
ACCOUNT_NAMES = dict(ACCOUNTS)
ACCOUNT_ORDER = {code: i for i, (code, _n) in enumerate(ACCOUNTS)}

SIDE_BOTH = "Buys and sells"
SIDE_BUY = "Buys only"
SIDE_SELL = "Sells only"
SIDE_OPTIONS = [SIDE_BOTH, SIDE_BUY, SIDE_SELL]

NET_TOLERANCE = 0.5      # PKR - Net Amount vs amount +/- commission +/- tax

ALIASES = {
    "Client Name": ["client name", "client", "account name"],
    "Stock Name": ["company", "stock name", "stock", "company name", "scrip name"],
    "Symbol": ["symbol", "stock symbol", "scrip", "ticker"],
    "Trade Date": ["td", "trade date", "trade dt"],
    "Settlement Date": ["sd", "settlement date", "settle date", "settlement dt"],
    "Buy Qty": ["buy qty", "buy quantity", "buy qty."],
    "Sell Qty": ["sell qty", "sell quantity", "sell qty."],
    "Rate": ["rate", "price"],
    "Buy Amount": ["buy amount", "buy amt"],
    "Sell Amount": ["sell amount", "sell amt"],
    "Comm Rate": ["comm rate", "commission rate", "comm. rate"],
    "Comm Amount": ["comm amount", "commission amount", "comm amt", "commission"],
    "Sales Tax": ["sst", "sales tax", "sales tax amount", "sst amount"],
    "Net Amount": ["net amount", "net amt", "net"],
}
REQUIRED = ["Stock Name", "Symbol", "Trade Date", "Settlement Date", "Buy Qty", "Sell Qty", "Rate",
            "Buy Amount", "Sell Amount", "Comm Rate", "Comm Amount", "Sales Tax", "Net Amount"]
NUMERIC = ["Buy Qty", "Sell Qty", "Rate", "Buy Amount", "Sell Amount", "Comm Rate", "Comm Amount",
           "Sales Tax", "Net Amount"]


class TradingReportsError(Exception):
    """A problem that should stop the run with a clear, actionable message."""
    pass


def _norm(value):
    return re.sub(r"\s+", " ", str(value).replace("\n", " ")).strip().lower()


_LOOKUP = {}
for _canon, _names in ALIASES.items():
    for _a in _names:
        _LOOKUP[_norm(_a)] = _canon


def _find_header_row(preview, max_scan=25, min_score=6):
    best, best_score = None, 0
    for i in range(min(max_scan, len(preview))):
        seen = set()
        for v in preview.iloc[i].tolist():
            if pd.isna(v):
                continue
            canon = _LOOKUP.get(_norm(v))
            if canon:
                seen.add(canon)
        if len(seen) > best_score:
            best, best_score = i, len(seen)
    return (best if best_score >= min_score else None), best_score


def _bank_name(code):
    return ACCOUNT_NAMES.get(code) or f"Account {code}"


def _code_from_text(text):
    """Last bracketed digit group in the text, e.g. '... LIMITED (22302) '."""
    found = re.findall(r"\((\d{3,})\)", str(text))
    return found[-1] if found else None


def load_report(path):
    """Load one trading report. Returns (df, meta).

    df columns: Account Code, Bank, Symbol, Stock Name, Trade Date,
    Settlement Date, Type (Buy/Sell), Buy Qty, Sell Qty, Rate, Buy Amount,
    Sell Amount, Comm Rate, Comm Amount, Sales Tax, Net Amount, Flag,
    Source File.
    """
    fname = os.path.basename(path)
    if not os.path.exists(path):
        raise TradingReportsError(f"File not found: {path}")
    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        raise TradingReportsError(f"Could not open '{fname}': {e}")

    best = None
    for sn in xls.sheet_names:
        preview = pd.read_excel(xls, sheet_name=sn, header=None, nrows=25)
        header_row, score = _find_header_row(preview)
        if header_row is not None and (best is None or score > best[0]):
            best = (score, sn, header_row)
    if best is None:
        raise TradingReportsError(
            f"'{fname}' does not look like a trading account report: no sheet has a header row with the expected "
            f"columns (Symbol, TD, SD, Buy Qty, Sell Qty, Rate, ...).")
    _score, sheet, header_row = best
    raw = pd.read_excel(xls, sheet_name=sheet, header=header_row)

    cols = {}
    for col in raw.columns:
        canon = _LOOKUP.get(_norm(col))
        if canon and canon not in cols:
            cols[canon] = col
    missing = [c for c in REQUIRED if c not in cols]
    if missing:
        raise TradingReportsError(
            f"'{fname}' is missing required column(s) {missing}. Columns found: {[str(c) for c in raw.columns]}.")

    out = pd.DataFrame(index=raw.index)
    out["Stock Name"] = raw[cols["Stock Name"]].astype("string").str.strip()
    out["Symbol"] = raw[cols["Symbol"]].astype("string").str.strip().str.upper()
    out["Trade Date"] = pd.to_datetime(raw[cols["Trade Date"]], errors="coerce")
    out["Settlement Date"] = pd.to_datetime(raw[cols["Settlement Date"]], errors="coerce")
    for c in NUMERIC:
        out[c] = pd.to_numeric(raw[cols[c]], errors="coerce")

    # Account code per row: client name -> sheet name -> file name.
    fallback = None
    for text in (sheet, fname):
        m = re.findall(r"(\d{4,})", str(text))
        if m:
            fallback = m[-1]
            break
    if "Client Name" in cols:
        out["Account Code"] = raw[cols["Client Name"]].map(_code_from_text)
    else:
        out["Account Code"] = None
    out["Account Code"] = out["Account Code"].where(out["Account Code"].notna(), fallback)

    # Real trade rows only: a symbol and some quantity or amount. (Blank
    # spacer / footer rows have neither.)
    has_symbol = out["Symbol"].notna() & (out["Symbol"] != "")
    has_activity = out[["Buy Qty", "Sell Qty", "Buy Amount", "Sell Amount"]].notna().any(axis=1)
    out = out.loc[has_symbol & has_activity].copy()
    if out.empty:
        raise TradingReportsError(f"'{fname}' has no trade rows.")
    if out["Account Code"].isna().any():
        raise TradingReportsError(
            f"'{fname}': could not work out the account code (no '(12345)' in the client name, sheet name or file name).")

    buy = out["Buy Qty"].notna() & (out["Buy Qty"] != 0)
    sell = out["Sell Qty"].notna() & (out["Sell Qty"] != 0)
    out["Type"] = None
    out.loc[buy & ~sell, "Type"] = "Buy"
    out.loc[sell & ~buy, "Type"] = "Sell"
    out.loc[buy & sell, "Type"] = "Buy"          # flagged below
    out["Type"] = out["Type"].where(out["Type"].notna(), "Buy")

    # ---- row checks (reported; figures are never altered) ----
    flags = pd.Series("", index=out.index, dtype=object)

    def add(mask, text):
        nonlocal flags
        flags = flags.where(~mask, flags.where(flags == "", flags + "; ") + text)

    add(buy & sell, "Both buy and sell quantity")
    add(~buy & ~sell, "No quantity")
    neg = (out["Rate"] < 0) | (out["Buy Amount"] < 0) | (out["Sell Amount"] < 0)
    add(neg, "Negative rate/amount")
    is_buy = out["Type"] == "Buy"
    amount = out["Buy Amount"].fillna(0).where(is_buy, out["Sell Amount"].fillna(0))
    comm = out["Comm Amount"].fillna(0)
    tax = out["Sales Tax"].fillna(0)
    expected = (amount + comm + tax).where(is_buy, -(amount - comm - tax))
    add((out["Net Amount"].fillna(0) - expected).abs() > NET_TOLERANCE, "Net amount does not reconcile")
    out["Flag"] = flags

    out["Bank"] = out["Account Code"].map(_bank_name)
    out["Source File"] = fname
    out = out[["Account Code", "Bank", "Symbol", "Stock Name", "Trade Date", "Settlement Date", "Type"] + NUMERIC
              + ["Flag", "Source File"]].reset_index(drop=True)
    meta = {
        "sheet": sheet,
        "rows": len(out),
        "codes": sorted(out["Account Code"].unique()),
        "date_from": out["Trade Date"].min(),
        "date_to": out["Trade Date"].max(),
        "flagged": int((out["Flag"] != "").sum()),
    }
    return out, meta


def load_reports(paths, report=None):
    """Load every report. Returns a bundle:
        {"data": DataFrame (all accounts), "files": [per-file info],
         "warnings": [str]}
    Raises TradingReportsError when an account appears in two files."""
    if not paths:
        raise TradingReportsError("No report files were selected.")
    frames, files, warnings = [], [], []
    owner = {}
    for path in paths:
        if report:
            report(f"Reading '{os.path.basename(path)}'...")
        df, meta = load_report(path)
        for code in meta["codes"]:
            if code in owner:
                raise TradingReportsError(
                    f"Account {code} ({_bank_name(code)}) is in both '{owner[code]}' and '{os.path.basename(path)}'. "
                    f"Each account's report should be loaded once - loading it twice would double-count every trade. "
                    f"Remove one of the files.")
            owner[code] = os.path.basename(path)
        frames.append(df)
        for code in meta["codes"]:
            part = df[df["Account Code"] == code]
            files.append({
                "file": os.path.basename(path), "code": code, "bank": _bank_name(code), "rows": len(part),
                "date_from": part["Trade Date"].min(), "date_to": part["Trade Date"].max(),
                "flagged": int((part["Flag"] != "").sum()),
            })
            if code not in ACCOUNT_NAMES:
                warnings.append(f"Account {code} (in '{os.path.basename(path)}') is not in the known account list - shown as "
                                f"'Account {code}'. Add it to ACCOUNTS in trading_reports_processor.py to give it a bank name.")
        if report:
            report(f"  {len(df):,} trade(s) read"
                   + (f"; {meta['flagged']} flagged." if meta["flagged"] else "."))
    files.sort(key=lambda f: (ACCOUNT_ORDER.get(f["code"], 99), f["code"]))
    data = pd.concat(frames, ignore_index=True)
    data["_order"] = data["Account Code"].map(lambda c: ACCOUNT_ORDER.get(c, 99))
    data = data.sort_values(["Trade Date", "_order", "Settlement Date"], kind="stable").drop(columns="_order").reset_index(drop=True)

    flagged = data[data["Flag"] != ""]
    if len(flagged):
        by = flagged.groupby("Flag").size()
        for text, n in by.items():
            warnings.append(f"{n} trade(s) flagged '{text}' - kept as reported; see the Flag column in the output.")
    return {"data": data, "files": files, "warnings": warnings}


# ---------------------------------------------------------------------------
# Date filter
# ---------------------------------------------------------------------------
DATE_FIELDS = ["Trade Date", "Settlement Date"]
_DATE_FORMATS = ["%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y", "%d %b %Y", "%d-%B-%Y", "%d %B %Y", "%d.%m.%Y"]


def parse_filter_date(text):
    """Read a date typed on the screen. Blank -> None (no limit). Day comes
    before month (30-10-2026 = 30 October); YYYY-MM-DD and 30-Oct-2026 work too."""
    text = (text or "").strip()
    if not text:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return pd.Timestamp(pd.to_datetime(text, format=fmt))
        except (ValueError, TypeError):
            continue
    raise TradingReportsError(f"Could not read the date '{text}'. Use DD-MM-YYYY (for example 20-10-2026) or YYYY-MM-DD.")


def apply_date_filter(bundle, date_from=None, date_to=None, field="Trade Date"):
    """Return a copy of the bundle limited to trades whose `field` (Trade
    Date or Settlement Date) falls in [date_from, date_to], both inclusive;
    either bound may be None. The returned bundle remembers the filter so
    the workings and workbooks can state the period."""
    if field not in DATE_FIELDS:
        raise TradingReportsError(f"Unknown date field '{field}'.")
    if date_from is None and date_to is None:
        out = dict(bundle)
        out["filter"] = None
        return out
    if date_from is not None and date_to is not None and date_from > date_to:
        raise TradingReportsError("The From date is after the To date.")
    d = bundle["data"]
    keep = pd.Series(True, index=d.index)
    if date_from is not None:
        keep &= d[field] >= date_from
    if date_to is not None:
        keep &= d[field] <= date_to
    filtered = d[keep].reset_index(drop=True)
    lo = f"{date_from:%d-%b-%Y}" if date_from is not None else "the start"
    hi = f"{date_to:%d-%b-%Y}" if date_to is not None else "the end"
    if filtered.empty:
        have = f"{d[field].min():%d-%b-%Y} to {d[field].max():%d-%b-%Y}"
        raise TradingReportsError(f"No trades with a {field.lower()} from {lo} to {hi}. The loaded reports cover {have}.")
    out = dict(bundle)
    out["data"] = filtered
    out["filter"] = {"from": date_from, "to": date_to, "field": field}
    return out


# ---------------------------------------------------------------------------
# Lists for the screen
# ---------------------------------------------------------------------------

def list_stocks(bundle):
    """[{symbol, company, trades, buy_qty, sell_qty}] sorted by symbol."""
    d = bundle["data"]
    rows = []
    for sym, g in d.groupby("Symbol"):
        rows.append({
            "symbol": sym,
            "company": g["Stock Name"].iloc[0],
            "trades": len(g),
            "buy_qty": float(g["Buy Qty"].sum()),
            "sell_qty": float(g["Sell Qty"].sum()),
        })
    return sorted(rows, key=lambda r: r["symbol"])


def list_banks(bundle):
    """[{code, bank, trades}] in the module's account order."""
    d = bundle["data"]
    rows = [{"code": c, "bank": _bank_name(c), "trades": int((d["Account Code"] == c).sum())}
            for c in d["Account Code"].unique()]
    return sorted(rows, key=lambda r: (ACCOUNT_ORDER.get(r["code"], 99), r["code"]))


# ---------------------------------------------------------------------------
# Workings
# ---------------------------------------------------------------------------

def _apply_side(df, side):
    if side == SIDE_BUY:
        return df[df["Type"] == "Buy"]
    if side == SIDE_SELL:
        return df[df["Type"] == "Sell"]
    return df


def _summarize(group):
    """Aggregate a set of trades into one summary line."""
    buy_qty = float(group["Buy Qty"].sum())
    sell_qty = float(group["Sell Qty"].sum())
    buy_amt = float(group["Buy Amount"].sum())
    sell_amt = float(group["Sell Amount"].sum())
    return {
        "buy_qty": buy_qty,
        "buy_amount": buy_amt,
        "buy_rate": (buy_amt / buy_qty) if buy_qty else None,
        "sell_qty": sell_qty,
        "sell_amount": sell_amt,
        "sell_rate": (sell_amt / sell_qty) if sell_qty else None,
        "net_qty": buy_qty - sell_qty,
        "comm": float(group["Comm Amount"].sum()),
        "tax": float(group["Sales Tax"].sum()),
        "net": float(group["Net Amount"].sum()),
        "trades": len(group),
    }


def _meta(df, bundle=None):
    """Dates for headings and file names. `date_from`/`date_to` are the first and
    last TRADE dates in the selection; `period_from`/`period_to` are what the
    report states as its period - the date filter when one is set, otherwise
    the actual span of the field it is describing."""
    flt = (bundle or {}).get("filter") or {}
    field = flt.get("field", "Trade Date")
    actual_from = df[field].min() if len(df) else None
    actual_to = df[field].max() if len(df) else None
    return {
        "date_from": df["Trade Date"].min() if len(df) else None,
        "date_to": df["Trade Date"].max() if len(df) else None,
        "period_from": flt.get("from") or actual_from,
        "period_to": flt.get("to") or actual_to,
        "date_label": "Settlement dates" if field == "Settlement Date" else "Trade dates",
        "flagged": int((df["Flag"] != "").sum()),
    }


def build_stock_working(bundle, symbol, side=SIDE_BOTH):
    """One stock across every account.

    Returns {"kind": "stock", "symbol", "company", "side", "rows": [one per
    account - bank, code + summary], "total": summary, "transactions": df,
    "meta": {...}}.
    """
    d = bundle["data"]
    stock = d[d["Symbol"] == symbol]
    if stock.empty:
        raise TradingReportsError(f"No trades found for {symbol}.")
    selected = _apply_side(stock, side)
    if selected.empty:
        raise TradingReportsError(f"{symbol} has no {side.lower()} in the loaded reports.")
    rows = []
    for code in sorted(d["Account Code"].unique(), key=lambda c: (ACCOUNT_ORDER.get(c, 99), c)):
        part = selected[selected["Account Code"] == code]
        line = _summarize(part)
        line.update({"code": code, "bank": _bank_name(code)})
        rows.append(line)
    tx = selected.copy()
    tx["_o"] = tx["Account Code"].map(lambda c: ACCOUNT_ORDER.get(c, 99))
    tx = tx.sort_values(["Trade Date", "_o", "Settlement Date"], kind="stable").drop(columns="_o").reset_index(drop=True)
    return {
        "kind": "stock", "symbol": symbol, "company": stock["Stock Name"].iloc[0], "side": side,
        "rows": rows, "total": _summarize(selected), "transactions": tx, "meta": _meta(selected, bundle),
    }


def build_bank_working(bundle, code, side=SIDE_BOTH):
    """One account/bank, summarised by stock.

    Returns {"kind": "bank", "code", "bank", "side", "rows": [one per stock -
    symbol, company + summary], "total": summary, "transactions": df,
    "meta": {...}}.
    """
    d = bundle["data"]
    account = d[d["Account Code"] == code]
    if account.empty:
        raise TradingReportsError(f"No trades found for account {code}.")
    selected = _apply_side(account, side)
    if selected.empty:
        raise TradingReportsError(f"{_bank_name(code)} has no {side.lower()} in the loaded reports.")
    rows = []
    for sym, g in selected.groupby("Symbol"):
        line = _summarize(g)
        line.update({"symbol": sym, "company": g["Stock Name"].iloc[0]})
        rows.append(line)
    rows.sort(key=lambda r: r["symbol"])
    tx = selected.sort_values(["Trade Date", "Symbol", "Settlement Date"], kind="stable").reset_index(drop=True)
    return {
        "kind": "bank", "code": code, "bank": _bank_name(code), "side": side,
        "rows": rows, "total": _summarize(selected), "transactions": tx, "meta": _meta(selected, bundle),
    }


def build_multi_stock_working(bundle, symbols, side=SIDE_BOTH):
    """Several stocks in ONE workbook.

    Returns {"kind": "multi", "side", "stocks": [stock workings, in symbol
    order], "skipped": [(symbol, reason)], "transactions": all the
    stocks' trades, "total": summary of everything, "meta": {...}}.
    A stock with no matching trades (date filter / buys-only etc.) is
    skipped and reported rather than failing the whole request."""
    stocks, skipped = [], []
    for symbol in sorted(set(symbols)):
        try:
            stocks.append(build_stock_working(bundle, symbol, side))
        except TradingReportsError as e:
            skipped.append((symbol, str(e)))
    if not stocks:
        detail = "; ".join(f"{sym}: {why}" for sym, why in skipped)
        raise TradingReportsError(f"None of the selected stocks has matching trades. {detail}")
    tx = pd.concat([w["transactions"] for w in stocks], ignore_index=True)
    tx["_o"] = tx["Account Code"].map(lambda c: ACCOUNT_ORDER.get(c, 99))
    tx = tx.sort_values(["Symbol", "Trade Date", "_o", "Settlement Date"], kind="stable").drop(columns="_o").reset_index(drop=True)
    return {
        "kind": "multi", "side": side, "stocks": stocks, "skipped": skipped,
        "transactions": tx, "total": _summarize(tx), "meta": _meta(tx, bundle),
    }


# ---------------------------------------------------------------------------
# Naming and default e-mail text
# ---------------------------------------------------------------------------

def _period_text(meta):
    a, b = meta.get("period_from"), meta.get("period_to")
    if a is None or b is None or pd.isna(a) or pd.isna(b):
        return ""
    return f"{a:%d-%b-%Y}" if a == b else f"{a:%d-%b-%Y} to {b:%d-%b-%Y}"


def working_title(working):
    if working["kind"] == "stock":
        return f"{working['symbol']} - {working['company']}"
    if working["kind"] == "multi":
        n = len(working["stocks"])
        return f"{n} stock{'s' if n != 1 else ''}: " + ", ".join(w["symbol"] for w in working["stocks"])
    return f"{working['bank']} (Account {working['code']})"


def _span(meta):
    a, b = meta.get("period_from"), meta.get("period_to")
    if a is None or pd.isna(a):
        return "report"
    return f"{a:%d%b%Y}" if a == b else f"{a:%d%b%Y}_to_{b:%d%b%Y}"


def build_filename(working):
    span = _span(working["meta"])
    if working["kind"] == "stock":
        name = f"Trading_Report_{working['symbol']}_{span}"
    elif working["kind"] == "multi":
        symbols = [w["symbol"] for w in working["stocks"]]
        label = "_".join(symbols) if len(symbols) <= 3 else f"{len(symbols)}_stocks"
        name = f"Trading_Report_{label}_{span}"
    else:
        name = f"Trading_Report_{working['bank']}_{working['code']}_{span}"
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", name) + ".xlsx"


def default_email_text(working, signature, separate=False):
    """(subject, body) - editable on the screen. `separate` is for several
    stocks sent as one e-mail with one workbook per stock."""
    period = _period_text(working["meta"])
    if working["kind"] == "stock":
        subject = f"Trading Report - {working['symbol']} - {period}"
        body = (
            "Dear Sir/Madam,\n\n"
            f"Please find attached the trading report for {working['company']} ({working['symbol']}) across our "
            f"trading accounts for the period {period}.\n\n"
            "The first sheet gives the summary by account and the second sheet lists the transactions.\n\n"
            "Regards,\n"
        )
    elif working["kind"] == "multi":
        symbols = ", ".join(w["symbol"] for w in working["stocks"])
        subject = f"Trading Report - {symbols} - {period}"
        if len(working["stocks"]) > 3:
            subject = f"Trading Report - {len(working['stocks'])} stocks - {period}"
        if separate:
            how = "A separate workbook is attached for each stock: the first sheet of each gives the summary by account and the second sheet lists the transactions."
            what = "the trading reports"
        else:
            how = "The first sheet gives the summary by stock and account and the second sheet lists the transactions."
            what = "the trading report"
        body = (
            "Dear Sir/Madam,\n\n"
            f"Please find attached {what} for {symbols} across our trading accounts for the period {period}.\n\n"
            f"{how}\n\n"
            "Regards,\n"
        )
    else:
        subject = f"Transaction Report - IGI General Insurance Limited ({working['code']}) - {period}"
        body = (
            "Dear Sir/Madam,\n\n"
            f"Please find attached the transaction report for IGI General Insurance Limited, trading account "
            f"{working['code']} ({working['bank']}), for the period {period}.\n\n"
            "The first sheet gives the summary by stock and the second sheet lists each transaction performed with "
            "its amount. Kindly review and confirm the details at your earliest convenience.\n\n"
            "Regards,\n"
        )
    return subject, body + signature
