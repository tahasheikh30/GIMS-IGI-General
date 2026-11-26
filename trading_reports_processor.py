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


def _meta(df):
    return {
        "date_from": df["Trade Date"].min() if len(df) else None,
        "date_to": df["Trade Date"].max() if len(df) else None,
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
        "rows": rows, "total": _summarize(selected), "transactions": tx, "meta": _meta(selected),
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
        "rows": rows, "total": _summarize(selected), "transactions": tx, "meta": _meta(selected),
    }


# ---------------------------------------------------------------------------
# Naming and default e-mail text
# ---------------------------------------------------------------------------

def _period_text(meta):
    a, b = meta["date_from"], meta["date_to"]
    if a is None or b is None or pd.isna(a) or pd.isna(b):
        return ""
    return f"{a:%d-%b-%Y} to {b:%d-%b-%Y}"


def working_title(working):
    if working["kind"] == "stock":
        return f"{working['symbol']} - {working['company']}"
    return f"{working['bank']} (Account {working['code']})"


def build_filename(working):
    meta = working["meta"]
    a, b = meta["date_from"], meta["date_to"]
    span = f"{a:%d%b%Y}_to_{b:%d%b%Y}" if a is not None and not pd.isna(a) else "report"
    if working["kind"] == "stock":
        name = f"Trading_Report_{working['symbol']}_{span}"
    else:
        name = f"Trading_Report_{working['bank']}_{working['code']}_{span}"
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", name) + ".xlsx"


def default_email_text(working, signature):
    """(subject, body) - editable on the screen."""
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
