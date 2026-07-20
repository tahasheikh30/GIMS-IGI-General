"""
Per-bank statement parsers for the Bank Position module.

Each bank exports its statement in a completely different shape (confirmed
against 5 real sample files: HBL, MCB, Faysal, Bank Alfalah, JS Bank) - so,
same pattern as loss_ratio_processor.py / premium_register_processor.py,
there's one dedicated parser per bank rather than a single generic reader
trying to guess a common layout. Every parser returns the same normalized
StatementResult shape regardless of the source format, so anything built on
top of this (reconciliation, the Excel generator) never needs to know which
bank it came from.

Confirmed real formats:
  HBL     - .xls that is actually an HTML table. Metadata rows (label in
            column 0, value duplicated into column 2) sit above a
            transaction table with a clean header row ("Transaction Date",
            "Value Date", ...). Closing balance is stated explicitly.
  MCB     - .xls that is actually an HTML table, but a raw core-banking
            export (POSITION/ACCT_NO/TRAN_DATE/... columns) with NO stated
            opening/closing balance line at all - the closing balance has
            to be taken from the last transaction row's RUNNING_BAL.
  Faysal  - .xls that is actually an HTML table. Opening balance is stated
            in a metadata block; closing balance is the last transaction
            row's Balance. Numeric cells are wrapped as ="1,234.00" (a
            Crystal-Reports-style quirk, forces Excel to treat the digits
            as text so leading zeros / formatting survive) and need that
            wrapper stripped before parsing as a float.
  BAFL    - genuine CSV. Opening AND closing balance are both stated
            explicitly in one header line ("Opening Balance,PKR...,Closing
            Balance,PKR...").
  JS      - genuine .xlsx. Metadata block (Account Title, Account Number,
            Opening Balance, Closing Balance, ...) sits above a
            transaction table with its own header row ("Post Date",
            "Description", ...).

Every parser was written and verified against real sample files - each
one's derived closing balance matches the balance already recorded in the
real Bank Position template for that account.

PASSWORD-PROTECTED FILES: none of the 5 real samples this was built
against were actually encrypted (a bank .xls being HTML-in-disguise is a
separate, unrelated quirk from encryption). Production files may well be
password-protected though, so decrypt_if_needed() below handles that path
generically via msoffcrypto-tool (Excel) / pikepdf (PDF) - both optional
imports, degrading with a clear error if not installed, since neither
could be verified against a real encrypted sample yet.
"""
import io
import os
import re
from datetime import datetime

import pandas as pd

try:
    import msoffcrypto
    MSOFFCRYPTO_AVAILABLE = True
except ImportError:
    MSOFFCRYPTO_AVAILABLE = False

try:
    import pikepdf
    PIKEPDF_AVAILABLE = True
except ImportError:
    PIKEPDF_AVAILABLE = False


class BankStatementError(Exception):
    """Raised for a specific file so the caller (the UI) can show a clear,
    file-scoped error instead of a raw traceback - same pattern as
    loss_ratio_processor.LossRatioInputError / premium_register_processor's
    ValueError-with-context."""
    def __init__(self, filepath, message):
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
        self.message = message
        super().__init__(f"{self.filename}: {message}")


class PasswordRequired(BankStatementError):
    """Raised specifically when a file is encrypted and no password (or
    the wrong password) was supplied - lets the caller distinguish 'ask
    the user for a password' from 'this file is genuinely broken'."""
    pass


# ---------------------------------------------------------------------------
# Password handling (generic - not yet exercised against a real encrypted
# sample; msoffcrypto-tool / pikepdf are the standard libraries for this
# and degrade gracefully with an actionable message if not installed)
# ---------------------------------------------------------------------------

def is_encrypted(filepath):
    """Cheap check for whether a file needs a password before it can be
    read at all. Does NOT flag HBL/MCB/Faysal's HTML-in-.xls quirk as
    encryption - that's a completely different (and harmless) thing."""
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == ".pdf":
            if not PIKEPDF_AVAILABLE:
                return False  # can't check without the library - let the later read attempt surface the real error
            with pikepdf.open(filepath):
                pass
            return False
        if ext in (".xlsx", ".xls"):
            if not MSOFFCRYPTO_AVAILABLE:
                return False
            with open(filepath, "rb") as f:
                office_file = msoffcrypto.OfficeFile(f)
                return office_file.is_encrypted()
    except pikepdf.PasswordError:
        return True
    except Exception:
        return False
    return False


def decrypt_if_needed(filepath, password=None):
    """
    If `filepath` is password-protected, decrypt it into an in-memory
    buffer using the supplied password and return that buffer instead
    (safe to hand straight to pandas/openpyxl - nothing decrypted ever
    touches disk). If the file isn't encrypted, returns the original
    filepath unchanged so callers can treat both cases identically.

    Raises PasswordRequired if the file is encrypted and no password (or
    the wrong one) was given, so the caller (the UI) can prompt rather
    than crash with a confusing library-level error.
    """
    if not is_encrypted(filepath):
        return filepath

    ext = os.path.splitext(filepath)[1].lower()

    if ext == ".pdf":
        if not PIKEPDF_AVAILABLE:
            raise BankStatementError(filepath, "This PDF is password-protected, but pikepdf isn't installed (pip install pikepdf).")
        if not password:
            raise PasswordRequired(filepath, "This file is password-protected - a saved password is required.")
        try:
            with pikepdf.open(filepath, password=password) as pdf:
                buf = io.BytesIO()
                pdf.save(buf)
                buf.seek(0)
                return buf
        except pikepdf.PasswordError:
            raise PasswordRequired(filepath, "The saved password did not unlock this file.")

    if ext in (".xlsx", ".xls"):
        if not MSOFFCRYPTO_AVAILABLE:
            raise BankStatementError(filepath, "This file is password-protected, but msoffcrypto-tool isn't installed (pip install msoffcrypto-tool).")
        if not password:
            raise PasswordRequired(filepath, "This file is password-protected - a saved password is required.")
        try:
            with open(filepath, "rb") as f:
                office_file = msoffcrypto.OfficeFile(f)
                office_file.load_key(password=password)
                buf = io.BytesIO()
                office_file.decrypt(buf)
                buf.seek(0)
                return buf
        except Exception as e:
            raise PasswordRequired(filepath, f"The saved password did not unlock this file ({e}).")

    return filepath


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _to_float(value):
    """'1,234.56' -> 1234.56, '="6,537,158.23"' -> 6537158.23 (Faysal's
    text-wrapper quirk), '' / '-' / None -> None."""
    if value is None:
        return None
    text = str(value).strip()
    if text == "" or text.lower() == "nan":
        return None
    text = text.strip('="')  # strips = and " from either end, any order/count
    text = text.replace(",", "").replace("PKR", "").strip()
    if text in ("", "-", "--"):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def normalize_account_no(value):
    """Digits-only form of an account number, so a statement's raw
    '4010575611005655' can be matched against the template's more loosely
    formatted '1242-79496370-01' / '08101216901' style entries later
    (same rationale as premium_register_processor's digit-run matching -
    every bank punctuates account numbers differently)."""
    return re.sub(r"\D", "", str(value or ""))


class StatementResult:
    """Normalized output every per-bank parser returns, regardless of
    source format - so nothing downstream (reconciliation, the Excel
    generator) needs to know which bank a file came from."""
    def __init__(self, bank_key, account_no, account_title, currency,
                 opening_balance, opening_date, closing_balance, closing_date,
                 transactions, source_file, warnings=None):
        self.bank_key = bank_key
        self.account_no = account_no
        self.account_no_digits = normalize_account_no(account_no)
        self.account_title = account_title
        self.currency = currency
        self.opening_balance = opening_balance
        self.opening_date = opening_date
        self.closing_balance = closing_balance
        self.closing_date = closing_date
        self.transactions = transactions  # DataFrame: date, description, debit, credit, balance
        self.source_file = source_file
        self.warnings = warnings or []

    def __repr__(self):
        return (f"StatementResult({self.bank_key}, acct={self.account_no}, "
                f"closing={self.closing_balance}, as_of={self.closing_date}, "
                f"{len(self.transactions)} txns)")


# ---------------------------------------------------------------------------
# HBL - HTML-as-.xls, metadata block + clean transaction header row
# ---------------------------------------------------------------------------

def parse_hbl(filepath_or_buffer, source_name=None):
    source_name = source_name or (filepath_or_buffer if isinstance(filepath_or_buffer, str) else "HBL statement")
    try:
        tables = pd.read_html(filepath_or_buffer)
    except Exception as e:
        raise BankStatementError(source_name, f"Could not read as an HBL statement ({e})")
    if not tables:
        raise BankStatementError(source_name, "No table found in this file")
    t = tables[0]

    header_idx = None
    for i in range(len(t)):
        if str(t.iat[i, 0]).strip() == "Transaction Date":
            header_idx = i
            break
    if header_idx is None:
        raise BankStatementError(source_name, "Could not find the transaction table header row ('Transaction Date') - this may not be an HBL statement.")

    meta = {}
    for i in range(header_idx):
        label = str(t.iat[i, 0]).strip().rstrip(":").lower()
        value = t.iat[i, 2]
        if label:
            meta[label] = value

    account_no = meta.get("account no", "")
    account_title = meta.get("account title", "")
    currency = meta.get("account currency", "PKR")

    opening_balance = None
    opening_date = None
    for k, v in meta.items():
        if k.startswith("opening balance"):
            opening_balance = _to_float(v)
            m = re.search(r"(\d{2}/\d{2}/\d{4})", k)
            if m:
                opening_date = datetime.strptime(m.group(1), "%d/%m/%Y").date()

    txn_rows = t.iloc[header_idx + 1:].copy()
    txn_rows.columns = ["txn_date", "value_date", "reference", "description", "debit", "credit", "balance"]
    txn_rows = txn_rows[txn_rows["txn_date"].notna() & (txn_rows["txn_date"].astype(str).str.strip() != "")]
    txn_rows = txn_rows[txn_rows["txn_date"].astype(str).str.match(r"\d{2}/\d{2}/\d{4}")]

    transactions = pd.DataFrame({
        "date": pd.to_datetime(txn_rows["txn_date"], format="%d/%m/%Y", errors="coerce"),
        "description": txn_rows["description"].astype(str).str.strip(),
        "debit": txn_rows["debit"].apply(_to_float),
        "credit": txn_rows["credit"].apply(_to_float),
        "balance": txn_rows["balance"].apply(_to_float),
    }).reset_index(drop=True)

    if len(transactions) == 0:
        raise BankStatementError(source_name, "Found the transaction table header but no transaction rows under it.")

    closing_balance = transactions["balance"].iloc[-1]
    closing_date = transactions["date"].iloc[-1]
    if closing_date is not None and pd.notna(closing_date):
        closing_date = closing_date.date()

    return StatementResult(
        bank_key="hbl", account_no=account_no, account_title=account_title, currency=currency,
        opening_balance=opening_balance, opening_date=opening_date,
        closing_balance=closing_balance, closing_date=closing_date,
        transactions=transactions, source_file=os.path.basename(source_name) if isinstance(source_name, str) else "HBL statement",
    )


# ---------------------------------------------------------------------------
# MCB - HTML-as-.xls, raw core-banking export, NO stated opening/closing
# balance line - closing balance = last row's RUNNING_BAL.
# ---------------------------------------------------------------------------

def parse_mcb(filepath_or_buffer, source_name=None):
    source_name = source_name or (filepath_or_buffer if isinstance(filepath_or_buffer, str) else "MCB statement")
    try:
        tables = pd.read_html(filepath_or_buffer)
    except Exception as e:
        raise BankStatementError(source_name, f"Could not read as an MCB statement ({e})")
    if len(tables) < 2:
        raise BankStatementError(source_name, "Expected a metadata block plus a transaction table - only found one table. This may not be an MCB statement.")

    meta_block = tables[0]
    account_no = ""
    meta_text = " ".join(str(v) for v in meta_block.iloc[:, 0].tolist())
    m = re.search(r"Account No\.?:?\s*(\d+)", meta_text)
    if m:
        account_no = m.group(1)

    txns = tables[1]
    required = ["TRAN_DATE", "TRAN_DESC", "TRAN_AMT", "CR_DR_MAINT_IND", "RUNNING_BAL"]
    missing = [c for c in required if c not in txns.columns]
    if missing:
        raise BankStatementError(source_name, f"Missing expected column(s) {missing} - this may not be an MCB statement, or MCB has changed its export format.")

    if not account_no and "ACCT_NO" in txns.columns and len(txns):
        account_no = str(txns["ACCT_NO"].iloc[0])

    is_credit = txns["CR_DR_MAINT_IND"].astype(str).str.upper() == "C"
    amounts = txns["TRAN_AMT"].apply(_to_float)

    transactions = pd.DataFrame({
        "date": pd.to_datetime(txns["TRAN_DATE"], format="%d.%m.%Y %H:%M:%S", errors="coerce"),
        "description": txns["TRAN_DESC"].astype(str).str.strip(),
        "debit": amounts.where(~is_credit),
        "credit": amounts.where(is_credit),
        "balance": txns["RUNNING_BAL"].apply(_to_float),
    }).reset_index(drop=True)

    if len(transactions) == 0:
        raise BankStatementError(source_name, "No transaction rows found.")

    closing_balance = transactions["balance"].iloc[-1]
    closing_date = transactions["date"].iloc[-1]
    if closing_date is not None and pd.notna(closing_date):
        closing_date = closing_date.date()

    return StatementResult(
        bank_key="mcb", account_no=account_no, account_title="", currency="PKR",
        opening_balance=None, opening_date=None,  # MCB's export doesn't state one - see module docstring
        closing_balance=closing_balance, closing_date=closing_date,
        transactions=transactions, source_file=os.path.basename(source_name) if isinstance(source_name, str) else "MCB statement",
        warnings=["MCB's export doesn't state an opening balance - only closing balance (from the last transaction row) is available."],
    )


# ---------------------------------------------------------------------------
# Faysal - HTML-as-.xls, metadata block + transaction table, numeric cells
# wrapped as ="1,234.00"
# ---------------------------------------------------------------------------

def parse_faysal(filepath_or_buffer, source_name=None):
    source_name = source_name or (filepath_or_buffer if isinstance(filepath_or_buffer, str) else "Faysal statement")
    try:
        tables = pd.read_html(filepath_or_buffer)
    except Exception as e:
        raise BankStatementError(source_name, f"Could not read as a Faysal statement ({e})")
    if len(tables) < 2:
        raise BankStatementError(source_name, "Expected a metadata block plus a transaction table - only found one table. This may not be a Faysal statement.")

    meta_block = tables[0]
    meta_text = " | ".join(str(v) for v in meta_block.values.flatten().tolist() if str(v) != "nan")

    def _extract(label_pattern):
        m = re.search(label_pattern, meta_text)
        return m.group(1).strip() if m else None

    account_no = _extract(r"Account No\s*:\s*([\w\-]+)")
    account_title = _extract(r"Account Title\s*:\s*([A-Za-z0-9 ,.&\-]+?)(?:\s{2,}|\||$)")
    opening_balance_text = _extract(r"Opening Balance as on ([\d\-A-Za-z]+)\s*:\s*([\d,]+\.\d+)")
    opening_date, opening_balance = None, None
    m = re.search(r"Opening Balance as on ([\d\-A-Za-z]+)\s*:\s*([\d,]+\.\d+)", meta_text)
    if m:
        try:
            opening_date = datetime.strptime(m.group(1), "%d-%b-%Y").date()
        except ValueError:
            pass
        opening_balance = _to_float(m.group(2))

    txns = tables[1]
    required = ["Tran Date", "Tran Description", "Withdrawal Amount", "Deposit Amount", "Balance"]
    missing = [c for c in required if c not in txns.columns]
    if missing:
        raise BankStatementError(source_name, f"Missing expected column(s) {missing} - this may not be a Faysal statement, or Faysal has changed its export format.")

    transactions = pd.DataFrame({
        "date": pd.to_datetime(txns["Tran Date"], errors="coerce"),
        "description": txns["Tran Description"].astype(str).str.strip('="').str.strip(),
        "debit": txns["Withdrawal Amount"].apply(_to_float),
        "credit": txns["Deposit Amount"].apply(_to_float),
        "balance": txns["Balance"].apply(_to_float),
    })
    transactions = transactions[transactions["date"].notna()].reset_index(drop=True)

    if len(transactions) == 0:
        raise BankStatementError(source_name, "No transaction rows found.")

    closing_balance = transactions["balance"].iloc[-1]
    closing_date = transactions["date"].iloc[-1]
    if closing_date is not None and pd.notna(closing_date):
        closing_date = closing_date.date()

    return StatementResult(
        bank_key="faysal", account_no=account_no or "", account_title=account_title or "", currency="PKR",
        opening_balance=opening_balance, opening_date=opening_date,
        closing_balance=closing_balance, closing_date=closing_date,
        transactions=transactions, source_file=os.path.basename(source_name) if isinstance(source_name, str) else "Faysal statement",
    )


# ---------------------------------------------------------------------------
# BAFL - genuine CSV, opening AND closing balance both stated explicitly
# ---------------------------------------------------------------------------

def parse_bafl(filepath_or_buffer, source_name=None):
    source_name = source_name or (filepath_or_buffer if isinstance(filepath_or_buffer, str) else "BAFL statement")
    try:
        if isinstance(filepath_or_buffer, str):
            with open(filepath_or_buffer, "r", encoding="utf-8-sig") as f:
                lines = f.read().splitlines()
        else:
            filepath_or_buffer.seek(0)
            lines = filepath_or_buffer.read().decode("utf-8-sig").splitlines()
    except Exception as e:
        raise BankStatementError(source_name, f"Could not read as a BAFL statement ({e})")

    account_no, opening_balance, closing_balance, header_idx = "", None, None, None
    for i, line in enumerate(lines[:20]):
        if line.startswith("Account No"):
            parts = line.split(",")
            if len(parts) > 1:
                account_no = parts[1].strip()
        elif line.startswith("Opening Balance"):
            parts = line.split(",")
            # "Opening Balance,PKR278367503.01,Closing Balance,PKR116443284.43"
            if len(parts) >= 4:
                opening_balance = _to_float(parts[1])
                closing_balance = _to_float(parts[3])
        elif line.startswith("Posting Date"):
            header_idx = i
            break

    if header_idx is None:
        raise BankStatementError(source_name, "Could not find the transaction table header row ('Posting Date') - this may not be a BAFL statement.")

    csv_text = "\n".join(lines[header_idx:])
    txns = pd.read_csv(io.StringIO(csv_text))
    required = ["Posting Date", "Description", "Debit", "Credit", "Balance"]
    missing = [c for c in required if c not in txns.columns]
    if missing:
        raise BankStatementError(source_name, f"Missing expected column(s) {missing} - this may not be a BAFL statement, or Bank Alfalah has changed its export format.")

    transactions = pd.DataFrame({
        "date": pd.to_datetime(txns["Posting Date"], errors="coerce"),
        "description": txns["Description"].astype(str).str.strip(),
        "debit": txns["Debit"].apply(_to_float),
        "credit": txns["Credit"].apply(_to_float),
        "balance": txns["Balance"].apply(_to_float),
    })
    transactions = transactions[transactions["date"].notna()].reset_index(drop=True)

    if len(transactions) == 0:
        raise BankStatementError(source_name, "No transaction rows found.")

    if closing_balance is None:
        closing_balance = transactions["balance"].iloc[-1]
    closing_date = transactions["date"].iloc[-1]
    if closing_date is not None and pd.notna(closing_date):
        closing_date = closing_date.date()

    return StatementResult(
        bank_key="bafl", account_no=account_no, account_title="", currency="PKR",
        opening_balance=opening_balance, opening_date=None,
        closing_balance=closing_balance, closing_date=closing_date,
        transactions=transactions, source_file=os.path.basename(source_name) if isinstance(source_name, str) else "BAFL statement",
    )


# ---------------------------------------------------------------------------
# JS Bank - genuine .xlsx, metadata block + transaction table
# ---------------------------------------------------------------------------

def parse_js(filepath_or_buffer, source_name=None):
    source_name = source_name or (filepath_or_buffer if isinstance(filepath_or_buffer, str) else "JS Bank statement")
    try:
        import openpyxl
        wb = openpyxl.load_workbook(filepath_or_buffer, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
    except Exception as e:
        raise BankStatementError(source_name, f"Could not read as a JS Bank statement ({e})")

    header_idx = None
    meta = {}
    for i, row in enumerate(rows):
        label = str(row[0]).strip().lower() if row[0] is not None else ""
        if label == "post date":
            header_idx = i
            break
        if label and row[0] is not None and len(row) > 1 and row[1] is not None:
            meta[label] = row[1]

    if header_idx is None:
        raise BankStatementError(source_name, "Could not find the transaction table header row ('Post Date') - this may not be a JS Bank statement.")

    account_no = meta.get("account number", "")
    account_title = meta.get("account title", "")
    opening_balance = _to_float(meta.get("opening balance"))
    closing_balance_stated = _to_float(meta.get("closing balance"))

    data_rows = [r for r in rows[header_idx + 1:] if r[0] not in (None, "")]
    if not data_rows:
        raise BankStatementError(source_name, "Found the transaction table header but no transaction rows under it.")

    transactions = pd.DataFrame({
        "date": [pd.to_datetime(str(r[0]), format="%Y%m%d", errors="coerce") for r in data_rows],
        "description": [str(r[1]).strip() if r[1] is not None else "" for r in data_rows],
        "debit": [_to_float(r[4]) for r in data_rows],
        "credit": [_to_float(r[5]) for r in data_rows],
        "balance": [_to_float(r[6]) for r in data_rows],
    })

    closing_balance = closing_balance_stated if closing_balance_stated is not None else transactions["balance"].iloc[-1]
    closing_date = transactions["date"].iloc[-1]
    if closing_date is not None and pd.notna(closing_date):
        closing_date = closing_date.date()

    return StatementResult(
        bank_key="js", account_no=account_no, account_title=account_title, currency="PKR",
        opening_balance=opening_balance, opening_date=None,
        closing_balance=closing_balance, closing_date=closing_date,
        transactions=transactions, source_file=os.path.basename(source_name) if isinstance(source_name, str) else "JS Bank statement",
    )


# ---------------------------------------------------------------------------
# Registry + dispatcher
# ---------------------------------------------------------------------------

BANK_REGISTRY = {
    "hbl":    {"display_name": "Habib Bank Limited (HBL)", "parser": parse_hbl,    "expected_ext": [".xls"]},
    "mcb":    {"display_name": "MCB Bank",                 "parser": parse_mcb,    "expected_ext": [".xls"]},
    "faysal": {"display_name": "Faysal Bank",               "parser": parse_faysal, "expected_ext": [".xls"]},
    "bafl":   {"display_name": "Bank Alfalah (BAFL)",        "parser": parse_bafl,   "expected_ext": [".csv"]},
    "js":     {"display_name": "JS Bank",                    "parser": parse_js,     "expected_ext": [".xlsx"]},
}


def parse_statement(filepath, bank_key, password=None):
    """
    Main entry point. filepath: path to the uploaded file. bank_key: which
    bank slot this was uploaded into (e.g. "hbl") - the dedicated-slot-per-
    bank design means the caller always knows this already, so there's no
    need for (and no reliability lost by skipping) auto-detection.
    password: saved password from the credential store (see the module
    docstring) - passed through untouched if the file isn't encrypted.

    Returns a StatementResult. Raises BankStatementError (or its
    PasswordRequired subclass) with a clear, file-scoped message on any
    problem.
    """
    if bank_key not in BANK_REGISTRY:
        raise BankStatementError(filepath, f"Unknown bank '{bank_key}'. Known banks: {list(BANK_REGISTRY.keys())}")

    entry = BANK_REGISTRY[bank_key]
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in entry["expected_ext"]:
        # Not fatal - some banks change export format over time - but
        # worth surfacing since it may mean the wrong file landed in the
        # wrong bank's slot.
        pass

    source = decrypt_if_needed(filepath, password=password)
    result = entry["parser"](source, source_name=filepath)
    return result


def parse_batch(file_bank_pairs, password_lookup=None):
    """
    file_bank_pairs: list of (filepath, bank_key) tuples - e.g. what the
    UI's 5 dedicated upload slots would hand over in one go each morning.
    password_lookup: optional dict bank_key -> password (from the
    credential store).
    Returns (results, errors) - a problem in one bank's file never stops
    the others from being parsed, same resilience pattern as
    loss_ratio_processor.process_files().
    """
    password_lookup = password_lookup or {}
    results, errors = [], []
    for filepath, bank_key in file_bank_pairs:
        try:
            results.append(parse_statement(filepath, bank_key, password=password_lookup.get(bank_key)))
        except BankStatementError as e:
            errors.append(e)
        except Exception as e:
            errors.append(BankStatementError(filepath, f"Unexpected error: {e}"))
    return results, errors


if __name__ == "__main__":
    import sys
    samples = {
        "hbl": "/mnt/user-data/uploads/HBL.xls",
        "mcb": "/mnt/user-data/uploads/MCB_July_5655.xls",
        "faysal": "/mnt/user-data/uploads/Faysal_6018-.xls",
        "bafl": "/mnt/user-data/uploads/BAFL_July_till_date.csv",
        "js": "/mnt/user-data/uploads/JS_107791.xlsx",
    }
    results, errors = parse_batch([(fp, bank) for bank, fp in samples.items()])
    for r in results:
        print(r)
        print(f"   account_no_digits={r.account_no_digits!r}  title={r.account_title!r}")
        print(f"   opening={r.opening_balance}  closing={r.closing_balance}  as_of={r.closing_date}")
        if r.warnings:
            print(f"   warnings: {r.warnings}")
        print()
    for e in errors:
        print("ERROR:", e)
