"""
Bank Position Excel generator.

Design: rather than building a workbook from scratch, this opens the
PREVIOUS day's real Bank Position file and updates only the specific
cells that change day to day. Everything else - Kibor rates, T-Bills,
the Investment portfolio block, the Annex-A / Annex-A (2) / Kibor_PQ
sheets - is carried forward completely untouched, per instruction.

Why this approach rather than regenerating the workbook fresh each time:
  - Summary is almost entirely formula-driven off 'IGI General '/'Takaful'/
    'Outward remittances' (confirmed by inspecting the real template's
    formulas) - so updating the source sheets' balance cells is enough;
    Summary recalculates itself in Excel on open, nothing needs to be
    computed here.
  - It guarantees byte-for-byte fidelity on everything we were told to
    leave alone, instead of us having to reverse-engineer and rebuild
    Kibor-table/T-Bills formatting we were explicitly told not to touch.

KNOWN STRUCTURAL LIMIT (surfaced deliberately, not silently worked
around): the Summary sheet's Remittances (3 rows), Major Payments
(1 row), and Major Receivables (2 rows) blocks have a fixed row count
baked into existing SUM formulas and sit directly above the Investment
section. Inserting/deleting rows there would need every downstream
formula and merged-cell layout in the workbook re-verified to avoid
silently corrupting the sheet - too risky to do unattended. So this
generator fills what fits into the existing capacity and raises a clear,
itemized warning for anything that doesn't fit, rather than guessing.
"""
import os
import shutil
from datetime import datetime, date

import openpyxl

# ---------------------------------------------------------------------------
# Row maps - which row in the template each account lives on. Built from
# inspecting the real Bank_Position_16th_July.xlsx template directly.
# ---------------------------------------------------------------------------

IGI_GENERAL_SHEET = "IGI General "  # trailing space is real, confirmed from the template
TAKAFUL_SHEET = "Takaful"
REMITTANCES_SHEET = "Outward remittances"
SUMMARY_SHEET = "Summary"

AMOUNT_COL = 4  # column D on both IGI General and Takaful

IGI_GENERAL_ROWS = {
    "standard_chartered": 6, "faysal": 7, "js": 8, "hbl": 9,
    "mcb_current": 10, "mcb_savings": 11, "bafl": 12, "alhabib": 13,
    "ubl": 14, "alhabib_seaview": 15, "mcb_stock_exchange": 16,
    "allied": 17, "alhabib_garden_town": 18,
}
IGI_GENERAL_TOTAL_ROW = 19

TAKAFUL_ROWS = {
    "meezan_ptf": 6, "meezan_ops": 7, "bankislami_ptf": 8, "mcb_islamic_ptf": 9,
    "alhabib_takaful_ptf": 10, "alfalah_islamic_ptf": 11, "ubl_ameen_ptf": 12,
    "scb_sadiq_ptf": 13, "scb_sadiq_ops": 14, "faysal_islamic_ptf": 15,
    "dib_ptf": 16, "dib_ops": 17,
}
TAKAFUL_TOTAL_ROW = 18

# Which of bank_statement_processor's bank_keys maps to which template row
# key, confirmed by matching account numbers between the real sample
# statements and the real template (see build discussion). Only banks we
# actually have a parser for today are listed - add more here as new
# parsers/accounts come online, per the "5 banks now, rest later" decision.
STATEMENT_BANK_TO_TEMPLATE_ROW = {
    "hbl": "hbl",
    "faysal": "faysal",
    "js": "js",
    "bafl": "bafl",
    "mcb": "mcb_savings",  # the MCB sample statement's account matches the savings account row
}

REMITTANCES_HEADER_ROW = 6
REMITTANCES_START_ROW = 7
REMITTANCES_CAPACITY = 3  # rows 7-9 - see module docstring on why this isn't auto-expanded
USD_RATE_CELL = "D2"

SUMMARY_REMITTANCES_START_ROW = 14
SUMMARY_MAJOR_PAYMENTS_START_ROW = 20
SUMMARY_MAJOR_PAYMENTS_CAPACITY = 1
SUMMARY_MAJOR_RECEIVABLES_START_ROW = 24
SUMMARY_MAJOR_RECEIVABLES_CAPACITY = 2


class BankPositionGenerationWarning:
    """Non-fatal issue worth surfacing to whoever reviews the file before
    it goes to the CFO (e.g. overflow beyond a fixed block's capacity)."""
    def __init__(self, message):
        self.message = message

    def __repr__(self):
        return f"Warning: {self.message}"


def _clear_row_range(ws, start_row, end_row, cols):
    for r in range(start_row, end_row + 1):
        for c in cols:
            ws.cell(row=r, column=c).value = None


def _write_conventional_balances(ws, balances, warnings):
    """balances: dict of template_row_key -> amount. Rows not present in
    `balances` are left exactly as they were (carry-forward by omission -
    we simply never touch that cell)."""
    for row_key, amount in balances.items():
        if row_key not in IGI_GENERAL_ROWS:
            warnings.append(BankPositionGenerationWarning(
                f"Unknown conventional account key '{row_key}' - no matching row in the template, skipped."))
            continue
        ws.cell(row=IGI_GENERAL_ROWS[row_key], column=AMOUNT_COL).value = amount


def _write_islamic_balances(ws, balances, warnings):
    for row_key, amount in balances.items():
        if row_key not in TAKAFUL_ROWS:
            warnings.append(BankPositionGenerationWarning(
                f"Unknown Takaful account key '{row_key}' - no matching row in the template, skipped."))
            continue
        ws.cell(row=TAKAFUL_ROWS[row_key], column=AMOUNT_COL).value = amount


def _write_remittances(ws, remittances, warnings):
    """remittances: list of dicts with keys reinsurer, insured_name, amount,
    currency (defaults to 'PKR'). Overwrites the existing fixed-capacity
    block; F (PKR equivalent) / G (USD-Euro equivalent) keep their
    existing per-row formulas (=D{row}, =F{row}/$D$2) so this stays
    consistent with how the template already computes them."""
    if len(remittances) > REMITTANCES_CAPACITY:
        overflow = remittances[REMITTANCES_CAPACITY:]
        names = ", ".join(r.get("reinsurer", "?") for r in overflow)
        warnings.append(BankPositionGenerationWarning(
            f"{len(overflow)} remittance(s) beyond the template's fixed {REMITTANCES_CAPACITY}-row capacity "
            f"were NOT written and need adding manually: {names}"))
        remittances = remittances[:REMITTANCES_CAPACITY]

    _clear_row_range(ws, REMITTANCES_START_ROW, REMITTANCES_START_ROW + REMITTANCES_CAPACITY - 1,
                      cols=[1, 2, 3, 4, 5, 6, 7])

    for i, r in enumerate(remittances):
        row = REMITTANCES_START_ROW + i
        ws.cell(row=row, column=1).value = r.get("reinsurer", "")
        ws.cell(row=row, column=2).value = r.get("insured_name", "")
        ws.cell(row=row, column=4).value = r["amount"]
        ws.cell(row=row, column=5).value = r.get("currency", "PKR")
        ws.cell(row=row, column=6).value = f"=D{row}"
        ws.cell(row=row, column=7).value = f"=F{row}/$D${USD_RATE_CELL[1:]}"


def _write_summary_remittances_block(ws, count, warnings):
    """Summary's remittance rows are cross-sheet formulas pointing at
    'Outward remittances' - rewritten here to match however many rows
    were actually written there (capped the same way)."""
    count = min(count, REMITTANCES_CAPACITY)
    _clear_row_range(ws, SUMMARY_REMITTANCES_START_ROW, SUMMARY_REMITTANCES_START_ROW + REMITTANCES_CAPACITY - 1,
                      cols=[2, 3, 5])
    for i in range(count):
        summary_row = SUMMARY_REMITTANCES_START_ROW + i
        source_row = REMITTANCES_START_ROW + i
        ws.cell(row=summary_row, column=2).value = f"='Outward remittances'!A{source_row}"
        ws.cell(row=summary_row, column=3).value = f"='Outward remittances'!D{source_row}"
        ws.cell(row=summary_row, column=5).value = f"='Outward remittances'!B{source_row}"
    total_row = SUMMARY_REMITTANCES_START_ROW + REMITTANCES_CAPACITY
    last_data_row = SUMMARY_REMITTANCES_START_ROW + REMITTANCES_CAPACITY - 1
    ws.cell(row=total_row, column=3).value = f"=SUM(C{SUMMARY_REMITTANCES_START_ROW}:C{last_data_row})"


def _write_summary_named_amount_block(ws, items, start_row, capacity, name_key, extra_key, extra_col,
                                       total_row_offset, warnings, block_label):
    """Shared logic for Major Payments / Major Receivables - both are a
    literal (not formula) Name/Amount/Bank/Date list directly on Summary,
    same shape, different capacity."""
    if len(items) > capacity:
        overflow = items[capacity:]
        names = ", ".join(i.get(name_key, "?") for i in overflow)
        warnings.append(BankPositionGenerationWarning(
            f"{len(overflow)} {block_label} item(s) beyond the template's fixed {capacity}-row capacity "
            f"were NOT written and need adding manually: {names}"))
        items = items[:capacity]

    _clear_row_range(ws, start_row, start_row + capacity - 1, cols=[2, 3, 4, 5])
    for i, item in enumerate(items):
        row = start_row + i
        ws.cell(row=row, column=2).value = item.get(name_key, "")
        ws.cell(row=row, column=3).value = item.get("amount", 0)
        ws.cell(row=row, column=4).value = item.get(extra_key, "")
        exp_date = item.get("expected_date")
        if exp_date is not None:
            ws.cell(row=row, column=5).value = exp_date if isinstance(exp_date, (date, datetime)) else exp_date

    total_row = start_row + capacity
    last_data_row = start_row + capacity - 1
    ws.cell(row=total_row, column=3).value = f"=SUM(C{start_row}:C{last_data_row})"


def get_account_labels(previous_file):
    """
    Read the actual bank/account display label for every known row
    DIRECTLY FROM the given template file's own column B, rather than
    hardcoding names here - this way the UI always reflects exactly what
    the real workbook calls each account (spacing, exact bank names,
    branch qualifiers, etc.) and stays correct even if a label is ever
    edited by hand in Excel later, instead of silently drifting out of
    sync with a hardcoded guess.

    Returns (conventional_labels, islamic_labels), each a dict of
    template_row_key -> display text (falls back to the row_key itself
    if that cell is somehow blank).
    """
    wb = openpyxl.load_workbook(previous_file, data_only=True)
    igi_ws = wb[IGI_GENERAL_SHEET]
    takaful_ws = wb[TAKAFUL_SHEET]
    conventional_labels = {k: (igi_ws.cell(row=r, column=2).value or k) for k, r in IGI_GENERAL_ROWS.items()}
    islamic_labels = {k: (takaful_ws.cell(row=r, column=2).value or k) for k, r in TAKAFUL_ROWS.items()}
    return conventional_labels, islamic_labels


def get_prior_balances(previous_file):
    """Read yesterday's own D-column balance for every known row, so a
    quick-entry UI can pre-fill each field with the last known value
    (clearly a starting point to overwrite, not today's figure) rather
    than starting blank. Returns (conventional_balances, islamic_balances)."""
    wb = openpyxl.load_workbook(previous_file, data_only=True)
    igi_ws = wb[IGI_GENERAL_SHEET]
    takaful_ws = wb[TAKAFUL_SHEET]
    conventional = {k: igi_ws.cell(row=r, column=AMOUNT_COL).value for k, r in IGI_GENERAL_ROWS.items()}
    islamic = {k: takaful_ws.cell(row=r, column=AMOUNT_COL).value for k, r in TAKAFUL_ROWS.items()}
    return conventional, islamic


def balances_from_statements(statement_results):
    """Convenience: turn a list of StatementResult (from
    bank_statement_processor.py) into the template_row_key -> amount dict
    _write_conventional_balances expects, using each result's closing
    balance. Manual/portal entries should still be merged in on top of
    this (portal preferred - see build discussion) rather than used alone."""
    out = {}
    for r in statement_results:
        row_key = STATEMENT_BANK_TO_TEMPLATE_ROW.get(r.bank_key)
        if row_key:
            out[row_key] = r.closing_balance
    return out


def generate_bank_position(previous_file, output_file,
                            conventional_balances=None, islamic_balances=None,
                            remittances=None, major_payments=None, major_receivables=None,
                            as_of_date=None):
    """
    previous_file: path to yesterday's Bank Position workbook - used as the
        base so everything not explicitly updated is carried forward
        exactly as-is (Kibor rates, T-Bills, Investment portfolio,
        Annex-A / Annex-A (2) / Kibor_PQ).
    output_file: where to save today's file.
    conventional_balances / islamic_balances: dict of template_row_key ->
        amount (see IGI_GENERAL_ROWS / TAKAFUL_ROWS for valid keys, or use
        balances_from_statements() then merge manual overrides on top).
        Any row key not present is left untouched (carried forward).
    remittances: list of dicts {reinsurer, insured_name, amount, currency}.
    major_payments / major_receivables: list of dicts {name (or
        counterparty), amount, bank, expected_date} - e.g. straight from
        bank_position_register.list_items().
    as_of_date: date for this report; defaults to today.

    Returns a list of BankPositionGenerationWarning - empty if everything
    fit cleanly. ALWAYS check this list; a non-empty list means something
    didn't make it into the file and needs manual attention before this
    goes to the CFO.
    """
    if not os.path.exists(previous_file):
        raise FileNotFoundError(f"Previous day's Bank Position file not found: {previous_file}")

    as_of_date = as_of_date or date.today()
    warnings = []

    shutil.copy(previous_file, output_file)
    wb = openpyxl.load_workbook(output_file, data_only=False)

    for sheet in (IGI_GENERAL_SHEET, TAKAFUL_SHEET, REMITTANCES_SHEET, SUMMARY_SHEET):
        if sheet not in wb.sheetnames:
            raise ValueError(f"Expected sheet '{sheet}' not found in {previous_file} - is this the right template?")

    igi_ws = wb[IGI_GENERAL_SHEET]
    takaful_ws = wb[TAKAFUL_SHEET]
    remit_ws = wb[REMITTANCES_SHEET]
    summary_ws = wb[SUMMARY_SHEET]

    # Date - IGI General's J1 drives A3 (='J1' formula) on both IGI
    # General and Takaful (Takaful's own date cells reference IGI
    # General's), so this one write updates the date everywhere it needs to.
    igi_ws["J1"].value = datetime.combine(as_of_date, datetime.min.time())

    if conventional_balances:
        _write_conventional_balances(igi_ws, conventional_balances, warnings)
    if islamic_balances:
        _write_islamic_balances(takaful_ws, islamic_balances, warnings)

    if remittances is not None:
        _write_remittances(remit_ws, remittances, warnings)
        _write_summary_remittances_block(summary_ws, min(len(remittances), REMITTANCES_CAPACITY), warnings)

    if major_payments is not None:
        _write_summary_named_amount_block(
            summary_ws, major_payments, SUMMARY_MAJOR_PAYMENTS_START_ROW, SUMMARY_MAJOR_PAYMENTS_CAPACITY,
            name_key="name", extra_key="bank", extra_col=4, total_row_offset=1, warnings=warnings,
            block_label="Major Payment")

    if major_receivables is not None:
        _write_summary_named_amount_block(
            summary_ws, major_receivables, SUMMARY_MAJOR_RECEIVABLES_START_ROW, SUMMARY_MAJOR_RECEIVABLES_CAPACITY,
            name_key="name", extra_key="bank", extra_col=4, total_row_offset=1, warnings=warnings,
            block_label="Major Receivable")

    wb.save(output_file)
    return warnings


if __name__ == "__main__":
    import sys
    sys.path.insert(0, os.path.dirname(__file__))
    from bank_statement_processor import parse_batch

    samples = {
        "hbl": "/mnt/user-data/uploads/HBL.xls",
        "mcb": "/mnt/user-data/uploads/MCB_July_5655.xls",
        "faysal": "/mnt/user-data/uploads/Faysal_6018-.xls",
        "bafl": "/mnt/user-data/uploads/BAFL_July_till_date.csv",
        "js": "/mnt/user-data/uploads/JS_107791.xlsx",
    }
    results, errors = parse_batch([(fp, bank) for bank, fp in samples.items()])
    conventional_balances = balances_from_statements(results)
    print("Derived conventional balances from statements:", conventional_balances)

    # Simulate a "manual portal override" for one bank, proving manual wins
    # over statement when both are supplied (build a dict with the manual
    # value taking precedence by merging it in last).
    conventional_balances["hbl"] = 3050000.00  # pretend the portal shows a slightly newer number

    remittances = [
        {"reinsurer": "Tysers Singapore", "insured_name": "BSPL", "amount": 61574160, "currency": "PKR"},
        {"reinsurer": "Marsh Sigorta", "insured_name": "Zorlu Energi", "amount": 999400, "currency": "PKR"},
        {"reinsurer": "Marsh Sigorta", "insured_name": "Zorlu Energi", "amount": 4052639, "currency": "PKR"},
        {"reinsurer": "TEST OVERFLOW Reinsurer", "insured_name": "Overflow Test", "amount": 123456, "currency": "PKR"},
    ]
    major_payments = [{"name": "Revenue Board", "amount": 80000000, "bank": "SCB", "expected_date": date(2026, 7, 17)}]
    major_receivables = [
        {"name": "Al Mumtaz Electronics", "amount": 3000000, "bank": "SCB", "expected_date": date(2026, 7, 9)},
        {"name": "CMS Pakistan", "amount": 7500000, "bank": "SCB", "expected_date": date(2026, 7, 9)},
    ]

    out_path = "/home/claude/Bank_Position_TEST_OUTPUT.xlsx"
    warnings = generate_bank_position(
        previous_file="/mnt/user-data/uploads/Bank_Position_16th_July.xlsx",
        output_file=out_path,
        conventional_balances=conventional_balances,
        islamic_balances={},
        remittances=remittances,
        major_payments=major_payments,
        major_receivables=major_receivables,
        as_of_date=date(2026, 7, 19),
    )
    print("\nWarnings:")
    for w in warnings:
        print(" -", w.message)
    print(f"\nGenerated: {out_path}")
