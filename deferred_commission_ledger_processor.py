"""
Deferred Commission Expense, Ledger Balance PDF parsing.

Parses the same "All Ledgers Balance" GL report (Report ID:
GL_LB_LEDGERBALANCE) UPR Automation's own upr_ledger_processor.py
already reads, just pulling out the three GL accounts this module cares
about instead of UPR's own four: Commission Expense - Direct, Commission
Expense - Coinsurance Accepted, and Service Charges - Co Insurance. This
lets the three ledger values shown on the Deferred Commission Summary
sheet be filled in automatically from a PDF upload. Only Opening
Deferred Commission Expense stays manual, since it is the prior period's
own Closing figure and is not on this report at all.

Kept as its own module rather than reused from upr_ledger_processor.py
because the sign convention has to be different. UPR's own accounts are
revenue accounts, where a credit (Cr) balance is the normal, positive
direction. This module's three accounts are expense side accounts,
where debit (Dr) is the normal, positive direction, matching
deferred_commission_processor.py's own "Balance = Debit - Credit"
convention exactly. Confirmed against a real sample PDF: Commission
Expense - Direct and Coinsurance both report Dr, while Service Charges,
Co Insurance reports Cr, meaning a net credit balance that reads as a
negative contribution to Debit minus Credit.

Also note this report is a whole company "All Ledgers Balance" export,
not one scoped to just these three accounts. The real sample PDF's own
Grand Total sums well over a hundred unrelated GL lines, salary, rent,
utilities, and so on, alongside the three this module actually needs.
So, unlike UPR's own narrower report, there is no meaningful "does the
matched total equal the report's own Grand Total" check to run here.
The Grand Total line is still parsed, purely for reference, never
compared against.
"""
import os
import re
from datetime import datetime

try:
    import pdfplumber
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False


class DeferredCommissionLedgerError(Exception):
    """Raised for a problem reading or parsing the Ledger Balance PDF,
    with a clear, actionable message, same pattern as
    upr_ledger_processor.UPRLedgerError."""
    pass


# Keyword markers checked in this order (most specific first) against
# each row's own "Particular / Title" text, uppercased. "COINSURANCE
# ACCEPTED" (not just "ACCEPTED") is deliberately the marker for the
# coinsurance field, since a plain "ACCEPTED" would also match the
# report's own separate "COMMISSION EXPENSE - REINSURANCE ACCEPTED"
# line, a real line on the sample PDF, but not one of this module's own
# three ledgers. The full phrase avoids that collision entirely.
#
# CONFIRMED REAL BUG, fixed here: a bare "SERVICE CHARGES" marker for
# the service_charges field also matched a completely unrelated staff-
# cost line on a real report - "0503001039 CONSULTANCY SERVICE CHARGES -
# HEALTH" (a Dr 4,079,280.00 line, nothing to do with coinsurance). Both
# that line AND the real "0503013004 SERVICE CHARGES - CO INSURANCE"
# line (Cr 1,383,303.00) matched the same bare marker, so they were
# silently summed together into one figure (Dr 4,079,280 + Cr -1,383,303
# = 2,695,977) instead of the real Service Charges - Coinsurance balance
# of Cr 1,383,303 (-1,383,303 in this module's Dr-positive convention) -
# a Health-department salary-adjacent cost inflating a coinsurance
# service-charge figure with no error or warning at all. The full phrase
# "SERVICE CHARGES - CO INSURANCE" (plus a comma/no-hyphen variant, in
# case a future export punctuates it slightly differently) only matches
# the genuine account, the same fix already applied to the coinsurance
# marker above for the identical class of collision.
_FIELD_MARKERS = [
    ("service_charges", ["SERVICE CHARGES - CO INSURANCE", "SERVICE CHARGES, CO INSURANCE", "SERVICE CHARGES - COINSURANCE"]),
    ("coinsurance", ["COINSURANCE ACCEPTED"]),
    ("direct", ["COMMISSION EXPENSE - DIRECT"]),
]

FIELD_LABELS = {
    "direct": "Commission Expense - Direct",
    "coinsurance": "Commission Expense - Coinsurance",
    "service_charges": "Service Charges - Coinsurance",
}

_DATA_LINE_RE = re.compile(
    r"^\s*(\d{6,15})\s+(.+?)\s+([\d,]+\.\d{2})\s+(Cr|Dr)\s*$", re.IGNORECASE
)
_GRAND_TOTAL_RE = re.compile(
    r"^\s*Grand\s+Total\s+([\d,]+\.\d{2})\s+(Cr|Dr)\s*$", re.IGNORECASE
)
_AS_AT_DATE_RE = re.compile(r"As\s*At\s*Date\s*:\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)


def _classify_description(description):
    text = description.upper()
    for field, markers in _FIELD_MARKERS:
        if any(m in text for m in markers):
            return field
    return None


def _signed_amount(amount_text, dr_cr):
    """Dr is positive here, Cr is negative. This is the opposite polarity
    of upr_ledger_processor._signed_amount, see the module docstring for
    why: these are expense side accounts, so Debit minus Credit, the
    convention deferred_commission_processor.py already uses, is the
    correct sign here, not the revenue side, credit positive convention
    UPR's own accounts use."""
    amount = float(amount_text.replace(",", ""))
    return amount if dr_cr.strip().upper() == "DR" else -amount


def parse_deferred_commission_ledger_pdf(filepath):
    """
    Parse one "All Ledgers Balance" PDF for the three GL accounts this
    module cares about. Returns a dict:
      {
        "direct", "coinsurance", "service_charges": float or None (None
            if that account was not found anywhere in the report),
        "as_at_date": date or None,
        "grand_total_reported": float or None, the report's own Grand
            Total line, shown for reference only, see the module
            docstring for why it is never compared against,
        "matched_lines": [{"gl_code", "description", "amount", "field"}],
        "unmatched_lines": [{"gl_code", "description", "amount"}], every
            other GL line found (Commission Expense, Reinsurance
            Accepted, salaries, rent, and so on). Not an error, this
            report always carries far more than these three accounts.
      }
    Raises DeferredCommissionLedgerError if pdfplumber is not installed,
    the file cannot be read, or none of the three expected accounts were
    found at all (almost certainly the wrong report).
    """
    if not PDF_AVAILABLE:
        raise DeferredCommissionLedgerError(
            "pdfplumber is not installed, reading a Ledger Balance PDF requires 'pip install pdfplumber'."
        )
    if not os.path.exists(filepath):
        raise DeferredCommissionLedgerError(f"File not found: {filepath}")

    try:
        lines = []
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                lines.extend(text.split("\n"))
    except Exception as e:
        raise DeferredCommissionLedgerError(f"Could not read '{os.path.basename(filepath)}' as a PDF ({e})")

    result = {field: None for field in FIELD_LABELS}
    matched_lines, unmatched_lines = [], []
    grand_total_reported = None
    as_at_date = None

    for line in lines:
        m = _DATA_LINE_RE.match(line)
        if m:
            gl_code, description, amount_text, dr_cr = m.groups()
            amount = _signed_amount(amount_text, dr_cr)
            field = _classify_description(description)
            if field:
                result[field] = (result[field] or 0.0) + amount
                matched_lines.append({"gl_code": gl_code, "description": description.strip(), "amount": amount, "field": field})
            else:
                unmatched_lines.append({"gl_code": gl_code, "description": description.strip(), "amount": amount})
            continue

        gt = _GRAND_TOTAL_RE.match(line)
        if gt:
            grand_total_reported = _signed_amount(gt.group(1), gt.group(2))
            continue

        if as_at_date is None:
            dm = _AS_AT_DATE_RE.search(line)
            if dm:
                try:
                    as_at_date = datetime.strptime(dm.group(1), "%d/%m/%Y").date()
                except ValueError:
                    pass

    if not matched_lines:
        raise DeferredCommissionLedgerError(
            "Could not find any of the expected accounts (Commission Expense - Direct, Commission "
            f"Expense - Coinsurance Accepted, Service Charges - Co Insurance) in "
            f"'{os.path.basename(filepath)}'. This may not be the Ledger Balance report, or its account "
            "descriptions use different wording than expected. Add the real wording to _FIELD_MARKERS "
            "in deferred_commission_ledger_processor.py."
        )

    return {
        **result,
        "as_at_date": as_at_date,
        "grand_total_reported": grand_total_reported,
        "matched_lines": matched_lines,
        "unmatched_lines": unmatched_lines,
    }