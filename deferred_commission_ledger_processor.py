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


# GL-CODE based matching (primary) - this module used to classify each
# row purely by matching keywords/phrases against the row's own
# "Particular / Title" text (see _FIELD_MARKERS_FALLBACK below, kept only
# as a safety net). Matching on the GL Account CODE instead is more
# reliable, since a code is a single fixed value with nothing to
# accidentally collide with (the description-based approach had already
# needed two separate confirmed-bug fixes for exactly that - a bare
# "ACCEPTED"/"SERVICE CHARGES" marker each coincidentally matching an
# unrelated account's own description text - see the git history / the
# comments that used to sit here for the specifics).
#
# Codes are normalized by stripping any leading zeros before comparison
# (int() round-trip), so "0502001004" (as printed in the PDF) and
# "502001004" both key to the same entry - the PDF's own GL Account
# column is always written with a leading zero, but this map is written
# without one so it reads the same way the codes are normally spoken/
# written elsewhere in Finance.
GL_CODE_FIELD_MAP = {
    "502001001": "direct",                  # Commission Expense - Direct
    "502001004": "coinsurance",             # Commission Expense - Coinsurance Accepted
    "502002001": "reinsurance_accepted",    # Commission Expense - Reinsurance Accepted
    "503013004": "service_charges",         # Service Charges - Co Insurance
    "503001033": "producer_allowance",      # Business Producer Allowance
}

FIELD_LABELS = {
    "direct": "Commission Expense - Direct",
    "coinsurance": "Commission Expense - Coinsurance",
    "reinsurance_accepted": "Commission Expense - Reinsurance Accepted",
    "service_charges": "Service Charges - Coinsurance",
    "producer_allowance": "Business Producer Allowance",
}

# The five fields above, in the fixed order their SUM is taken in
# wherever "the sum of all five ledger balances" is needed (currently:
# deferred_commission_generator.py's own Commission Paid/Payable figure
# on the Summary sheet, computed from this PDF instead of the three
# per-policy-matched Excel ledgers whenever a PDF has been uploaded for
# the run - see that module's own build_deferred_commission_workbook).
LEDGER_TOTAL_FIELDS = ["direct", "coinsurance", "reinsurance_accepted", "service_charges", "producer_allowance"]

# Description-keyword fallback, used ONLY for a row whose GL code isn't
# in GL_CODE_FIELD_MAP above (e.g. a future export that renumbers the
# chart of accounts) - kept as a safety net so the report doesn't go
# from "matches by code" to "matches nothing" the moment a code changes,
# but GL_CODE_FIELD_MAP is always checked first and wins whenever it has
# an entry.
_FIELD_MARKERS_FALLBACK = [
    ("service_charges", ["SERVICE CHARGES - CO INSURANCE", "SERVICE CHARGES, CO INSURANCE", "SERVICE CHARGES - COINSURANCE"]),
    ("reinsurance_accepted", ["REINSURANCE ACCEPTED"]),
    ("coinsurance", ["COINSURANCE ACCEPTED"]),
    ("direct", ["COMMISSION EXPENSE - DIRECT"]),
]

_DATA_LINE_RE = re.compile(
    r"^\s*(\d{6,15})\s+(.+?)\s+([\d,]+\.\d{2})\s+(Cr|Dr)\s*$", re.IGNORECASE
)
_GRAND_TOTAL_RE = re.compile(
    r"^\s*Grand\s+Total\s+([\d,]+\.\d{2})\s+(Cr|Dr)\s*$", re.IGNORECASE
)
_AS_AT_DATE_RE = re.compile(r"As\s*At\s*Date\s*:\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)


def _normalize_gl_code(gl_code):
    """Strip leading zeros for comparison against GL_CODE_FIELD_MAP - e.g.
    '0502001004' -> '502001004'. Falls back to the original string if it
    isn't purely numeric (shouldn't happen given _DATA_LINE_RE's own
    \\d{6,15} capture, but never worth raising over)."""
    try:
        return str(int(gl_code))
    except ValueError:
        return gl_code


def _classify_row(gl_code, description):
    """GL code first (see GL_CODE_FIELD_MAP's own docstring for why),
    falling back to the old description-keyword match only if this row's
    code isn't one of the four expected ones."""
    field = GL_CODE_FIELD_MAP.get(_normalize_gl_code(gl_code))
    if field:
        return field
    text = description.upper()
    for field, markers in _FIELD_MARKERS_FALLBACK:
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
    Parse one "All Ledgers Balance" PDF for the five GL accounts this
    module cares about, matched by GL Account CODE (see
    GL_CODE_FIELD_MAP): 502001001 (Direct), 502001004 (Coinsurance
    Accepted), 502002001 (Reinsurance Accepted), 503013004 (Service
    Charges - Co Insurance), 503001033 (Business Producer Allowance).
    Returns a dict:
      {
        "direct", "coinsurance", "reinsurance_accepted", "service_charges",
            "producer_allowance": float or None (None if that account
            was not found anywhere in the report),
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
            field = _classify_row(gl_code, description)
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
            "Could not find any of the expected accounts (GL codes 502001001 Commission Expense - Direct, "
            "502001004 Commission Expense - Coinsurance Accepted, 502002001 Commission Expense - Reinsurance "
            "Accepted, 503013004 Service Charges - Co Insurance, 503001033 Business Producer Allowance) in "
            f"'{os.path.basename(filepath)}'. This may not be the Ledger Balance report, or its GL codes have "
            "changed. Update GL_CODE_FIELD_MAP in deferred_commission_ledger_processor.py if the codes have "
            "genuinely changed."
        )

    return {
        **result,
        "as_at_date": as_at_date,
        "grand_total_reported": grand_total_reported,
        "matched_lines": matched_lines,
        "unmatched_lines": unmatched_lines,
    }