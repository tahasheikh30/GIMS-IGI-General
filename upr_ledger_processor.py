"""
UPR Automation - Ledger Balance PDF parsing.

Parses the "All Ledgers Balance" report (Report ID: GL_LB_LEDGERBALANCE -
a Crystal Reports-style PDF export of the four GL accounts that feed
Section 1 of the UPR working: Premium Direct, Premium Income Coinsurance
- Accepted, Premium Income Reinsurance - Accepted, and Administrative
Surcharge) so those four figures can be pulled automatically from the
PDF instead of typed in by hand - per the module's own "manual for now,
automated later" trajectory, this is the first piece of that automation:
Section 1's four premium-income figures now come from this PDF if one is
uploaded, while Opening Unearned Premium Revenue stays manual for now
(it isn't on this report at all - it comes from the PRIOR period's own
closing figure, which is a separate future automation).

REPORT SHAPE (confirmed against a real sample export): a short header
block (Voucher Status, As At Date, Location/Account/Fiscal Year filter
values, Report ID, run Date/Time/User), then a "GL Account / Particular
/ Title / Balance" table with one row per account:

    0401001001   PREMIUM DIRECT                          5,842,287,567.00 Cr
    0401001004   PREMIUM INCOME COINSURANCE - ACCEPTED    1,035,145,247.00 Cr
    0401001006   PREMIUM INCOME REINSURANCE - ACCEPTED      8,926,539.00 Cr
    0401002001   ADMINISTRATIVE SURCHARGE                   77,318,060.00 Cr
                                            Grand Total   6,963,677,413.00 Cr

pdfplumber's own plain page.extract_text() (no special layout mode
needed - confirmed against a real sample) returns this as clean,
single-space-separated lines, one data row per line, so this is parsed
with a straightforward per-line regex rather than anything Crystal-
Reports-jumbling-aware like loss_ratio_processor.py's PDF path has to be
for ITS report layout.

MATCHING: each row's own "Particular / Title" text is matched to one of
the four canonical Section 1 fields by keyword (see
_classify_description) rather than by GL code - the account CODE isn't
guaranteed stable across companies/chart-of-accounts, but the account
NAME text ("PREMIUM DIRECT", "...COINSURANCE...", "...REINSURANCE...",
"...SURCHARGE...") is a much more portable signal. A row whose
description doesn't match any of the four is not silently dropped - it's
kept as an "unmatched line" and surfaced to the caller, since a real
ledger-balance run scoped too broadly (e.g. covering GL accounts outside
the four expected ones) could otherwise mean a Section 1 figure quietly
comes out wrong with nothing visibly different about the run.

Cr/Dr: a normal premium-income GL account carries a credit (Cr) balance
- that's the "positive" direction for revenue. If a row is ever reported
Dr instead (a reversal/contra entry outweighing the account's normal
balance), that flips the sign so Section 1 still reflects the true net
figure rather than silently treating a Dr balance as if it were Cr.
"""
import os
import re
from datetime import datetime

try:
    import pdfplumber
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False


class UPRLedgerError(Exception):
    """Raised for a problem reading/parsing the Ledger Balance PDF, with
    a clear, actionable message - pdfplumber not installed, an unreadable
    file, or a PDF that doesn't look like this report at all (none of the
    four expected accounts found)."""
    pass


# Keyword markers checked in this order (most specific first) against
# each row's own "Particular / Title" text, uppercased - see the module
# docstring for why substring/keyword matching is used instead of
# matching on the GL account code.
_FIELD_MARKERS = [
    ("admin_surcharge", ["SURCHARGE"]),
    ("premium_coinsurance", ["COINSURANCE"]),
    ("premium_reinsurance", ["REINSURANCE"]),
    ("premium_direct", ["PREMIUM DIRECT"]),
]

FIELD_LABELS = {
    "premium_direct": "Premium Direct",
    "premium_coinsurance": "Premium Income Coinsurance",
    "premium_reinsurance": "Premium Income Reinsurance",
    "admin_surcharge": "Administrative Surcharge",
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
    amount = float(amount_text.replace(",", ""))
    return amount if dr_cr.strip().upper() == "CR" else -amount


def parse_ledger_balance_pdf(filepath):
    """
    Parse one Ledger Balance PDF. Returns a dict:
      {
        "premium_direct", "premium_coinsurance", "premium_reinsurance",
            "admin_surcharge": float or None (None if that account
            wasn't found anywhere in the report),
        "as_at_date": date or None - parsed from the report's own
            "As At Date" header, for cross-checking against the UPR
            period actually selected,
        "grand_total_reported": float or None - the report's own
            "Grand Total" line, if present,
        "grand_total_matched": float - sum of the four matched fields
            above (None values treated as 0),
        "matched_lines": [{"gl_code", "description", "amount", "field"}],
        "unmatched_lines": [{"gl_code", "description", "amount"}] - GL
            rows found in the report that didn't match any of the four
            expected accounts - not silently dropped, since a report run
            for the wrong account range could otherwise leave a figure
            wrong with nothing visibly different.
      }
    Raises UPRLedgerError if pdfplumber isn't installed, the file can't
    be read, or NONE of the four expected accounts were found at all
    (almost certainly the wrong report).
    """
    if not PDF_AVAILABLE:
        raise UPRLedgerError("pdfplumber isn't installed - reading a Ledger Balance PDF requires 'pip install pdfplumber'.")
    if not os.path.exists(filepath):
        raise UPRLedgerError(f"File not found: {filepath}")

    try:
        lines = []
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                lines.extend(text.split("\n"))
    except Exception as e:
        raise UPRLedgerError(f"Could not read '{os.path.basename(filepath)}' as a PDF ({e})")

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
                # If the same field appears on more than one row (unusual,
                # but possible with a wider account-range run), amounts
                # are summed rather than the later row silently
                # overwriting the earlier one.
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
        raise UPRLedgerError(
            f"Could not find any of the expected accounts (Premium Direct, Premium Income Coinsurance, "
            f"Premium Income Reinsurance, Administrative Surcharge) in '{os.path.basename(filepath)}' - "
            "this may not be the Ledger Balance report, or its account descriptions use different wording "
            "than expected (add the real wording to _FIELD_MARKERS in upr_ledger_processor.py)."
        )

    grand_total_matched = sum(v or 0.0 for v in result.values())

    return {
        **result,
        "as_at_date": as_at_date,
        "grand_total_reported": grand_total_reported,
        "grand_total_matched": grand_total_matched,
        "matched_lines": matched_lines,
        "unmatched_lines": unmatched_lines,
    }


if __name__ == "__main__":
    import sys
    fp = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/UPR_-_All_Ledgers_Balance.pdf"
    result = parse_ledger_balance_pdf(fp)
    for field, label in FIELD_LABELS.items():
        val = result[field]
        print(f"  {label:30s} {val:>18,.2f}" if val is not None else f"  {label:30s} NOT FOUND")
    print(f"  {'Matched total':30s} {result['grand_total_matched']:>18,.2f}")
    if result["grand_total_reported"] is not None:
        diff = result["grand_total_matched"] - result["grand_total_reported"]
        print(f"  {'Report Grand Total':30s} {result['grand_total_reported']:>18,.2f}  (diff: {diff:,.2f})")
    print(f"  As At Date: {result['as_at_date']}")
    if result["unmatched_lines"]:
        print("\n  Unmatched line(s):")
        for u in result["unmatched_lines"]:
            print(f"    {u['gl_code']}  {u['description']}  {u['amount']:,.2f}")
