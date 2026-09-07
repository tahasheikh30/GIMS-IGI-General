"""
Deferred Commission Expense (Finance module).

WHAT THIS DOES: given the Premium Register file(s) (the same "Gross
Premium - Conventional Register" files UPR Automation already consumes)
and three commission/service-charge LEDGER exports - Commission Expense
Direct, Commission Expense Coinsurance Accepted, and Service Charges
Coinsurance - this computes, for every policy in the Premium Register,
how much commission expense/service charge was actually posted against
it, sums the three ledgers into one "Total Commission" figure per
policy, and then amortizes that Total Commission the same way UPR
Automation amortizes Gross Premium (Previous / Jan..Dec / Future
Policies bucketing by Comm Date, 1/24 for most departments, 1/6 for
Marine/Travel, zero for BIMA) to arrive at Closing Deferred Commission
Expense.

MATCHING, per ledger row:
  1. The ledger's own "Manual / Ref. No." column is the policy reference
     - but a single cell can carry TWO references glued together with a
       "~" separator (e.g. the current transaction's own doc ref, then
       the original/base doc ref it renewed from -
       "2025/01/HOVPCDE00024~2024/10/HOVPCDP05157"). Only the FIRST
       reference (before the "~") is the one that actually matches the
       Premium Register's own "Doc Ref #" column - confirmed directly
       against a real finished working file, where the "Particulars"
       column used for matching is exactly the text before the "~".
       A cell with no "~" at all is used as-is.
  2. Balance = Debit - Credit for that row (a debit increases the
     commission expense balance, a credit reduces it - same sign
     convention as any other GL ledger in this project).
  3. Every ledger row's Balance is summed by its own (post-split) policy
     reference - a policy can have several ledger rows (e.g. an original
     posting plus a later reversal/adjustment), and all of them belong
     to that one policy's total.
  4. That per-policy sum is matched against the Premium Register's own
     "Doc Ref #" and added as a new column on that register row - one
     column per ledger (Commission Expense - Direct / Commission Expense
     - Coinsurance / Service Charges - Coinsurance). A register row with
     no matching ledger entry at all gets 0, not a blank - a policy
     genuinely can have no commission posted against it yet.

Once every register row has all three ledger columns, "Total Commission"
= the sum of the three, and everything downstream (period bounds,
Previous/Jan..Dec/Future bucketing by Comm Date, the per-department
1/24 vs 1/6 vs zero amortization method) is IDENTICAL to UPR Automation
and Reinsurance Automation - reused directly from upr_processor.py
rather than re-implemented a third time, the same way reinsurance_
processor.py already reuses it (see that module's own docstring for the
full rationale: none of that logic cares what the amortized figure
actually represents, only that the source data carries a "Department"
and a "Comm Date").
"""
import os
import re

import pandas as pd

from premium_register_processor import _read_workbook_any_engine
import upr_processor as up

try:
    from perf_utils import fast_read_excel_rows, estimate_row_count
    _PERF_AVAILABLE = True
except ImportError:  # perf_utils.py not yet added to this checkout, degrade gracefully
    _PERF_AVAILABLE = False

FAST_READ_ROW_THRESHOLD = 100000

TOTAL_COMMISSION_LABEL = "Total Commission"

LEDGER_COLUMNS = {
    "direct": "Commission Expense - Direct",
    "coinsurance": "Commission Expense - Coinsurance",
    "service_charges": "Service Charges - Coinsurance",
}


class DeferredCommissionError(Exception):
    """Raised for a problem that should stop the run with a clear,
    actionable message - a missing required column in the register or a
    ledger file, or (reusing upr_processor.UPRError's own check directly)
    the register(s) not covering the full selected period."""
    pass


# ---------------------------------------------------------------------------
# Ledger loading
# ---------------------------------------------------------------------------

_LEDGER_ALIASES = {
    "Manual Ref No": ["manual / ref no", "manual/ref no", "manual ref no"],
    "Debit": ["debit"],
    "Credit": ["credit"],
}
_LEDGER_ALIAS_LOOKUP = {}
for _canon, _aliases in _LEDGER_ALIASES.items():
    for _a in _aliases:
        _LEDGER_ALIAS_LOOKUP[_a] = _canon


def _normalize_header(value):
    """Same normalization as every other GIMS loader (premium_register_
    processor._normalize_header et al.) - periods stripped, whitespace
    collapsed, lowercased - so 'Manual / Ref. No.' matches the alias
    'manual / ref no' regardless of exactly how the period is punctuated."""
    text = str(value)
    text = re.sub(r"[()]", " ", text)
    text = text.replace(".", "")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def _find_ledger_header(preview_df, max_scan=10):
    """Scan the first `max_scan` rows for the one carrying 'Manual / Ref.
    No.', 'Debit', and 'Credit' all together - returns (header_row_index,
    (manual_col, debit_col, credit_col)) as 0-indexed column positions
    within that row, or (None, None) if no such row is found in this
    sheet (e.g. a headerless continuation sheet - see load_ledger_balances'
    own docstring for how that's handled)."""
    for r in range(min(max_scan, len(preview_df))):
        norm = {}
        for c, v in enumerate(preview_df.iloc[r].tolist()):
            if pd.isna(v):
                continue
            key = _normalize_header(v)
            if key in _LEDGER_ALIAS_LOOKUP:
                norm[_LEDGER_ALIAS_LOOKUP[key]] = c
        if "Manual Ref No" in norm and "Debit" in norm and "Credit" in norm:
            return r, (norm["Manual Ref No"], norm["Debit"], norm["Credit"])
    return None, None


def _read_full_sheet(xls, resolved_filepath, filepath, sheet_name, report=None):
    ext = os.path.splitext(filepath)[1].lower()
    if _PERF_AVAILABLE and ext == ".xlsx":
        est_rows = estimate_row_count(resolved_filepath, sheet_name)
        if est_rows and est_rows > FAST_READ_ROW_THRESHOLD:
            if report:
                report(f"    Large sheet detected (~{est_rows:,} rows) - using the fast streaming reader.")
            rows, _n = fast_read_excel_rows(resolved_filepath, sheet_name, report=report)
            return pd.DataFrame(rows)
    return pd.read_excel(xls, sheet_name=sheet_name, header=None)


def _split_policy_ref(series):
    """Vectorized equivalent of str(v).split('~')[0].strip() - keeps only
    the FIRST reference in a '~'-joined cell (see the module docstring's
    MATCHING section for why: the finished reference workbook's own
    matching column is exactly this first segment). A cell with no '~'
    at all passes through unchanged, just stripped."""
    return series.astype(str).str.split("~", n=1).str[0].str.strip()


def load_ledger_balances(filepaths, report=None):
    """
    Load one or more ledger export files (all sharing the same real
    layout: Account Code / Name / Voucher Date / ... / Manual / Ref. No.
    / ... / Debit / Credit / Balance / Create User) and return the
    NET Balance (Debit - Credit) summed per policy reference, across
    every row in every file/sheet given.

    A file can have more than one sheet where only the FIRST carries a
    real header (confirmed real case: a ledger split into Sheet1/Sheet2
    with no header of its own on Sheet2, same "continuation sheet"
    shape every other GIMS register loader already handles) - a sheet
    with no detectable header reuses whichever header/column layout was
    last found earlier in the SAME file, on the assumption it's a bare
    continuation of the same table. A sheet where no header has ever
    been found yet (nothing to inherit) is skipped with a warning rather
    than guessed at.

    A row is skipped only if its own Manual/Ref No cell is blank - this
    is what excludes the account-subtotal/header-repeat rows real
    exports carry (e.g. a plain "0502001004  COMMISSION EXPENSE -
    COINSURANCE ACCEPTED" row with every other column blank) without
    needing to special-case them by name.

    Returns (balances, meta):
      balances - {policy_ref: net_balance} - a plain dict, one entry per
          distinct policy reference found across every file.
      meta - {"total_rows": int, "sheets_used": [(filename, sheet), ...],
          "sheets_skipped": [(filename, sheet, reason), ...]}
    """
    if not filepaths:
        return {}, {"total_rows": 0, "sheets_used": [], "sheets_skipped": []}

    balances = {}
    total_rows = 0
    sheets_used = []
    sheets_skipped = []

    for filepath in filepaths:
        if not os.path.exists(filepath):
            raise DeferredCommissionError(f"File not found: {filepath}")
        fname = os.path.basename(filepath)
        xls, resolved_filepath = _read_workbook_any_engine(filepath)

        colmap = None  # (manual_col, debit_col, credit_col) - carries over to a headerless continuation sheet
        for sheet_name in xls.sheet_names:
            preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=10)
            header_row, found_colmap = _find_ledger_header(preview)

            if header_row is not None:
                colmap = found_colmap
                data_start = header_row + 1
            elif colmap is not None:
                data_start = 0  # a bare continuation of the previous sheet's own table - every row is data
            else:
                sheets_skipped.append((fname, sheet_name, "no recognizable 'Manual / Ref. No.' / Debit / Credit header found"))
                continue

            if report:
                report(f"  Reading '{fname}' - sheet '{sheet_name}'...")
            raw = _read_full_sheet(xls, resolved_filepath, filepath, sheet_name, report=report)
            sub = raw.iloc[data_start:]
            if len(sub) == 0:
                sheets_used.append((fname, sheet_name))
                continue

            manual_col, debit_col, credit_col = colmap
            if max(manual_col, debit_col, credit_col) >= sub.shape[1]:
                sheets_skipped.append((fname, sheet_name, "column layout from an earlier sheet doesn't fit this sheet's own width"))
                continue

            manual_series = sub.iloc[:, manual_col]
            mask = manual_series.notna() & (manual_series.astype(str).str.strip() != "")
            if not mask.any():
                sheets_used.append((fname, sheet_name))
                continue

            policy_ref = _split_policy_ref(manual_series[mask])
            debit = pd.to_numeric(sub.iloc[:, debit_col][mask], errors="coerce").fillna(0.0)
            credit = pd.to_numeric(sub.iloc[:, credit_col][mask], errors="coerce").fillna(0.0)
            balance = debit - credit

            grouped = balance.groupby(policy_ref.values).sum()
            for ref, amt in grouped.items():
                balances[ref] = balances.get(ref, 0.0) + float(amt)

            total_rows += int(mask.sum())
            sheets_used.append((fname, sheet_name))

    meta = {"total_rows": total_rows, "sheets_used": sheets_used, "sheets_skipped": sheets_skipped}
    return balances, meta


# ---------------------------------------------------------------------------
# Matching ledger balances onto the Premium Register, then reusing UPR's
# own period/bucketing/amortization logic unchanged.
# ---------------------------------------------------------------------------

def build_commission_table(bucketed_df, departments):
    """Table 1 equivalent for this module: Department x [Previous,
    Jan..Dec, Future Policies] grid of summed Total Commission - same
    shape/contract as upr_processor.build_gross_premium_table /
    reinsurance_processor.build_ceded_table, just summing the
    TOTAL_COMMISSION_LABEL column this module computes instead."""
    grouped = bucketed_df.groupby(["Department", "Bucket"])[TOTAL_COMMISSION_LABEL].sum()
    table = {}
    for dept in departments:
        row = {}
        for col in up.TABLE_COLUMNS:
            row[col] = float(grouped.get((dept, col), 0.0))
        table[dept] = row
    return table


def match_commission_to_register(register_df, direct_balances, coinsurance_balances, service_charge_balances):
    """
    Add the three ledger columns plus TOTAL_COMMISSION_LABEL onto a COPY
    of register_df, matched by its own "Doc Ref" column (as loaded by
    upr_processor.load_premium_registers - the canonical name for the
    register's own "Doc Ref #" / "Document No." column) against each
    ledger's own {policy_ref: balance} dict. A register row with no
    matching entry in a given ledger gets 0 for that ledger's column,
    not a blank.
    """
    if "Doc Ref" not in register_df.columns:
        raise DeferredCommissionError(
            "The Premium Register file(s) don't have a recognizable Doc Ref # / Document No. column - "
            "this is required to match commission/service-charge ledger balances to a policy. Add the "
            "real header text to upr_processor.ALIASES's own 'Doc Ref' entry if this is just a naming mismatch."
        )
    out = register_df.copy()
    doc_ref = out["Doc Ref"].astype(str).str.strip()
    out[LEDGER_COLUMNS["direct"]] = doc_ref.map(direct_balances).fillna(0.0)
    out[LEDGER_COLUMNS["coinsurance"]] = doc_ref.map(coinsurance_balances).fillna(0.0)
    out[LEDGER_COLUMNS["service_charges"]] = doc_ref.map(service_charge_balances).fillna(0.0)
    out[TOTAL_COMMISSION_LABEL] = (
        out[LEDGER_COLUMNS["direct"]] + out[LEDGER_COLUMNS["coinsurance"]] + out[LEDGER_COLUMNS["service_charges"]]
    )
    return out


# ---------------------------------------------------------------------------
# Period selection / bucketing / amortization - reused directly from
# upr_processor.py, unchanged (see reinsurance_processor.py for the exact
# same reuse pattern and rationale - neither of these cares what the
# amortized figure represents, only that the data carries a "Department"
# and a "Comm Date").
# ---------------------------------------------------------------------------

get_available_period_options = up.get_available_period_options
compute_period_bounds = up.compute_period_bounds
validate_coverage = up.validate_coverage
assign_buckets = up.assign_buckets
resolve_department_list = up.resolve_department_list
row_total = up.row_total
build_deferred_table = up.build_unearned_table  # produces the *Deferred* (still-unamortized) commission table
column_totals = up.column_totals
grand_total = up.grand_total
TABLE_COLUMNS = up.TABLE_COLUMNS
MONTH_NAMES = up.MONTH_NAMES
PREVIOUS_LABEL = up.PREVIOUS_LABEL
FUTURE_LABEL = up.FUTURE_LABEL
TOTAL_LABEL = up.TOTAL_LABEL


def run_deferred_commission_calculation(register_filepaths, direct_ledger_filepaths, coinsurance_ledger_filepaths,
                                         service_charge_ledger_filepaths, period_year, period_month, report=None):
    """
    Full pipeline: load the Premium Register(s) -> load & aggregate all
    three ledgers -> match each ledger's balance onto every register row
    by Doc Ref # -> sum into Total Commission -> validate period coverage
    -> bucket by Comm Date -> amortize per department (1/24 / 1/6 / zero)
    -> Closing Deferred Commission Expense.

    Returns a result dict:
      {
        "period_start", "period_end": date,
        "departments": [str, ...],
        "methods": {department: "1/24"|"1/6"|"zero"},
        "commission_table": {department: {column: amount}},   # Table 1
        "deferred_table": {department: {column: amount}},     # Table 2
        "total_commission": float,      # grand total of commission_table
        "closing_deferred_commission": float,   # grand total of deferred_table
        "warnings": [str, ...],
        "excluded_rows": int,
        "sheets_used": [(filename, sheet_name), ...],
        "ledger_meta": {"direct": {...}, "coinsurance": {...}, "service_charges": {...}},
      }
    Raises DeferredCommissionError / upr_processor.UPRError with a clear,
    user-facing message if a file can't be read, is missing a required
    column, or the register(s) don't cover the full selected period.
    """
    df, meta = up.load_premium_registers(register_filepaths, report=report)
    if report:
        report(f"Loaded {len(df):,} usable premium register row(s) from {len(meta['sheets_used'])} sheet(s) "
               f"({meta['excluded_rows']:,} row(s) excluded - missing Department/Comm Date/Gross Premium).")

    if report:
        report("\nLoading Commission Expense - Direct ledger(s)...")
    direct_balances, direct_meta = load_ledger_balances(direct_ledger_filepaths, report=report)
    if report:
        report(f"  {len(direct_balances):,} distinct policy reference(s), {direct_meta['total_rows']:,} ledger row(s).")

    if report:
        report("\nLoading Commission Expense - Coinsurance Accepted ledger(s)...")
    coinsurance_balances, coinsurance_meta = load_ledger_balances(coinsurance_ledger_filepaths, report=report)
    if report:
        report(f"  {len(coinsurance_balances):,} distinct policy reference(s), {coinsurance_meta['total_rows']:,} ledger row(s).")

    if report:
        report("\nLoading Service Charges - Coinsurance ledger(s)...")
    service_charge_balances, service_charge_meta = load_ledger_balances(service_charge_ledger_filepaths, report=report)
    if report:
        report(f"  {len(service_charge_balances):,} distinct policy reference(s), {service_charge_meta['total_rows']:,} ledger row(s).")

    matched_df = match_commission_to_register(df, direct_balances, coinsurance_balances, service_charge_balances)
    num_matched_rows = int((matched_df[TOTAL_COMMISSION_LABEL] != 0).sum())
    if report:
        report(f"\n{num_matched_rows:,} of {len(matched_df):,} premium register row(s) matched at least one "
               f"ledger balance.")

    period_start, period_end = up.compute_period_bounds(period_year, period_month)
    if report:
        report(f"Selected period: {period_start:%d-%b-%Y} to {period_end:%d-%b-%Y}.")

    up.validate_coverage(matched_df, period_start, period_end)

    bucketed = up.assign_buckets(matched_df, period_start, period_end)
    departments, dept_warnings = up.resolve_department_list(bucketed)

    commission_table = build_commission_table(bucketed, departments)
    deferred_table, methods = up.build_unearned_table(commission_table, departments, period_start)
    closing_deferred_commission = up.grand_total(deferred_table, departments)
    total_commission = up.grand_total(commission_table, departments)

    if report:
        report(f"\nTotal Commission (period): {total_commission:,.2f}")
        report(f"Closing Deferred Commission Expense: {closing_deferred_commission:,.2f}")

    return {
        "period_start": period_start,
        "period_end": period_end,
        "departments": departments,
        "methods": methods,
        "commission_table": commission_table,
        "deferred_table": deferred_table,
        "total_commission": total_commission,
        "closing_deferred_commission": closing_deferred_commission,
        "warnings": dept_warnings,
        "excluded_rows": meta["excluded_rows"],
        "sheets_used": meta["sheets_used"],
        "ledger_meta": {"direct": direct_meta, "coinsurance": coinsurance_meta, "service_charges": service_charge_meta},
    }


if __name__ == "__main__":
    register_files = ["/mnt/user-data/uploads/Gross_Premium_-_Conventional_Register_2025_v2.xls",
                       "/mnt/user-data/uploads/Gross_Premium_-_Conventional_Register_June_2026_v2.xls"]
    direct_files = ["/mnt/user-data/uploads/Conv_-_Commission_Exp_Direct__2025_.xlsx",
                     "/mnt/user-data/uploads/Conv_-_Commission_Exp_Direct__Jan_2026_-_June_2026_.xlsx"]
    coinsurance_files = ["/mnt/user-data/uploads/Conv_-_Commission_Exp_Coinsurance_Accepted__2025_.xlsx",
                          "/mnt/user-data/uploads/Conv_-_Commission_Exp_Coinsurance_Accepted__Jan_2026_-_June_2026_.xlsx"]
    service_charge_files = ["/mnt/user-data/uploads/Conv_-_Service_Charges_Coinsurance__Dec_2025_.xlsx",
                             "/mnt/user-data/uploads/Conv_-_Service_Charges_Coinsurance__Jan_2026_-_June_2026_.xlsx"]

    result = run_deferred_commission_calculation(
        register_files, direct_files, coinsurance_files, service_charge_files, 2026, 6, report=print)
    print("\nDepartments:", result["departments"])
    print("Methods:", result["methods"])
    print("Total Commission:", f"{result['total_commission']:,.2f}")
    print("Closing Deferred Commission Expense:", f"{result['closing_deferred_commission']:,.2f}")
    for dept in result["departments"]:
        print(f"  {dept:15s} commission_total={row_total(result['commission_table'][dept]):>15,.0f}  "
              f"deferred_total={row_total(result['deferred_table'][dept]):>15,.0f}")
