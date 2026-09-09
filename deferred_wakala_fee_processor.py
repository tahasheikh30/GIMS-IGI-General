"""
Deferred Wakala Fee (Takaful Finance module).

WHAT THIS DOES: given one or more uploaded Contribution Register files
(the same Takaful Contribution Register upr_processor.py's own Takaful
mode already knows how to read - Department/Class, Comm Date, Gross
Contribution, plus this module's own additional required column, Admin)
and a Wakala Fee % entered by hand for each department, this computes,
for every register row:

    Wakala Fee = (Gross Contribution + Admin) x Wakala Fee %
                 (that row's own Department)

- mirroring the real, manually-built "Deferred Wakala Fees Working"
reference workbook this module was reverse-engineered from exactly (its
own per-row "Wakala Fee" column is an XLOOKUP of the row's Department
against a small Department -> % table, multiplied against
Gross Contribution + Admin).

Once every row has its own Wakala Fee figure, the REST of the working is
deliberately identical to UPR Automation (upr_processor.py) - same
period selection, same "Previous / Jan..Dec / Future Policies" bucketing
by Comm Date, and the same per-department 1/24 (most departments) / 1/6
(Marine, Travel) / zero (BIMA) amortization method, producing a Closing
Deferred Wakala Fees figure the same way UPR produces a Closing Unearned
Premium/Contribution Revenue figure. Rather than re-implement that
shared logic a second time, this module imports and reuses
upr_processor's own period/bucketing/amortization functions directly -
see reinsurance_processor.py / deferred_commission_processor.py for the
identical reuse pattern, applied here to a third kind of amortized
figure.

A department with no percentage entered (left blank, typed 0, or a
genuinely unrecognized department name found in the source data)
defaults to 0% rather than raising - matches the reference workbook's
own XLOOKUP(...,0,0) fallback for an unmapped department, and means a
new/unexpected department name in the data never silently blocks the
whole run.
"""
import os
import re

import pandas as pd

from premium_register_processor import _read_workbook_any_engine
from claims_register_processor import _read_full_table_with_continuations
import upr_processor as up

WAKALA_FEE_LABEL = "Wakala Fee"


class DeferredWakalaFeeError(Exception):
    """Raised for a problem that should stop the run with a clear,
    actionable message - a missing required column in the Contribution
    Register, or (reusing upr_processor.UPRError's own check directly)
    the register(s) not covering the full selected period."""
    pass


# ---------------------------------------------------------------------------
# Column alias matching - same pattern as every other GIMS loader. Only
# Department, Comm Date, Gross Contribution, and Admin are REQUIRED (the
# four inputs the Wakala Fee formula itself needs); Doc Ref is carried
# through if present, purely for context, never required.
#
# "Gross Contribution" accepts both this module's own Takaful-only
# wording AND the combined "Gross Premium/Contribution" wording (some
# real Contribution Register exports use one, some the other - confirmed
# against two different real sample shapes) - either way the figure
# lands in the same canonical column below.
# ---------------------------------------------------------------------------
ALIASES = {
    "Department": ["department/class", "department / class", "department", "class", "class of business"],
    "Comm Date": ["comm date", "commencement date", "comm. date", "comm  date"],
    "Gross Contribution": [
        "gross contribution", "gross  contribution", "grosscontribution", "contribution",
        "gross premium/contribution", "gross premium / contribution", "gross premium contribution",
    ],
    "Admin": ["admin", "admin fee", "admin charges", "admin  fee"],
    "Doc Ref": ["doc ref #", "doc ref#", "doc ref number", "document no", "document number", "doc ref"],
}

REQUIRED_CANONICAL = ["Department", "Comm Date", "Gross Contribution", "Admin"]


def _normalize_header(value):
    text = str(value).replace("\n", " ")
    text = re.sub(r"[()]", " ", text)
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def _build_alias_lookup():
    lookup = {}
    for canonical, aliases in ALIASES.items():
        for a in aliases:
            lookup[a] = canonical
    return lookup


_ALIAS_LOOKUP = _build_alias_lookup()


def _map_columns_to_canonical(columns):
    mapping = {}
    for col in columns:
        norm = _normalize_header(col)
        if norm in _ALIAS_LOOKUP:
            mapping[col] = _ALIAS_LOOKUP[norm]
    return mapping


def _find_header_row(preview_df, max_scan=25, min_score=3):
    """Same banner/title-row-skipping scan as every other GIMS loader
    (upr_processor._find_header_row et al.) - returns
    (row_index_or_None, score)."""
    best_row, best_score = None, 0
    for i in range(min(max_scan, len(preview_df))):
        score = 0
        for v in preview_df.iloc[i].tolist():
            if pd.isna(v):
                continue
            if _normalize_header(v) in _ALIAS_LOOKUP:
                score += 1
        if score > best_score:
            best_score, best_row = score, i
    if best_score < min_score:
        return None, best_score
    return best_row, best_score


def _pick_best_sheet(xls, max_scan=25, min_score=3):
    best = None  # (score, sheet_name, header_row)
    for sn in xls.sheet_names:
        preview = pd.read_excel(xls, sheet_name=sn, header=None, nrows=max_scan)
        header_row, score = _find_header_row(preview, max_scan=max_scan, min_score=min_score)
        if header_row is None:
            continue
        if best is None or score > best[0]:
            best = (score, sn, header_row)
    if best is None:
        raise DeferredWakalaFeeError(
            "Could not auto-detect the header row in any sheet of this workbook "
            f"(no sheet had a row matching at least {min_score} of the expected columns: "
            "Department/Class, Comm Date, Gross Contribution, Admin)."
        )
    _, sheet_name, header_row = best
    return sheet_name, header_row


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

def load_contribution_registers(filepaths, report=None):
    """
    Load and concatenate every uploaded Contribution Register file. Each
    file's own real header row is auto-detected (skipping any banner/
    title rows above it, same as every other GIMS register loader), and
    a legacy .xls file split across several sheets at Excel's old 65,536-
    row-per-sheet cap is merged back into one table via
    claims_register_processor._read_full_table_with_continuations
    (reused directly - same fix already applied everywhere else in GIMS
    that reads this kind of register).

    Returns (df, meta):
      df - concatenated rows with canonical columns "Department",
          "Comm Date" (datetime), "Gross Contribution" (float), "Admin"
          (float, blank treated as 0 - a genuinely missing Admin figure
          on an otherwise valid row isn't reason enough to drop the
          whole row), and "Doc Ref" if that column was found in at least
          one file. A row is excluded only for a missing/unparseable
          Comm Date, a missing/non-numeric Gross Contribution, a blank
          Department, or a Comm Date implausibly far in the past (same
          corruption guard as upr_processor.load_premium_registers,
          reused via the same MIN_PLAUSIBLE_COMM_YEAR bound).
      meta - {"sheets_used": [(filename, sheet_name), ...],
          "excluded_rows": int, "total_rows_read": int}

    Raises DeferredWakalaFeeError if a file is missing one of the
    required columns (Department, Comm Date, Gross Contribution, Admin).
    """
    if not filepaths:
        raise DeferredWakalaFeeError("No Contribution Register file(s) provided.")

    frames = []
    sheets_used = []

    for filepath in filepaths:
        if not os.path.exists(filepath):
            raise DeferredWakalaFeeError(f"File not found: {filepath}")
        fname = os.path.basename(filepath)

        xls, _resolved_filepath = _read_workbook_any_engine(filepath)
        sheet_name, header_row = _pick_best_sheet(xls)
        if report:
            report(f"Reading '{fname}' - sheet '{sheet_name}'...")

        data, _source_note = _read_full_table_with_continuations(
            filepath, xls, sheet_name, header_row, _ALIAS_LOOKUP, report=report)
        sheets_used.append((fname, sheet_name))

        colmap = _map_columns_to_canonical(data.columns)
        inv = {}
        for orig, canon in colmap.items():
            inv.setdefault(canon, orig)

        missing = [c for c in REQUIRED_CANONICAL if c not in inv]
        if missing:
            raise DeferredWakalaFeeError(
                f"'{fname}' is missing required column(s) {missing}. Recognized headers were: "
                f"{list(data.columns)}. Add the real header text to ALIASES in "
                "deferred_wakala_fee_processor.py."
            )

        keep_cols = {inv[c]: c for c in REQUIRED_CANONICAL}
        if "Doc Ref" in inv:
            keep_cols[inv["Doc Ref"]] = "Doc Ref"
        sub = data[list(keep_cols.keys())].rename(columns=keep_cols).copy()
        sub["Source File"] = fname
        frames.append(sub)

    df = pd.concat(frames, ignore_index=True, sort=False)
    df["Comm Date"] = pd.to_datetime(df["Comm Date"], errors="coerce")
    gross_numeric = pd.to_numeric(df["Gross Contribution"], errors="coerce")
    admin_numeric = pd.to_numeric(df["Admin"], errors="coerce")
    df["Gross Contribution"] = gross_numeric
    df["Admin"] = admin_numeric.fillna(0.0)
    df["Department"] = df["Department"].astype(str).str.strip()

    base_valid_mask = (
        df["Comm Date"].notna()
        & df["Gross Contribution"].notna()
        & (df["Department"] != "")
        & (df["Department"].str.lower() != "nan")
    )
    low_plausible_mask = df["Comm Date"].dt.year >= up.MIN_PLAUSIBLE_COMM_YEAR
    valid_mask = base_valid_mask & low_plausible_mask

    excluded_rows = int((~valid_mask).sum())
    clean_df = df.loc[valid_mask].reset_index(drop=True)

    meta = {
        "sheets_used": sheets_used,
        "excluded_rows": excluded_rows,
        "total_rows_read": len(df),
    }
    return clean_df, meta


# ---------------------------------------------------------------------------
# Wakala Fee computation
# ---------------------------------------------------------------------------

def compute_wakala_fee_column(df, wakala_pct_by_department):
    """
    Adds the WAKALA_FEE_LABEL ("Wakala Fee") column to a COPY of df:
    (Gross Contribution + Admin) x that row's own Department's Wakala
    Fee % - mirrors the real reference workbook's own per-row XLOOKUP
    formula exactly (see the module docstring).

    wakala_pct_by_department: {department_name: percent_as_number}, e.g.
        {"Fire": 5.0} for 5% - PERCENT values, not fractions (divided by
        100 internally, same convention as every other percentage entry
        field in GIMS, e.g. Loss Ratio's own threshold field). Department
        name matching is case/whitespace-insensitive (normalized the same
        way upr_processor.resolve_department_list already does for its
        own DEPARTMENT_ORDER matching). A department with no entry, or a
        department name found in the data that isn't in this dict at
        all, defaults to 0% - see the module docstring for why this is
        the correct fallback, not an error.
    """
    out = df.copy()
    pct_lookup = {str(k).strip().upper(): v for k, v in (wakala_pct_by_department or {}).items()}
    dept_upper = out["Department"].astype(str).str.strip().str.upper()
    pct_series = pd.to_numeric(dept_upper.map(pct_lookup), errors="coerce").fillna(0.0)
    out[WAKALA_FEE_LABEL] = (out["Gross Contribution"] + out["Admin"]) * (pct_series / 100.0)
    return out


# ---------------------------------------------------------------------------
# Period selection / bucketing / amortization - reused directly from
# upr_processor.py, unchanged (same reuse pattern as reinsurance_
# processor.py / deferred_commission_processor.py - none of this cares
# what the amortized figure represents, only that the data carries a
# "Department" and a "Comm Date").
# ---------------------------------------------------------------------------

get_available_period_options = up.get_available_period_options
compute_period_bounds = up.compute_period_bounds
validate_coverage = up.validate_coverage
assign_buckets = up.assign_buckets
resolve_department_list = up.resolve_department_list
row_total = up.row_total
build_deferred_wakala_table = up.build_unearned_table  # produces the *Deferred* (still-unamortized) Wakala Fee table
column_totals = up.column_totals
grand_total = up.grand_total
TABLE_COLUMNS = up.TABLE_COLUMNS
MONTH_NAMES = up.MONTH_NAMES
PREVIOUS_LABEL = up.PREVIOUS_LABEL
FUTURE_LABEL = up.FUTURE_LABEL
TOTAL_LABEL = up.TOTAL_LABEL
DEPARTMENT_ORDER = up.DEPARTMENT_ORDER


def build_wakala_table(bucketed_df, departments):
    """Table 1 equivalent for this module: Department x [Previous,
    Jan..Dec, Future Policies] grid of summed Wakala Fee - same shape/
    contract as upr_processor.build_gross_premium_table /
    reinsurance_processor.build_ceded_table /
    deferred_commission_processor.build_commission_table, just summing
    the WAKALA_FEE_LABEL column this module computes instead."""
    grouped = bucketed_df.groupby(["Department", "Bucket"])[WAKALA_FEE_LABEL].sum()
    table = {}
    for dept in departments:
        row = {}
        for col in up.TABLE_COLUMNS:
            row[col] = float(grouped.get((dept, col), 0.0))
        table[dept] = row
    return table


def run_deferred_wakala_fee_calculation(filepaths, wakala_pct_by_department, period_year, period_month, report=None):
    """
    Full pipeline: load the Contribution Register(s) -> compute each
    row's own Wakala Fee (Gross Contribution + Admin, x its Department's
    own %) -> validate period coverage -> bucket by Comm Date -> amortize
    per department (1/24 / 1/6 / zero) -> Closing Deferred Wakala Fees.

    Returns a result dict:
      {
        "period_start", "period_end": date,
        "departments": [str, ...],
        "methods": {department: "1/24"|"1/6"|"zero"},
        "wakala_table": {department: {column: amount}},           # Table 1
        "deferred_table": {department: {column: amount}},         # Table 2
        "gross_wakala_total": float,      # grand total of wakala_table
        "closing_deferred_wakala": float, # grand total of deferred_table
        "warnings": [str, ...],
        "excluded_rows": int,
        "sheets_used": [(filename, sheet_name), ...],
        "wakala_pct_by_department": {department: percent},
      }
    Raises DeferredWakalaFeeError / upr_processor.UPRError with a clear,
    user-facing message if a file can't be read, is missing a required
    column, or the register(s) don't cover the full selected period.
    """
    df, meta = load_contribution_registers(filepaths, report=report)
    if report:
        report(f"Loaded {len(df):,} usable row(s) from {len(meta['sheets_used'])} sheet(s) "
               f"({meta['excluded_rows']:,} row(s) excluded - missing Department/Comm Date/"
               f"Gross Contribution, or an implausibly old Comm Date).")

    df = compute_wakala_fee_column(df, wakala_pct_by_department)
    if report:
        report(f"Computed Wakala Fee for every row: (Gross Contribution + Admin) x that row's own "
               f"Department %  -  total {df[WAKALA_FEE_LABEL].sum():,.2f} across {len(df):,} row(s).")

    period_start, period_end = up.compute_period_bounds(period_year, period_month)
    if report:
        report(f"Selected period: {period_start:%d-%b-%Y} to {period_end:%d-%b-%Y}.")

    up.validate_coverage(df, period_start, period_end)

    bucketed = up.assign_buckets(df, period_start, period_end)
    departments, dept_warnings = up.resolve_department_list(bucketed)

    wakala_table = build_wakala_table(bucketed, departments)
    deferred_table, methods = up.build_unearned_table(wakala_table, departments, period_start)
    closing_deferred_wakala = up.grand_total(deferred_table, departments)
    gross_wakala_total = up.grand_total(wakala_table, departments)

    if report:
        report(f"Gross Wakala Fee (period): {gross_wakala_total:,.2f}")
        report(f"Closing Deferred Wakala Fees: {closing_deferred_wakala:,.2f}")

    return {
        "period_start": period_start,
        "period_end": period_end,
        "departments": departments,
        "methods": methods,
        "wakala_table": wakala_table,
        "deferred_table": deferred_table,
        "gross_wakala_total": gross_wakala_total,
        "closing_deferred_wakala": closing_deferred_wakala,
        "warnings": dept_warnings,
        "excluded_rows": meta["excluded_rows"],
        "sheets_used": meta["sheets_used"],
        "wakala_pct_by_department": dict(wakala_pct_by_department or {}),
    }


if __name__ == "__main__":
    import sys

    files = sys.argv[1:] or [
        "/mnt/user-data/uploads/Contribution_Register_2025.xls",
        "/mnt/user-data/uploads/Contribution_Register_June_2026.xls",
    ]
    wakala_pct = {
        "Fire": 20.0, "Marine": 20.0, "Motor": 20.0, "Miscellaneous": 20.0,
        "Travel": 20.0, "Engineering": 20.0, "Bond": 20.0, "Health": 20.0, "BIMA": 0.0,
    }

    df, meta = load_contribution_registers(files, report=print)
    print("Rows loaded:", len(df), "| excluded:", meta["excluded_rows"])

    options = get_available_period_options(df)
    print(f"\n{len(options)} period option(s), most recent first (showing 5):")
    for o in options[:5]:
        print(" ", o["label"])

    if options:
        latest = options[0]
        result = run_deferred_wakala_fee_calculation(files, wakala_pct, latest["year"], latest["month"], report=print)
        print("\nDepartments:", result["departments"])
        print("Methods:", result["methods"])
        print("\nGross Wakala Fee (period):", f"{result['gross_wakala_total']:,.2f}")
        print("Closing Deferred Wakala Fees:", f"{result['closing_deferred_wakala']:,.2f}")
        for dept in result["departments"]:
            print(f"  {dept:15s} wakala_total={row_total(result['wakala_table'][dept]):>15,.0f}  "
                  f"deferred_total={row_total(result['deferred_table'][dept]):>15,.0f}")
