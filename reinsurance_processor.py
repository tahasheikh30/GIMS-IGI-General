"""
Reinsurance Automation (Finance module) - Prepaid RI Premium Ceded.

WHAT THIS DOES: given one or more uploaded Reinsurance (RI) Premium Ceded
register files - per the notes this module was built from, there is a
SEPARATE register per department/class, and the register itself carries
NO department column at all, so the department is determined from the
FILE NAME instead (e.g. "Eng__June_2026_v2.xls" -> Engineering,
"RI Fire 2026.xlsx" -> Fire) - this computes, for every policy/row:

    Total Reinsurance Ceded = Treaty Premium + Local Premium + Foreign Premium
                             = A + B + C  (per the notes' own column lettering)

(the register's real header for the first of these three is genuinely
misspelled "Treaty Premiumium" in the source system - confirmed against
a real register export - so that exact spelling is matched here too,
alongside the correctly-spelled form in case a future export fixes it).

Once every row across every uploaded register has its own Total
Reinsurance Ceded figure and its own Department (from its file name),
the REST of the working is deliberately identical to UPR Automation
(upr_processor.py) - same period selection, same "Previous / Jan..Dec /
Future Policies" bucketing by Comm Date, and the same per-department
1/24 (most departments) / 1/6 (Marine, Travel) amortization method,
producing a Closing Prepaid Reinsurance Premium Ceded figure the same
way UPR produces a Closing Unearned Premium Revenue figure. Rather than
re-implement that shared logic a second time, this module imports and
reuses upr_processor's own period/bucketing/amortization functions
directly wherever they're already generic enough to apply unchanged
(they only ever look for a "Department" + "Comm Date" column, both of
which this module also produces) - see the docstring on each function
below for exactly what's reused vs. newly written.
"""
import os
import re

import pandas as pd

from premium_register_processor import _read_workbook_any_engine, _dedupe_columns
from claims_register_processor import _read_full_table_with_continuations
import upr_processor as up

try:
    from perf_utils import fast_read_excel_rows, estimate_row_count
    _PERF_AVAILABLE = True
except ImportError:  # perf_utils.py not yet added to this checkout, degrade gracefully
    _PERF_AVAILABLE = False

FAST_READ_ROW_THRESHOLD = 100000

CEDED_LABEL = "Total Reinsurance Ceded"


class ReinsuranceError(Exception):
    """Raised for a problem that should stop the run with a clear,
    actionable message - a missing required column, a file whose
    department couldn't be determined from its own name, or (reusing
    upr_processor.UPRError's own check directly) the register(s) not
    covering the full selected period."""
    pass


# ---------------------------------------------------------------------------
# Column alias matching - same pattern as every other GIMS loader. Only
# the three amount columns + Comm Date are REQUIRED (per the notes);
# Insured / Sum Insured / Gross Premium are carried through if present,
# purely for context, but never required.
# ---------------------------------------------------------------------------

REGISTER_ALIASES = {
    "Policy": ["policy", "policy no", "policy number", "policy/pmd #", "policy /pmd#"],
    "Comm Date": ["comm date", "commencement date", "comm. date", "comm  date"],
    # "Treaty Premiumium" - the real, misspelled header confirmed on the
    # actual source register export (see module docstring). The correctly
    # spelled form is also listed in case a future export fixes the typo.
    "Treaty Premium": ["treaty premiumium", "treaty premium", "treaty premimium"],
    "Local Premium": ["local premium"],
    "Foreign Premium": ["foreign premium"],
    "Insured": ["insured"],
    "Sum Insured": ["sum insured"],
    "Gross Premium": ["gross premium"],
}

REQUIRED_CANONICAL = ["Comm Date", "Treaty Premium", "Local Premium", "Foreign Premium"]


def _normalize_header(value):
    text = str(value).replace("\n", " ")
    text = re.sub(r"[()]", " ", text)
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def _build_alias_lookup():
    lookup = {}
    for canonical, aliases in REGISTER_ALIASES.items():
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
    (premium_register_processor.find_header_row et al.) - returns
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
        raise ReinsuranceError(
            "Could not auto-detect the header row in any sheet of this workbook "
            f"(no sheet had a row matching at least {min_score} of the expected columns: "
            "Comm Date, Treaty Premiumium, Local Premium, Foreign Premium)."
        )
    _, sheet_name, header_row = best
    return sheet_name, header_row


# ---------------------------------------------------------------------------
# Department from filename - there is no Department column in the
# register at all (per the notes: "Column 'Department/Class' would be
# determined by file name"), since every department has its own,
# separate register file. Matched by whole word first (e.g. the file
# "Eng__2025_v2.xls" tokenizes to ["Eng", "2025", "v2"], and "Eng"
# matches the Engineering keyword exactly), falling back to a plain
# substring check on the whole filename for a keyword that got glued
# directly onto another word with no separator at all.
# ---------------------------------------------------------------------------

DEPARTMENT_FILENAME_KEYWORDS = {
    "Fire": ["fire"],
    "Marine": ["marine"],
    "Motor": ["motor"],
    "Miscellaneous": ["miscellaneous", "misc"],
    "Travel": ["travel"],
    "Engineering": ["engineering", "eng"],
    "Bond": ["bond"],
    "Health": ["health"],
    "BIMA": ["bima"],
}

# Checked in this order so a more specific/longer keyword (e.g.
# "engineering") is tried before a short one that could otherwise
# collide with an unrelated word (e.g. "eng") - matters only for the
# substring fallback below, since the whole-word pass already matches
# every keyword for a department together regardless of order.
_DEPT_KEYWORDS_BY_LENGTH = sorted(
    ((dept, kw) for dept, kws in DEPARTMENT_FILENAME_KEYWORDS.items() for kw in kws),
    key=lambda pair: -len(pair[1]),
)


def detect_department_from_filename(filepath):
    """Return the department name this register file belongs to, or None
    if nothing in the filename matches a known department keyword."""
    base = os.path.splitext(os.path.basename(filepath))[0]
    tokens_lower = [t.lower() for t in re.findall(r"[A-Za-z]+", base)]

    for dept, keywords in DEPARTMENT_FILENAME_KEYWORDS.items():
        if any(kw in tokens_lower for kw in keywords):
            return dept

    base_lower = base.lower()
    for dept, kw in _DEPT_KEYWORDS_BY_LENGTH:
        if kw in base_lower:
            return dept
    return None


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

def _read_full_sheet(xls, resolved_filepath, filepath, sheet_name, report=None):
    ext = os.path.splitext(filepath)[1].lower()
    if _PERF_AVAILABLE and ext == ".xlsx":
        est_rows = estimate_row_count(resolved_filepath, sheet_name)
        if est_rows and est_rows > FAST_READ_ROW_THRESHOLD:
            if report:
                report(f"  Large sheet detected (~{est_rows:,} rows) - using the fast streaming reader.")
            rows, _n = fast_read_excel_rows(resolved_filepath, sheet_name, report=report)
            return pd.DataFrame(rows)
    return pd.read_excel(xls, sheet_name=sheet_name, header=None)


def load_reinsurance_registers(filepaths, report=None):
    """
    Load and concatenate every uploaded RI Premium Ceded register file.
    Each file is its own department's register (see module docstring) -
    the department is determined once per file (from its own filename),
    then stamped onto every row read from that file.

    A file split across several legacy-.xls sheets at Excel's old 65,536-
    row-per-sheet cap is merged back into one table the same way every
    other GIMS register loader already does (reuses
    claims_register_processor._read_full_table_with_continuations
    directly).

    Returns (df, meta):
      df - concatenated rows with canonical columns "Department",
          "Comm Date" (datetime), "Treaty Premium", "Local Premium",
          "Foreign Premium" (all numeric), and "Total Reinsurance Ceded"
          (= the sum of those three, per row - the notes' own A+B+C).
          A row is excluded only for a missing/unparseable Comm Date, or
          one implausibly far in the past (same corruption guard as
          upr_processor.load_premium_registers, reused via the same
          MIN_PLAUSIBLE_COMM_YEAR bound) - a missing amount column value
          is simply treated as 0 for that row rather than excluding the
          whole row, since a real register can legitimately have e.g. no
          Foreign Premium on a purely local-market policy.
      meta - {"departments_by_file": {filename: department}, "sheets_used":
          [(filename, sheet_name), ...], "excluded_rows": int,
          "implausible_date_rows": int}

    Raises ReinsuranceError if a file's department can't be determined
    from its name, or if a file is missing one of the required columns
    (Comm Date, Treaty Premium, Local Premium, Foreign Premium).
    """
    if not filepaths:
        raise ReinsuranceError("No reinsurance premium ceded register file(s) provided.")

    frames = []
    sheets_used = []
    departments_by_file = {}
    implausible_date_rows_total = 0

    for filepath in filepaths:
        if not os.path.exists(filepath):
            raise ReinsuranceError(f"File not found: {filepath}")
        fname = os.path.basename(filepath)

        department = detect_department_from_filename(filepath)
        if department is None:
            raise ReinsuranceError(
                f"Could not determine a department/class for '{fname}' from its own file name. "
                f"Expected one of: {', '.join(DEPARTMENT_FILENAME_KEYWORDS.keys())} to appear "
                "somewhere in the filename (e.g. 'RI Fire 2026.xlsx', 'Eng__2025_v2.xls')."
            )
        departments_by_file[fname] = department

        xls, resolved_filepath = _read_workbook_any_engine(filepath)
        sheet_name, header_row = _pick_best_sheet(xls)
        if report:
            report(f"Reading '{fname}' - sheet '{sheet_name}' (detected department: {department})...")

        data, _source_note = _read_full_table_with_continuations(
            filepath, xls, sheet_name, header_row, _ALIAS_LOOKUP, report=report)
        sheets_used.append((fname, sheet_name))

        colmap = _map_columns_to_canonical(data.columns)
        inv = {}
        for orig, canon in colmap.items():
            inv.setdefault(canon, orig)

        missing = [c for c in REQUIRED_CANONICAL if c not in inv]
        if missing:
            raise ReinsuranceError(
                f"'{fname}' is missing required column(s) {missing}. Recognized headers were: "
                f"{list(data.columns)}. Add the real header text to REGISTER_ALIASES in "
                "reinsurance_processor.py."
            )

        keep_cols = {inv[c]: c for c in REQUIRED_CANONICAL}
        for optional in ("Policy", "Insured", "Sum Insured", "Gross Premium"):
            if optional in inv:
                keep_cols[inv[optional]] = optional
        sub = data[list(keep_cols.keys())].rename(columns=keep_cols).copy()
        sub["Department"] = department
        sub["Source File"] = fname
        frames.append(sub)

    df = pd.concat(frames, ignore_index=True, sort=False)
    df["Comm Date"] = pd.to_datetime(df["Comm Date"], errors="coerce")
    for col in ("Treaty Premium", "Local Premium", "Foreign Premium"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    has_date = df["Comm Date"].notna()
    low_plausible = df["Comm Date"].dt.year >= up.MIN_PLAUSIBLE_COMM_YEAR
    valid_mask = has_date & low_plausible

    implausible_date_rows_total = int((has_date & ~low_plausible).sum())
    excluded_rows = int((~valid_mask).sum())

    clean_df = df.loc[valid_mask].reset_index(drop=True)
    clean_df[CEDED_LABEL] = (
        clean_df["Treaty Premium"] + clean_df["Local Premium"] + clean_df["Foreign Premium"]
    )

    meta = {
        "departments_by_file": departments_by_file,
        "sheets_used": sheets_used,
        "excluded_rows": excluded_rows,
        "implausible_date_rows": implausible_date_rows_total,
        "total_rows_read": len(df),
    }
    return clean_df, meta


# ---------------------------------------------------------------------------
# Period selection, coverage validation, and bucketing - all reused
# directly from upr_processor.py, unchanged. Every one of these already
# operates generically on any dataframe carrying "Department" and
# "Comm Date" columns (which this module's own load_reinsurance_registers
# produces), so there is nothing reinsurance-specific about them that
# would justify a second copy - see each function's own docstring in
# upr_processor.py for the full detail (period math, Previous/Future
# Policies bucketing, the 1/24 vs 1/6 amortization rule per department).
# ---------------------------------------------------------------------------

get_available_period_options = up.get_available_period_options
compute_period_bounds = up.compute_period_bounds
validate_coverage = up.validate_coverage
assign_buckets = up.assign_buckets
resolve_department_list = up.resolve_department_list
build_unearned_table = up.build_unearned_table  # produces the *Prepaid* (still-unamortized) table
row_total = up.row_total
column_totals = up.column_totals
grand_total = up.grand_total
TABLE_COLUMNS = up.TABLE_COLUMNS
MONTH_NAMES = up.MONTH_NAMES
PREVIOUS_LABEL = up.PREVIOUS_LABEL
FUTURE_LABEL = up.FUTURE_LABEL
TOTAL_LABEL = up.TOTAL_LABEL


def build_ceded_table(bucketed_df, departments):
    """
    Table 1 equivalent for this module: Department x [Previous, Jan..Dec,
    Future Policies] grid of summed Total Reinsurance Ceded (rather than
    Gross Premium, which is what upr_processor.build_gross_premium_table
    sums - the only reason this isn't reused directly from there). Same
    shape/contract otherwise: a plain nested dict {department: {column:
    amount}}, every column always present (0.0 if nothing landed there).
    """
    grouped = bucketed_df.groupby(["Department", "Bucket"])[CEDED_LABEL].sum()
    table = {}
    for dept in departments:
        row = {}
        for col in TABLE_COLUMNS:
            row[col] = float(grouped.get((dept, col), 0.0))
        table[dept] = row
    return table


def build_component_totals(period_df, departments):
    """
    Per-department totals of the three underlying components (Treaty /
    Local / Foreign Premium) PLUS their combined Total Reinsurance Ceded -
    feeds the Summary sheet's own breakdown table, the same kind of
    at-a-glance "how much of this is treaty vs facultative" figure the
    real manually-built working file shows.

    period_df: rows ALREADY restricted to the selected 12-month period
    (Comm Date between period_start and period_end - see
    run_reinsurance_calculation, which is the only caller). This is
    deliberately NOT the full multi-year clean_df: "Premium Ceded" on
    the real Summary sheet means premium on policies incepting DURING
    the period being reported on, not the register's entire cumulative
    history - a department can easily carry several years of old,
    still-amortizing business in its own register (landing in the
    Previous bucket) plus already-booked future-dated business (the
    Future Policies bucket), and neither belongs in this period's own
    "Premium Ceded" figure. Summing the full, unfiltered register here
    (an earlier version of this function did exactly that) silently
    inflated every department's figure by however much Previous/Future
    business happened to be sitting in ITS OWN register - a different
    amount per department, which is exactly why a real run showed each
    department's total off by a different, inconsistent multiple rather
    than one uniform factor.

    Returns {department: {"Treaty Premium": x, "Local Premium": y,
    "Foreign Premium": z, "Total Reinsurance Ceded": x+y+z}}.
    """
    totals = {}
    for dept in departments:
        sub = period_df[period_df["Department"] == dept]
        totals[dept] = {
            "Treaty Premium": float(sub["Treaty Premium"].sum()),
            "Local Premium": float(sub["Local Premium"].sum()),
            "Foreign Premium": float(sub["Foreign Premium"].sum()),
            CEDED_LABEL: float(sub[CEDED_LABEL].sum()),
        }
    return totals


def run_reinsurance_calculation(filepaths, period_year, period_month, report=None):
    """
    Full pipeline: load -> compute Total Reinsurance Ceded per row ->
    validate coverage -> bucket -> build both tables. Mirrors
    upr_processor.run_upr_calculation's own shape exactly (see that
    function's docstring), just built from a Total Reinsurance Ceded
    figure this module computes itself rather than a Gross Premium
    column already present in the source file.

    Returns a result dict:
      {
        "period_start", "period_end": date,
        "departments": [str, ...],
        "methods": {department: "1/24"|"1/6"|"zero"},
        "ceded_table": {department: {column: amount}},          # Table 1
        "prepaid_table": {department: {column: amount}},        # Table 2
        "component_totals": {department: {"Treaty Premium", "Local
            Premium", "Foreign Premium", "Total Reinsurance Ceded"}},
        "closing_prepaid": float,   # grand total of prepaid_table
        "premium_ceded": float,     # period-scoped: sum of Total Reinsurance Ceded for rows whose own Comm Date falls inside the selected period (excludes Previous/Future Policies business)
        "warnings": [str, ...],
        "excluded_rows": int,
        "sheets_used": [(filename, sheet_name), ...],
        "departments_by_file": {filename: department},
      }
    Raises ReinsuranceError with a clear, user-facing message if the
    file(s) can't be read, are missing a required column, a department
    can't be determined from a filename, or the register(s) don't cover
    the full selected period.
    """
    df, meta = load_reinsurance_registers(filepaths, report=report)
    if report:
        report(f"Loaded {len(df):,} usable row(s) from {len(meta['sheets_used'])} sheet(s) "
               f"({meta['excluded_rows']:,} row(s) excluded - missing/implausible Comm Date).")

    period_start, period_end = compute_period_bounds(period_year, period_month)
    if report:
        report(f"Selected period: {period_start:%d-%b-%Y} to {period_end:%d-%b-%Y}.")

    try:
        validate_coverage(df, period_start, period_end)
    except up.UPRError as e:
        raise ReinsuranceError(str(e)) from e

    bucketed = assign_buckets(df, period_start, period_end)
    departments, dept_warnings = resolve_department_list(bucketed)

    ceded_table = build_ceded_table(bucketed, departments)
    prepaid_table, methods = build_unearned_table(ceded_table, departments, period_start)
    closing_prepaid = grand_total(prepaid_table, departments)

    # "Premium Ceded" is business INCEPTING during the selected period
    # only - i.e. just the Jan..Dec bucket columns of Table 1, never the
    # Previous bucket (older business still amortizing from before the
    # period) or the Future Policies bucket (business already booked for
    # after it). grand_total(ceded_table, ...) sums every column
    # including Previous/Future, which is exactly right for Table 1's
    # own "Total" column (a genuine, complete distribution of everything
    # loaded) but wrong for this figure - restrict to the same period
    # window used for bucketing, on the row data directly, rather than
    # summing a subset of Table 1's own columns (equivalent, but avoids
    # any coupling to TABLE_COLUMNS' exact ordering).
    period_mask = (df["Comm Date"] >= pd.Timestamp(period_start)) & (df["Comm Date"] <= pd.Timestamp(period_end))
    period_df = df.loc[period_mask]
    premium_ceded = float(period_df[CEDED_LABEL].sum())
    component_totals = build_component_totals(period_df, departments)

    if report:
        report(f"Premium Ceded (selected period): {premium_ceded:,.2f}")
        report(f"Closing Prepaid Reinsurance Premium Ceded: {closing_prepaid:,.2f}")

    return {
        "period_start": period_start,
        "period_end": period_end,
        "departments": departments,
        "methods": methods,
        "ceded_table": ceded_table,
        "prepaid_table": prepaid_table,
        "component_totals": component_totals,
        "closing_prepaid": closing_prepaid,
        "premium_ceded": premium_ceded,
        "warnings": dept_warnings,
        "excluded_rows": meta["excluded_rows"],
        "sheets_used": meta["sheets_used"],
        "departments_by_file": meta["departments_by_file"],
    }


if __name__ == "__main__":
    import sys
    files = sys.argv[1:] or ["Eng__2025_v2.xls", "Eng__June_2026_v2.xls"]

    df, meta = load_reinsurance_registers(files, report=print)
    print("\nDepartments by file:", meta["departments_by_file"])
    print("Rows loaded:", len(df), "| excluded:", meta["excluded_rows"])
    print(df[["Department", "Comm Date", "Treaty Premium", "Local Premium", "Foreign Premium", CEDED_LABEL]].head(10))

    options = get_available_period_options(df)
    print(f"\n{len(options)} period option(s), most recent first (showing 5):")
    for o in options[:5]:
        print(" ", o["label"])

    if options:
        latest = options[0]
        result = run_reinsurance_calculation(files, latest["year"], latest["month"], report=print)
        print("\nDepartments:", result["departments"])
        print("Methods:", result["methods"])
        print("\nPremium Ceded (period):", f"{result['premium_ceded']:,.2f}")
        print("Closing Prepaid Reinsurance Premium Ceded:", f"{result['closing_prepaid']:,.2f}")
        for dept in result["departments"]:
            print(f"  {dept:15s} ceded_total={row_total(result['ceded_table'][dept]):>15,.0f}  "
                  f"prepaid_total={row_total(result['prepaid_table'][dept]):>15,.0f}")
