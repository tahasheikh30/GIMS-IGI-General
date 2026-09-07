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

from premium_register_processor import _read_workbook_any_engine
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

# ---------------------------------------------------------------------------
# Register "as-of" year, from the filename - separate from the department
# keyword match above. A real department can have MORE THAN ONE register
# file uploaded at once - a full prior calendar year's file (e.g.
# "Bond_2025.xls") plus the CURRENT, still-open year's file (e.g.
# "Bond_June_2026.xls") - and "Premium Ceded" (see build_component_totals's
# own docstring below) needs to know which uploaded file(s) belong to the
# period's own ending year, as opposed to earlier, purely-historical
# files kept around only to seed the "Previous" bucket. This is a property
# of the FILE (which register snapshot it is), not of any one row's own
# Comm Date - a current-year file can and does legitimately contain a row
# whose own Comm Date falls in the prior year (an endorsement against an
# older policy), and that row still belongs to the current file's own
# total.
# ---------------------------------------------------------------------------
_YEAR_RE = re.compile(r"(20\d{2})")


def detect_register_year_from_filename(filepath):
    """
    The calendar year this register file is "as of", parsed from its own
    filename - e.g. "Bond_2025.xls" -> 2025, "Bond_June_2026.xls" -> 2026,
    "Eng__June_2026_v2.xls" -> 2026. If more than one 20xx-shaped year
    appears in the name (unusual), the LAST one is used, since a trailing
    year is far more likely to be the file's own "as of" year than an
    earlier one that's incidentally part of some other token. Returns
    None if no such year is found anywhere in the name.
    """
    matches = _YEAR_RE.findall(os.path.basename(filepath))
    if not matches:
        return None
    return int(matches[-1])


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
      meta - {"departments_by_file": {filename: department},
          "register_years_by_file": {filename: year_or_None}, "sheets_used":
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
    register_years_by_file = {}
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

        register_year = detect_register_year_from_filename(filepath)
        register_years_by_file[fname] = register_year

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
        sub["Register Year"] = register_year
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
        "register_years_by_file": register_years_by_file,
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
row_total = up.row_total
build_unearned_table = up.build_unearned_table  # produces the *Prepaid* (still-unamortized) table - Travel uses UPR's own 1/6 method here too (same as Marine), only BIMA is 0
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


def build_component_totals(current_year_df, departments):
    """
    Per-department totals of the three underlying components (Treaty /
    Local / Foreign Premium) PLUS their combined Total Reinsurance Ceded -
    feeds the Summary sheet's own breakdown table AND is what "Premium
    Ceded" itself is (see run_reinsurance_calculation - premium_ceded is
    just the grand total of this same table).

    current_year_df: rows already restricted to whichever uploaded
    register file(s) are "as of" the selected period's own ENDING year
    (register_years_by_file / detect_register_year_from_filename - see
    run_reinsurance_calculation, which is the only caller) - i.e. every
    row from that file, taken WHOLE, with no further Comm Date filtering
    of any kind.

    THIS IS NOT A COMM-DATE FILTER, AND THAT'S DELIBERATE. "Premium Ceded"
    means the reinsurance premium recorded on the department's own CURRENT
    register (the file that's still open/being maintained as of the
    report date) - confirmed directly against a real reconciliation: a
    row can legitimately sit inside the current-year file with a Comm
    Date from the PRIOR year (an endorsement against an older policy,
    dated whenever that endorsement itself was raised) and it still
    belongs in Premium Ceded, exactly as recorded in that file - while a
    row with a 2026 Comm Date that happens to live in the OLDER,
    already-closed prior-year file must NOT be pulled in, because that
    file is only kept for the "Previous" bucket's own amortization
    history, not as a second, overlapping source of current premium.
    Filtering by Comm Date range instead of by source file (an earlier
    version of this function did exactly that) silently mixed exactly
    that: it pulled in stray in-period rows from the prior-year file
    while missing out-of-period rows genuinely recorded in the current
    file, and the two errors didn't cancel out - confirmed against a real
    reconciliation where every department came out understated by a
    different, inconsistent amount.

    Returns {department: {"Treaty Premium": x, "Local Premium": y,
    "Foreign Premium": z, "Total Reinsurance Ceded": x+y+z}}.
    """
    totals = {}
    for dept in departments:
        sub = current_year_df[current_year_df["Department"] == dept]
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
        "premium_ceded": float,     # grand total of component_totals - see build_component_totals's own docstring for exactly what this does and doesn't include
        "warnings": [str, ...],
        "excluded_rows": int,
        "sheets_used": [(filename, sheet_name), ...],
        "departments_by_file": {filename: department},
        "register_years_by_file": {filename: year_or_None},
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

    # "Premium Ceded" = every row belonging to whichever uploaded
    # register file(s) are "as of" the period's own ENDING year (e.g. for
    # a Jul-2025..Jun-2026 period, the "*_June_2026" file) - taken WHOLE,
    # not filtered by each row's own Comm Date. See
    # build_component_totals's own docstring for the full rationale and
    # the real reconciliation this was verified against; a Comm-Date
    # range filter (this function's own earlier approach) looks
    # superficially similar but is a different, wrong quantity - it
    # double-dips into the prior-year file's own in-period rows while
    # missing out-of-period rows the current file legitimately carries
    # (e.g. an endorsement dated in the prior year against a policy the
    # current file is otherwise tracking).
    current_year = period_end.year
    current_year_mask = df["Register Year"] == current_year
    current_year_df = df.loc[current_year_mask]
    if len(current_year_df) == 0:
        dept_warnings = dept_warnings + [
            f"No uploaded register file was detected as belonging to {current_year} (the selected "
            f"period's own ending year) - Premium Ceded and the Treaty/FAC/Foreign FAC breakdown will "
            f"show as 0. Make sure the current year's own register file(s) are uploaded alongside any "
            f"prior-year file(s)."
        ]
    component_totals = build_component_totals(current_year_df, departments)
    premium_ceded = sum(t[CEDED_LABEL] for t in component_totals.values())

    if report:
        report(f"Premium Ceded ({current_year} register file(s)): {premium_ceded:,.2f}")
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
        "register_years_by_file": meta["register_years_by_file"],
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
