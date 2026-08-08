"""
UPR (Unearned Premium Revenue) Automation - Finance module.

WHAT THIS DOES: given one or more uploaded Premium Register files (a
single consolidated file spanning several years, or several separate
per-year files - either works, since everything is concatenated before
any calculation happens) and a selected PERIOD (a month/year - e.g.
"June 2026" - meaning the 12-month period ending on that month), this
computes the CLOSING UNEARNED PREMIUM REVENUE for that period using the
same two-table method as the real, manually-built working file this
module was reverse-engineered from (see the uploaded example workbook):

  TABLE 1 (Gross Premium Distribution): every premium register row is
  bucketed by its own Comm Date into one of 14 columns - "Previous" (any
  date before the period started), the 12 calendar months Jan-Dec (each
  representing whichever year that calendar month actually falls in
  within the selected period), or "Future Policies" (any date after the
  period ended) - and summed per Department. Departments are the rows.

  TABLE 2 (Unearned Premium Working): the SAME department x column grid,
  but every cell is multiplied by a fraction that reflects how much of
  that month's premium is still unearned as of the period end date:
    - Every department except Marine/Travel/BIMA follows the "1/24"
      method: the Nth month of the 12-month period (N=1 for the first
      month, N=12 for the period's own last/current month) is weighted
      (2N-1)/24 - i.e. 1/24, 3/24, 5/24, ... 23/24 - and "Future
      Policies" is weighted the full 24/24 (entirely unearned, since no
      time has passed against it at all).
    - Marine and Travel follow the "1/6" method: every month is weighted
      0 EXCEPT the last three months of the period (N=10 -> 1/6, N=11 ->
      3/6, N=12 -> 5/6), and "Future Policies" is weighted 6/6.
    - BIMA is excluded entirely - every cell is 0, per instruction (BIMA
      business isn't included in this working at all).
    - "Previous" is always weighted 0 for every department (that
      premium's amortization period has already fully elapsed by the
      period end date).
  The grand total of Table 2 (the sum of every department's own row
  total) is the CLOSING UNEARNED PREMIUM REVENUE for the period.

COVERAGE VALIDATION: the whole point of the "Previous"/"Future Policies"
buckets is that they represent premium OUTSIDE the 12-month period, which
this method treats as fully earned/fully unearned respectively without
needing to know precisely how much of it there is beyond "some". But if
the uploaded register(s) simply don't reach far enough back (no rows
before the period's own start) or far enough forward (no rows at or past
the period's own end), that's not really a "Previous"/"Future Policies"
situation at all - it means the file is missing real, in-period data,
and the whole calculation would be silently wrong in a way that's very
easy to miss. So this is checked explicitly and raises a clear,
actionable error rather than quietly producing an understated/overstated
number.
"""
import calendar
import os
import re
from datetime import date

import pandas as pd

from premium_register_processor import _read_workbook_any_engine, _dedupe_columns

try:
    from perf_utils import fast_read_excel_rows, estimate_row_count
    _PERF_AVAILABLE = True
except ImportError:  # perf_utils.py not yet added to this checkout, degrade gracefully
    _PERF_AVAILABLE = False

FAST_READ_ROW_THRESHOLD = 100000

# ---------------------------------------------------------------------------
# Departments and their amortization method - fixed list/order, confirmed
# against the real working file and Notes_for_UPR_Automation.xlsx. Any
# department found in the source data that ISN'T on this list is still
# included (appended after the known ones, sorted) using the "1/24"
# method by default, with a warning surfaced to the caller - rather than
# silently dropping unrecognized business or guessing it should be
# excluded entirely.
# ---------------------------------------------------------------------------
DEPARTMENT_ORDER = ["Fire", "Marine", "Motor", "Miscellaneous", "Travel", "Engineering", "Bond", "Health", "BIMA"]

METHOD_1_24 = "1/24"
METHOD_1_6 = "1/6"
METHOD_ZERO = "zero"

DEPARTMENT_METHOD = {
    "FIRE": METHOD_1_24,
    "MARINE": METHOD_1_6,
    "MOTOR": METHOD_1_24,
    "MISCELLANEOUS": METHOD_1_24,
    "TRAVEL": METHOD_1_6,
    "ENGINEERING": METHOD_1_24,
    "BOND": METHOD_1_24,
    "HEALTH": METHOD_1_24,
    "BIMA": METHOD_ZERO,
}

# See load_premium_registers' own comment above the plausibility check for
# the full rationale - this is a corruption filter (catches a Comm Date
# that's decades or centuries off, from a bad Excel date serial in the
# source export), not a "recent years only" filter, so it's kept
# deliberately generous rather than tied to "this year".
MIN_PLAUSIBLE_COMM_YEAR = 1990
MAX_PLAUSIBLE_COMM_YEAR_AHEAD = 15  # years beyond today - wide enough for legitimate multi-year installment/travel schedules

MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]
PREVIOUS_LABEL = "Previous"
FUTURE_LABEL = "Future Policies"
TABLE_COLUMNS = [PREVIOUS_LABEL] + MONTH_NAMES + [FUTURE_LABEL]
TOTAL_LABEL = "Total"


class UPRError(Exception):
    """Raised for a problem that should stop the run with a clear,
    actionable message - a missing required column, or (most
    importantly) the register(s) not covering the full selected period."""
    pass


# ---------------------------------------------------------------------------
# Column alias matching (same pattern as every other GIMS loader)
# ---------------------------------------------------------------------------

ALIASES = {
    "Department": ["department/class", "department / class", "department", "class", "class of business"],
    "Comm Date": ["comm date", "commencement date", "comm. date", "comm  date"],
    "Gross Premium": [
        "gross premium/contribution", "gross premium / contribution", "gross premium contribution",
        "gross premium", "grosspremium", "premium/contribution", "premium",
    ],
    "Doc Ref": ["doc ref #", "doc ref#", "doc ref number", "document no", "document number", "doc ref"],
}


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


def _find_header_row(preview_df, max_scan=25, min_score=2):
    """Scan the first `max_scan` rows for whichever one scores highest
    against the alias list - handles banner/title rows sitting above the
    real header (confirmed: the real Premium Register has 5 such rows
    above its own header). Returns (row_index_or_None, score)."""
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


def _read_full_sheet(xls, resolved_filepath, filepath, sheet_name, report=None):
    """Read every raw row of one sheet (header=None) - uses the fast
    streaming reader above FAST_READ_ROW_THRESHOLD rows, same pattern as
    claims_register_processor.py / premium_register_processor.py."""
    ext = os.path.splitext(filepath)[1].lower()
    if _PERF_AVAILABLE and ext == ".xlsx":
        est_rows = estimate_row_count(resolved_filepath, sheet_name)
        if est_rows and est_rows > FAST_READ_ROW_THRESHOLD:
            if report:
                report(f"  Large sheet detected (~{est_rows:,} rows) - using the fast streaming reader.")
            rows, _n = fast_read_excel_rows(resolved_filepath, sheet_name, report=report)
            return pd.DataFrame(rows)
    return pd.read_excel(xls, sheet_name=sheet_name, header=None)


def load_premium_registers(filepaths, report=None):
    """
    Load and concatenate every qualifying sheet across every given
    filepath. "Qualifying" means: its own header row (auto-detected,
    skipping any banner/title rows above it) scores at least 2 against
    the alias list AND, once mapped, actually has Department, Comm Date,
    and Gross Premium columns identified.

    Every sheet in every uploaded workbook that qualifies is included -
    not just the single best-scoring one - since the whole point of this
    loader is to support the register being uploaded as one consolidated
    multi-year file (several qualifying sheets in one workbook) OR as
    several separate per-year files (one qualifying sheet per file) with
    no difference in outcome either way.

    Returns (df, meta):
      df - concatenated rows with canonical columns "Department",
          "Comm Date" (datetime), "Gross Premium" (float), and "Doc Ref"
          if that column was found in at least one sheet (else omitted).
          Only rows with a usable Department + Comm Date + Gross Premium,
          AND a plausible Comm Date (see the plausibility check further
          down - catches a corrupted Excel date serial in the source
          export, e.g. a Comm Date landing in the year 5026), are
          included - excluded_rows in meta reports how many weren't.
      meta - dict: {"sheets_used": [(filename, sheet_name), ...],
          "sheets_skipped": [(filename, sheet_name, reason), ...],
          "excluded_rows": int, "implausible_date_rows": int,
          "implausible_date_samples": [{"department", "comm_date",
              "doc_ref"}, ...] (up to 10, for surfacing in the UI),
          "total_rows_read": int}
    """
    if not filepaths:
        raise UPRError("No premium register file(s) provided.")

    frames = []
    sheets_used, sheets_skipped = [], []
    total_rows_read = 0

    for filepath in filepaths:
        if not os.path.exists(filepath):
            raise UPRError(f"File not found: {filepath}")
        fname = os.path.basename(filepath)
        xls, resolved_filepath = _read_workbook_any_engine(filepath)

        for sheet_name in xls.sheet_names:
            preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=25)
            header_row, score = _find_header_row(preview)
            if header_row is None:
                sheets_skipped.append((fname, sheet_name, "no recognizable header row found"))
                continue

            if report:
                report(f"Reading '{fname}' - sheet '{sheet_name}'...")
            raw = _read_full_sheet(xls, resolved_filepath, filepath, sheet_name, report=report)
            columns = _dedupe_columns(raw.iloc[header_row].tolist())
            data = raw.iloc[header_row + 1:].reset_index(drop=True)
            data.columns = columns
            data = data.dropna(how="all")

            colmap = _map_columns_to_canonical(data.columns)
            inv = {}
            for orig, canon in colmap.items():
                inv.setdefault(canon, orig)

            missing = [c for c in ("Department", "Comm Date", "Gross Premium") if c not in inv]
            if missing:
                sheets_skipped.append((fname, sheet_name, f"missing required column(s): {missing}"))
                continue

            keep_cols = {inv["Department"]: "Department", inv["Comm Date"]: "Comm Date", inv["Gross Premium"]: "Gross Premium"}
            if "Doc Ref" in inv:
                keep_cols[inv["Doc Ref"]] = "Doc Ref"
            sub = data[list(keep_cols.keys())].rename(columns=keep_cols)
            total_rows_read += len(sub)
            frames.append(sub)
            sheets_used.append((fname, sheet_name))

    if not frames:
        raise UPRError(
            "Could not find a usable Premium Register sheet in any of the uploaded file(s) - "
            "none had a recognizable Department/Comm Date/Gross Premium header. "
            f"Sheets checked: {sheets_skipped}"
        )

    df = pd.concat(frames, ignore_index=True, sort=False)
    df["Comm Date"] = pd.to_datetime(df["Comm Date"], errors="coerce")
    df["Gross Premium"] = pd.to_numeric(df["Gross Premium"], errors="coerce")
    df["Department"] = df["Department"].astype(str).str.strip()

    base_valid_mask = df["Comm Date"].notna() & df["Gross Premium"].notna() & (df["Department"] != "") & (df["Department"].str.lower() != "nan")

    # ---- Comm Date plausibility check ----
    # CONFIRMED against a real Crystal Reports export: a handful of rows
    # carry a Comm Date cell that Excel itself already stored as a
    # genuine date type, but whose underlying numeric date SERIAL is
    # corrupted (a real example: raw serial 1,141,848, which is what a
    # normal Excel date serial looks like roughly 3,000 years from now).
    # pandas/xlrd faithfully convert that into a nonsense calendar date -
    # year 5026, 2302, 2206, 2066, 2062, etc. This isn't a parsing bug on
    # this module's side, the file's own stored value is wrong, but left
    # unfiltered it does two real things downstream: every one of those
    # nonsense years shows up as its own separate entry in the period
    # dropdown (get_available_period_options lists every distinct
    # (year, month) actually present in Comm Date, which is what actually
    # surfaced this), and that row's premium silently lands in the
    # "Future Policies" bucket in Table 1/2 at the wrong weight.
    #
    # This is deliberately NOT the same thing as filtering out a
    # genuinely far-future but real Comm Date. Multi-year installment or
    # travel policies in real data legitimately carry Comm Dates several
    # years out (confirmed: several groups of rows share the same
    # day/month across 3-4 consecutive future years under sibling Doc
    # Refs, e.g. 2027-04-12 / 2028-04-12 / 2029-04-12 / 2030-04-12 -
    # clearly an intentional installment schedule, not a typo), so the
    # bound below is a wide corruption filter, not a "recent years only"
    # filter - it only excludes a date that's decades or centuries out,
    # which no legitimate policy commencement date can ever be.
    today_year = date.today().year
    plausible_mask = df["Comm Date"].dt.year.between(MIN_PLAUSIBLE_COMM_YEAR, today_year + MAX_PLAUSIBLE_COMM_YEAR_AHEAD)
    valid_mask = base_valid_mask & plausible_mask

    implausible_date_rows = int((base_valid_mask & ~plausible_mask).sum())
    implausible_date_samples = []
    if implausible_date_rows:
        sample_df = df.loc[base_valid_mask & ~plausible_mask].head(10)
        for _, r in sample_df.iterrows():
            implausible_date_samples.append({
                "department": r["Department"],
                "comm_date": r["Comm Date"],
                "doc_ref": r["Doc Ref"] if "Doc Ref" in df.columns else None,
            })

    excluded_rows = int((~valid_mask).sum())
    clean_df = df.loc[valid_mask].reset_index(drop=True)

    meta = {
        "sheets_used": sheets_used,
        "sheets_skipped": sheets_skipped,
        "excluded_rows": excluded_rows,
        "implausible_date_rows": implausible_date_rows,
        "implausible_date_samples": implausible_date_samples,
        "total_rows_read": total_rows_read,
    }
    return clean_df, meta


# ---------------------------------------------------------------------------
# Period selection
# ---------------------------------------------------------------------------

def get_available_period_options(df, date_col="Comm Date"):
    """
    Every distinct (year, month) found in the loaded data's Comm Date
    column, most recent first - each one is a candidate period END month
    for the dropdown (e.g. year=2026, month=6 -> "June 2026", meaning the
    12-month period July 2025 - June 2026). Whether a given option
    actually has enough surrounding data to be valid is checked
    separately at generation time (validate_coverage) - every distinct
    month present is offered here rather than pre-filtering, since a
    file with a narrower range than expected should surface as a clear
    validation error when picked, not silently disappear from the list.

    Returns a list of dicts: {"year": int, "month": int, "label": str}.
    """
    dates = df[date_col].dropna()
    if len(dates) == 0:
        return []
    pairs = sorted({(d.year, d.month) for d in dates}, reverse=True)
    return [{"year": y, "month": m, "label": f"{calendar.month_name[m]} {y}"} for y, m in pairs]


def compute_period_bounds(period_end_year, period_end_month):
    """Returns (period_start_date, period_end_date) - both plain `date`
    objects. period_end is the LAST calendar day of the selected month;
    period_start is the FIRST calendar day of the month 11 months
    earlier (so the period is exactly 12 calendar months long, ending on
    the selected month)."""
    end_last_day = calendar.monthrange(period_end_year, period_end_month)[1]
    period_end = date(period_end_year, period_end_month, end_last_day)

    start_month_index = period_end_month - 11  # may go <= 0, normalized below
    start_year = period_end_year
    while start_month_index <= 0:
        start_month_index += 12
        start_year -= 1
    period_start = date(start_year, start_month_index, 1)
    return period_start, period_end


def validate_coverage(df, period_start, period_end, date_col="Comm Date"):
    """
    Raises UPRError if the loaded register(s) don't actually reach back
    to period_start and forward to period_end - i.e. the "Previous" and
    "Future Policies" buckets would silently be standing in for real
    missing in-period data rather than genuinely out-of-period business.
    """
    dates = df[date_col].dropna()
    if len(dates) == 0:
        raise UPRError("The premium register has no usable Comm Date values to check coverage against.")

    min_date, max_date = dates.min().date(), dates.max().date()
    problems = []
    if min_date > period_start:
        problems.append(
            f"the earliest Comm Date in the uploaded register(s) is {min_date:%b %Y}, but the selected "
            f"period needs data going back to {period_start:%b %Y}"
        )
    if max_date < period_end:
        problems.append(
            f"the latest Comm Date in the uploaded register(s) is {max_date:%b %Y}, but the selected "
            f"period runs through {period_end:%b %Y}"
        )
    if problems:
        raise UPRError(
            "The premium register doesn't have values for the entire selected year: " + "; ".join(problems) + "."
        )


# ---------------------------------------------------------------------------
# Bucketing
# ---------------------------------------------------------------------------

def _month_to_period_position(period_start):
    """{calendar_month_number (1-12): period_position (1-12)} - which
    position within the 12-month period each calendar month falls at,
    given the period's own start month. E.g. period starting July -> July=1,
    Aug=2, ..., June=12."""
    mapping = {}
    m = period_start.month
    for position in range(1, 13):
        mapping[m] = position
        m = 1 if m == 12 else m + 1
    return mapping


def assign_buckets(df, period_start, period_end, date_col="Comm Date"):
    """Adds two columns to a COPY of df: "Bucket" (Previous / one of the
    12 month names / Future Policies) and "Period Position" (1-12 for a
    month bucket, else None)."""
    out = df.copy()
    month_to_position = _month_to_period_position(period_start)

    dates = out[date_col]
    is_previous = dates.dt.date < period_start
    is_future = dates.dt.date > period_end

    bucket = pd.Series(MONTH_NAMES[0], index=out.index, dtype=object)
    bucket = dates.dt.month.map(lambda m: MONTH_NAMES[m - 1])
    bucket = bucket.mask(is_previous, PREVIOUS_LABEL).mask(is_future, FUTURE_LABEL)

    position = dates.dt.month.map(month_to_position)
    position = position.mask(is_previous | is_future, other=pd.NA)

    out["Bucket"] = bucket
    out["Period Position"] = position
    return out


def resolve_department_list(df):
    """DEPARTMENT_ORDER (canonical order) followed by any other distinct
    department names actually present in the data, sorted, appended at
    the end. Returns (departments, warnings) - warnings lists any
    unrecognized department name (which will default to the "1/24"
    method, since silently dropping real business would be worse)."""
    present = set(df["Department"].unique())
    canonical_present_map = {}
    warnings = []
    ordered = []
    used_raw_names = set()

    for canon in DEPARTMENT_ORDER:
        match = next((p for p in present if p.strip().upper() == canon.upper()), None)
        if match is not None:
            ordered.append(match)
            used_raw_names.add(match)

    extras = sorted(p for p in present if p not in used_raw_names)
    for e in extras:
        ordered.append(e)
        warnings.append(f"Unrecognized department '{e}' - not on the standard list, defaulting to the 1/24 method.")

    return ordered, warnings


def _method_for(department_name):
    key = department_name.strip().upper()
    return DEPARTMENT_METHOD.get(key, METHOD_1_24)


# ---------------------------------------------------------------------------
# Table building
# ---------------------------------------------------------------------------

def build_gross_premium_table(bucketed_df, departments):
    """
    Table 1: Department x [Previous, Jan..Dec, Future Policies] grid of
    summed Gross Premium, as a plain nested dict {department: {column:
    amount}} (every column always present, 0.0 if no rows landed there).
    Row/column totals are computed separately by the caller (kept out of
    this dict so it stays a clean department x column grid).
    """
    grouped = bucketed_df.groupby(["Department", "Bucket"])["Gross Premium"].sum()
    table = {}
    for dept in departments:
        row = {}
        for col in TABLE_COLUMNS:
            row[col] = float(grouped.get((dept, col), 0.0))
        table[dept] = row
    return table


def _fraction(method, bucket, position):
    if method == METHOD_ZERO:
        return 0.0
    if bucket == PREVIOUS_LABEL:
        return 0.0
    if bucket == FUTURE_LABEL:
        return 1.0
    # bucket is a month name; position is 1-12 (guaranteed non-null here)
    n = int(position)
    if method == METHOD_1_24:
        return (2 * n - 1) / 24.0
    if method == METHOD_1_6:
        if n == 10:
            return 1 / 6.0
        if n == 11:
            return 3 / 6.0
        if n == 12:
            return 5 / 6.0
        return 0.0
    return 0.0


def build_unearned_table(gross_table, departments, period_start):
    """
    Table 2: same department x column grid as Table 1, with each cell
    multiplied by its amortization fraction (see the module docstring).
    Returns (table, methods) where methods = {department: "1/24"|"1/6"|"zero"}.
    """
    month_to_position = _month_to_period_position(period_start)
    methods = {dept: _method_for(dept) for dept in departments}

    table = {}
    for dept in departments:
        method = methods[dept]
        row = {}
        for col in TABLE_COLUMNS:
            position = None
            if col not in (PREVIOUS_LABEL, FUTURE_LABEL):
                position = month_to_position[MONTH_NAMES.index(col) + 1]
            frac = _fraction(method, col, position)
            row[col] = gross_table[dept][col] * frac
        table[dept] = row
    return table, methods


def row_total(table_row):
    return sum(table_row[c] for c in TABLE_COLUMNS)


def column_totals(table, departments):
    return {col: sum(table[d][col] for d in departments) for col in TABLE_COLUMNS}


def grand_total(table, departments):
    return sum(row_total(table[d]) for d in departments)


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def run_upr_calculation(filepaths, period_year, period_month, report=None):
    """
    Full pipeline: load -> validate coverage -> bucket -> build both
    tables. Returns a result dict:
      {
        "period_start", "period_end": date,
        "departments": [str, ...],
        "methods": {department: "1/24"|"1/6"|"zero"},
        "gross_table": {department: {column: amount}},
        "unearned_table": {department: {column: amount}},
        "closing_upr": float,   # grand total of unearned_table
        "warnings": [str, ...],
        "excluded_rows": int,
        "sheets_used": [(filename, sheet_name), ...],
      }
    Raises UPRError (with a clear, user-facing message) if the file(s)
    can't be read, don't have the required columns, or don't cover the
    full selected period.
    """
    df, meta = load_premium_registers(filepaths, report=report)
    if report:
        report(f"Loaded {len(df):,} usable row(s) from {len(meta['sheets_used'])} sheet(s) "
               f"({meta['excluded_rows']:,} row(s) excluded - missing Department/Comm Date/Gross Premium).")

    period_start, period_end = compute_period_bounds(period_year, period_month)
    if report:
        report(f"Selected period: {period_start:%d-%b-%Y} to {period_end:%d-%b-%Y}.")

    validate_coverage(df, period_start, period_end)

    bucketed = assign_buckets(df, period_start, period_end)
    departments, dept_warnings = resolve_department_list(bucketed)

    gross_table = build_gross_premium_table(bucketed, departments)
    unearned_table, methods = build_unearned_table(gross_table, departments, period_start)
    closing_upr = grand_total(unearned_table, departments)

    if report:
        report(f"Closing Unearned Premium Revenue: {closing_upr:,.2f}")

    return {
        "period_start": period_start,
        "period_end": period_end,
        "departments": departments,
        "methods": methods,
        "gross_table": gross_table,
        "unearned_table": unearned_table,
        "closing_upr": closing_upr,
        "warnings": dept_warnings,
        "excluded_rows": meta["excluded_rows"],
        "sheets_used": meta["sheets_used"],
    }


if __name__ == "__main__":
    import sys
    files = sys.argv[1:] or ["/mnt/user-data/uploads/Gross_Premium_and_Unearned_Premium_Working_-_Conventional_June_2026.xlsx"]

    df, meta = load_premium_registers(files, report=print)
    print("Sheets used:", meta["sheets_used"])
    print("Sheets skipped:", meta["sheets_skipped"])

    options = get_available_period_options(df)
    print(f"\n{len(options)} period option(s), most recent first (showing 5):")
    for o in options[:5]:
        print(" ", o["label"])

    result = run_upr_calculation(files, 2026, 6, report=print)
    print("\nDepartments:", result["departments"])
    print("Methods:", result["methods"])
    print("\nClosing UPR:", f"{result['closing_upr']:,.2f}")
    for dept in result["departments"]:
        print(f"  {dept:15s} gross_total={row_total(result['gross_table'][dept]):>15,.0f}  "
              f"unearned_total={row_total(result['unearned_table'][dept]):>15,.0f}")
