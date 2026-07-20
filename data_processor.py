"""
Core data processing for Outstanding Balance Reporting.
Loads the raw transaction sheet, validates it, and produces
per-client / per-department / per-year aggregates.
"""
import pandas as pd
import numpy as np

try:
    from perf_utils import fast_read_excel_rows, rows_to_dataframe, estimate_row_count
    _PERF_AVAILABLE = True
except ImportError:  # perf_utils.py not yet added to this checkout, degrade gracefully
    _PERF_AVAILABLE = False

REQUIRED_COLS = ["Client/Participant", "Department", "O/s"]
DEFAULT_DATE_COL_FOR_YEAR = "Issue Date"  # alternative: "Comm.  Date"
BRANCH_COL = "Branch desc"  # optional column - branch filter is skipped if not present in the source file

# Above this row count, switch from pandas' own read_excel() to
# openpyxl's read_only streaming reader (perf_utils.fast_read_excel_rows)
# - on a real ~1,000,000-row export this cuts both load time and peak
# memory substantially, since read_only mode never builds openpyxl's
# full in-memory cell object graph before pandas gets to see it.
FAST_READ_ROW_THRESHOLD = 100000


def load_raw_data(filepath, sheet_name=0, report=None):
    """
    Load the raw transaction sheet (e.g. 'Sheet1').
    report: optional callable(str) for progress on huge files (see
        perf_utils.run_in_background) - safe to leave as None.
    Automatically switches to a streaming openpyxl reader for files above
    FAST_READ_ROW_THRESHOLD rows; small/normal files are unaffected and
    still go through the same pd.read_excel() path as before.
    """
    if _PERF_AVAILABLE:
        est_rows = estimate_row_count(filepath, sheet_name if isinstance(sheet_name, str) else None)
        if est_rows and est_rows > FAST_READ_ROW_THRESHOLD:
            if report:
                report(f"Large file detected (~{est_rows:,} rows) - using the fast streaming reader.")
            resolved_sheet = sheet_name if isinstance(sheet_name, str) else 0
            rows, n_rows = fast_read_excel_rows(filepath, resolved_sheet, report=report)
            if not rows:
                return pd.DataFrame()
            df = rows_to_dataframe(rows, header_row_index=0)
            return df
    df = pd.read_excel(filepath, sheet_name=sheet_name)
    return df


def validate_data(df, date_col_for_year=DEFAULT_DATE_COL_FOR_YEAR):
    """
    Validate the raw dataframe. Returns (clean_df, issues_df, dropped_columns).
    issues_df lists rows flagged with a problem and why.
    dropped_columns lists source columns that were entirely empty across
    the whole dataset and were removed (they'd add nothing but clutter to
    every generated report).
    date_col_for_year: which date column to derive the "Year" from
        ("Issue Date" or "Comm.  Date" - note the double space in the
        source column name).
    """
    df = df.copy()
    issues = []

    missing_cols = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")

    # Drop columns that are 100% empty in the source sheet (never drop the
    # columns we require for validation/aggregation, even if somehow blank).
    dropped_columns = [
        c for c in df.columns
        if c not in REQUIRED_COLS and df[c].isna().all()
    ]
    if dropped_columns:
        df = df.drop(columns=dropped_columns)

    # Flag missing client name
    mask_no_client = df["Client/Participant"].isna() | (df["Client/Participant"].astype(str).str.strip() == "")
    # Flag missing department
    mask_no_dept = df["Department"].isna()
    # Flag non-numeric / missing O/s
    os_numeric = pd.to_numeric(df["O/s"], errors="coerce")
    mask_bad_os = os_numeric.isna()

    for idx in df.index[mask_no_client]:
        issues.append({"row": idx, "issue": "Missing client/participant name", "value": df.at[idx, "Client/Participant"]})
    for idx in df.index[mask_no_dept]:
        issues.append({"row": idx, "issue": "Missing department", "value": df.at[idx, "Client/Participant"]})
    for idx in df.index[mask_bad_os]:
        issues.append({"row": idx, "issue": "Non-numeric outstanding balance", "value": df.at[idx, "O/s"]})

    issues_df = pd.DataFrame(issues)

    bad_rows = mask_no_client | mask_no_dept | mask_bad_os
    clean_df = df.loc[~bad_rows].copy()
    clean_df["O/s"] = pd.to_numeric(clean_df["O/s"], errors="coerce")

    # Derive Year from the chosen date column (fallback: "Unknown").
    # Vectorized (no per-row Python apply()): pandas' nullable Int64
    # dtype lets a column of years survive NaN through an int cast, so
    # the whole column converts to text in one C-level pass instead of a
    # Python function call per row.
    if date_col_for_year in clean_df.columns:
        years = pd.to_datetime(clean_df[date_col_for_year], errors="coerce").dt.year.astype("Int64")
        clean_df["Year"] = years.astype(str).where(years.notna(), "Unknown")
    else:
        clean_df["Year"] = "Unknown"

    return clean_df, issues_df, dropped_columns


def build_pivot(clean_df, company_filter=None, department_filter=None, agent_filter=None, branch_filter=None):
    """
    Build a client x department x year pivot (sum of O/s).
    agent_filter: optional list of Agent Name values to restrict to
        (e.g. show only clients serviced by a given agency, then check
        which of those clients individually exceed the threshold).
    branch_filter: optional list of Branch desc values to restrict to
        (same pattern as department_filter/agent_filter).
    Returns the filtered raw dataframe plus a summary series with grand
    totals per client (for threshold filtering).
    """
    df = clean_df.copy()

    if company_filter:
        df = df[df["Client/Participant"].isin(company_filter)]
    if department_filter:
        df = df[df["Department"].isin(department_filter)]
    if agent_filter and "Agent Name" in df.columns:
        df = df[df["Agent Name"].isin(agent_filter)]
    if branch_filter and BRANCH_COL in df.columns:
        df = df[df[BRANCH_COL].isin(branch_filter)]

    # Grand total per client (across all departments/years) - used for threshold check
    grand_totals = df.groupby("Client/Participant")["O/s"].sum().rename("Grand Total")

    return df, grand_totals


def client_agents(df, client_name):
    """Return the distinct agent name(s) associated with a client in df.

    PERFORMANCE NOTE: this re-scans the whole (potentially huge) df for
    a single client. Calling this once per client in a loop over many
    qualifying clients is O(num_clients x num_rows) - use
    client_agents_branches_bulk() below instead when you need this for
    every qualifying client at once."""
    if "Agent Name" not in df.columns:
        return []
    sub = df[df["Client/Participant"] == client_name]
    return sorted(sub["Agent Name"].dropna().unique().tolist())


def client_branches(df, client_name):
    """Return the distinct branch(es) associated with a client in df.
    See the performance note on client_agents() above - the same applies
    here."""
    if BRANCH_COL not in df.columns:
        return []
    sub = df[df["Client/Participant"] == client_name]
    return sorted(sub[BRANCH_COL].dropna().unique().tolist())


def client_agents_branches_bulk(df):
    """
    Compute the distinct Agent Name(s) and Branch desc(s) for EVERY
    client in df in a single pass, instead of calling client_agents()/
    client_branches() once per client in a loop (each of which rescans
    the whole dataframe). On a real ~1,000,000-row export with hundreds
    of qualifying clients, that loop pattern is O(num_clients x
    num_rows); this is one groupby pass over the data regardless of how
    many clients there are.

    Returns (agents_by_client, branches_by_client), each a
    {client_name: [sorted values]} dict. A client with no Agent Name /
    Branch desc column present, or no non-blank values, simply won't
    have a key (callers should use .get(client, []) - see analyze() in
    app.py).
    """
    agents_by_client = {}
    if "Agent Name" in df.columns:
        agents_by_client = (
            df.dropna(subset=["Agent Name"])
            .groupby("Client/Participant")["Agent Name"]
            .apply(lambda s: sorted(s.unique().tolist()))
            .to_dict()
        )
    branches_by_client = {}
    if BRANCH_COL in df.columns:
        branches_by_client = (
            df.dropna(subset=[BRANCH_COL])
            .groupby("Client/Participant")[BRANCH_COL]
            .apply(lambda s: sorted(s.unique().tolist()))
            .to_dict()
        )
    return agents_by_client, branches_by_client


def get_client_raw_rows(clean_df, client_name):
    """Return the original transaction rows for a single client (for the
    'Raw Data' sheet in that client's report).

    PERFORMANCE NOTE: same caveat as client_agents() above - calling this
    once per client in the Generate Reports loop is O(num_clients x
    num_rows). Use group_raw_rows_by_client() below instead to pre-split
    once and look each client up in O(1)."""
    return clean_df[clean_df["Client/Participant"] == client_name].copy()


def group_raw_rows_by_client(df):
    """
    Pre-split df into one sub-dataframe per client via a SINGLE groupby
    pass, instead of the get_client_raw_rows() pattern of re-scanning the
    whole dataframe once per client in a loop. On a real ~1,000,000-row
    export with hundreds of qualifying clients, the per-client-scan
    pattern is O(num_clients x num_rows); this is a single O(num_rows)
    pass regardless of how many clients there are, then an O(1)
    dictionary lookup per client afterward.

    Returns {client_name: sub_df}. Every client actually present in df
    gets an entry (not just qualifying ones) - callers look up only the
    ones they need.
    """
    return {name: sub.copy() for name, sub in df.groupby("Client/Participant", sort=False)}


def qualifying_clients(df, grand_totals, threshold):
    """Return list of client names whose Grand Total meets/exceeds threshold (abs value)."""
    qualifying = grand_totals[grand_totals.abs() >= threshold].sort_values(ascending=False)
    return qualifying


def client_pivot_table(df, client_name):
    """
    Build the Department (rows) x Year (columns) pivot table for a single client,
    with Grand Total row and column, matching the style of the existing Sheet2 pivot.

    PERFORMANCE NOTE: this filters the whole (potentially huge) df down
    to one client every time it's called. In the Generate Reports loop
    that client's raw rows have usually already been isolated (via
    group_raw_rows_by_client()) - pass those directly to
    client_pivot_table_from_rows() instead to skip re-filtering the full
    dataset per client.
    """
    sub = df[df["Client/Participant"] == client_name]
    return client_pivot_table_from_rows(sub)


def client_pivot_table_from_rows(raw_rows):
    """
    Same Department x Year margin-pivot as client_pivot_table(), built
    directly from a client's own ALREADY-FILTERED raw rows (e.g. from
    group_raw_rows_by_client()) rather than re-filtering the full
    dataset - this is what avoids the O(num_clients x num_rows) scan
    client_pivot_table() would otherwise do once per client in the
    Generate Reports loop (used for the non-native-Excel fallback path,
    which is the ONLY path when Excel/pywin32 isn't installed at all).
    """
    pivot = pd.pivot_table(
        raw_rows, values="O/s", index="Department", columns="Year",
        aggfunc="sum", fill_value=0, margins=True, margins_name="Grand Total"
    )
    # Order year columns chronologically, keep Grand Total last
    year_cols = [c for c in pivot.columns if c != "Grand Total"]
    def sort_key(y):
        return (0, int(y)) if str(y).isdigit() else (1, str(y))
    year_cols_sorted = sorted(year_cols, key=sort_key)
    pivot = pivot[year_cols_sorted + ["Grand Total"]]
    return pivot


if __name__ == "__main__":
    import sys
    fp = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/Karachi_Conventional_Outstanding.xlsx"
    raw = load_raw_data(fp, sheet_name="Sheet1")
    clean, issues, _dropped = validate_data(raw)
    print(f"Raw rows: {len(raw)}, Clean rows: {len(clean)}, Issues: {len(issues)}")
    print(issues.head(10))
    df, totals = build_pivot(clean)
    q = qualifying_clients(df, totals, threshold=2_000_000)
    print(f"\nClients >= 2,000,000 threshold: {len(q)}")
    print(q.head(10))
    if len(q) > 0:
        sample_client = q.index[0]
        print(f"\nSample pivot for: {sample_client}")
        print(client_pivot_table(df, sample_client))
