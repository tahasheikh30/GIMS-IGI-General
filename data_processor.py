"""
Core data processing for Outstanding Balance Reporting.
Loads the raw transaction sheet, validates it, and produces
per-client / per-department / per-year aggregates.
"""
import pandas as pd
import numpy as np

REQUIRED_COLS = ["Client/Participant", "Department", "O/s"]
DEFAULT_DATE_COL_FOR_YEAR = "Issue Date"  # alternative: "Comm.  Date"
BRANCH_COL = "Branch desc"  # optional column - branch filter is skipped if not present in the source file


def load_raw_data(filepath, sheet_name=0):
    """Load the raw transaction sheet (e.g. 'Sheet1')."""
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

    # Derive Year from the chosen date column (fallback: "Unknown")
    if date_col_for_year in clean_df.columns:
        years = pd.to_datetime(clean_df[date_col_for_year], errors="coerce").dt.year
        clean_df["Year"] = years.apply(lambda y: str(int(y)) if pd.notna(y) else "Unknown")
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
    """Return the distinct agent name(s) associated with a client in df."""
    if "Agent Name" not in df.columns:
        return []
    sub = df[df["Client/Participant"] == client_name]
    return sorted(sub["Agent Name"].dropna().unique().tolist())


def client_branches(df, client_name):
    """Return the distinct branch(es) associated with a client in df."""
    if BRANCH_COL not in df.columns:
        return []
    sub = df[df["Client/Participant"] == client_name]
    return sorted(sub[BRANCH_COL].dropna().unique().tolist())


def get_client_raw_rows(clean_df, client_name):
    """Return the original transaction rows for a single client (for the
    'Raw Data' sheet in that client's report)."""
    return clean_df[clean_df["Client/Participant"] == client_name].copy()


def qualifying_clients(df, grand_totals, threshold):
    """Return list of client names whose Grand Total meets/exceeds threshold (abs value)."""
    qualifying = grand_totals[grand_totals.abs() >= threshold].sort_values(ascending=False)
    return qualifying


def client_pivot_table(df, client_name):
    """
    Build the Department (rows) x Year (columns) pivot table for a single client,
    with Grand Total row and column, matching the style of the existing Sheet2 pivot.
    """
    sub = df[df["Client/Participant"] == client_name]
    pivot = pd.pivot_table(
        sub, values="O/s", index="Department", columns="Year",
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
