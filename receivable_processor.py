"""
Core data processing for Premium & Claims Receivable Email Automation.

REAL SOURCE FORMAT - two variants confirmed so far:

(1) Outstanding_Auto_Email.xlsx (the original sample this module was built
    against) - a single-sheet workbook whose one sheet carries a plain
    "Company" column already grouped to the correct co-insurance company
    name.

(2) A real production master export (e.g. "Co_Insurance_<date>.xlsx") -
    a MULTI-SHEET workbook. The actual transaction data lives on a sheet
    conventionally named "Detail"; the workbook also carries unrelated
    tabs alongside it - e.g. "List" (a raw-branch-name -> company lookup
    table), "Ageing", "PIVOT", and/or monthly snapshot tabs ("June-26",
    "May-26", ...). Reading sheet index 0 blindly (the old behavior)
    picks whichever of these happens to be first - typically NOT the
    transaction sheet - and header detection then fails outright, since
    none of that sheet's columns match anything in ALIASES. To handle
    this robustly (rather than hardcoding the sheet name "Detail", which
    isn't guaranteed to be stable across exports), load_raw_data() scans
    every sheet in the workbook and picks whichever one has the
    highest-scoring header row - see _pick_best_sheet().

    On this real "Detail" sheet, there is no column literally called
    "Company" - the properly-grouped company name instead lives in a
    column called "BOD" (already cleanly normalized down to ~27 real
    companies, vs. 500+ raw spelling variants in the sheet's own
    "Insurance Company" column - see below), and it already carries the
    literal value "OFC (OTHER FOLLOWER COMPANIES)" for every row that
    belongs in the pooled OFC bucket - so it lines up perfectly with the
    existing OFC substring-detection logic once mapped to the "Company"
    canonical name (see ALIASES).

    IMPORTANT: this sheet ALSO has a separate "Insurance Company" column,
    which looks superficially similar but is NOT the same thing - it's
    the raw, un-normalized branch/sub-entity name as originally entered
    per policy (e.g. "ACE INSURANCE", "AGI Faisalabad Branch", "(old
    code) MASJID-E-KHIZRA SERVICE STATION"), with 500+ distinct spellings
    for what is really only ~27 actual companies, and it does NOT reliably
    contain the substring "OFC" for rows that should be pooled there.
    Grouping on it directly would silently explode into hundreds of wrong
    per-branch output files and break OFC pooling. For that reason
    "Insurance Company" is intentionally left OUT of ALIASES - if a
    future export drops the clean "BOD"/"Company" column entirely,
    load_raw_data() should fail loudly (as it already does) rather than
    silently group at the wrong granularity.

- Each transaction row carries its own "Premium Receivable" amount column
  and its own "Claim Receivable" amount column side by side. In the real
  sample exactly one of the two is non-zero per row (641 premium-only
  rows, 460 claims-only rows, 0 rows with both populated) - so no separate
  Type/Category column is needed to route a row to a sheet; the non-zero
  amount column IS the category. This module detects that layout
  automatically ("dual" mode).
- "Client Name" is the insured. "Company" is the co-insurance company
  group (e.g. "EFU GENERAL INSURANCE COMPANY", "ADAMJEE INSURANCE
  COMPANY LTD"). This is also the key used in the separate email lookup
  file the user browses to on tab 1 - see
  receivable_automation.load_email_lookup_from_file().
- The real header text is "Leader Ref. No" (note the period), matched via
  the alias list below like every other column.

A legacy fallback ("single" mode: one "Amount" column + a "Type"/"Category"
column) is also supported in case a different export format shows up
later - see _detect_mode().

IMPORTANT - Leader Ref Number coverage: in the real sample, 460 of 461
blank-Leader-Ref-Number rows are exactly the Claims Receivable rows (claims
transactions in this export don't carry a Leader Ref Number the way
premium transactions do). Since validate_data() flags and excludes blank
Leader Ref Number rows per spec, this means - on a file shaped like the
real sample - Claims Receivable would come out empty for every group
unless that's addressed. validate_data() still flags these (as
instructed), but run_pipeline() surfaces it as a prominent warning so
this can be confirmed as intended (or the check relaxed for Claims rows)
before it reaches production use.
"""
import os
import re
import pandas as pd

# ---------------------------------------------------------------------------
# Column alias matching
# ---------------------------------------------------------------------------

ALIASES = {
    "Class": ["class", "department", "class of business", "line of business"],
    "Leader Ref Number": [
        "leader ref number", "leader ref no", "leader ref. no", "leader reference number",
        "ref number", "ref no", "reference number", "leader ref",
    ],
    "Code": ["code", "client code", "account code"],
    "Client Name": ["client name", "client", "insured", "insured name"],
    # "bod" (Business/Book-of-Business) is the real production export's
    # own name for the already-normalized company-grouping column - see
    # the module docstring for why this is mapped here but the sheet's
    # separate, un-normalized "Insurance Company" column deliberately is not.
    "Company": ["company", "company name", "group company", "bod"],
    # Dual-amount-column mode (confirmed real format)
    "Premium Receivable Amount": [
        "premium receivable", "premium receivable amount", "premium amount",
        "gross premium receivable",
    ],
    "Claims Receivable Amount": [
        "claim receivable", "claims receivable", "claim receivable amount",
        "claims receivable amount", "claims amount", "claim amount",
    ],
    # Legacy single-amount-column mode (fallback)
    "Type": ["type", "category", "receivable type", "receivable category", "type/category"],
    "Amount": ["amount", "receivable amount", "outstanding amount", "balance", "o/s", "value"],
}

_ALIAS_LOOKUP = {}
for canonical, aliases in ALIASES.items():
    for a in aliases:
        _ALIAS_LOOKUP[a] = canonical

PREMIUM_LABEL = "Premium Receivable"
CLAIMS_LABEL = "Claims Receivable"

# Legacy mode: values in the Type/Category column that mean "premium" or
# "claims" (case-insensitive substring match).
TYPE_MARKERS = {
    PREMIUM_LABEL: ["premium"],
    CLAIMS_LABEL: ["claim"],
}

REQUIRED_BASE = ["Client Name", "Company", "Leader Ref Number"]


def _normalize_header(header):
    h = str(header)
    h = re.sub(r"[()]", " ", h)
    h = re.sub(r"\s+", " ", h).strip().lower()
    return h


def _map_columns_to_canonical(columns):
    mapping = {}
    for col in columns:
        norm = _normalize_header(col)
        if norm in _ALIAS_LOOKUP:
            mapping[col] = _ALIAS_LOOKUP[norm]
    return mapping


def find_header_row(raw_df, max_scan=25, min_score=3, return_score=False):
    """
    Scan the first `max_scan` rows of a header=None read for the row that
    looks most like a real header (most cells matching a known alias).
    Handles title-banner rows above the real header, e.g.:
        Row 0: "IGI General Insurance Limited"
        Row 1: "Co Insurance Outstanding Position"
        Row 2: "As at May 31, 2026 (2020-2026)"
        Row 3: (mostly blank)
        Row 4: Br. Code | Branch Name | ... | Client Name | ... <- real header
    return_score: if True, return (row_or_None, score) instead of raising
        on a bad match - used by _pick_best_sheet() to compare candidate
        sheets against each other rather than failing on the first one
        that doesn't look right. If False (default, used everywhere else),
        raises ValueError if nothing scores highly enough to be confident.
    """
    best_row, best_score = None, 0
    for i in range(min(max_scan, len(raw_df))):
        score = 0
        for v in raw_df.iloc[i].tolist():
            if pd.isna(v):
                continue
            if _normalize_header(v) in _ALIAS_LOOKUP:
                score += 1
        if score > best_score:
            best_score, best_row = score, i
    if best_row is None or best_score < min_score:
        if return_score:
            return None, best_score
        raise ValueError(
            "Could not auto-detect the header row in this file (no row in the "
            f"first {max_scan} rows matched at least {min_score} known column names). "
            "If this file uses different headers than expected, add them to ALIASES "
            "in receivable_processor.py."
        )
    if return_score:
        return best_row, best_score
    return best_row


def _pick_best_sheet(filepath, max_scan=25, min_score=3):
    """
    Scan EVERY sheet in the workbook and return (sheet_name, header_row)
    for whichever one has the highest-scoring header row - i.e. whichever
    sheet actually looks like the real transaction data, rather than
    assuming it's sheet index 0. Real master exports are often multi-sheet
    workbooks that bundle the real data (conventionally on a sheet named
    "Detail") alongside unrelated reference/lookup tabs ("List",
    "Ageing"), a "PIVOT" tab, and/or monthly snapshot tabs - blindly
    reading the first sheet can silently pick one of those instead and
    then fail with a confusing "could not detect header row" error on a
    sheet that was never meant to hold the transaction data.

    Only the first `max_scan` rows of each sheet are read for this scan
    (not the whole sheet), so this stays fast even when the winning sheet
    itself has hundreds of thousands of rows.
    """
    xls = pd.ExcelFile(filepath)
    best = None  # (score, sheet_name, header_row)
    for sn in xls.sheet_names:
        preview = pd.read_excel(xls, sheet_name=sn, header=None, nrows=max_scan)
        header_row, score = find_header_row(preview, max_scan=max_scan, min_score=min_score, return_score=True)
        if header_row is None:
            continue
        if best is None or score > best[0]:
            best = (score, sn, header_row)
    if best is None:
        raise ValueError(
            "Could not auto-detect the header row in any sheet of this workbook "
            f"(no sheet had a row in its first {max_scan} rows matching at least "
            f"{min_score} known column names). If this file uses different headers "
            "than expected, add them to ALIASES in receivable_processor.py."
        )
    _, sheet_name, header_row = best
    return sheet_name, header_row


def load_raw_data(filepath, sheet_name=None):
    """
    Load the master receivable sheet, auto-detecting the real header row
    (title/banner rows above it are skipped) and dropping fully-blank
    spacer rows.

    sheet_name: leave as None (the default - used by the GUI and by
        run_pipeline()) to scan every sheet in the workbook and
        automatically use whichever one actually looks like the real
        transaction data (see _pick_best_sheet) - this handles both a
        single-sheet file (the original sample) and a real multi-sheet
        master export (transaction data on a "Detail"-style sheet,
        alongside unrelated reference/pivot/monthly tabs) without the
        caller needing to know or specify which sheet to use. Pass an
        explicit sheet name/index to skip that scan and read one sheet
        directly instead (e.g. for scripted/test use against a file
        that's known to be single-sheet).

    The dataframe's .attrs["source_sheet"] is set to the sheet actually
    read, so callers (e.g. app.py / receivable_automation.py) can log
    which sheet was auto-selected for transparency.
    """
    if sheet_name is None:
        sheet_name, header_row = _pick_best_sheet(filepath)
    else:
        preview = pd.read_excel(filepath, sheet_name=sheet_name, header=None, nrows=25)
        header_row = find_header_row(preview)

    raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
    columns = raw.iloc[header_row].tolist()
    data = raw.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = columns
    data = data.dropna(how="all").reset_index(drop=True)
    data.attrs["source_sheet"] = sheet_name
    return data


def _detect_mode(colmap):
    """colmap: original_col_name -> canonical. Returns (mode, {...}) naming
    the real column names to use, or raises with a clear message if
    neither recognizable layout is present."""
    inv = {}
    for orig, canon in colmap.items():
        inv.setdefault(canon, orig)  # first match wins if a canonical appears twice

    missing_base = [c for c in REQUIRED_BASE if c not in inv]
    if missing_base:
        raise ValueError(
            f"Missing required column(s): {missing_base}. "
            f"Recognized headers were: {list(colmap.keys())}. "
            "Add the real header text to ALIASES in receivable_processor.py if this is just a naming mismatch."
        )

    if "Premium Receivable Amount" in inv and "Claims Receivable Amount" in inv:
        return "dual", {
            "client": inv["Client Name"], "company": inv["Company"], "leader_ref": inv["Leader Ref Number"],
            "premium_amount": inv["Premium Receivable Amount"], "claims_amount": inv["Claims Receivable Amount"],
        }
    if "Type" in inv and "Amount" in inv:
        return "single", {
            "client": inv["Client Name"], "company": inv["Company"], "leader_ref": inv["Leader Ref Number"],
            "type": inv["Type"], "amount": inv["Amount"],
        }
    raise ValueError(
        "Could not find either a Premium Receivable + Claim Receivable amount-column pair, "
        "or a Type/Category + Amount column pair. Recognized headers were: "
        f"{list(colmap.keys())}. Add the real header text to ALIASES in receivable_processor.py."
    )


def _classify_type(value):
    if pd.isna(value):
        return None
    text = str(value).strip().lower()
    for label, markers in TYPE_MARKERS.items():
        if any(m in text for m in markers):
            return label
    return None


def validate_data(df):
    """
    Validate the raw dataframe (as returned by load_raw_data). Returns
    (clean_df, issues_df, dropped_columns, colmap).

    Leader Ref Number is the OTHER co-insurance company's own policy
    reference - it only exists (and is only meaningful) on Premium
    Receivable rows, where IGI is the follower being billed against a
    policy the other company leads. On Claims Receivable rows IGI is the
    primary payer being reimbursed, so there is no such reference to
    expect from the other side - a blank Leader Ref Number there is
    normal and correct, not a data problem, and is therefore never
    flagged or excluded. (Confirmed against the real sample file: all 460
    claims-only rows had a blank Leader Ref Number, and every single
    premium-only row had one - this is structural, not messy data.)
    Leader Ref Number is required only on rows carrying a non-zero
    Premium Receivable amount (dual mode) or classified as Premium
    Receivable (single/legacy mode).

    clean_df keeps every ORIGINAL column name and ORIGINAL value untouched
    (nothing is renamed or reformatted) - only bad rows are excluded, and
    four internal helper columns are added (prefixed "_rec_" so they never
    collide with a real column) for grouping/splitting downstream:
        _rec_client, _rec_company, _rec_premium_amt, _rec_claims_amt

    colmap: {"mode": "dual"|"single", "client": <original col name>,
             "company": ..., "leader_ref": ..., plus mode-specific keys}

    Flags and excludes rows with:
      - blank Client Name
      - blank Company
      - blank Leader Ref Number
      - (dual mode) both Premium and Claims amounts blank/zero, or either
        one non-numeric
      - (single mode) unrecognized Type/Category value, or non-numeric Amount
    """
    df = df.copy()
    issues = []

    colmap = _map_columns_to_canonical(df.columns)
    mode, cols = _detect_mode(colmap)

    dropped_columns = [c for c in df.columns if df[c].isna().all()]
    if dropped_columns:
        df = df.drop(columns=dropped_columns)

    client_col, company_col, ref_col = cols["client"], cols["company"], cols["leader_ref"]

    mask_no_client = df[client_col].isna() | (df[client_col].astype(str).str.strip() == "")
    mask_no_company = df[company_col].isna() | (df[company_col].astype(str).str.strip() == "")
    mask_no_ref = df[ref_col].isna() | (df[ref_col].astype(str).str.strip() == "")

    def _add_issues(mask, label, value_col):
        for idx in df.index[mask]:
            issues.append({"row": idx, "issue": label, "value": df.at[idx, value_col] if value_col in df.columns else None})

    _add_issues(mask_no_client, "Missing Client Name", client_col)
    _add_issues(mask_no_company, "Missing Company", company_col)

    if mode == "dual":
        premium_col, claims_col = cols["premium_amount"], cols["claims_amount"]
        premium_numeric = pd.to_numeric(df[premium_col], errors="coerce")
        claims_numeric = pd.to_numeric(df[claims_col], errors="coerce")
        mask_bad_premium = df[premium_col].notna() & premium_numeric.isna()
        mask_bad_claims = df[claims_col].notna() & claims_numeric.isna()
        mask_no_amount = premium_numeric.fillna(0).eq(0) & claims_numeric.fillna(0).eq(0)

        _add_issues(mask_bad_premium, "Non-numeric Premium Receivable amount", premium_col)
        _add_issues(mask_bad_claims, "Non-numeric Claim Receivable amount", claims_col)
        _add_issues(mask_no_amount & ~(mask_bad_premium | mask_bad_claims), "No Premium or Claims amount on this row", premium_col)

        # Leader Ref Number is only required (and only flagged/excluded when
        # missing) on rows carrying a Premium Receivable amount - see the
        # docstring above for why a blank ref on a Claims-only row is
        # expected, not an error.
        mask_premium_row = premium_numeric.fillna(0) != 0
        mask_no_ref_on_premium = mask_no_ref & mask_premium_row
        _add_issues(mask_no_ref_on_premium, "Missing Leader Ref Number (Premium Receivable row)", ref_col)

        bad_rows = mask_no_client | mask_no_company | mask_no_ref_on_premium | mask_bad_premium | mask_bad_claims | mask_no_amount
        clean_df = df.loc[~bad_rows].copy()
        clean_df["_rec_premium_amt"] = pd.to_numeric(clean_df[premium_col], errors="coerce").fillna(0)
        clean_df["_rec_claims_amt"] = pd.to_numeric(clean_df[claims_col], errors="coerce").fillna(0)
    else:
        type_col, amount_col = cols["type"], cols["amount"]
        normalized_type = df[type_col].apply(_classify_type)
        mask_bad_type = normalized_type.isna()
        amount_numeric = pd.to_numeric(df[amount_col], errors="coerce")
        mask_bad_amount = amount_numeric.isna()

        _add_issues(mask_bad_type, "Type/Category value did not match Premium or Claims", type_col)
        _add_issues(mask_bad_amount, "Non-numeric Amount", amount_col)

        # Same rule as dual mode: Leader Ref Number is only required on
        # Premium Receivable rows, never on Claims Receivable rows.
        mask_premium_row = normalized_type == PREMIUM_LABEL
        mask_no_ref_on_premium = mask_no_ref & mask_premium_row
        _add_issues(mask_no_ref_on_premium, "Missing Leader Ref Number (Premium Receivable row)", ref_col)

        bad_rows = mask_no_client | mask_no_company | mask_no_ref_on_premium | mask_bad_type | mask_bad_amount
        clean_df = df.loc[~bad_rows].copy()
        norm_type_clean = normalized_type.loc[clean_df.index]
        amt_clean = pd.to_numeric(clean_df[amount_col], errors="coerce").fillna(0)
        clean_df["_rec_premium_amt"] = amt_clean.where(norm_type_clean == PREMIUM_LABEL, 0)
        clean_df["_rec_claims_amt"] = amt_clean.where(norm_type_clean == CLAIMS_LABEL, 0)

    clean_df["_rec_client"] = clean_df[client_col].astype(str).str.strip()
    clean_df["_rec_company"] = clean_df[company_col].astype(str).str.strip()

    issues_df = pd.DataFrame(issues)
    return clean_df, issues_df, dropped_columns, {"mode": mode, "client": client_col, "company": company_col, "leader_ref": ref_col, **cols}


_HELPER_COLS = ["_rec_client", "_rec_company", "_rec_premium_amt", "_rec_claims_amt"]

OFC_BUCKET_LABEL = "OFC"


def _is_ofc_company(company_name):
    """OFC companies are identified purely by their Company name
    containing the text "OFC" (case-insensitive) - confirmed directly,
    since nothing in the source data itself flags this any other way
    (no dedicated column, no fixed name list)."""
    return "ofc" in str(company_name).strip().lower()


def _bucket_series(clean_df):
    """Map every row's Company to its output bucket: its own exact
    Company name, or the shared OFC_BUCKET_LABEL if that name contains
    "OFC". This is the one place that decides bucketing, so every
    grouping/totals/breakdown function below stays in sync automatically."""
    return clean_df["_rec_company"].apply(lambda c: OFC_BUCKET_LABEL if _is_ofc_company(c) else c)


def build_company_groups(clean_df):
    """
    Group clean_df into one bucket per output workbook. Per spec, that's
    one file per insurance company, covering every client for that
    company - NOT one file per (client, company) pair like the module
    used to produce. Returns an ordered dict:
        {bucket_label: {"premium_rows": df, "claims_rows": df}}

    Every company whose name contains "OFC" is pooled into a single
    shared bucket (OFC_BUCKET_LABEL = "OFC") instead of getting its own
    per-company bucket - i.e. every OFC-named company ends up together in
    one combined file, rather than each having its own (confirmed: "OFC
    companies get their own file" means one shared file for the whole OFC
    set, not one each).

    Each sub-dataframe keeps every ORIGINAL source column (including
    Client Name and Company - both stay visible on every row, which
    matters most for the pooled OFC bucket since it mixes rows from
    several distinct real companies into one workbook) - the helper
    columns are stripped back out before being returned.
    """
    df = clean_df.copy()
    df["_rec_bucket"] = _bucket_series(df)

    groups = {}
    for bucket, sub in df.groupby("_rec_bucket", sort=False):
        premium_rows = sub[sub["_rec_premium_amt"] != 0].drop(columns=_HELPER_COLS + ["_rec_bucket"])
        claims_rows = sub[sub["_rec_claims_amt"] != 0].drop(columns=_HELPER_COLS + ["_rec_bucket"])
        groups[bucket] = {
            "premium_rows": premium_rows.reset_index(drop=True),
            "claims_rows": claims_rows.reset_index(drop=True),
        }
    return groups


def company_totals(clean_df, groups):
    """Return {bucket_label: {"premium_total": x, "claims_total": y,
    "num_clients": n}} computed straight from the helper amount columns
    (before they're stripped), so totals are correct even though the
    output sheets show the original source column names. num_clients
    counts distinct clients within that bucket - for a normal company
    that's just its own client count; for the pooled OFC bucket it's the
    distinct client count across every OFC-named company combined."""
    df = clean_df.copy()
    df["_rec_bucket"] = _bucket_series(df)

    totals = {}
    for bucket, sub in df.groupby("_rec_bucket", sort=False):
        totals[bucket] = {
            "premium_total": sub["_rec_premium_amt"].sum(),
            "claims_total": sub["_rec_claims_amt"].sum(),
            "num_clients": int(sub["_rec_client"].nunique()),
        }
    return totals


def client_breakdown_for_bucket(clean_df, bucket_label):
    """Per-client Premium/Claims Receivable totals within one output
    bucket (a single company, or the pooled OFC set) - this feeds that
    workbook's own Summary sheet chart, sorted largest Premium Receivable
    first so the chart reads most-to-least at a glance."""
    df = clean_df.copy()
    df["_rec_bucket"] = _bucket_series(df)
    sub = df[df["_rec_bucket"] == bucket_label]
    g = (
        sub.groupby("_rec_client")
        .agg(premium_total=("_rec_premium_amt", "sum"), claims_total=("_rec_claims_amt", "sum"))
        .reset_index()
        .rename(columns={"_rec_client": "Client Name"})
    )
    return g.sort_values("premium_total", ascending=False).reset_index(drop=True)
