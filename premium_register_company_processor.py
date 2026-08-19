"""
Premium Register (Coinsurance) - Company Split module.

Purpose: take IGI's own Premium Register export exactly as-is (the same
raw file Policy Check and Claims Intimation already know how to read -
"Ins", "Document No.", "Base Doc No.", "Leader", "Business Class",
"Insured", "Issue Date", "Comm Date", "Expiry Date", "Share", "Sum
Insured", "Gross Premium", "Net Premium", etc.), split it into one
bucket per coinsurer (the "Ins" column, on every row where "Leader" is
"No" - i.e. every OTHER company's own share of a policy IGI leads, which
is exactly what should be billed/reported back to them), and produce one
Excel workbook per coinsurer plus a draft/send email per file - the same
three-step shape as Premium/Claims Email Automation (Receivable) and
Claims Intimation (Claims Register), just sourced directly from the raw
Premium Register export instead of a separately-prepared master file.

WHY LEADER = "No" ONLY: a row with Leader = "Yes" is IGI's OWN share of
that policy - there is nobody to email it to, and including it would
mean IGI's own book showing up as if it were a coinsurer's file. Every
"No" row is a real coinsurer's own share, already carrying that
company's own Gross Premium / Net Premium / Share % for that policy
(unlike Claims Register, nothing needs to be computed proportionally
here - each row already IS that company's own amount).

Reuses the exact same loading building blocks already proven against
this file shape in claims_register_processor.py / premium_register_
processor.py: header-row auto-detection past banner rows, legacy .xls
continuation-sheet merging (this project's real files are routinely
split across Sheet1-4 at Excel's old 65,536-row-per-sheet cap - verified
against the real upload this module was built against), and the fast
streaming reader for the full read.

Bucketing, OFC pooling, "Merge files for the same company" support, and
canonical output filenames all reuse company_alias_resolver.py directly,
so this module's grouping behaves identically to Claims Register's and
Receivable's.
"""
import os
import re

import pandas as pd

from premium_register_processor import (
    _read_workbook_any_engine, _build_alias_lookup, _map_columns_to_canonical,
    find_header_row, _pick_best_sheet, _dedupe_columns,
)
from claims_register_processor import (
    _read_full_table_with_continuations,
)
import company_alias_resolver as car

try:
    from perf_utils import RowProgress
    _PERF_AVAILABLE = True
except ImportError:
    _PERF_AVAILABLE = False


REGISTER_ALIASES = {
    # "Insurance Company" - confirmed real header on a second, structurally
    # different Premium Register export shape: a "Co Insurance Ceded
    # Register" (e.g. "Co Insurance Register 2026 Conventional July",
    # "Detail" sheet). This is the RAW/branch-level company name column on
    # that export (e.g. "AIC Karachi Main Branch") - see "Company Name"
    # below for that same export's own already-normalized counterpart,
    # which is preferred over this one when both are present (see
    # load_premium_register's own handling, right below this dict).
    "Ins": ["ins", "insurer", "insurance company"],
    "Document No": ["document no", "doc no"],
    "Base Doc No": ["base doc no"],
    # "Policy Number" - this export's own single policy-reference column
    # (no separate Base Doc No / Document No split at all on this export
    # shape) - not required by this module (bucketing never needs it),
    # just recognized so it counts toward header-row detection and so a
    # future caller has a documented canonical name to look it up by if
    # ever needed, the same way Claims Register's own optional columns
    # are handled.
    "Policy Number": ["policy number", "policy no"],
    "Leader": ["leader"],
    "Business Class": ["business class", "class", "class of business"],
    "Insured": ["insured", "insured/participant", "insured name", "client name"],
    "Issue Date": ["issue date"],
    "Comm Date": ["comm date", "commencement date", "comm. date", "comm  date"],
    "Expiry Date": ["expiry date"],
    # "Share %" - same field as "Share" on the original export shape, just
    # punctuated with a literal "%" in the header text on this one.
    "Share": ["share", "share %"],
    "Sum Insured": ["sum insured", "suminsured/covered"],
    "Gross Premium": ["gross premium"],
    "Net Premium": ["net premium"],
    # "Company Name" - the Co Insurance Ceded Register export's own
    # ALREADY-NORMALIZED company name column (e.g. "ADAMJEE INSURANCE
    # COMPANY LTD", "JUBILEE GENERAL INSURANCE COMPANY LTD" - a clean,
    # small, consistent set of real company names), sitting alongside
    # that same export's own raw/branch-level "Insurance Company" column
    # (e.g. "AIC Karachi Main Branch", "Century Insurance Co. Karachi
    # Br.") - structurally the same relationship as receivable_processor's
    # own "BOD" (normalized) vs. "Insurance Company" (raw, 500+ spelling
    # variants) columns on ITS own real export. Kept as its own separate
    # canonical name (not folded directly into "Ins") so
    # load_premium_register can explicitly PREFER it over "Ins" for
    # bucketing whenever both are present - see right below this dict.
    "Company Name": ["company name"],
}

LEADER_NO_VALUES = {"no", "n"}
LEADER_YES_VALUES = {"yes", "y"}

OFC_BUCKET_LABEL = "OFC"

# "Leader" is intentionally NOT required here (unlike the original export
# shape) - see load_premium_register's own docstring and
# filter_to_coinsurer_rows's own docstring below for why: the Co
# Insurance Ceded Register export shape has no Leader column at all,
# because every row on it already IS ceded (coinsurer) business by
# construction - there's nothing further to filter by Leader status on
# that shape, so requiring the column here would reject a perfectly
# valid, real file.
REQUIRED_COLUMNS = ["Ins", "Gross Premium"]


def load_premium_register(filepath, sheet_name=None, report=None):
    """
    Load the raw Premium Register export. Auto-detects the real header
    row (skipping banner rows above it), and merges in any legacy .xls
    continuation sheets (Sheet2/3/4/... with no header of their own -
    Excel's old 65,536-row-per-sheet cap) as more rows of the same
    table - confirmed necessary against a real 4-sheet, ~330,000-row
    export.

    Returns (df, colmap) where colmap is {canonical_name: original_column_name}
    for every REGISTER_ALIASES column actually found.

    Accepts EITHER real Premium Register export shape this module has
    been confirmed against:
      1. The original "Leader"-bearing register (Ins/Leader/Base Doc No/
         Document No, etc.) - every real company AND IGI's own rows are
         both present, so filter_to_coinsurer_rows() below still needs
         to filter down to Leader="No" rows.
      2. A "Co Insurance Ceded Register" export (e.g. "Co Insurance
         Register 2026 Conventional July", "Detail" sheet) - has no
         Leader column at all, because every row on this export shape
         already IS ceded (coinsurer) business by construction (it's a
         register of what's been ceded OUT, not IGI's whole book) - see
         filter_to_coinsurer_rows's own docstring for how that's handled.
         This shape also carries the company name under TWO different
         columns - "Insurance Company" (raw/branch-level, e.g. "AIC
         Karachi Main Branch") and "Company Name" (already normalized,
         e.g. "ADAMJEE INSURANCE COMPANY LTD") - "Company Name" is
         preferred for bucketing whenever both are present, same
         rationale as receivable_processor.py preferring its own "BOD"
         column over "Insurance Company" on ITS real export.
    """
    alias_lookup = _build_alias_lookup(REGISTER_ALIASES)
    xls, _resolved_filepath = _read_workbook_any_engine(filepath)

    if sheet_name is None:
        sheet_name, header_row = _pick_best_sheet(xls, alias_lookup)
    else:
        preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=25)
        header_row = find_header_row(preview, alias_lookup)

    data, source_sheet_note = _read_full_table_with_continuations(
        filepath, xls, sheet_name, header_row, alias_lookup, report=report)
    data.attrs["source_sheet"] = source_sheet_note

    colmap = _map_columns_to_canonical(data.columns, alias_lookup)
    inv = {}
    for orig, canon in colmap.items():
        inv.setdefault(canon, orig)

    # Prefer the Co Insurance Ceded Register's own already-normalized
    # "Company Name" column for bucketing over the raw "Ins"/"Insurance
    # Company" column, whenever both are present on this export shape -
    # see this function's own docstring above for the full rationale.
    # Every downstream consumer (filter_to_coinsurer_rows,
    # build_company_buckets, app.py's merge-checkbox handler) reads the
    # bucketing column via colmap["Ins"], so overriding it here is
    # sufficient - the raw "Insurance Company" column, if present, is
    # still preserved untouched as its own regular data column, just not
    # used as the bucket key.
    if "Company Name" in inv:
        inv["Ins"] = inv["Company Name"]

    missing = [c for c in REQUIRED_COLUMNS if c not in inv]
    if missing:
        raise ValueError(
            f"Could not find required column(s) {missing} in '{os.path.basename(filepath)}'. "
            f"Recognized headers were: {list(data.columns)}. Add the real header text to "
            "REGISTER_ALIASES in premium_register_company_processor.py."
        )
    return data, inv


OWN_COMPANY_MARKER = "IGI"


def filter_to_coinsurer_rows(df, colmap, exclude_own_company=True):
    """
    Keep only rows that represent a coinsurer's own share of a policy
    (a coinsurer's own share of a policy IGI leads - see the module
    docstring for why Leader = "Yes" rows, IGI's own share, are always
    excluded here). Also drops rows with a blank Ins value, since there
    is no company to bucket those under.

    If colmap has no "Leader" column at all (see load_premium_register's
    own docstring - a "Co Insurance Ceded Register" export shape has no
    Leader column because every row on it already IS ceded/coinsurer
    business by construction, not IGI's whole book), every row with a
    real company name is treated as already qualifying - there's no
    Leader status left to filter by on that shape, and requiring the
    column would incorrectly reject every row of an otherwise perfectly
    valid file.

    exclude_own_company: also drops any row whose Ins value contains
        "IGI" (case-insensitive) - confirmed against a real register:
        IGI itself can show up as a Leader="No" follower row on a policy
        another company leads (and, on the Ceded Register shape, as a
        follower row directly, e.g. a window takaful unit following a
        conventional unit's policy). That is real, legitimate data on
        IGI's own book, but there is nobody to email a "your premium
        share" statement to for IGI's own name, so it is excluded from
        the company split by default. Set to False to keep those rows in
        (e.g. for a manual review export).
    """
    leader_col = colmap.get("Leader")
    ins_col = colmap["Ins"]
    if leader_col:
        is_coinsurer = df[leader_col].astype(str).str.strip().str.lower().isin(LEADER_NO_VALUES)
    else:
        is_coinsurer = pd.Series(True, index=df.index)
    has_company = df[ins_col].notna() & (df[ins_col].astype(str).str.strip() != "")
    keep = is_coinsurer & has_company
    if exclude_own_company:
        is_own = df[ins_col].astype(str).str.contains(OWN_COMPANY_MARKER, case=False, na=False)
        keep = keep & ~is_own
    return df.loc[keep].reset_index(drop=True)


def filter_by_date(df, colmap, date_field, start_date=None, end_date=None):
    """
    Restrict df to rows whose date_field ("Issue Date" or "Comm Date")
    falls within [start_date, end_date] (either bound may be None).
    Mirrors claims_register_processor.filter_claims_by_date - a row with
    a blank/unparseable date is excluded whenever a filter is active.
    Returns (filtered_df, num_excluded_no_date).
    """
    if start_date is None and end_date is None:
        return df, 0
    date_col = colmap.get(date_field)
    if not date_col or date_col not in df.columns:
        return df, 0

    parsed = pd.to_datetime(df[date_col], errors="coerce")
    has_date = parsed.notna()
    in_range = pd.Series(True, index=df.index)
    if start_date is not None:
        in_range &= parsed >= pd.Timestamp(start_date)
    if end_date is not None:
        in_range &= parsed <= pd.Timestamp(end_date)
    keep = has_date & in_range
    num_excluded_no_date = int((~has_date).sum())
    return df.loc[keep].reset_index(drop=True), num_excluded_no_date


def filter_by_fields(df, field_filters):
    """Same exact-value, AND-combined filter convention as every other
    GIMS module (see claims_register_processor.filter_claims_by_fields).
    field_filters: list of (original_column_name, [exact_value, ...])."""
    for col, values in field_filters:
        if not values or col is None or col not in df.columns:
            continue
        df = df[df[col].astype(str).str.strip().isin(values)]
    return df.reset_index(drop=True)


def _bucket_series(df, ins_col):
    company_series = df[ins_col].astype(str).str.strip()
    is_ofc = company_series.str.contains("ofc", case=False, na=False, regex=False)
    return company_series.where(~is_ofc, OFC_BUCKET_LABEL)


def build_company_buckets(df, colmap):
    """
    Group the (already Leader="No"-filtered) rows into one bucket per
    coinsurer, OFC-named companies pooled into a single shared bucket -
    identical pattern to receivable_processor.build_company_groups /
    claims_register_processor.build_coinsurer_buckets.

    Returns (groups, totals, breakdowns):
      groups: {bucket_label: df} - every original column, untouched.
      totals: {bucket_label: {"num_policies", "total_gross_premium",
          "total_net_premium"}}
      breakdowns: {bucket_label: per-insured breakdown df} (Insured Name /
          gross_total / net_total / num_policies, sorted largest gross
          premium first) - feeds that bucket's Summary sheet chart.
    """
    ins_col = colmap["Ins"]
    insured_col = colmap.get("Insured")
    gross_col = colmap["Gross Premium"]
    net_col = colmap.get("Net Premium")

    work = df.copy()
    work["_bucket"] = _bucket_series(work, ins_col)
    work["_gross_numeric"] = pd.to_numeric(work[gross_col], errors="coerce").fillna(0)
    work["_net_numeric"] = pd.to_numeric(work[net_col], errors="coerce").fillna(0) if net_col else 0.0

    groups, totals, breakdowns = {}, {}, {}
    for bucket, sub in work.groupby("_bucket", sort=False):
        groups[bucket] = sub.drop(columns=["_bucket", "_gross_numeric", "_net_numeric"]).reset_index(drop=True)
        totals[bucket] = {
            "num_policies": int(len(sub)),
            "total_gross_premium": sub["_gross_numeric"].sum(),
            "total_net_premium": sub["_net_numeric"].sum(),
        }
        if insured_col and insured_col in sub.columns:
            g = (
                sub.groupby(insured_col)
                .agg(gross_total=("_gross_numeric", "sum"),
                     net_total=("_net_numeric", "sum"),
                     num_policies=("_gross_numeric", "size"))
                .reset_index()
                .rename(columns={insured_col: "Insured Name"})
                .sort_values("gross_total", ascending=False)
                .reset_index(drop=True)
            )
            breakdowns[bucket] = g
    return groups, totals, breakdowns


def get_column(colmap, canonical_name):
    """Convenience passthrough - the real original column name for a
    canonical field, or None if this source file didn't have it."""
    return colmap.get(canonical_name)
