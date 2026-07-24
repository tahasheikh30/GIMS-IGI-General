"""
Claims Register (Coinsurance) - Claim Intimation Register matching against
the Premium Register to recover each claim's Coinsurance name(s), then
fanning each claim out to every coinsurer that shares the underlying
policy (a policy can be shared by several coinsurers at once - confirmed
against a real premium register sample, where one policy had 11 separate
coinsurer rows).

WHY MATCHING IS NEEDED: the Claim Intimation Register's own "Coinsurance"
column is sparsely and manually filled (confirmed against a real sample:
1,761 of 1,917 rows blank) - Finance needs it populated automatically by
matching each claim's "Policy/PMD #" against the Premium Register, which
DOES carry the coinsurer name for every row (in its "Ins" column).

MATCHING mirrors premium_register_processor.py's approach exactly (same
fuzzy digit-run blocking + longest-complete-run scoring, since reference
number formats between two different Crystal Reports exports are never
identical character-for-character), reusing its blocking/scoring helpers
directly rather than re-implementing them. The PRIMARY matching column is
the Premium Register's "Base Doc No." column - a claim's "Policy/PMD #" is
matched against it first; the register's "Document No." column is an
OPTIONAL secondary column still checked as a fallback (a claim can
reference either the original policy it renewed from, or the specific
transaction), but "Base Doc No." is what's used both to accept a match and
to fan a claim out to every coinsurer sharing that policy. The Premium
Register carries the SAME Base Doc No. on multiple rows - one row per
company sharing that policy (see PREMIUM REGISTER STRUCTURE below) - so
once a query resolves to a best-matching row, every row sharing that row's
own Base Doc No. is pulled in as a sibling. That's what produces the
coinsurer fan-out.

PREMIUM REGISTER STRUCTURE (confirmed against a real sample):
  - "Base Doc No." - the primary matching/fan-out column: the original
    policy a transaction was issued/renewed from. "Document No." - the
    specific transaction - is an optional secondary column still checked
    as a fallback. A claim's "Policy/PMD #" can reference either one.
  - "Ins" - the company on THIS row's share of the policy. IGI's own
    share AND every coinsurer's share each get their own row under the
    same Base Doc No.
  - "Leader" - "Yes" on IGI's own row, "No" on every coinsurer's row.
  - "Share" - that row's own percentage share of the policy.
  A policy with several coinsurers therefore has several "No" rows all
  sharing one Base Doc No. - matching a claim to that Base Doc No. and
  collecting every "No" row under it identifies ALL of that claim's
  coinsurers at once, not just one.

FAN-OUT: since a claim can rightfully belong to several coinsurers'
recovery files at once (each owes their own share), match_claims() does
not resolve "the" coinsurer per claim - it emits one (claim, coinsurer)
row per matched coinsurer, each carrying that coinsurer's own Share % and
their proportional cut of Claim Paid / Claim O/S (share% x that claim's
own amount). A claim with 3 coinsurers therefore appears in 3 different
output rows (and, later, 3 different companies' output workbooks).

UNMATCHED: a claim with no Premium Register match at all, or one whose
matched policy turns out to have no coinsurer row (100% IGI, nothing to
recover), is collected separately for manual review rather than silently
dropped or guessed at - same never-force-a-bad-match principle as
premium_register_processor.match_one(). Per spec, ALL claims are matched
regardless of Status (Outstanding/Paid/No Loss) - a coinsurer still owes
their share on a claim IGI has already paid out, so Status is preserved
as an output column but never used to filter which claims get matched.

PERFORMANCE: this module was written after the rest of GIMS already went
through the O(items x rows) / iterrows() cleanup documented in
PERFORMANCE_NOTES.md, and had picked up the same class of issue on its
own: _build_policy_index() and match_claims() were both iterating the
register/claims dataframes with .iterrows(), which materializes a full
pandas Series (every column, not just the one or two actually needed) on
every single row - the dominant cost of both functions on a real
9,000-24,000 row register / ~1,900 row claim file. Both now read column
values directly instead. build_coinsurer_buckets() had the same shape of
issue one level down: its per-insured breakdown re-ran pd.to_numeric()
inside a groupby().agg() lambda, so the same Share of Claim Paid/O-S
values were re-parsed once per insured, for every coinsurer's file - now
parsed once up front. See the docstrings on each function below for the
specific before/after.
"""
import os
import re

import pandas as pd

from premium_register_processor import (
    _read_workbook_any_engine, _build_alias_lookup, _map_columns_to_canonical,
    find_header_row, _pick_best_sheet, _dedupe_columns,
    _digits_only, BLOCK_WINDOW, FALLBACK_WINDOW,
    MIN_RUN_MATCH, MIN_RATIO_FALLBACK,
)
from difflib import SequenceMatcher

try:
    from perf_utils import fast_read_excel_rows, estimate_row_count
    _PERF_AVAILABLE = True
except ImportError:  # perf_utils.py not yet added to this checkout, degrade gracefully
    _PERF_AVAILABLE = False

# Above this row count, the full-sheet read switches from pandas' own
# read_excel() to openpyxl's read_only streaming reader - same threshold
# and rationale as data_processor.py / receivable_processor.py. The
# header-row preview scan (find_header_row / _pick_best_sheet) already
# only reads a cheap 25-row preview regardless of file size, so this only
# changes how the FULL sheet gets pulled in afterwards.
FAST_READ_ROW_THRESHOLD = 100000

# ---------------------------------------------------------------------------
# Blocking keys - a LOCAL variant of premium_register_processor._block_keys,
# not the shared one (deliberately not imported/reused for this one piece).
#
# WHY: both this module's Policy/PMD # and the Premium Register's own
# Document No / Base Doc No share an identical "YYYY/MM/CODE#####" shape
# (confirmed against real samples - unlike Premium Register's own NUMS-vs-
# Leader# pairing, where the two sides' formats differ enough that a bare
# 4-digit run is rarely a problem). That shared shape means the 4-digit
# YEAR segment collides constantly here: querying with a key derived from
# "2025" alone matched over 6,000 of 9,065 real register rows in testing,
# since "2025" also appears inside countless OTHER policies' Base Doc No
# (their much earlier original issue year). The premium_register_processor
# module's own min-run-length of 3 was tuned for its own file pair and is
# deliberately left untouched there.
#
# Fix: raise the minimum indexable digit-run length to 5. Checked against
# both real sample files - the actual serial number at the end of the
# reference (the part that's genuinely distinctive, e.g. "00746") is 5
# digits on 97%+ of real rows in both files (1,863/1,917 claims and
# 9,065/9,065 + 9,058/9,065 register rows). A run below 5 digits is
# therefore always noise here (the year or the month, never the serial),
# so excluding it from blocking removes the collision entirely without
# meaningfully reducing real match coverage - the small remainder of
# claims with a shorter last run correctly fall through to "Not Matched"
# rather than being forced through a same-year-only false candidate pool.
# ---------------------------------------------------------------------------
_DIGIT_RUN_RE = re.compile(r"\d+")
MIN_BLOCK_RUN_LEN = 5


def _block_keys(value, window):
    """
    Unlike premium_register_processor._block_keys (which slides a fixed
    window across a run because the real match is often only a SUBSTRING
    embedded inside a longer, dirtier digit run - see that module's own
    docstring), this module's two sides don't have that problem: the
    distinguishing serial number is a clean, standalone digit run of its
    own (confirmed: 5 digits on 97%+ of real rows both sides, never
    embedded inside something longer). Sliding a 4-char window across a
    5-digit serial instead produces two OVERLAPPING partial keys (e.g.
    "00746" -> "0074"/"0746") that collide heavily with unrelated nearby
    serial numbers from the same sequential numbering range (tested: a
    mean candidate set of ~1,644 rows per claim, some over 2,800) -
    exactly the kind of blow-up blocking is meant to avoid, just
    triggered by a different cause than the year-collision problem
    above. Using each qualifying run WHOLE as its own key (no windowing)
    turns this back into a near-exact serial-number join - verified this
    drops the mean candidate set from ~1,644 to a handful.
    """
    keys = set()
    for run in _DIGIT_RUN_RE.findall(str(value)):
        if len(run) < MIN_BLOCK_RUN_LEN:
            continue
        keys.add(run)
    return keys

# ---------------------------------------------------------------------------
# Column alias matching (same pattern as premium_register_processor.ALIASES)
# ---------------------------------------------------------------------------

CLAIM_ALIASES = {
    "Claim Ref Number": ["claim/benefit ref #", "claim/benefit ref", "claim ref number", "claim ref #", "claim number", "claim/benefit ref no"],
    # "policy /pmd#" (space before the slash, none before "#") is the
    # exact real header text seen on a real coinsurance settlement
    # statement's own "IGI Claims" sheet - the "tidier" spacing already
    # listed here ("policy/pmd #") does NOT match it, since alias
    # matching is exact-string after normalization, not fuzzy - so this
    # variant needed adding explicitly rather than assuming the existing
    # entries covered it.
    "Policy/PMD Number": ["policy/pmd #", "policy /pmd#", "policy/pmd#", "policy /pmd #",
                           "policy pmd #", "policy no", "policy number", "pmd #", "policy/pmd number"],
    "Insured Name": ["insured/participant", "insured", "insured name", "participant"],
    "Business Class": ["business class", "class", "class of business"],
    "Cause of Loss": ["cause of loss", "cause of  loss"],
    "Loss Date": ["loss date"],
    "Intimation Date": ["intimation date"],
    "Status": ["status"],
    "Claim Paid": ["claim paid"],
    "Claim O/S": ["claim o/s", "claim os"],
    "Sum Insured": ["suminsured/covered", "sum insured/covered", "sum insured"],
    "Coinsurance": ["coinsurance"],
    "IGI Share Percentage": ["igi share percentage"],
    "Branch Name": ["branch name"],
    "Department Name": ["department name"],
}

REGISTER_ALIASES = {
    "Ins": ["ins", "insurer"],
    "Document No": ["document no", "doc no"],
    "Base Doc No": ["base doc no"],
    "Leader": ["leader"],
    "Share": ["share"],
    "Insured": ["insured"],
    "Business Class": ["business class"],
}

LEADER_NO_VALUES = {"no", "n"}
LEADER_YES_VALUES = {"yes", "y"}

NEW_COLUMNS = [
    "Coinsurer Name", "Coinsurer Share %", "Coinsurer Share of Claim Paid",
    "Coinsurer Share of Claim O/S", "Matched Policy Doc No.", "Match Score",
]

REASON_NOT_MATCHED = "Not matched to any policy in the Premium Register"
REASON_NO_COINSURER = "Matched policy has no coinsurer on it (100% IGI - nothing to recover)"


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

def _read_full_sheet(xls, filepath, sheet_name, report=None):
    """
    Read the full sheet's raw rows (header=None). Above
    FAST_READ_ROW_THRESHOLD rows, switches to openpyxl's read_only
    streaming reader (perf_utils.fast_read_excel_rows) instead of
    pandas' own read_excel() - the same optimization already applied to
    Outstanding Premium's and the Receivable module's loaders, now
    extended to Claims Register.

    Only takes the fast path for a genuine .xlsx read directly off disk.
    A legacy .xls file that went through _read_workbook_any_engine's
    LibreOffice conversion is read through the already-open `xls`
    (pd.ExcelFile) instead - the streaming reader needs the converted
    file's own path, not the original .xls path, and that conversion is
    a one-off cost regardless (every real Premium Register / Claim
    Intimation sample seen so far is comfortably under this threshold).
    """
    ext = os.path.splitext(filepath)[1].lower()
    if _PERF_AVAILABLE and ext == ".xlsx":
        est_rows = estimate_row_count(filepath, sheet_name)
        if est_rows and est_rows > FAST_READ_ROW_THRESHOLD:
            if report:
                report(f"Large file detected (~{est_rows:,} rows) - using the fast streaming reader.")
            rows, _n = fast_read_excel_rows(filepath, sheet_name, report=report)
            return pd.DataFrame(rows)
    return pd.read_excel(xls, sheet_name=sheet_name, header=None)


def load_claim_intimation_file(filepath, sheet_name=None, report=None):
    """
    Load the Claim Intimation Register. Only the reference-number column
    (aliased to "Policy/PMD Number") is required for matching - every
    other original column is left exactly as-is and carried through to
    the output untouched.

    report: optional callable(str) - only used to surface the "large
        file, using the fast reader" note on a genuinely huge file; safe
        to leave as None otherwise.

    Returns (df, pmd_col) where pmd_col is the ORIGINAL column name that
    was matched for "Policy/PMD Number".
    """
    alias_lookup = _build_alias_lookup(CLAIM_ALIASES)
    xls = _read_workbook_any_engine(filepath)

    if sheet_name is None:
        sheet_name, header_row = _pick_best_sheet(xls, alias_lookup)
    else:
        preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=25)
        header_row = find_header_row(preview, alias_lookup)

    raw = _read_full_sheet(xls, filepath, sheet_name, report=report)
    columns = raw.iloc[header_row].tolist()
    data = raw.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = _dedupe_columns(columns)
    data = data.dropna(how="all").reset_index(drop=True)
    data.attrs["source_sheet"] = sheet_name

    colmap = _map_columns_to_canonical(data.columns, alias_lookup)
    inv = {}
    for orig, canon in colmap.items():
        inv.setdefault(canon, orig)
    if "Policy/PMD Number" not in inv:
        raise ValueError(
            f"Could not find the claim's own policy reference column (e.g. 'Policy/PMD #') in "
            f"'{os.path.basename(filepath)}'. Recognized headers were: {list(data.columns)}. "
            "Add the real header text to CLAIM_ALIASES in claims_register_processor.py."
        )
    return data, inv["Policy/PMD Number"]


def get_claim_colmap(claims_df):
    """Return {canonical_name: original_column_name} for every CLAIM_ALIASES
    column actually present in claims_df - convenience for callers (the
    GUI) that need to look up e.g. the real 'Claim Paid'/'Claim O/S'
    column names without re-deriving the alias lookup themselves."""
    alias_lookup = _build_alias_lookup(CLAIM_ALIASES)
    colmap = _map_columns_to_canonical(claims_df.columns, alias_lookup)
    inv = {}
    for orig, canon in colmap.items():
        inv.setdefault(canon, orig)
    return inv


def filter_claims_by_date(claims_df, date_col, start_date=None, end_date=None):
    """
    Restrict claims_df to rows whose date_col falls within [start_date,
    end_date] (either bound may be None to leave that side open). This is
    what lets a run be scoped to a selected period even though the source
    Claim Intimation Register itself covers a much wider date range -
    filtering happens here, BEFORE matching, so everything downstream
    (the matched/unmatched counts, the per-coinsurer buckets, and every
    generated Excel workbook) only ever sees claims inside the window,
    rather than filtering the output after the fact.

    date_col: the ORIGINAL column name in claims_df to filter on (i.e.
        get_claim_colmap(claims_df)["Intimation Date"] or ["Loss Date"] -
        whichever the caller lets the user choose), NOT the canonical name.
    start_date / end_date: date objects (or anything pandas can parse via
        pd.Timestamp), or None.

    A row whose date_col is blank or unparseable is EXCLUDED whenever a
    filter is active - there's no reliable way to know if it belongs
    inside the window or not, and silently including it either way would
    make the reported claim count wrong in one direction or the other.
    If both start_date and end_date are None, claims_df is returned
    unchanged (no rows excluded for a missing date either, since no
    filter was actually requested).

    Returns (filtered_df, num_excluded_no_date) - the second value lets
    the caller log how many rows were dropped specifically for having no
    usable date, separately from rows dropped for being outside the range.
    """
    if start_date is None and end_date is None:
        return claims_df, 0

    parsed = pd.to_datetime(claims_df[date_col], errors="coerce")
    has_date = parsed.notna()

    in_range = pd.Series(True, index=claims_df.index)
    if start_date is not None:
        in_range &= parsed >= pd.Timestamp(start_date)
    if end_date is not None:
        in_range &= parsed <= pd.Timestamp(end_date)

    keep = has_date & in_range
    num_excluded_no_date = int((~has_date).sum())
    return claims_df.loc[keep].reset_index(drop=True), num_excluded_no_date


def load_premium_register_file(filepath, sheet_name=None, report=None):
    """
    Load the Premium Register. Identifies "Ins" (the coinsurer/company
    name for that row's share), "Base Doc No", "Document No", "Leader",
    and "Share". Returns (df, ins_col, doc_col, base_doc_col, leader_col,
    share_col) - NOTE: per the current matching rule, "doc_col" here
    holds the "Base Doc No" column (the PRIMARY matching/fan-out column)
    and "base_doc_col" holds "Document No" (an OPTIONAL secondary
    matching column) - i.e. the two are swapped relative to their plain
    English names, so that every downstream consumer (match_one_claim,
    _build_policy_index, match_claims) - which already treats its first
    doc-column argument as primary and its second as an optional
    fallback - automatically matches/fans-out on Base Doc No without
    needing its own logic changed.

    report: optional callable(str) - see _read_full_sheet; safe to leave
        as None.
    """
    alias_lookup = _build_alias_lookup(REGISTER_ALIASES)
    xls = _read_workbook_any_engine(filepath)

    if sheet_name is None:
        sheet_name, header_row = _pick_best_sheet(xls, alias_lookup)
    else:
        preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=25)
        header_row = find_header_row(preview, alias_lookup)

    raw = _read_full_sheet(xls, filepath, sheet_name, report=report)
    columns = raw.iloc[header_row].tolist()
    data = raw.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = _dedupe_columns(columns)
    data = data.dropna(how="all").reset_index(drop=True)
    data.attrs["source_sheet"] = sheet_name

    colmap = _map_columns_to_canonical(data.columns, alias_lookup)
    inv = {}
    for orig, canon in colmap.items():
        inv.setdefault(canon, orig)
    required = ["Ins", "Base Doc No", "Leader", "Share"]
    missing = [c for c in required if c not in inv]
    if missing:
        raise ValueError(
            f"Could not find required column(s) {missing} in '{os.path.basename(filepath)}'. "
            f"Recognized headers were: {list(data.columns)}. Add the real header text to "
            "REGISTER_ALIASES in claims_register_processor.py."
        )
    # Document No is now the optional secondary matching column - falls
    # back to Base Doc No matching alone if absent (see the module/
    # function docstrings above for why the two are swapped here).
    document_no_col = inv.get("Document No")
    return data, inv["Ins"], inv["Base Doc No"], document_no_col, inv["Leader"], inv["Share"]


# ---------------------------------------------------------------------------
# Fuzzy matching (blocking/scoring reused from premium_register_processor.py)
# ---------------------------------------------------------------------------

def _build_policy_index(register_df, doc_col, base_doc_col, window=BLOCK_WINDOW):
    """{block_key: set(row_index)} built from doc_col (the PRIMARY
    matching column - Base Doc No, see load_premium_register_file) and,
    if present, base_doc_col (the optional secondary column - Document
    No) of every register row, so a claim's own reference can land a
    match whichever of the two it actually refers to.

    PERFORMANCE: reads each column's own values directly via zip(), not
    register_df.iterrows(). iterrows() boxes EVERY column of a register
    row into a full pandas Series on each iteration, even though this
    function only ever reads one column value per pass - on a real
    9,000-24,000 row, many-column register, that per-row Series
    materialization was the dominant cost of building this index.
    Reading the target column's own values straight out of the Series
    (register_df[col]) skips that entirely and is a plain, fast
    column-at-a-time scan.
    """
    index = {}
    cols = [doc_col] + ([base_doc_col] if base_doc_col else [])
    for col in cols:
        for idx, value in zip(register_df.index, register_df[col]):
            for key in _block_keys(value, window):
                index.setdefault(key, set()).add(idx)
    return index


_ALPHA_ONLY_RE = re.compile(r"[^A-Za-z]")


def _alpha_only(value):
    """Letters-only, uppercased, form of a reference (digits/slashes/dashes
    stripped) - e.g. "2019/04/HOFFNIP00161" -> "HOFFNIP". This is the
    branch+class transaction code embedded in the reference, and it's
    what actually disambiguates matches here (see the module-level note
    above _score_row for why digits alone are NOT enough)."""
    return _ALPHA_ONLY_RE.sub("", str(value)).upper()


def _score_row(query_digits, query_runs, query_alpha, idx, doc_digits, base_digits, doc_alpha, base_alpha):
    """
    Score one candidate row against a pre-digested query (digits-only
    string, qualifying digit runs, and letters-only code, all computed
    ONCE per claim by the caller, not per candidate).

    IMPORTANT: the serial-number digit run alone is NOT a reliable match
    signal here, unlike premium_register_processor's own NUMS-vs-Leader#
    pairing. Checked directly against the real sample files: the same
    5-digit serial recurs across many DIFFERENT branches/classes (each
    has its own independent numbering sequence), so a query can score an
    identical, full-length digit-run match against a dozen candidates
    that are actually completely unrelated policies belonging to
    different companies - over half of all claims in the real sample hit
    this exact ambiguity. The branch+class letters embedded in the
    reference (e.g. "HOFFNIP" out of "2019/04/HOFFNIP00161") are what
    actually narrow this down, so they're scored as a second, higher-
    priority signal ABOVE the generic digits-only ratio (but still below
    the digit run length, since an exact serial-number match is still
    the strongest single signal when the letters also roughly agree).

    doc_digits/base_digits/doc_alpha/base_alpha: pandas Series of the
    register's own precomputed values (see match_claims) - avoids
    re-deriving the same register row's digits/letters via regex
    thousands of times across a run.
    """
    best_run_len, best_alpha_ratio, best_digit_ratio = 0, 0.0, 0.0
    pairs = [(doc_digits, doc_alpha)] + ([(base_digits, base_alpha)] if base_digits is not None else [])
    for digits_series, alpha_series in pairs:
        cand_digits = digits_series.at[idx]
        cand_alpha = alpha_series.at[idx]
        run_len = 0
        for run in query_runs:
            if len(run) > run_len and run in cand_digits:
                run_len = len(run)
        alpha_ratio = SequenceMatcher(None, query_alpha, cand_alpha).ratio()
        digit_ratio = SequenceMatcher(None, query_digits, cand_digits).ratio()
        if (run_len, alpha_ratio, digit_ratio) > (best_run_len, best_alpha_ratio, best_digit_ratio):
            best_run_len, best_alpha_ratio, best_digit_ratio = run_len, alpha_ratio, digit_ratio
    return best_run_len, best_alpha_ratio, best_digit_ratio


# A full serial-number match with a near-unrelated branch/class code is a
# false positive (the whole point of scoring the code at all) - so
# acceptance requires BOTH a qualifying digit run AND a plausibly similar
# code, not the digit run alone. Calibrated against the real sample: the
# true match's code always scores well above this on real overlapping
# data, while a coincidental same-serial/different-branch collision
# typically scores far below it.
MIN_ALPHA_RATIO = 0.5


def match_one_claim(pmd_value, register_df, ins_col, doc_col, base_doc_col, leader_col, share_col, index,
                     doc_digits=None, base_digits=None, doc_alpha=None, base_alpha=None):
    """
    Match a single claim's Policy/PMD # against the Premium Register.
    Returns a dict:
      {"status": "Matched" | "No Coinsurer" | "Not Matched",
       "policy_doc_no": <value or None>, "match_score": float,
       "coinsurers": [{"company": ..., "share_pct": ...}, ...]}
    "coinsurers" is only non-empty when status == "Matched".

    doc_digits/base_digits/doc_alpha/base_alpha: optional precomputed
    register-side caches (see match_claims) - built on the fly from
    register_df if not supplied, for standalone/one-off use.
    """
    if doc_digits is None:
        doc_digits = register_df[doc_col].astype(str).map(_digits_only)
    if base_digits is None and base_doc_col:
        base_digits = register_df[base_doc_col].astype(str).map(_digits_only)
    if doc_alpha is None:
        doc_alpha = register_df[doc_col].astype(str).map(_alpha_only)
    if base_alpha is None and base_doc_col:
        base_alpha = register_df[base_doc_col].astype(str).map(_alpha_only)

    if pd.isna(pmd_value) or str(pmd_value).strip() == "":
        return {"status": "Not Matched", "policy_doc_no": None, "match_score": 0.0, "coinsurers": []}

    candidates = set()
    for k in _block_keys(pmd_value, BLOCK_WINDOW):
        candidates |= index.get(k, set())
    if not candidates:
        for k in _block_keys(pmd_value, FALLBACK_WINDOW):
            candidates |= index.get(k, set())
    if not candidates:
        return {"status": "Not Matched", "policy_doc_no": None, "match_score": 0.0, "coinsurers": []}

    query_digits = _digits_only(pmd_value)
    query_runs = [r for r in re.findall(r"\d+", str(pmd_value)) if len(r) >= 3]
    query_alpha = _alpha_only(pmd_value)

    best_idx, best_run_len, best_alpha_ratio, best_digit_ratio = None, 0, 0.0, 0.0
    for idx in candidates:
        run_len, alpha_ratio, digit_ratio = _score_row(
            query_digits, query_runs, query_alpha, idx, doc_digits, base_digits, doc_alpha, base_alpha)
        if (run_len, alpha_ratio, digit_ratio) > (best_run_len, best_alpha_ratio, best_digit_ratio):
            best_idx, best_run_len, best_alpha_ratio, best_digit_ratio = idx, run_len, alpha_ratio, digit_ratio

    accepted = (
        best_idx is not None
        and (best_run_len >= MIN_RUN_MATCH or best_digit_ratio >= MIN_RATIO_FALLBACK)
        and best_alpha_ratio >= MIN_ALPHA_RATIO
    )
    if not accepted:
        return {"status": "Not Matched", "policy_doc_no": None, "match_score": round(best_digit_ratio, 3), "coinsurers": []}

    policy_doc_no = register_df.at[best_idx, doc_col]
    siblings = register_df[register_df[doc_col] == policy_doc_no]
    is_coinsurer = siblings[leader_col].astype(str).str.strip().str.lower().isin(LEADER_NO_VALUES)
    coinsurer_rows = siblings[is_coinsurer]

    if len(coinsurer_rows) == 0:
        return {"status": "No Coinsurer", "policy_doc_no": policy_doc_no, "match_score": round(best_digit_ratio, 3), "coinsurers": []}

    coinsurers = [
        {"company": str(r[ins_col]).strip(), "share_pct": pd.to_numeric(r[share_col], errors="coerce")}
        for _, r in coinsurer_rows.iterrows()
    ]
    return {"status": "Matched", "policy_doc_no": policy_doc_no, "match_score": round(best_digit_ratio, 3), "coinsurers": coinsurers}


def match_claims(claims_df, pmd_col, register_df, ins_col, doc_col, base_doc_col, leader_col, share_col,
                  claim_paid_col=None, claim_os_col=None, report=None):
    """
    Match every claim in claims_df against register_df. Returns
    (matched_df, unmatched_df):

    matched_df - one row PER (claim, coinsurer) pair, i.e. a claim with 3
        coinsurers appears 3 times. Every original claims_df column is
        preserved, plus NEW_COLUMNS: Coinsurer Name, Coinsurer Share %,
        Coinsurer Share of Claim Paid, Coinsurer Share of Claim O/S,
        Matched Policy Doc No., Match Score.

    unmatched_df - one row PER CLAIM (no fan-out - there's no coinsurer
        to fan out to), every original claims_df column preserved plus
        "Reason" explaining why it needs manual review.

    claim_paid_col / claim_os_col: original claims_df column names to
        multiply by each coinsurer's Share % for the two "Coinsurer Share
        of ..." columns. If either is missing/None, that computed column
        is left as None rather than raising, so a source file without
        those columns still matches successfully.

    report: optional callable(str) - if it's a perf_utils.Reporter,
        row-level progress also drives the numeric progress bar via
        perf_utils.RowProgress. Safe to leave as None.

    PERFORMANCE: the main loop used to be claims_df.iterrows() with a
    per-row claim_row.to_dict() call. iterrows() boxes every column of
    the claim row into a full Series on each pass (again, most of which
    is never read - only pmd_col is used for matching), and to_dict() on
    that boxed Series repeats work that pandas can do once, in bulk, far
    more cheaply via claims_df.to_dict("records"). Both are now done
    once, up front, instead of once per claim.
    """
    index = _build_policy_index(register_df, doc_col, base_doc_col)
    # Precompute the register's own digits-only AND letters-only strings
    # for both columns ONCE for the whole register, rather than
    # re-deriving the same row's digits/letters via regex on every
    # (claim, candidate) combination - see _score_row's docstring for why
    # this was the dominant cost, and why the letters-only cache exists
    # at all (digits alone are not a reliable match signal here).
    doc_digits = register_df[doc_col].astype(str).map(_digits_only)
    base_digits = register_df[base_doc_col].astype(str).map(_digits_only) if base_doc_col else None
    doc_alpha = register_df[doc_col].astype(str).map(_alpha_only)
    base_alpha = register_df[base_doc_col].astype(str).map(_alpha_only) if base_doc_col else None

    try:
        from perf_utils import RowProgress
        progress = RowProgress(report, total=len(claims_df), label="claims matched") if report else None
    except ImportError:
        progress = None

    matched_rows = []
    unmatched_rows = []

    claim_paid_series = pd.to_numeric(claims_df[claim_paid_col], errors="coerce") if claim_paid_col and claim_paid_col in claims_df.columns else None
    claim_os_series = pd.to_numeric(claims_df[claim_os_col], errors="coerce") if claim_os_col and claim_os_col in claims_df.columns else None

    # Bulk, one-shot equivalents of what iterrows() + per-row .to_dict()
    # used to do inside the loop - see the function docstring above.
    claim_records = claims_df.to_dict("records")
    pmd_values = claims_df[pmd_col]

    for i, (idx, pmd_value, base) in enumerate(zip(claims_df.index, pmd_values, claim_records), start=1):
        result = match_one_claim(pmd_value, register_df, ins_col, doc_col, base_doc_col, leader_col, share_col, index,
                                  doc_digits=doc_digits, base_digits=base_digits, doc_alpha=doc_alpha, base_alpha=base_alpha)

        if result["status"] == "Matched":
            claim_paid = claim_paid_series.at[idx] if claim_paid_series is not None else None
            claim_os = claim_os_series.at[idx] if claim_os_series is not None else None
            for c in result["coinsurers"]:
                share_pct = c["share_pct"]
                out = dict(base)
                out["Coinsurer Name"] = c["company"]
                out["Coinsurer Share %"] = share_pct
                out["Coinsurer Share of Claim Paid"] = (
                    round(claim_paid * share_pct / 100, 2)
                    if claim_paid is not None and pd.notna(claim_paid) and pd.notna(share_pct) else None
                )
                out["Coinsurer Share of Claim O/S"] = (
                    round(claim_os * share_pct / 100, 2)
                    if claim_os is not None and pd.notna(claim_os) and pd.notna(share_pct) else None
                )
                out["Matched Policy Doc No."] = result["policy_doc_no"]
                out["Match Score"] = result["match_score"]
                matched_rows.append(out)
        else:
            out = dict(base)
            out["Reason"] = REASON_NOT_MATCHED if result["status"] == "Not Matched" else REASON_NO_COINSURER
            out["Matched Policy Doc No."] = result["policy_doc_no"]
            out["Match Score"] = result["match_score"]
            unmatched_rows.append(out)

        if progress:
            progress.maybe(i)
    if progress:
        progress.final()

    claim_cols = list(claims_df.columns)
    matched_df = pd.DataFrame(matched_rows, columns=claim_cols + NEW_COLUMNS) if matched_rows else \
        pd.DataFrame(columns=claim_cols + NEW_COLUMNS)
    unmatched_df = pd.DataFrame(unmatched_rows, columns=claim_cols + ["Reason", "Matched Policy Doc No.", "Match Score"]) if unmatched_rows else \
        pd.DataFrame(columns=claim_cols + ["Reason", "Matched Policy Doc No.", "Match Score"])
    return matched_df, unmatched_df


# ---------------------------------------------------------------------------
# Bucketing (one output file per coinsurer, OFC-named companies pooled -
# same pattern as receivable_processor.py's build_company_groups/OFC pooling)
# ---------------------------------------------------------------------------

OFC_BUCKET_LABEL = "OFC"


def _bucket_series(matched_df, company_col="Coinsurer Name"):
    is_ofc = matched_df[company_col].astype(str).str.contains("ofc", case=False, na=False, regex=False)
    return matched_df[company_col].where(~is_ofc, OFC_BUCKET_LABEL)


def build_coinsurer_buckets(matched_df, company_col="Coinsurer Name",
                             insured_col=None, paid_col="Coinsurer Share of Claim Paid",
                             os_col="Coinsurer Share of Claim O/S"):
    """
    Group matched_df (one row per claim-coinsurer pair) into one bucket
    per output workbook - one per coinsurer, with every company whose
    name contains "OFC" pooled into a single shared "OFC" bucket, same
    rationale/pattern as receivable_processor.build_company_groups.

    Returns (groups, totals, breakdowns):
      groups: {bucket_label: claims_df} - every row belonging to that
          bucket, original columns + fan-out columns, unchanged.
      totals: {bucket_label: {"num_claims": n, "total_share_paid": x,
          "total_share_os": y}}
      breakdowns: {bucket_label: per-insured breakdown df} (Insured Name /
          share_paid_total / share_os_total, sorted largest paid-share
          first) - feeds that bucket's Summary sheet chart. Empty/omitted
          if insured_col isn't found in matched_df.

    PERFORMANCE: paid_col/os_col are converted to numeric ONCE, over the
    whole dataframe, before any grouping happens. The previous version
    ran pd.to_numeric(...).fillna(0) inside a per-INSURED-group lambda
    passed to .agg() - so for a coinsurer with, say, 40 distinct insureds,
    the same Share of Claim Paid/O-S values got re-parsed 40 separate
    times (once per insured group), on every one of the coinsurer's
    output files. A plain "sum" aggregation over an already-numeric
    column runs as a single vectorized pandas groupby; a Python lambda
    re-invoked per group does not.
    """
    df = matched_df.copy()
    if len(df) == 0:
        return {}, {}, {}
    df["_bucket"] = _bucket_series(df, company_col)
    df["_paid_numeric"] = pd.to_numeric(df[paid_col], errors="coerce").fillna(0) if paid_col in df.columns else 0.0
    df["_os_numeric"] = pd.to_numeric(df[os_col], errors="coerce").fillna(0) if os_col in df.columns else 0.0

    groups, totals, breakdowns = {}, {}, {}
    for bucket, sub in df.groupby("_bucket", sort=False):
        groups[bucket] = sub.drop(columns=["_bucket", "_paid_numeric", "_os_numeric"]).reset_index(drop=True)
        totals[bucket] = {
            "num_claims": int(len(sub)),
            "total_share_paid": sub["_paid_numeric"].sum(),
            "total_share_os": sub["_os_numeric"].sum(),
        }
        if insured_col and insured_col in sub.columns:
            g = (
                sub.groupby(insured_col)
                .agg(share_paid_total=("_paid_numeric", "sum"),
                     share_os_total=("_os_numeric", "sum"),
                     num_claims=("_paid_numeric", "size"))
                .reset_index()
                .rename(columns={insured_col: "Insured Name"})
                .sort_values("share_paid_total", ascending=False)
                .reset_index(drop=True)
            )
            breakdowns[bucket] = g
    return groups, totals, breakdowns
