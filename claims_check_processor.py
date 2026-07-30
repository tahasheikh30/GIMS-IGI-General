"""
Claims Check (Coinsurance) - verifies whether a co-insurance leader's own
claims file (e.g. the "IGI Claims" / "Claim Rec" sheet inside their
settlement workbook) references a policy that genuinely exists in IGI's
own Premium Register, by matching the claim's own "Policy /PMD#" against
the register's "Base Doc No." (falling back to "Document No." as a
secondary column) - confirmed against a real leader workbook
(Indus_Settlement_F.xlsx, "IGI Claims" sheet) and IGI's own real register
export (Indus_Motor_Premium.xls, "Base doc#" / "Doc Ref #" columns).

This is the claims-side counterpart to Policy Check
(premium_register_processor.py / premium_register_generator.py), which
does the same job for the leader's own *premium* file against the
register's Leader Ref Number. The purpose here is narrower than the
Claims Intimation -> coinsurer-recovery workflow in
claims_register_processor.py: Claims Check isn't trying to identify
which coinsurer(s) share a policy (and doesn't need "Ins"/"Leader"/
"Share" columns on the register at all, which a real leader-vs-register
pair like the one this was built against may not even carry under those
names) - it's answering one question per claim: does this policy exist
in our own books at all? So a leader's claims file gets exactly the same
treatment their premium file already gets in Policy Check: every row
either finds a confident match or is flagged "Not Matched" for manual
entry, one row per claim, no fan-out.

MATCHING ENGINE: the blocking (5+ digit run + compound branch/class-code
key) and scoring (digit-run length, then branch/class-code similarity,
then overall digit ratio) are reused directly from
claims_register_processor.py rather than reimplemented - that module's
approach was specifically tuned against this exact Policy/PMD# vs Base
Doc No/Document No reference shape (see its own docstrings for the full
rationale), so re-deriving a second, separately-tuned matcher here would
only risk drifting out of sync with it over time.

LEADER'S FILE: same multi-sheet-settlement-workbook situation Policy
Check already handles (a Summary/Premium Payable/Claim Rec/snapshot-tab
workbook, where more than one tab can independently look claims-shaped
by column content alone) - so sheet selection reuses
premium_register_processor's generic load_raw_data()/sheet-name-hint
machinery rather than reimplementing multi-sheet detection a third time.
"""
import os
import re

import pandas as pd

import premium_register_processor as prp
import claims_register_processor as crp

# ---------------------------------------------------------------------------
# Column alias matching
# ---------------------------------------------------------------------------

# The leader's own claims sheet - only the policy/PMD reference column is
# identified; every other column (Claim #, Loss Date, Claim Paid 100%,
# Client/Participant, Registration No., etc.) is left exactly as-is and
# carried through to the output untouched.
LEADER_CLAIMS_ALIASES = {
    "Policy/PMD Number": [
        # "Policy /PMD#" (space before the slash, none before "#") is the
        # exact real header text confirmed on a real leader settlement
        # workbook's own "IGI Claims" sheet - several punctuation/spacing
        # variants are listed since real exports are inconsistent about
        # where the space falls (same situation already documented in
        # premium_register_processor.LEADER_FILE_ALIASES for the premium
        # side of this exact kind of statement).
        "policy /pmd#", "policy/pmd#", "policy /pmd #", "policy/pmd #",
        "policy pmd #", "policy/pmd number", "pmd #", "policy no", "policy number",
    ],
}

# Sheet-name hints for the leader's own claims sheet, most specific first
# - same rationale as premium_register_processor.LEADER_SHEET_NAME_HINTS:
# a multi-sheet settlement workbook (Summary / Premium Payable / Claim
# Rec / dated snapshot tabs) can have more than one tab that looks
# claims-shaped by column content alone, so the sheet's own NAME is
# checked first, before falling back to full content-based scanning.
LEADER_CLAIMS_SHEET_NAME_HINTS = [
    "igiclaims", "claimrec", "claimspayable", "claimpayable", "claimsrec",
    "claimintimation", "claims",
]

# IGI's own Premium Register, for THIS module's purposes: only "Base Doc
# No" (primary matching column - the original policy) and "Document No"
# (optional secondary/fallback column - the specific transaction) are
# needed. Deliberately does NOT require "Ins"/"Leader"/"Share" the way
# claims_register_processor.REGISTER_ALIASES does - this module never
# needs to identify a coinsurer, only confirm a policy exists, and a real
# register export may not carry those columns under those exact names
# (confirmed: a real register file used for this module had "Base doc#"
# and "Doc Ref #" but no column literally named "Ins"/"Leader"/"Share").
REGISTER_ALIASES = {
    "Base Doc No": [
        "base doc no", "base doc#", "base doc #", "base docno", "base document no",
    ],
    "Document No": [
        "document no", "doc no", "doc ref #", "doc ref#", "docref", "doc reference",
    ],
}

REGISTER_SHEET_NAME_HINTS = ["premiumregister", "policyregister", "register"]

NEW_COLUMNS = ["IGI Base Doc No.", "Match Score", "Match Status"]


def load_leader_claims_file(filepath, sheet_name=None, report=None):
    """
    Load the leader's own claims sheet from their settlement workbook.
    sheet_name=None (the default) first looks for a sheet whose own NAME
    identifies it as the claims sheet (see LEADER_CLAIMS_SHEET_NAME_HINTS),
    falling back to scanning every sheet's header content if no name
    matches. Pass an explicit sheet_name to skip auto-detection entirely.

    Returns (df, pmd_col) where pmd_col is the ORIGINAL column name that
    was matched for "Policy/PMD Number".
    """
    df = prp.load_raw_data(filepath, LEADER_CLAIMS_ALIASES, sheet_name=sheet_name,
                            prefer_sheet_name_hints=LEADER_CLAIMS_SHEET_NAME_HINTS)
    alias_lookup = prp._build_alias_lookup(LEADER_CLAIMS_ALIASES)
    colmap = prp._map_columns_to_canonical(df.columns, alias_lookup)
    inv = {}
    for orig, canon in colmap.items():
        inv.setdefault(canon, orig)
    if "Policy/PMD Number" not in inv:
        raise ValueError(
            f"Could not find the claim's own policy reference column (e.g. 'Policy /PMD#') in "
            f"'{os.path.basename(filepath)}'. Recognized headers were: {list(df.columns)}. "
            "Add the real header text to LEADER_CLAIMS_ALIASES in claims_check_processor.py."
        )
    return df, inv["Policy/PMD Number"]


def load_register_file(filepath, sheet_name=None, report=None):
    """
    Load IGI's own Premium Register for Claims Check purposes. Identifies
    "Base Doc No" (required - the PRIMARY matching column) and
    "Document No" (optional - a secondary fallback column, e.g. the
    specific transaction reference rather than the original policy).

    Returns (df, base_doc_col, document_no_col) - document_no_col is None
    if that column isn't present (matching then falls back to Base Doc
    No alone, same as claims_register_processor's own behavior when
    Document No is absent).
    """
    df = prp.load_raw_data(filepath, REGISTER_ALIASES, sheet_name=sheet_name,
                            prefer_sheet_name_hints=REGISTER_SHEET_NAME_HINTS)
    alias_lookup = prp._build_alias_lookup(REGISTER_ALIASES)
    colmap = prp._map_columns_to_canonical(df.columns, alias_lookup)
    inv = {}
    for orig, canon in colmap.items():
        inv.setdefault(canon, orig)
    if "Base Doc No" not in inv:
        raise ValueError(
            f"Could not find the register's 'Base Doc No' column (e.g. 'Base doc#') in "
            f"'{os.path.basename(filepath)}'. Recognized headers were: {list(df.columns)}. "
            "Add the real header text to REGISTER_ALIASES in claims_check_processor.py."
        )
    return df, inv["Base Doc No"], inv.get("Document No")


# ---------------------------------------------------------------------------
# Matching - reuses claims_register_processor's own blocking/scoring
# machinery directly (see module docstring), just without the coinsurer
# fan-out step that module needs and this one doesn't.
# ---------------------------------------------------------------------------

def match_one(pmd_value, register_df, doc_col, base_doc_col, index,
              doc_digits, base_digits, doc_alpha, base_alpha):
    """
    Match a single claim's Policy/PMD # against the register. Returns a
    dict: {"status": "Matched" | "Not Matched", "policy_doc_no": <value
    or None>, "match_score": float}.

    This is the same acceptance logic as
    claims_register_processor.match_one_claim() (identical blocking,
    identical (run_len, alpha_ratio, digit_ratio) scoring and acceptance
    thresholds), just stopping once a policy is identified rather than
    also resolving its coinsurer siblings - Claims Check has no need for
    that step.
    """
    if pd.isna(pmd_value) or str(pmd_value).strip() == "":
        return {"status": "Not Matched", "policy_doc_no": None, "match_score": 0.0}

    candidates = set()
    for k in crp._all_block_keys(pmd_value, crp.BLOCK_WINDOW):
        candidates |= index.get(k, set())
    if not candidates:
        for k in crp._all_block_keys(pmd_value, crp.FALLBACK_WINDOW):
            candidates |= index.get(k, set())
    if not candidates:
        return {"status": "Not Matched", "policy_doc_no": None, "match_score": 0.0}

    query_digits = crp._digits_only(pmd_value)
    query_runs = [r for r in re.findall(r"\d+", str(pmd_value)) if len(r) >= 3]
    query_alpha = crp._alpha_only(pmd_value)

    best_idx, best_run_len, best_alpha_ratio, best_digit_ratio = None, 0, 0.0, 0.0
    for idx in candidates:
        run_len, alpha_ratio, digit_ratio = crp._score_row(
            query_digits, query_runs, query_alpha, idx, doc_digits, base_digits, doc_alpha, base_alpha)
        if (run_len, alpha_ratio, digit_ratio) > (best_run_len, best_alpha_ratio, best_digit_ratio):
            best_idx, best_run_len, best_alpha_ratio, best_digit_ratio = idx, run_len, alpha_ratio, digit_ratio

    accepted = (
        best_idx is not None
        and (best_run_len >= crp.MIN_RUN_MATCH or best_digit_ratio >= crp.MIN_RATIO_FALLBACK)
        and best_alpha_ratio >= crp.MIN_ALPHA_RATIO
    )
    if not accepted:
        return {"status": "Not Matched", "policy_doc_no": None, "match_score": round(best_digit_ratio, 3)}

    policy_doc_no = register_df.at[best_idx, doc_col]
    return {"status": "Matched", "policy_doc_no": policy_doc_no, "match_score": round(best_digit_ratio, 3)}


def match_leader_claims_file(leader_df, pmd_col, register_df, doc_col, base_doc_col, report=None):
    """
    Match every row of leader_df against register_df. Returns a NEW
    dataframe = leader_df (all original columns, untouched) plus
    NEW_COLUMNS: "IGI Base Doc No.", "Match Score", "Match Status"
    ("Matched" or "Not Matched").

    report: optional callable(str) - if it's a perf_utils.Reporter,
        row-level progress also drives the numeric progress bar via
        perf_utils.RowProgress. Safe to leave as None.
    """
    index = crp._build_policy_index(register_df, doc_col, base_doc_col)
    doc_digits = register_df[doc_col].astype(str).map(crp._digits_only)
    base_digits = register_df[base_doc_col].astype(str).map(crp._digits_only) if base_doc_col else None
    doc_alpha = register_df[doc_col].astype(str).map(crp._alpha_only)
    base_alpha = register_df[base_doc_col].astype(str).map(crp._alpha_only) if base_doc_col else None

    try:
        from perf_utils import RowProgress
        progress = RowProgress(report, total=len(leader_df), label="claims checked") if report else None
    except ImportError:
        progress = None

    doc_nos, scores, statuses = [], [], []
    for i, val in enumerate(leader_df[pmd_col], start=1):
        result = match_one(val, register_df, doc_col, base_doc_col, index,
                            doc_digits, base_digits, doc_alpha, base_alpha)
        doc_nos.append(result["policy_doc_no"])
        scores.append(result["match_score"])
        statuses.append(result["status"])
        if progress:
            progress.maybe(i)
    if progress:
        progress.final()

    out = leader_df.copy()
    out["IGI Base Doc No."] = doc_nos
    out["Match Score"] = scores
    out["Match Status"] = statuses
    return out


if __name__ == "__main__":
    import sys

    leader_fp = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/Indus_Settlement_F.xlsx"
    register_fp = sys.argv[2] if len(sys.argv) > 2 else "/mnt/user-data/uploads/Indus_Motor_Premium.xls"

    leader_df, pmd_col = load_leader_claims_file(leader_fp)
    print(f"Leader claims file: {len(leader_df)} rows, policy-reference column detected as '{pmd_col}'")

    register_df, doc_col, base_doc_col = load_register_file(register_fp)
    print(f"Register file: {len(register_df)} rows, Base Doc No column='{doc_col}'"
          + (f", Document No column='{base_doc_col}'" if base_doc_col else " (no Document No column found)"))

    result = match_leader_claims_file(leader_df, pmd_col, register_df, doc_col, base_doc_col)
    matched = (result["Match Status"] == "Matched").sum()
    print(f"\nMatched: {matched} / {len(result)}  ({matched / len(result):.1%})")
    print(result[[pmd_col, "IGI Base Doc No.", "Match Score", "Match Status"]].head(20).to_string())
