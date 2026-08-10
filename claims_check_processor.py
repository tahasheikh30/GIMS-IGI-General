"""
Claims Check (Coinsurance) - verifies whether a co-insurance leader's own
claims file (e.g. the "IGI Claims" sheet inside their settlement workbook)
references a policy that genuinely exists in IGI's own Premium Register,
by matching each claim's "Policy /PMD#" against the register's own
"Leader Ref Number" column (e.g. "Leader#") and pulling in the matched
"IGI Policy Number" (e.g. "Doc Ref #") wherever a confident match is
found - confirmed against a real leader workbook (Indus_Settlement_F.xlsx,
"IGI Claims" sheet) and IGI's own real register export
(Indus_Motor_Premium.xls, "Leader#" / "Doc Ref #" columns).

This is the claims-side counterpart to Policy Check
(premium_register_processor.py / premium_register_generator.py), and now
uses that module's matching engine directly rather than a separate one:
the claim's own "Policy /PMD#" is matched against the SAME register
column (Leader Ref Number, e.g. "Leader#") Policy Check already matches
the leader's policy numbers against, and the SAME column (IGI Policy
Number, e.g. "Doc Ref #") is pulled in as the result - this is exactly
the relationship a claim has to its own originating policy in the
register, the same relationship Policy Check already resolves for a
premium row. Since premium_register_processor's fuzzy digit-run blocking
and longest-complete-run scoring was specifically built and tuned against
this exact Leader#/Doc Ref# reference shape, reusing it here (load_
register_file + match_leader_file, unchanged) means Claims Check inherits
that tuning automatically rather than drifting out of sync with a second,
separately-tuned matcher.

SHEET DETECTION: the leader's own claims sheet is found by scanning the
workbook for a sheet whose NAME identifies it as the claims sheet (see
LEADER_CLAIMS_SHEET_NAME_HINTS below - "IGI Claims" -> compacted
"igiclaims" -> matches), falling back to full column-content scanning if
no sheet name matches - same multi-sheet-settlement-workbook handling
Policy Check already does for the leader's premium sheet, reused directly
via premium_register_processor.load_raw_data() rather than reimplemented
a third time.

Per spec, every claim either finds a confident match or is flagged "Not
Matched" for manual entry - one row per claim, no fan-out (Claims Check
never needs to identify a coinsurer, only confirm a policy exists).
"""
import os
import re
from difflib import SequenceMatcher

import pandas as pd

import premium_register_processor as prp

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
        # "POLICY" - confirmed real header on a GIMS/Centegy system-
        # generated claims recovery statement (e.g. a
        # "CLM_RCOV_COINS_..." export, single "Data" sheet, no "IGI
        # Claims"-named sheet at all). This export's own "POLICY" column
        # carries the reference to match against the register (e.g.
        # "2691121895/1/0/12/2024" = DocNo/CertNo/EndtNo/MM/YYYY) - a bare
        # "policy" alias is added here (distinct from "policy no"/"policy
        # number" above, which require a trailing word) so this exact
        # single-word header is recognized too.
        "policy",
    ],
}

# Sheet-name hints for the leader's own claims sheet, most specific first
# - same rationale as premium_register_processor.LEADER_SHEET_NAME_HINTS:
# a multi-sheet settlement workbook (Summary / Premium Payable / Claim
# Rec / dated snapshot tabs) can have more than one tab that looks
# claims-shaped by column content alone, so the sheet's own NAME is
# checked first, before falling back to full content-based scanning.
# "igiclaims" is what a sheet literally named "IGI Claims" compacts down
# to (see premium_register_processor._compact) - confirmed this is the
# real sheet name in a real leader settlement workbook.
LEADER_CLAIMS_SHEET_NAME_HINTS = [
    "igiclaims", "claimrec", "claimspayable", "claimpayable", "claimsrec",
    "claimintimation", "claims",
]

# NEW_COLUMNS mirrors premium_register_generator.NEW_COLUMNS exactly -
# same matching engine, same shape of result, just applied to a claim's
# own Policy/PMD # instead of a premium row's own policy number.
NEW_COLUMNS = ["IGI Policy Number", "Matched Leader Ref (Register)", "Match Score", "Match Status"]


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
    Load IGI's own Premium Register for Claims Check purposes - identical
    to Policy Check's own register load (premium_register_processor.
    load_register_file), reused directly rather than wrapped: identifies
    "Leader Ref Number" (e.g. "Leader#" - the column being matched
    against) and "IGI Policy Number" (e.g. "Doc Ref #" - the column
    pulled into the output), applies the same Crystal Reports label-
    prefix cleanup, and handles the same legacy-.xls / multi-sheet-split
    situations Policy Check already does.

    Returns (df, leader_ref_col, igi_policy_col).
    """
    return prp.load_register_file(filepath, sheet_name=sheet_name)


# ---------------------------------------------------------------------------
# Matching - this is Policy Check's own matching engine
# (premium_register_processor.match_leader_file), unchanged. A claim's
# "Policy /PMD#" plays exactly the role a premium row's own policy number
# plays there: matched against the register's "Leader Ref Number" column,
# with the register's "IGI Policy Number" pulled in on a confident match.
# ---------------------------------------------------------------------------

def match_leader_claims_file(leader_df, pmd_col, register_df, leader_ref_col, igi_policy_col, report=None):
    """
    Match every claim in leader_df against register_df. Returns a NEW
    dataframe = leader_df (every original column, untouched) plus
    NEW_COLUMNS: "IGI Policy Number", "Matched Leader Ref (Register)",
    "Match Score", "Match Status" ("Matched" or "Not Matched").

    report: optional callable(str) - if it's a perf_utils.Reporter,
        row-level progress also drives the numeric progress bar via
        perf_utils.RowProgress. Safe to leave as None.

    Thin pass-through to premium_register_processor.match_leader_file -
    see that function's own docstring for the full matching/performance
    detail (blocking, two-pass scoring, precomputed register-side
    caches). Nothing about it is claims-specific; a claim's Policy/PMD #
    is matched exactly the way a premium row's own policy number is.
    """
    return prp.match_leader_file(leader_df, pmd_col, register_df, leader_ref_col, igi_policy_col, report=report)


# ---------------------------------------------------------------------------
# Strict Doc-No matching - for leader/claim files whose own reference
# column has NO letters at all (e.g. a GIMS/Centegy SYSTEM-GENERATED
# claims recovery statement, "CLM_RCOV_COINS_..." export, whose own
# "POLICY" column is a plain "DocNo/CertNo/EndtNo/MM/YYYY" string - e.g.
# "269327938/1/18269/04/2024" - confirmed against a real such file).
#
# WHY THE NORMAL ENGINE (match_leader_claims_file, above) ISN'T SAFE ON
# THIS SHAPE: premium_register_processor.match_leader_file's own
# branch/class-code safeguard (MIN_ALPHA_RATIO) exists specifically to
# stop a coincidentally-shared digit run (e.g. a claim's own endorsement
# number happening to also appear inside a totally unrelated register
# row's Doc-No) from being accepted as a match on digits alone - see
# that module's own MIN_ALPHA_RATIO docstring. That safeguard compares
# the LETTERS embedded in each reference (_alpha_only) - but a reference
# like "269327938/1/18269/04/2024" has no letters at all, so both sides
# of that comparison are always the empty string, SequenceMatcher scores
# an empty-vs-empty comparison as a perfect 1.0 match, and the safeguard
# silently passes every candidate rather than actually filtering
# anything. Confirmed via a real sanity check against a real KU4 Premium
# Register (20,851 rows) and a real 81-claim Centegy export: 25 of 74
# "Matched" rows the normal engine produced turned out to have been
# accepted purely because the claim's own CERT or ENDORSEMENT number
# happened to appear as a digit run somewhere inside a completely
# unrelated register row's own Doc-No - one confirmed case, claim
# "269327938/1/18269/04/2024", was matched to a register row whose real
# Doc-No has no relationship to the claim's own true Doc-No
# ("269327938") at all - that Doc-No doesn't exist ANYWHERE in the
# register, yet the claim came back "Matched" with a 0.811 score that
# looks confident.
#
# FIX: only ever accept a match on the strength of the claim's own
# LEADING Doc-No segment (the text before the first "/" - this format's
# own real policy reference, confirmed structurally consistent across
# every real sample row) being found as a genuine embedded digit run
# inside the CANDIDATE register row's own Leader Ref Number / IGI Policy
# Number digit string - never on a cert/endorsement number, and never on
# any other digit run found elsewhere in the claim's own reference.
# MIN_DOC_NO_LEN guards against a false-generic Doc-No (a very short
# leading run would collide constantly) - real Doc-Nos on this file
# shape run 8-10 digits, comfortably clear of it. A claim whose own
# Doc-No isn't found anywhere in the register - even if some OTHER part
# of its reference happens to coincidentally match something - is left
# "Not Matched" for manual review, rather than falling back to a weaker
# signal the way the normal engine's own MIN_RATIO_FALLBACK does. Never
# force a bad match - same "wrong is worse than none" principle as every
# other matcher in this project.
# ---------------------------------------------------------------------------

MIN_DOC_NO_LEN = 6  # a shorter leading run is too generic to trust as the sole match signal


def _leading_doc_no(pmd_value):
    """The claim's own true policy Doc-No: the digits-only form of the
    text before the FIRST '/' in its reference - e.g.
    '269327938/1/18269/04/2024' -> '269327938'. Returns '' for a blank/
    unparseable value (caller treats that the same as "no candidates")."""
    if pd.isna(pmd_value):
        return ""
    first_segment = str(pmd_value).split("/", 1)[0]
    return re.sub(r"\D", "", first_segment)


def match_leader_claims_file_strict(leader_df, pmd_col, register_df, leader_ref_col, igi_policy_col, report=None):
    """
    Strict Doc-No-only counterpart to match_leader_claims_file() - see the
    module-level note above for why this exists and when to use it
    (any claims file whose own reference column has no letters in it at
    all, e.g. a GIMS/Centegy system-generated 'CLM_RCOV_COINS_...' claims
    recovery statement's own 'POLICY' column).

    SAME parameter shape as match_leader_claims_file (this is a drop-in
    swap, not a different call convention): leader_ref_col is the
    register's own single matching column (e.g. 'Leader#' - what the
    claim's Doc-No is searched against), igi_policy_col is the column
    PULLED IN as the result once a match is accepted (e.g. 'Doc Ref #') -
    igi_policy_col is never itself searched as a second matching column
    (unlike claims_register_processor's fan-out matcher, which has a
    genuinely different Base-Doc-No/Document-No pairing - conflating the
    two here was an earlier bug in this function, caught before this
    module shipped, by matching claim Doc-Nos against the IGI Policy
    Number's own different format, e.g. '2024/12/KU4MInIP05584', which
    would only ever coincidentally collide, not genuinely match).

    Returns a NEW dataframe = leader_df (every original column,
    untouched) plus NEW_COLUMNS (IDENTICAL shape/column names to
    match_leader_claims_file's own output), so this is a drop-in
    replacement wherever that function's result is consumed downstream
    (claims_check_generator.build_match_output, the app's own summary/
    log lines, etc.) - only the ACCEPTANCE rule differs, not the output
    contract.
    """
    from premium_register_processor import _digits_only

    leader_ref_digits = register_df[leader_ref_col].astype(str).map(_digits_only)

    try:
        from perf_utils import RowProgress
        progress = RowProgress(report, total=len(leader_df), label="claims matched") if report else None
    except ImportError:
        progress = None

    igi_numbers, matched_refs, scores, statuses = [], [], [], []
    for i, val in enumerate(leader_df[pmd_col], start=1):
        doc_no = _leading_doc_no(val)
        best_idx, best_ratio = None, 0.0

        if len(doc_no) >= MIN_DOC_NO_LEN:
            mask = leader_ref_digits.str.contains(doc_no, regex=False, na=False)
            candidates = register_df.index[mask]

            if len(candidates):
                query_digits = _digits_only(val)
                for idx in candidates:
                    # Among genuine Doc-No matches (there's almost always
                    # exactly one, occasionally a small handful sharing the
                    # same Doc-No across different endorsements/cert
                    # numbers on the same policy), the overall digit-string
                    # similarity ratio breaks the tie in favor of whichever
                    # candidate's own full reference (incl. its own
                    # cert/endt/date suffix) most closely resembles the
                    # claim's own full reference - not just the Doc-No
                    # prefix both already share.
                    cand_digits = leader_ref_digits.at[idx]
                    ratio = SequenceMatcher(None, query_digits, cand_digits).ratio()
                    if ratio > best_ratio:
                        best_idx, best_ratio = idx, ratio

        if best_idx is not None:
            igi_numbers.append(register_df.at[best_idx, igi_policy_col])
            matched_refs.append(register_df.at[best_idx, leader_ref_col])
            scores.append(round(best_ratio, 3))
            statuses.append("Matched")
        else:
            igi_numbers.append(None)
            matched_refs.append(None)
            scores.append(round(best_ratio, 3))
            statuses.append("Not Matched")

        if progress:
            progress.maybe(i)
    if progress:
        progress.final()

    out = leader_df.copy()
    out["IGI Policy Number"] = igi_numbers
    out["Matched Leader Ref (Register)"] = matched_refs
    out["Match Score"] = scores
    out["Match Status"] = statuses
    return out


if __name__ == "__main__":
    import sys

    leader_fp = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/Indus_Settlement_F.xlsx"
    register_fp = sys.argv[2] if len(sys.argv) > 2 else "/mnt/user-data/uploads/Indus_Motor_Premium.xls"

    leader_df, pmd_col = load_leader_claims_file(leader_fp)
    print(f"Leader claims file: {len(leader_df)} rows, policy-reference column detected as '{pmd_col}'")

    register_df, leader_ref_col, igi_policy_col = load_register_file(register_fp)
    print(f"Register file: {len(register_df)} rows, Leader Ref Number column='{leader_ref_col}', "
          f"IGI Policy Number column='{igi_policy_col}'")

    result = match_leader_claims_file(leader_df, pmd_col, register_df, leader_ref_col, igi_policy_col)
    matched = (result["Match Status"] == "Matched").sum()
    print(f"\nMatched: {matched} / {len(result)}  ({matched / len(result):.1%})")
    print(result[[pmd_col, "IGI Policy Number", "Matched Leader Ref (Register)", "Match Score", "Match Status"]].head(20).to_string())
