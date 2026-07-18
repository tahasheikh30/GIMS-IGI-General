"""
Premium Register (Coinsurance) - Policy Number Matching.

Matches each policy in a co-insurance LEADER's file (their own policy
number, in a column such as "NUMS") against IGI's own Premium Register
(which carries the leader's policy number back to us in a column such as
"Leader #"), and pulls in the corresponding IGI Policy Number wherever a
match is found. Unmatched rows are flagged so they can be entered into
the Premium Register manually.

Both source files are handled the same way as the rest of GIMS
(receivable_processor.py, loss_ratio_processor.py): header text varies by
partner/export, so columns are matched via a case-insensitive alias list
rather than fixed positions, and - since real exports often have a few
banner/title rows above the real header - the header row is auto-detected
by scanning for whichever row scores highest against the alias list
(same approach as receivable_processor.find_header_row /
_pick_best_sheet).

MATCHING IS FUZZY, NOT EXACT STRING EQUALITY, per spec: the leader's own
policy number format essentially never matches our "Leader #" column
character-for-character (different partners format their numbers very
differently, e.g. "1/114591/25585/11006" vs. our "2691025585/12/2023").
In practice the two numbers reliably share at least one distinctive
digit run (the policy/endorsement's own numeric core) even when
everything else around it - separators, prefixes, date-like suffixes -
differs. So matching works in two stages:
  1. BLOCKING - index every register row by its own digit runs (4+ digit
     substrings), then look up candidate rows for a leader value by its
     own digit runs. This keeps matching fast (no need to compare every
     leader row against all ~24k register rows) and, in practice, is
     also what actually identifies the right candidate(s).
  2. SCORING - among the (usually small) candidate set, pick whichever
     register row shares the longest COMPLETE digit run (its own
     policy/endorsement number as a whole token) with the query, using
     overall digit-string similarity only as a tiebreaker. A row with no
     candidates, or whose best candidate doesn't clear the acceptance
     thresholds (see MIN_RUN_MATCH / MIN_RATIO_FALLBACK), is left
     unmatched rather than forced onto a poor guess - a wrong policy
     number would be worse than none, since someone would act on it.
"""
import os
import re
import subprocess
import tempfile
from difflib import SequenceMatcher

import pandas as pd

# ---------------------------------------------------------------------------
# Column alias matching (same pattern as receivable_processor.ALIASES)
# ---------------------------------------------------------------------------

LEADER_FILE_ALIASES = {
    "Leader Policy Number": [
        "nums", "num", "policy no", "policy number", "leader policy no",
        "leader policy number", "leader ref no", "leader reference",
    ],
}

REGISTER_ALIASES = {
    "Leader Ref Number": [
        "leader#", "leader #", "leader no", "leader number", "leader ref no",
        "leader ref. no", "leader ref number", "leader reference number",
        "leader ref", "leaderref",
    ],
    "IGI Policy Number": [
        "doc ref #", "doc ref#", "doc reference", "doc ref number",
        "igi policy number", "igi policy no", "igi doc ref", "doc no",
    ],
}


def _normalize_header(header):
    h = str(header)
    h = re.sub(r"[()]", " ", h)
    # Real exports vary in how they punctuate a header (e.g. "POLICY NO."
    # vs "Policy No", "Leader Ref. No." vs "Leader Ref No") - periods are
    # dropped entirely so both forms normalize to the same string, rather
    # than requiring every alias to be listed twice (with and without the
    # trailing period). "#" is deliberately NOT stripped here, since it's
    # used literally in some aliases below (e.g. "leader#", "doc ref #").
    h = h.replace(".", "")
    h = re.sub(r"\s+", " ", h).strip().lower()
    return h


def _compact(text):
    """Lowercase with every non-alphanumeric character removed - used only
    for matching a WORKBOOK SHEET NAME against a hint phrase (e.g. sheet
    'Premium Payable', 'PremiumPayable', 'Prem. Payable', and 'PREM
    PAYABLE - AIC' should all be recognized as the same intent), since
    sheet-naming conventions vary far more than column headers do across
    different co-insurance partners' exports."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


# Sheet-name hints for the leader's own "Premium Payable" sheet, most
# specific first. A workbook from a co-insurance leader is commonly a
# multi-sheet settlement statement (e.g. Summary / Premium Payable / Claim
# Rec / Prem Rec / monthly snapshot tabs) - several of those OTHER sheets
# can easily score just as high (or higher) on generic column-header
# matching alone as the real Premium Payable sheet does (e.g. a "Prem Rec"
# or dated snapshot tab also has its own "POLICY NO." column), so relying
# on header content alone is not reliable enough to pick the right one.
# Checking the sheet's own name first resolves that ambiguity directly
# whenever the exporting company has named it anything recognizably close
# to "Premium Payable" - which, in practice, most do.
LEADER_SHEET_NAME_HINTS = [
    "premiumpayable", "prempayable", "premiumaccountpayable",
    "premiumapayable", "payabletoigi", "coinsurancepremiumpayable",
]

# Similarly, IGI's own Premium Register export can itself be a multi-sheet
# workbook (e.g. a raw data tab alongside a pivot/summary tab) - these
# hints let the real register data sheet be found by name first, the same
# way LEADER_SHEET_NAME_HINTS does for the leader's own file.
REGISTER_SHEET_NAME_HINTS = ["premiumregister", "policyregister", "register"]


def _find_sheet_by_name(xls, hints):
    """Return the first sheet whose (compacted) name contains one of
    `hints`, checked hint-by-hint in priority order so a more specific
    hint always wins over a looser one if both would otherwise match.
    Falls back to any sheet whose name contains "payable" but not "claim"
    (guards against a workbook that also has a "Claim Payable" tab).
    Returns None if nothing matches, so the caller can fall back to
    full content-based scanning."""
    compact_names = {sn: _compact(sn) for sn in xls.sheet_names}
    for hint in hints:
        for sn, compact in compact_names.items():
            if hint in compact:
                return sn
    for sn, compact in compact_names.items():
        if "payable" in compact and "claim" not in compact:
            return sn
    return None


def _build_alias_lookup(aliases_dict):
    lookup = {}
    for canonical, aliases in aliases_dict.items():
        for a in aliases:
            lookup[a] = canonical
    return lookup


def _map_columns_to_canonical(columns, alias_lookup):
    mapping = {}
    for col in columns:
        norm = _normalize_header(col)
        if norm in alias_lookup:
            mapping[col] = alias_lookup[norm]
    return mapping


def find_header_row(raw_df, alias_lookup, max_scan=25, min_score=1, return_score=False):
    """Scan the first `max_scan` rows of a header=None read for whichever
    row scores highest against `alias_lookup` (most cells matching a
    known alias) - handles title/banner rows sitting above the real
    header, same approach as receivable_processor.find_header_row."""
    best_row, best_score = None, 0
    for i in range(min(max_scan, len(raw_df))):
        score = 0
        for v in raw_df.iloc[i].tolist():
            if pd.isna(v):
                continue
            if _normalize_header(v) in alias_lookup:
                score += 1
        if score > best_score:
            best_score, best_row = score, i
    if best_row is None or best_score < min_score:
        if return_score:
            return None, best_score
        raise ValueError(
            f"Could not auto-detect the header row in this file (no row in the first "
            f"{max_scan} rows matched at least {min_score} known column name(s))."
        )
    if return_score:
        return best_row, best_score
    return best_row


def _pick_best_sheet(xls, alias_lookup, max_scan=25, min_score=1):
    """Scan every sheet and return (sheet_name, header_row) for whichever
    sheet's header row scores highest - same rationale as
    receivable_processor._pick_best_sheet: real exports are sometimes
    multi-sheet workbooks bundling the real data alongside unrelated tabs.
    `xls` is an already-open pd.ExcelFile (see _read_workbook_any_engine)."""
    best = None  # (score, sheet_name, header_row)
    for sn in xls.sheet_names:
        preview = pd.read_excel(xls, sheet_name=sn, header=None, nrows=max_scan)
        header_row, score = find_header_row(preview, alias_lookup, max_scan=max_scan,
                                             min_score=min_score, return_score=True)
        if header_row is None:
            continue
        if best is None or score > best[0]:
            best = (score, sn, header_row)
    if best is None:
        raise ValueError(
            "Could not auto-detect the header row in any sheet of this workbook. "
            "If this file uses different headers than expected, add them to the "
            "alias list in premium_register_processor.py."
        )
    _, sheet_name, header_row = best
    return sheet_name, header_row


def _read_workbook_any_engine(filepath):
    """pandas/openpyxl can't read the legacy binary .xls format at all
    (only .xlsx) - that requires the separate 'xlrd' package. Try the
    normal read first; if it fails specifically because xlrd isn't
    installed, fall back to converting the file to .xlsx with a local
    LibreOffice install (if one is present on this machine), which
    produces an equivalent, cleanly readable file. If neither xlrd nor
    LibreOffice is available, raise a clear, actionable error rather than
    a confusing pandas/import traceback."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext != ".xls":
        return pd.ExcelFile(filepath)

    try:
        return pd.ExcelFile(filepath, engine="xlrd")
    except ImportError:
        pass  # xlrd not installed - try the LibreOffice fallback below
    except Exception:
        pass  # any other xlrd read issue - still worth trying the fallback

    import shutil
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError(
            f"Could not read '{os.path.basename(filepath)}' - it's a legacy .xls file, which "
            "needs either the 'xlrd' package (pip install xlrd>=2.0.1) or a local LibreOffice "
            "install to convert it. Neither was found on this machine."
        )

    tmp_dir = tempfile.mkdtemp(prefix="premreg_xls_")
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", "xlsx", "--outdir", tmp_dir, filepath],
        capture_output=True, timeout=180,
    )
    converted = os.path.join(tmp_dir, os.path.splitext(os.path.basename(filepath))[0] + ".xlsx")
    if result.returncode != 0 or not os.path.exists(converted):
        raise RuntimeError(
            f"Could not read '{os.path.basename(filepath)}' - LibreOffice conversion from .xls "
            f"to .xlsx failed: {result.stderr.decode(errors='ignore')[:300]}"
        )
    return pd.ExcelFile(converted)


def load_raw_data(filepath, aliases_dict, sheet_name=None, prefer_sheet_name_hints=None):
    """
    Load a source workbook, auto-detecting its real header row (skipping
    any banner/title rows above it) and dropping fully-blank rows.
    sheet_name: leave as None to auto-detect which sheet to use; pass an
        explicit name/index to read one sheet directly instead.
    prefer_sheet_name_hints: optional list of compact name hints (see
        _find_sheet_by_name) checked FIRST, before falling back to
        scanning every sheet's header content - use this when the target
        sheet is commonly identifiable by name (e.g. "Premium Payable"),
        since column-header content alone can be ambiguous across a
        multi-sheet settlement workbook where several tabs share similar
        columns. Ignored if sheet_name is given explicitly.
    """
    alias_lookup = _build_alias_lookup(aliases_dict)
    xls = _read_workbook_any_engine(filepath)

    if sheet_name is None:
        picked = None
        if prefer_sheet_name_hints:
            name_candidate = _find_sheet_by_name(xls, prefer_sheet_name_hints)
            if name_candidate is not None:
                preview = pd.read_excel(xls, sheet_name=name_candidate, header=None, nrows=25)
                header_row, score = find_header_row(preview, alias_lookup, max_scan=25,
                                                      min_score=1, return_score=True)
                if header_row is not None:
                    picked = (name_candidate, header_row)
                # else: the name-matched sheet didn't actually contain a
                # recognizable header row (e.g. empty, or laid out
                # differently than expected) - fall through to the
                # full content-based scan below rather than failing outright.
        if picked is None:
            picked = _pick_best_sheet(xls, alias_lookup)
        sheet_name, header_row = picked
    else:
        preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=25)
        header_row = find_header_row(preview, alias_lookup)

    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    columns = raw.iloc[header_row].tolist()
    data = raw.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = _dedupe_columns(columns)
    data = data.dropna(how="all").reset_index(drop=True)
    data.attrs["source_sheet"] = sheet_name
    return data


def _dedupe_columns(columns):
    """Mirror pandas' own read_excel convention for a header row that (as
    read via header=None, so pandas' usual auto-handling doesn't apply)
    may contain blank/NaN cells or repeated text - blanks become
    'Unnamed: N' and repeats get a '.1', '.2', ... suffix, so every
    column name ends up unique and non-null. Without this, a real export
    with a repeated or blank header (seen in practice - e.g. "NUMS"
    appearing twice) silently collides into a single duplicate-labeled
    column, which breaks any later `df[col_name] = ...` assignment."""
    seen = {}
    result = []
    for i, c in enumerate(columns):
        name = f"Unnamed: {i}" if pd.isna(c) or str(c).strip() == "" else str(c).strip()
        if name in seen:
            seen[name] += 1
            name = f"{name}.{seen[name]}"
        else:
            seen[name] = 0
        result.append(name)
    return result


# ---------------------------------------------------------------------------
# Crystal Reports export quirk: some source columns carry the field's own
# caption baked into every cell's value (e.g. "Leader Ref. No.  PL-0720-
# 200501-F01-000053" instead of just "PL-0720-200501-F01-000053"),
# separated from the real value by a run of 2+ spaces. Stripping this is
# safe and generic: real reference numbers in these exports don't
# themselves contain runs of 2+ spaces, so splitting on the first such
# run and keeping only what's after it reliably recovers the real value
# and is a no-op on values that were already clean.
# ---------------------------------------------------------------------------

_LABEL_PREFIX_RE = re.compile(r"\s{2,}")


def strip_label_prefix(value):
    if pd.isna(value):
        return value
    text = str(value)
    parts = _LABEL_PREFIX_RE.split(text, maxsplit=1)
    if len(parts) == 2 and parts[1].strip():
        return parts[1].strip()
    return text.strip()


# ---------------------------------------------------------------------------
# Loading each file
# ---------------------------------------------------------------------------

def load_leader_file(filepath, sheet_name=None):
    """
    Load the leader's own policy file. Only the policy-number column
    (aliased to "Leader Policy Number", e.g. "NUMS") is identified and
    used for matching - every other column is left exactly as-is (name
    and value), since the leader's own file layout varies by partner and
    none of the rest needs to be understood to do the match.

    Real leader files are frequently multi-sheet settlement workbooks
    (e.g. a "Summary" tab, a "Premium Payable" tab, "Claim Rec"/"Prem Rec"
    tabs, and one or more dated snapshot tabs) rather than a single plain
    sheet, and more than one of those tabs can independently contain a
    "POLICY NO."-style column - so column content alone isn't always
    enough to reliably pick the right one. sheet_name=None (the default)
    therefore first looks for a sheet whose own NAME identifies it as the
    Premium Payable sheet (see LEADER_SHEET_NAME_HINTS - matches regardless
    of spacing/punctuation, e.g. "Premium Payable", "PremPayable", "Prem.
    Payable - AIC"), and only falls back to scanning every sheet's header
    content (the original behavior) if no sheet name matches or the
    name-matched sheet's header can't be located. Pass an explicit
    sheet_name to skip auto-detection entirely.

    Returns (df, nums_col) where nums_col is the ORIGINAL column name that
    was matched, so the caller can reference the real data.
    """
    df = load_raw_data(filepath, LEADER_FILE_ALIASES, sheet_name=sheet_name,
                        prefer_sheet_name_hints=LEADER_SHEET_NAME_HINTS)
    alias_lookup = _build_alias_lookup(LEADER_FILE_ALIASES)
    colmap = _map_columns_to_canonical(df.columns, alias_lookup)
    inv = {}
    for orig, canon in colmap.items():
        inv.setdefault(canon, orig)
    if "Leader Policy Number" not in inv:
        raise ValueError(
            f"Could not find the leader's own policy-number column (e.g. 'NUMS') in "
            f"'{os.path.basename(filepath)}'. Recognized headers were: {list(df.columns)}. "
            "Add the real header text to LEADER_FILE_ALIASES in premium_register_processor.py."
        )
    return df, inv["Leader Policy Number"]


def load_register_file(filepath, sheet_name=None):
    """
    Load IGI's own Premium Register. Identifies the "Leader Ref Number"
    column (the leader's policy number as recorded on our side - this is
    what NUMS is matched against) and the "IGI Policy Number" column (what
    gets pulled into the output). The Leader Ref Number column is cleaned
    of the Crystal Reports label-prefix quirk (see strip_label_prefix).
    Returns (df, leader_ref_col, igi_policy_col).
    """
    df = load_raw_data(filepath, REGISTER_ALIASES, sheet_name=sheet_name,
                        prefer_sheet_name_hints=REGISTER_SHEET_NAME_HINTS)
    alias_lookup = _build_alias_lookup(REGISTER_ALIASES)
    colmap = _map_columns_to_canonical(df.columns, alias_lookup)
    inv = {}
    for orig, canon in colmap.items():
        inv.setdefault(canon, orig)
    missing = [c for c in ("Leader Ref Number", "IGI Policy Number") if c not in inv]
    if missing:
        raise ValueError(
            f"Could not find required column(s) {missing} in '{os.path.basename(filepath)}'. "
            f"Recognized headers were: {list(df.columns)}. Add the real header text to "
            "REGISTER_ALIASES in premium_register_processor.py."
        )
    leader_ref_col, igi_policy_col = inv["Leader Ref Number"], inv["IGI Policy Number"]
    df = df.copy()
    df[leader_ref_col] = df[leader_ref_col].apply(strip_label_prefix)
    # Drop rows with no real leader ref / IGI policy number to match against
    # (e.g. a trailing "<<< End of Report >>>" footer row some Crystal
    # Reports exports leave behind).
    df = df[df[leader_ref_col].notna() & (df[leader_ref_col].astype(str).str.strip() != "")]
    df = df.reset_index(drop=True)
    return df, leader_ref_col, igi_policy_col


# ---------------------------------------------------------------------------
# Fuzzy matching
# ---------------------------------------------------------------------------

BLOCK_WINDOW = 4           # blocking key length (sliding window over each digit run) -
                           # any longer shared substring necessarily also shares a
                           # window of this length, so one window size covers both
FALLBACK_WINDOW = 3        # used only if no candidates turn up at BLOCK_WINDOW (short numbers)
_DIGIT_RUN_RE = re.compile(r"\d+")


def _block_keys(value, window):
    """
    Blocking keys for one value: a sliding window of `window` digits over
    every contiguous digit run in the value. A single maximal digit run
    (e.g. "2691025585" from "2691025585/12/2023") is NOT itself a useful
    key on its own - the number that actually needs to match is often
    only a SUBSTRING of a longer run on one side (our own "Leader#"
    column tends to glue a policy's core number together with an
    adjoining date/suffix with no separator, e.g. "...025585..." where
    "25585" is the leader's own reference embedded inside a longer
    digit-only run). Sliding a fixed-width window across every run - on
    both the query and the index side - means a shared embedded number
    still produces at least one identical key, regardless of what
    surrounds it on either side. Runs shorter than `window` are used
    whole (nothing shorter to slide).
    """
    keys = set()
    for run in _DIGIT_RUN_RE.findall(str(value)):
        if len(run) < 3:
            continue
        if len(run) <= window:
            keys.add(run)
        else:
            for i in range(len(run) - window + 1):
                keys.add(run[i:i + window])
    return keys


def build_match_index(register_df, leader_ref_col, window=BLOCK_WINDOW):
    """{block_key: set(row_index)} built from every register row's Leader
    Ref Number - the blocking step that keeps matching fast and, in
    practice, is also what actually narrows in on the right candidate(s)."""
    index = {}
    for idx, val in register_df[leader_ref_col].items():
        for key in _block_keys(val, window):
            index.setdefault(key, set()).add(idx)
    return index


def _digits_only(value):
    return re.sub(r"\D", "", str(value))


def _score_candidate(query_value, candidate_value):
    """
    Score one candidate register row against a query value.

    The primary signal is NOT overall character-similarity (a generic
    ratio over the whole string is easily fooled - two numbers that just
    happen to share a lot of common digits in roughly the same positions
    can out-score the actual matching pair, e.g. "...115585..." vs the
    true "...025585..." both partially resemble a query containing
    "25585"). What actually identifies the right row is whether one of
    the query's own COMPLETE digit runs (its policy/endorsement number as
    a whole token, not a fragment) shows up intact inside the candidate -
    that's a much more specific, much less coincidental signal. So:
      1. primary score = the length of the longest of the query's own
         digit runs (3+ digits) that appears verbatim as a substring of
         the candidate's digit string
      2. tiebreaker = a generic digits-only SequenceMatcher ratio, only
         used to rank among candidates tied on (1) - e.g. when nothing
         matched exactly, or several candidates share an equally long run
    Returns (primary_score, ratio).
    """
    cand_digits = _digits_only(candidate_value)
    query_digits = _digits_only(query_value)
    best_run_len = 0
    for run in re.findall(r"\d+", str(query_value)):
        if len(run) >= 3 and len(run) > best_run_len and run in cand_digits:
            best_run_len = len(run)
    ratio = SequenceMatcher(None, query_digits, cand_digits).ratio()
    return best_run_len, ratio


MIN_RUN_MATCH = 4   # a fully-contained query run shorter than this is too common to trust alone
MIN_RATIO_FALLBACK = 0.75  # only used to accept a match with no qualifying full-run overlap


def match_one(nums_value, register_df, leader_ref_col, igi_policy_col, index):
    """
    Try to match a single leader policy number against the register.
    Returns a dict: {"igi_policy_number", "matched_leader_ref", "score",
    "status"} where status is "Matched" or "Not Matched" (igi_policy_number
    is None when not matched).
    """
    if pd.isna(nums_value) or str(nums_value).strip() == "":
        return {"igi_policy_number": None, "matched_leader_ref": None, "score": 0.0, "status": "Not Matched"}

    candidates = set()
    for k in _block_keys(nums_value, BLOCK_WINDOW):
        candidates |= index.get(k, set())
    if not candidates:
        # Short numbers (all runs under BLOCK_WINDOW digits) may not have
        # produced any BLOCK_WINDOW-length key at all - retry narrower.
        for k in _block_keys(nums_value, FALLBACK_WINDOW):
            candidates |= index.get(k, set())

    if not candidates:
        return {"igi_policy_number": None, "matched_leader_ref": None, "score": 0.0, "status": "Not Matched"}

    best_idx, best_run_len, best_ratio = None, 0, 0.0
    for idx in candidates:
        cand_val = register_df.at[idx, leader_ref_col]
        run_len, ratio = _score_candidate(nums_value, cand_val)
        if (run_len, ratio) > (best_run_len, best_ratio):
            best_idx, best_run_len, best_ratio = idx, run_len, ratio

    accepted = best_idx is not None and (
        best_run_len >= MIN_RUN_MATCH or best_ratio >= MIN_RATIO_FALLBACK
    )
    if accepted:
        return {
            "igi_policy_number": register_df.at[best_idx, igi_policy_col],
            "matched_leader_ref": register_df.at[best_idx, leader_ref_col],
            "score": round(best_ratio, 3),
            "status": "Matched",
        }
    return {"igi_policy_number": None, "matched_leader_ref": None, "score": round(best_ratio, 3), "status": "Not Matched"}


def match_leader_file(leader_df, nums_col, register_df, leader_ref_col, igi_policy_col):
    """
    Match every row of leader_df against register_df. Returns a NEW
    dataframe = leader_df (all original columns, untouched) plus four
    added columns: "IGI Policy Number", "Matched Leader Ref (Register)",
    "Match Score", "Match Status".
    """
    index = build_match_index(register_df, leader_ref_col)

    igi_numbers, matched_refs, scores, statuses = [], [], [], []
    for val in leader_df[nums_col]:
        result = match_one(val, register_df, leader_ref_col, igi_policy_col, index)
        igi_numbers.append(result["igi_policy_number"])
        matched_refs.append(result["matched_leader_ref"])
        scores.append(result["score"])
        statuses.append(result["status"])

    out = leader_df.copy()
    # Some real leader files already ship a blank "IGI Policy Number"
    # column as a placeholder for exactly this - fill that one in rather
    # than adding a second, duplicate-named column next to it.
    out["IGI Policy Number"] = igi_numbers
    out["Matched Leader Ref (Register)"] = matched_refs
    out["Match Score"] = scores
    out["Match Status"] = statuses
    return out


if __name__ == "__main__":
    import sys

    leader_fp = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/Premium_EFU_Jinnah_Division__statement.xlsx"
    register_fp = sys.argv[2] if len(sys.argv) > 2 else "/mnt/user-data/uploads/Premium_Register_KU4.xls"

    leader_df, nums_col = load_leader_file(leader_fp)
    print(f"Leader file: {len(leader_df)} rows, policy-number column detected as '{nums_col}'")

    register_df, leader_ref_col, igi_policy_col = load_register_file(register_fp)
    print(f"Register file: {len(register_df)} rows, Leader Ref column='{leader_ref_col}', "
          f"IGI Policy Number column='{igi_policy_col}'")

    result = match_leader_file(leader_df, nums_col, register_df, leader_ref_col, igi_policy_col)
    matched = (result["Match Status"] == "Matched").sum()
    print(f"\nMatched: {matched} / {len(result)}  ({matched / len(result):.1%})")
    print(result[[nums_col, "IGI Policy Number", "Matched Leader Ref (Register)", "Match Score", "Match Status"]].head(15).to_string())
