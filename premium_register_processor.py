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
differs. So matching works in three stages:
  1. BLOCKING - index every register row by its own digit runs (4+ digit
     substrings), then look up candidate rows for a leader value by its
     own digit runs. This keeps matching fast (no need to compare every
     leader row against all ~24k register rows) and, in practice, is
     also what actually identifies the right candidate(s).
  2. SCORING - among the (usually small) candidate set, pick whichever
     register row shares the longest COMPLETE digit run (its own
     policy/endorsement number as a whole token) with the query, using
     branch/class-code similarity as the primary tiebreaker and overall
     digit-string similarity as a secondary one (see MIN_ALPHA_RATIO
     below for why the branch/class code has to be checked at all).
  3. ACCEPTANCE - a row with no candidates, or whose best candidate
     doesn't clear the acceptance thresholds (see MIN_RUN_MATCH /
     MIN_RATIO_FALLBACK / MIN_ALPHA_RATIO), is left unmatched rather than
     forced onto a poor guess - a wrong policy number would be worse
     than none, since someone would act on it.

BRANCH/CLASS-CODE SAFEGUARD (MIN_ALPHA_RATIO): a real leader file whose
own reference is itself a self-numbering "YYYY/MM/CODE#####" style value
(rather than a plain numeric NUMS-style reference) restarts its serial
numbering back at 1 PER DEPARTMENT/BRANCH - confirmed against a real
Indus Motor Company settlement statement, where digit-only matching
alone had matched a Fire transaction to a Marine register row, and
(system-wide, once checked) roughly a quarter of every "Matched" row in
that run turned out to be matched into the wrong department purely
because two unrelated transactions in different departments happened to
share the same restarting-at-1 serial number. Matching now also requires
the candidate's own branch/class code (the letters embedded in the
reference - see _alpha_only / MIN_ALPHA_RATIO's own docstring, further
down, for the full detail and the real numbers this was calibrated
against) to plausibly agree before a shared digit run is trusted.
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
        # "Docref" - confirmed against a real settlement statement (Indus
        # Motor Company Ltd's coinsurance leader, "Premium Payable" sheet):
        # this partner's own column name for the exact same thing NUMS is
        # elsewhere - their own reference number for the transaction, in
        # a format ("2024/06/GroupEEEOE00035") that lines up structurally
        # with what IGI's own Premium Register stores back in its own
        # "Leader#" column for this same partner. Verified end-to-end:
        # 2,457 of 2,464 rows matched (99.7%) - the remainder are blank
        # department subtotal rows ("TOTAL ENGINEERING" etc.), which
        # correctly and harmlessly fall through to "Not Matched" the same
        # way any other leader file's non-data rows already do.
        "docref",
        # "Policy/PMD #" - the same coinsurance company (or another
        # partner using the same settlement-statement convention) may
        # export their own premium/policy-payable sheet with this header
        # instead of "Docref" - it's the same field CLAIM_ALIASES already
        # recognizes on the claims side of this kind of statement (see
        # claims_register_processor.CLAIM_ALIASES's "Policy/PMD Number"),
        # just needed here too for the premium side. Several
        # punctuation/spacing variants are listed since real exports are
        # inconsistent about where the space falls around the slash and
        # the "#" - confirmed directly against a real file: the actual
        # header text there is "Policy /PMD#" (space BEFORE the slash,
        # none before "#"), which doesn't match the more "tidy" spacing
        # any alias list in this codebase had before - that specific
        # variant is included below, not just the tidier-looking ones.
        "policy/pmd #", "policy /pmd#", "policy/pmd#", "policy /pmd #",
        "policy pmd #", "policy/pmd number", "pmd #",
        # "Documents" - confirmed real header on another leader's own
        # Premium Payable sheet (Jubilee General Insurance's own
        # "Coinsurance Settlement with IGI Insurance Ltd" workbook,
        # Sheet1) - this is that leader's own per-transaction reference
        # number for a specific premium item (e.g.
        # "2023-07-104-E02001OP0000127"), sitting alongside a second,
        # related "Base Documents #" column (the original/base policy an
        # endorsement stems from - not aliased here, since "Documents"
        # alone is the right column to match on: the specific
        # transaction reference, not the base policy it originated
        # from). WHY THIS MATTERS: without this alias, this leader's
        # real Premium Payable sheet (Sheet1, 202 rows) scores 0 against
        # every alias here, while a totally different sheet in the same
        # workbook (Sheet2, "Coinsurance Claims Recoverable", whose own
        # "Policy /PMD#" header already matches an alias above) scores 1
        # and wins by default - so Policy Check was silently matching
        # against 9 usable CLAIMS rows instead of the real 202-row
        # premium file, with no error raised at all. Confirmed against
        # this real file: with "documents" recognized, Sheet1 now
        # correctly outscores Sheet2 for sheet selection (tie-broken in
        # its favor anyway, since it appears first in the workbook), and
        # Policy Check picks up the real premium data instead.
        "documents",
    ],
}

REGISTER_ALIASES = {
    "Leader Ref Number": [
        "leader#", "leader #", "leader no", "leader number", "leader ref no",
        "leader ref. no", "leader ref number", "leader reference number",
        "leader ref", "leaderref",
        # "Base Doc No." - confirmed real header on the current Premium
        # Register export (same file/column Claims Register's own
        # REGISTER_ALIASES already recognizes as its primary matching
        # column) - this module was still only matching the older
        # "Leader#"-style header until this was added, so it silently
        # failed to load the current export format entirely.
        "base doc no", "base doc no.", "base document no",
    ],
    "IGI Policy Number": [
        "doc ref #", "doc ref#", "doc reference", "doc ref number",
        "igi policy number", "igi policy no", "igi doc ref", "doc no",
        # "Document No." - the current export's own name for the specific
        # transaction/document, paired with "Base Doc No." above the same
        # way Claims Register already treats the two columns.
        "document no", "document no.", "document number",
    ],
}

# ---------------------------------------------------------------------------
# "Premium Payable" lookup - for the OUTSTANDING PREMIUM file variant of
# the "register" input (e.g. IGI's own "Co Insurance Outstanding
# Position" export, Detail sheet - confirmed real header "Premium
# Payable"), as opposed to a plain Premium Register export, which has no
# such column at all. The GUI's own "Register file type" radio button
# (Premium Register / Outstanding Premium File) decides whether this is
# looked up and pulled through into the match output as an extra column
# - see app.py's _run_premium_register_match and
# match_leader_file's own premium_payable_col parameter below.
#
# Deliberately kept as its own small, separate alias list rather than
# folded into REGISTER_ALIASES: REGISTER_ALIASES feeds _pick_best_sheet's
# own header-scoring (which sheet "looks like" the real register/
# outstanding-premium data), and this column has no role in MATCHING at
# all - it's purely an extra value pulled from whichever row a match
# already landed on. Keeping it separate means turning "Outstanding
# Premium File" mode on/off can never change which sheet gets picked or
# how matching itself behaves, only whether this one extra column shows
# up in the output.
# ---------------------------------------------------------------------------
PREMIUM_PAYABLE_ALIASES = ["premium payable", "premium payable amount"]


def find_premium_payable_column(df):
    """
    Best-effort lookup of a 'Premium Payable' column in an Outstanding
    Premium (Co-Insurance Outstanding Position) file - e.g. a real
    "Co_Insurance_<date>.xlsx" export's own "Detail" sheet, confirmed
    real header text "Premium Payable" - as opposed to a plain Premium
    Register export, which doesn't carry this column at all.

    Returns the ORIGINAL column name if found, else None - callers
    should treat None as "this file doesn't have a Premium Payable
    column" (e.g. someone picked "Outstanding Premium File" but actually
    browsed to a plain register) and continue without it rather than
    raising, since Premium Payable is an enhancement to the output, not
    a column matching itself depends on.
    """
    for col in df.columns:
        if _normalize_header(col) in PREMIUM_PAYABLE_ALIASES:
            return col
    return None


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


def _find_continuation_sheets(xls, primary_sheet, expected_ncols, alias_lookup, max_scan=25, min_score=1):
    """
    Legacy .xls workbooks are capped at 65,536 rows PER SHEET - a
    register bigger than that (confirmed against a real Premium Register
    export - 204,220 rows split across 4 sheets) gets split across
    Sheet1/Sheet2/Sheet3/... by whatever produced the file, with the real
    header row living ONLY on the first sheet and every later sheet being
    a bare continuation of the same table (no header row of its own, same
    column count). Reading only the sheet _pick_best_sheet()/name-hint
    matching picks (the only one that scores anything against the alias
    list, since it's the only one with a header) silently drops every row
    on every later sheet - a real policy could be sitting in Sheet2/3/4
    and simply never get read, with no error or warning at all.

    Returns an ordered list of sheet names (in workbook order, excluding
    primary_sheet) that look like pure-data continuations of
    primary_sheet: same raw column count, and NO row in the first
    `max_scan` rows scores as a header against alias_lookup (a sheet that
    DOES score is presumably a genuinely different table with its own
    header and must never be merged in as if it were more rows of the
    same table). Same logic as claims_register_processor.py's own
    _find_continuation_sheets, which this was adapted from.
    """
    continuations = []
    for sn in xls.sheet_names:
        if sn == primary_sheet:
            continue
        preview = pd.read_excel(xls, sheet_name=sn, header=None, nrows=max_scan)
        if preview.shape[1] != expected_ncols:
            continue
        _, score = find_header_row(preview, alias_lookup, max_scan=max_scan, min_score=min_score, return_score=True)
        if score == 0:
            continuations.append(sn)
    return continuations


def _read_workbook_any_engine(filepath):
    """pandas/openpyxl can't read the legacy binary .xls format at all
    (only .xlsx) - that requires the separate 'xlrd' package. Try the
    normal read first; if it fails specifically because xlrd isn't
    installed, fall back to converting the file to .xlsx with a local
    LibreOffice install (if one is present on this machine), which
    produces an equivalent, cleanly readable file. If neither xlrd nor
    LibreOffice is available, raise a clear, actionable error rather than
    a confusing pandas/import traceback.

    Returns (xls, resolved_filepath) - xls is a pd.ExcelFile (used for
    the cheap header-row preview scans, which stay on pandas' own
    reader), resolved_filepath is the path load_raw_data should hand to
    perf_utils.fast_read_excel_rows() for the actual full-sheet read
    (the original path for .xlsx or an xlrd-readable .xls, or the
    LibreOffice-converted temp .xlsx path otherwise - python-calamine,
    fast_read_excel_rows()'s preferred reader, can read a genuine .xls
    directly too, so this only takes the LibreOffice detour when even
    that fails)."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext != ".xls":
        return pd.ExcelFile(filepath), filepath

    try:
        return pd.ExcelFile(filepath, engine="xlrd"), filepath
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
    return pd.ExcelFile(converted), converted


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
    xls, resolved_filepath = _read_workbook_any_engine(filepath)

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

    # The header-row scan above always stays a cheap pandas nrows=25
    # preview - only the FULL sheet read goes through the fast path.
    # perf_utils.fast_read_excel_rows() prefers python-calamine (a
    # Rust-based reader, ~7x faster than pandas'/openpyxl's own reader on
    # a real 500,000-row file - see PERFORMANCE_NOTES.md's Part 4),
    # falling back to openpyxl's own read_only streaming mode if
    # calamine isn't installed, and further falling back to plain
    # pd.read_excel() below if perf_utils itself isn't available at all.
    #
    # Continuation sheets (see _find_continuation_sheets' docstring - a
    # legacy .xls register split across Sheet1/2/3/4 at Excel's old
    # 65,536-row-per-sheet cap, header only on the first) are detected
    # and merged in as more rows of the same table, exactly the same fix
    # already applied in claims_register_processor.py - confirmed against
    # a real 204,220-row Premium Register that this module would
    # otherwise have silently read only its first 65,523 rows and missed
    # the remaining 138,697 with no warning at all.
    ext = os.path.splitext(filepath)[1].lower()
    # Fast path only for a genuine .xlsx file. A non-.xlsx file (a legacy
    # .xls, or its LibreOffice-converted .xlsx stand-in) stays on the SAME
    # reader (xlrd/pandas) used for the header-row PREVIEW scan above -
    # this mirrors the identical .xlsx-only gate claims_register_processor.py's
    # own _read_full_sheet() already uses, for the same reason: python-calamine
    # (perf_utils.fast_read_excel_rows' preferred reader) trims genuinely-
    # blank leading/trailing columns that xlrd/pandas does not. Confirmed
    # against a real 204,220-row register split across 4 sheets at Excel's
    # old 65,536-row cap: Sheet1's own leading blank column made calamine
    # report 18 columns where the preview scan reported 19, so
    # _find_continuation_sheets()'s column-count check silently rejected
    # Sheet2/3/4 as "not real continuations" and this function returned
    # only Sheet1's ~65,520 rows - 32% of the real register - with no
    # error or warning at all. Keeping both reads on the same library
    # removes that discrepancy entirely.
    use_fast_path = (ext == ".xlsx")
    if use_fast_path:
        try:
            from perf_utils import fast_read_excel_rows
        except ImportError:
            use_fast_path = False

    if use_fast_path:
        rows, _n = fast_read_excel_rows(resolved_filepath, sheet_name)
        # Re-locate the header row directly against THESE rows, rather
        # than blindly reusing the row index found by the separate
        # xlrd/pandas preview scan above.
        #
        # CONFIRMED REAL BUG this guards against: for a real legacy .xls
        # file with a fully-blank leading row (a Crystal Reports export
        # quirk - a title/banner row so empty even the row itself is
        # blank), xlrd (used for the preview scan above) and
        # python-calamine (perf_utils.fast_read_excel_rows' preferred
        # reader, used for this full read) disagreed on whether that
        # blank row exists at all: xlrd reported it (902 rows), calamine
        # silently dropped it (901 rows) - a genuine cross-library
        # discrepancy, not a bug in either reader's row COUNT/CONTENT
        # otherwise. Reusing the xlrd-side header_row index against
        # calamine's own, one-row-shorter row list then landed one row
        # too early: the first real DATA row got read as if it were the
        # header, and every subsequent row's "Base Doc No"/"Leader Ref
        # Number"/etc. value silently paired with the WRONG row's other
        # columns from then on - wrong, but with no error raised, so it
        # would have gone completely unnoticed. Re-scanning the actual
        # rows about to be used removes any dependency on the two readers
        # agreeing on row alignment; if it can't find a header (e.g. a
        # very short/unusual sheet), it safely falls back to the
        # xlrd-side header_row found above rather than raising.
        preview_from_rows = pd.DataFrame(rows[:25])
        fast_header_row, _fast_score = find_header_row(preview_from_rows, alias_lookup, max_scan=25,
                                                         min_score=1, return_score=True)
        if fast_header_row is not None:
            header_row = fast_header_row
        columns = _dedupe_columns(list(rows[header_row]))
        data = pd.DataFrame(rows[header_row + 1:], columns=columns)

        continuation_sheets = _find_continuation_sheets(xls, sheet_name, len(columns), alias_lookup)
        if continuation_sheets:
            frames = [data]
            for sn in continuation_sheets:
                cont_rows, _n2 = fast_read_excel_rows(resolved_filepath, sn)
                frames.append(pd.DataFrame(cont_rows, columns=columns))
            data = pd.concat(frames, ignore_index=True)
    else:
        raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        columns = _dedupe_columns(raw.iloc[header_row].tolist())
        data = raw.iloc[header_row + 1:].reset_index(drop=True)
        data.columns = columns

        continuation_sheets = _find_continuation_sheets(xls, sheet_name, len(columns), alias_lookup)
        if continuation_sheets:
            frames = [data]
            for sn in continuation_sheets:
                cont_raw = pd.read_excel(xls, sheet_name=sn, header=None)
                cont_raw.columns = columns
                frames.append(cont_raw)
            data = pd.concat(frames, ignore_index=True)

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

# ---------------------------------------------------------------------------
# Block-purge cap - fixes a confirmed performance blow-up.
#
# A 4-digit sliding-window key that happens to be a calendar year (e.g.
# "2025") is common to a huge fraction of the register, since almost
# every real reference embeds a date somewhere. Verified against a
# synthetic-but-realistic ~24,000-row register spanning a handful of
# policy years: a handful of these year-shaped keys each collided with
# 2,000-4,000 register rows, and every leader row that happened to share
# one had to be fuzzy-scored (SequenceMatcher) against that entire
# bloated candidate pool - 100-200ms per leader row, which alone was
# enough to take a single "Run Matching" click from seconds to tens of
# minutes on an ordinary production-sized file.
#
# This is the exact same collision problem already diagnosed and fixed
# in claims_register_processor.py for ITS OWN blocking - but that
# module's fix (raising the minimum indexable digit-run length to 5, and
# dropping the sliding window entirely) can't just be copy-pasted here:
# THIS module's matching specifically depends on the sliding window to
# catch a short embedded number glued inside a longer, dirtier digit run
# with no separator (see _block_keys' docstring above) - a leader's own
# 5-digit serial can appear embedded inside our register's own longer
# run. Raising the minimum run length would lose exactly the matches
# this module exists to catch.
#
# Instead, MAX_BLOCK_SIZE purges any block key AFTER the index is built
# that turned out to match more than this many register rows - i.e. a
# key so generic (a bare year, a common short prefix) that it provides
# essentially zero discriminating power anyway, at the cost of dominating
# match time. This keeps the sliding window (and therefore the embedded-
# substring matching) fully intact: a genuine match's real discriminating
# substring - the policy/endorsement's own short numeric core - is
# virtually never shared by hundreds of unrelated register rows, so it
# survives the purge untouched. Since _block_keys() already emits one key
# per window position (multiple keys per reference), losing one overly
# generic key (e.g. the embedded year) still leaves the specific ones in
# place to do the real matching - nothing needs to fall back to
# FALLBACK_WINDOW just because one of several keys got purged.
MAX_BLOCK_SIZE = 300


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


def build_match_index(register_df, leader_ref_col, window=BLOCK_WINDOW, max_block_size=MAX_BLOCK_SIZE):
    """{block_key: set(row_index)} built from every register row's Leader
    Ref Number - the blocking step that keeps matching fast and, in
    practice, is also what actually narrows in on the right candidate(s).

    max_block_size: after the index is built, any key whose candidate set
        grew larger than this is dropped entirely - see MAX_BLOCK_SIZE's
        docstring above for why (a key that common is providing no real
        filtering, only cost, and the reference it came from will almost
        always still have OTHER, more specific keys surviving in the
        index). Pass None/0 to disable purging and restore the original,
        unbounded behavior."""
    index = {}
    for idx, val in register_df[leader_ref_col].items():
        for key in _block_keys(val, window):
            index.setdefault(key, set()).add(idx)
    if max_block_size:
        for key in [k for k, candidates in index.items() if len(candidates) > max_block_size]:
            del index[key]
    return index


def _digits_only(value):
    return re.sub(r"\D", "", str(value))


_ALPHA_ONLY_RE = re.compile(r"[^A-Za-z]")


def _alpha_only(value):
    """Letters-only, uppercased form of a reference (digits/slashes/dashes
    stripped) - e.g. "2025/03/GroupMIPOE00033" -> "GROUPMIPOE". This is
    the branch/department/class code embedded in the reference - see
    MIN_ALPHA_RATIO's own docstring below for why comparing this is what
    fixes a confirmed real cross-department mismatch bug."""
    return _ALPHA_ONLY_RE.sub("", str(value)).upper()


def _strip_common_prefix(a, b):
    """Return (a, b) with whatever LITERAL prefix they share verbatim
    removed from both - e.g. ("GROUPMIPOE", "GROUPSMTOE") -> ("MIPOE",
    "SMTOE"). Used only for the alpha/class-code comparison (see
    MIN_ALPHA_RATIO) to stop a boilerplate word shared by nearly every
    reference in a file (confirmed real case: "Group", present at the
    start of virtually every reference in a real Indus Motor Company
    settlement statement) from inflating SequenceMatcher's ratio and
    masking the ACTUAL class-code difference that sits right after it.
    Computed per-pair, dynamically, from whatever the two strings being
    compared actually share - nothing about "Group" (or any other
    specific word) is hardcoded, so this works the same way for a
    partner file that has no shared boilerplate at all (nothing gets
    stripped, and the comparison falls back to the plain full-string
    ratio unchanged)."""
    n = 0
    limit = min(len(a), len(b))
    while n < limit and a[n] == b[n]:
        n += 1
    return a[n:], b[n:]


def _score_candidate(query_runs, cand_digits):
    """
    Cheap first-pass score: the length of the longest of the query's own
    digit runs (3+ digits) that appears verbatim as a substring of the
    candidate's digit string - no SequenceMatcher call here at all.

    The primary signal for a real match is NOT overall character-
    similarity (a generic ratio over the whole string is easily fooled -
    two numbers that just happen to share a lot of common digits in
    roughly the same positions can out-score the actual matching pair,
    e.g. "...115585..." vs the true "...025585..." both partially
    resemble a query containing "25585"). What actually identifies the
    right row is whether one of the query's own COMPLETE digit runs (its
    policy/endorsement number as a whole token, not a fragment) shows up
    intact inside the candidate - that's a much more specific, much less
    coincidental signal, and it's cheap to check (a plain substring test)
    compared to a full SequenceMatcher ratio.

    query_runs: the query's own qualifying 3+ digit runs, precomputed
        ONCE per query by the caller (match_one), not re-derived via
        regex on every candidate.
    cand_digits: this candidate's own precomputed digits-only string (see
        match_leader_file's register_digits_series) - also precomputed
        rather than re-derived per candidate.

    Returns best_run_len only. The SequenceMatcher-based scoring (alpha
    similarity, then digit-ratio tiebreaker - see MIN_ALPHA_RATIO and
    match_one) is scored separately, and only for whichever candidates
    tie for the best run_len found across the WHOLE candidate set - see
    match_one. A candidate with a shorter run_len can never win
    regardless of its alpha/digit ratio, since acceptance/ranking always
    compares (run_len, alpha_ratio, digit_ratio) with run_len first, so
    there is nothing to gain by also scoring it - this two-pass split is
    what keeps matching fast: it's the same (run_len, ...) comparison as
    before, just skipping the expensive part for candidates that can't
    possibly win it.
    """
    best_run_len = 0
    for run in query_runs:
        if len(run) > best_run_len and run in cand_digits:
            best_run_len = len(run)
    return best_run_len


MIN_RUN_MATCH = 4   # a fully-contained query run shorter than this is too common to trust alone
MIN_RATIO_FALLBACK = 0.75  # only used to accept a match with no qualifying full-run overlap

# ---------------------------------------------------------------------------
# Branch/class-code safeguard - fixes a confirmed real mismatch bug.
#
# CONFIRMED BUG: matching a real leader file (Indus_Settlement_F.xlsx)
# against a real register (Indus_Motor_Premium.xls) produced Policy
# Check output where a Fire-department transaction had been matched to a
# Marine register row, and - once checked properly - roughly a quarter
# of ALL "Matched" rows in that run turned out to be matched to a
# register row in a DIFFERENT department (Marine matched to Engineering,
# Fire, or Miscellaneous) than the leader's own transaction, plus a
# further, larger set matched to a register row in the SAME broad
# department but a genuinely different, unrelated sub-class/transaction.
#
# ROOT CAUSE: this leader's own reference (e.g. "2025/01/GroupMIPOC00024")
# and the register's own Leader Ref Number/Base Doc No (which records
# that same leader-assigned reference back, verbatim) share the exact
# "YYYY/MM/CODE#####" shape claims_register_processor.py's own module
# docstring already documents as prone to this collision: every branch/
# department/class restarts its OWN serial numbering back at 1, so two
# completely unrelated transactions in two different departments can
# carry the IDENTICAL digit string (confirmed directly: a real Marine
# transaction's serial "00024" and a real, unrelated Miscellaneous
# transaction's serial "00024" are digit-for-digit identical -
# "20250100024" - even though the register's own Department/Class field
# says "Miscellaneous"/"ALL RISK INSURANCE" for that row, nothing to do
# with the leader's actual Marine transaction). The matcher above only
# ever compares DIGITS, so a shared serial number alone was enough to
# force a confident match across completely unrelated departments -
# exactly the "a wrong policy number would be worse than none" failure
# this module's own docstring warns against, just not actually prevented
# for this reference shape.
#
# FIX: require the candidate's own branch/class CODE (the letters
# embedded in the reference, see _alpha_only) to also plausibly agree
# before a digit-run match is accepted - the same safeguard
# claims_register_processor.py already applies via its own
# MIN_ALPHA_RATIO, adapted here for this module's specific data shape.
#
# WHY THE COMPARISON IS DONE DIFFERENTLY THAN CLAIMS REGISTER'S OWN:
# claims_register_processor.py compares the FULL alpha string as-is,
# which works there because its real reference format has no big shared
# boilerplate word diluting the comparison. Confirmed this ISN'T true
# here: this leader's own convention prefixes virtually EVERY reference
# with the literal word "Group" before the actual class code (e.g.
# "GroupMIPOC", "GroupSMTOE", "GroupEEEOE" - Marine, Miscellaneous, and
# Engineering respectively), so a plain whole-string SequenceMatcher
# ratio between two DIFFERENT classes still comes out misleadingly high
# (~0.6-0.7) purely from that shared "Group" prefix, nowhere near
# specific enough to catch the real bug above. _strip_common_prefix()
# removes whatever literal prefix the two strings being compared happen
# to share (not hardcoded to "Group" specifically, so this degrades
# gracefully to the plain full-string comparison for a partner file with
# no such boilerplate) before scoring similarity on what's left - i.e.
# the part that actually carries the class-code information.
#
# CALIBRATED against the real file pair above, using a ground-truth
# check independent of this fix (whether the leader's own exact class
# code + serial number combination is present ANYWHERE in the register
# at all - if it isn't, no possible match for that row can be correct):
# at this threshold, 125 of 126 genuinely-correct matches are kept
# (99.2%) while 2,210 of 2,331 genuinely-wrong "matches" the old,
# digits-only logic had produced are correctly rejected (94.8%) - a
# clearly favorable trade, and a large step up from the 0.5 threshold
# claims_register_processor.py uses for its own (differently-shaped,
# non-prefix-diluted) comparison, which is why this module keeps its own
# separate constant rather than importing that one.
MIN_ALPHA_RATIO = 0.7


def match_one(nums_value, register_df, leader_ref_col, igi_policy_col, index, leader_ref_series=None, igi_policy_series=None,
              register_digits_series=None, register_alpha_series=None, premium_payable_series=None):
    """
    Try to match a single leader policy number against the register.
    Returns a dict: {"igi_policy_number", "matched_leader_ref", "score",
    "status"} where status is "Matched" or "Not Matched" (igi_policy_number
    is None when not matched). If premium_payable_series is given, the
    dict also carries a "premium_payable" key (that matched row's own
    Premium Payable value, or None if not matched) - see
    match_leader_file's own premium_payable_col parameter.

    leader_ref_series / igi_policy_series: optional precomputed
    register_df[leader_ref_col] / register_df[igi_policy_col] Series (see
    match_leader_file) - built on the fly from register_df if not
    supplied, for standalone/one-off use.
    register_digits_series: optional precomputed
    register_df[leader_ref_col].map(_digits_only) Series (see
    match_leader_file) - built on the fly if not supplied.
    register_alpha_series: optional precomputed
    register_df[leader_ref_col].map(_alpha_only) Series (see
    match_leader_file) - built on the fly if not supplied. Used for the
    branch/class-code safeguard - see MIN_ALPHA_RATIO's own docstring.
    premium_payable_series: optional precomputed register_df[premium_payable_col]
    Series - only present at all when the person selected "Outstanding
    Premium File" as the register file type AND that column was actually
    found in it (see find_premium_payable_column). None (the default)
    means this run isn't carrying Premium Payable through at all, in
    which case the returned dict has no "premium_payable" key.

    PERFORMANCE (verified via cProfile against a real-shaped 200,000-row
    register / 15,000-row leader file):
      1. Candidate lookups use precomputed SERIES via Series.at[idx], not
         register_df.at[idx, col] on the full DataFrame. DataFrame.at[]
         on a 2D frame goes through a much heavier path than Series.at[]
         on a 1D column (column lookup, scalar "boxing", and a pandas
         __finalize__ call that deep-copies frame metadata on every
         single access) - repeating that once per CANDIDATE, for every
         leader row, was one of the two dominant costs of matching
         (measured at ~13s of a ~34s run on a 2,000-row subset, before
         this fix). Roughly 5x faster for this "same column, many
         different row indices" access pattern.
      2. Scoring is two passes: a cheap run_len-only check (plain
         substring test, no SequenceMatcher) across every candidate
         first, THEN SequenceMatcher-based alpha/digit scoring only for
         whichever candidates tie for the best run_len found - see
         _score_candidate's own docstring for why this is provably the
         same result as scoring every candidate fully, just faster. This
         was the OTHER dominant cost (~18s of the same ~34s run, mostly
         inside Python's own difflib module) even before the alpha check
         was added - adding one more SequenceMatcher call per finalist
         (there are normally only a handful of these, tied-for-best
         candidates, not the whole candidate pool) keeps this fast.
    """
    if leader_ref_series is None:
        leader_ref_series = register_df[leader_ref_col]
    if igi_policy_series is None:
        igi_policy_series = register_df[igi_policy_col]
    if register_digits_series is None:
        register_digits_series = leader_ref_series.astype(str).map(_digits_only)
    if register_alpha_series is None:
        register_alpha_series = leader_ref_series.astype(str).map(_alpha_only)

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

    query_digits = _digits_only(nums_value)
    query_runs = [r for r in re.findall(r"\d+", str(nums_value)) if len(r) >= 3]
    query_alpha = _alpha_only(nums_value)

    # Pass 1: cheap run_len-only scoring for every candidate.
    best_run_len = 0
    run_len_by_idx = {}
    for idx in candidates:
        rl = _score_candidate(query_runs, register_digits_series.at[idx])
        run_len_by_idx[idx] = rl
        if rl > best_run_len:
            best_run_len = rl
    finalists = [idx for idx, rl in run_len_by_idx.items() if rl == best_run_len]

    # Pass 2: for the tied-for-best-run_len finalists only, score BOTH the
    # branch/class-code similarity (alpha_ratio - see MIN_ALPHA_RATIO's
    # docstring for why this now comes first) and the overall digit
    # similarity (digit_ratio, the pre-existing tiebreaker). Ranking by
    # (alpha_ratio, digit_ratio) - rather than digit_ratio alone, as
    # before - means that among two candidates sharing the same digit run
    # (the exact collision MIN_ALPHA_RATIO exists to catch), whichever
    # one's branch/class code actually agrees with the query wins, not
    # whichever happens to have marginally more similar digits overall.
    best_idx, best_alpha_ratio, best_ratio = None, 0.0, 0.0
    for idx in finalists:
        cand_alpha = register_alpha_series.at[idx]
        a_stripped, q_stripped = _strip_common_prefix(cand_alpha, query_alpha)
        alpha_ratio = SequenceMatcher(None, q_stripped, a_stripped).ratio()
        ratio = SequenceMatcher(None, query_digits, register_digits_series.at[idx]).ratio()
        if (alpha_ratio, ratio) > (best_alpha_ratio, best_ratio):
            best_idx, best_alpha_ratio, best_ratio = idx, alpha_ratio, ratio

    accepted = (
        best_idx is not None
        and (best_run_len >= MIN_RUN_MATCH or best_ratio >= MIN_RATIO_FALLBACK)
        and best_alpha_ratio >= MIN_ALPHA_RATIO
    )
    if accepted:
        result = {
            "igi_policy_number": igi_policy_series.at[best_idx],
            "matched_leader_ref": leader_ref_series.at[best_idx],
            "score": round(best_ratio, 3),
            "status": "Matched",
        }
        if premium_payable_series is not None:
            result["premium_payable"] = premium_payable_series.at[best_idx]
        return result
    result = {"igi_policy_number": None, "matched_leader_ref": None, "score": round(best_ratio, 3), "status": "Not Matched"}
    if premium_payable_series is not None:
        result["premium_payable"] = None
    return result


def match_leader_file(leader_df, nums_col, register_df, leader_ref_col, igi_policy_col, report=None,
                       premium_payable_col=None):
    """
    Match every row of leader_df against register_df. Returns a NEW
    dataframe = leader_df (all original columns, untouched) plus four
    added columns: "IGI Policy Number", "Matched Leader Ref (Register)",
    "Match Score", "Match Status" - and, when premium_payable_col is
    given, a fifth added column "Premium Payable" (that matched row's
    own Premium Payable value from the register, blank for an unmatched
    row).

    premium_payable_col: the ORIGINAL column name in register_df holding
        Premium Payable (see find_premium_payable_column) - only ever
        passed when the person is using an Outstanding Premium file as
        the register input (see app.py's "Register file type" radio
        button) rather than a plain Premium Register, which has no such
        column. Leave as None (the default) to get the original 4-column
        output unchanged - this is purely additive and never affects
        matching itself.

    report: optional callable(str) - if it's a perf_utils.Reporter (i.e.
        came from perf_utils.run_in_background), row-level matching
        progress is also pushed to the numeric progress bar channel via
        perf_utils.RowProgress. Safe to leave as None for normal,
        non-GUI use.
    """
    index = build_match_index(register_df, leader_ref_col)
    # Precompute these ONCE for the whole register, rather than letting
    # match_one() default them per call - see match_one's own docstring
    # for why passing the same Series in on every call (instead of
    # DataFrame.at[idx, col], and instead of re-deriving digits-only/
    # letters-only strings via regex on every single candidate check) is
    # the fix that actually matters here.
    leader_ref_series = register_df[leader_ref_col]
    igi_policy_series = register_df[igi_policy_col]
    register_digits_series = leader_ref_series.astype(str).map(_digits_only)
    register_alpha_series = leader_ref_series.astype(str).map(_alpha_only)
    premium_payable_series = register_df[premium_payable_col] if premium_payable_col else None

    try:
        from perf_utils import RowProgress
        progress = RowProgress(report, total=len(leader_df), label="leader rows matched") if report else None
    except ImportError:
        progress = None

    igi_numbers, matched_refs, scores, statuses = [], [], [], []
    premium_payables = [] if premium_payable_col else None
    for i, val in enumerate(leader_df[nums_col], start=1):
        result = match_one(val, register_df, leader_ref_col, igi_policy_col, index,
                            leader_ref_series=leader_ref_series, igi_policy_series=igi_policy_series,
                            register_digits_series=register_digits_series, register_alpha_series=register_alpha_series,
                            premium_payable_series=premium_payable_series)
        igi_numbers.append(result["igi_policy_number"])
        matched_refs.append(result["matched_leader_ref"])
        scores.append(result["score"])
        statuses.append(result["status"])
        if premium_payable_col:
            premium_payables.append(result.get("premium_payable"))
        if progress:
            progress.maybe(i)
    if progress:
        progress.final()

    out = leader_df.copy()
    # Some real leader files already ship a blank "IGI Policy Number"
    # column as a placeholder for exactly this - fill that one in rather
    # than adding a second, duplicate-named column next to it.
    out["IGI Policy Number"] = igi_numbers
    out["Matched Leader Ref (Register)"] = matched_refs
    out["Match Score"] = scores
    out["Match Status"] = statuses
    if premium_payable_col:
        out["Premium Payable"] = premium_payables
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
