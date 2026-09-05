"""
Withholding Tax (Income Tax) - General Ledger enrichment.

WHAT THIS DOES, in order:
  1. Load the General Ledger (one row per WHT voucher/transaction). The
     only column this module strictly needs is the payee's own name
     ("Payee Name") - every other column is preserved untouched and
     carried straight through to the output.
  2. Match each GL row's Payee Name against the NTN Locator file (a
     curated Taxpayer Name -> NTN/CNIC lookup) and pull in that payee's
     NTN and CNIC.
  3. Using the NTN/CNIC just resolved, match against the Status file
     (NTN or CNIC -> Status/Address/Location) - NTN is tried first,
     falling back to CNIC only if the NTN itself isn't found there
     (real data confirmed: some Status rows only carry a CNIC, no NTN,
     and vice versa).
  4. Using the same resolved NTN/CNIC, check FBR's own Active Taxpayer
     List (ATL) export - a large multi-sheet workbook (a real one seen
     was ~8.9 million rows across 23 sheets) - and add ONE Filer/
     Non-Filer column PER SHEET, since each sheet is its own list.

NEVER GUESS, EVER: this feeds withholding tax rates directly, so a wrong
NTN/CNIC has real compliance and financial consequences - far more so
than a wrong policy number elsewhere in GIMS. Payee-name matching
(step 2) is the only fuzzy-adjacent step (a name has to be normalized to
even have a chance of matching), and even there matching is EXACT after
normalization, never approximate/fuzzy. Any payee whose name isn't found
at all, or whose name maps to more than one conflicting NTN/CNIC in the
Locator (ambiguous), is left for manual review rather than guessed at -
see resolve_final_ntn_cnic's own docstring for the exact fallback rule
(the source GL file's own, already-filled-in NTL/CNIC is used ONLY when
the Locator step couldn't confidently resolve one itself).

PERFORMANCE: the ATL file is the one genuinely large input here. Plain
openpyxl (even in its own read_only streaming mode) is far too slow at
that scale - confirmed directly: it took 98 seconds just to reach the
5th row of a single ~400,000-row sheet. python-calamine (a Rust-based
reader, already used elsewhere in GIMS via perf_utils.fast_read_excel_rows)
reads that same sheet in about a second. Since all this module actually
needs from a huge ATL sheet is a SET of normalized NTN/CNIC digit-strings
(not a full table), load_atl_sets() reads each sheet as plain Python rows
via calamine directly and builds a set from them - never materializing a
pandas DataFrame for a sheet that can run into the hundreds of thousands
of rows.
"""
import os
import re

import pandas as pd

from premium_register_processor import (
    _read_workbook_any_engine, _build_alias_lookup, _map_columns_to_canonical,
    find_header_row, _pick_best_sheet, _dedupe_columns,
)

try:
    from python_calamine import CalamineWorkbook
    _CALAMINE_AVAILABLE = True
except ImportError:  # python-calamine not installed - the ATL loader raises a clear, actionable error (see load_atl_sets)
    _CALAMINE_AVAILABLE = False

try:
    from perf_utils import fast_read_excel_rows
    _PERF_AVAILABLE = True
except ImportError:  # perf_utils.py not yet added to this checkout, degrade gracefully
    _PERF_AVAILABLE = False


class WithholdingTaxError(Exception):
    """Raised for a problem that should stop the run with a clear,
    actionable message - a missing required column, a file that can't be
    read, or (for the ATL file) python-calamine not being installed."""
    pass


# ---------------------------------------------------------------------------
# Normalization helpers - used everywhere a name or an NTN/CNIC needs to
# be compared across two files that were never guaranteed to format it
# identically.
# ---------------------------------------------------------------------------

def _is_blank(value):
    """
    True for None, NaN, or an empty/whitespace-only string.

    WHY THIS EXISTS: pandas silently converts a Python None into a float
    NaN the moment it's stored in a DataFrame column (confirmed directly:
    pd.DataFrame({"a": [None, "x"]})["a"].tolist() comes back as
    [nan, 'x'], not [None, 'x']) - and NaN is TRUTHY in plain Python
    (bool(float('nan')) is True), so a naive `if not value` check on a
    value read back out of a DataFrame column silently treats "genuinely
    missing" as "present". That's exactly backwards for this module: an
    ATL/Status match has to tell "we checked and this NTN/CNIC really
    isn't on the list" (Non-Filer) apart from "there was never an NTN/
    CNIC to check in the first place" (Not Determined) - conflating the
    two would misrepresent a data gap as a real determination on a
    figure that directly drives the withholding tax rate applied. Every
    truthiness check on an NTN/CNIC value pulled back out of a DataFrame
    column goes through this helper instead of a bare `if value`/`if not
    value`.
    """
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(value, str) and not value.strip():
        return True
    return False


_DIGITS_RE = re.compile(r"\d+")


def digits_only(value):
    """Every digit in value, concatenated, everything else (dashes,
    letters like the 'I' individual-registration prefix FBR's own ATL
    export uses, spaces) stripped - e.g. '2494918-3' -> '24949183',
    'I218749-4' -> '2187494', '42301-0984531-5' -> '423010984531 5'... no,
    concatenated: '4230109845315'. This is deliberately simple (not
    "drop the check digit" or any other NTN/CNIC-format-specific parsing)
    because it only has to be applied CONSISTENTLY on both sides of a
    comparison to work - the NTN Locator, Status, and ATL files all
    format the same real NTN/CNIC differently (with/without dashes, with/
    without a leading letter), but a digits-only reduction of the same
    underlying number always produces the same string regardless of
    which file it came from. Returns None for a blank/unparseable value,
    never an empty string, so callers can use a plain truthiness check.
    """
    if _is_blank(value):
        return None
    s = "".join(_DIGITS_RE.findall(str(value)))
    return s or None


_NAME_PUNCT_RE = re.compile(r"[.,]")
_NAME_WS_RE = re.compile(r"\s+")


def normalize_name(value):
    """Uppercase, periods/commas stripped, whitespace collapsed - e.g.
    'A. R. Associates (Private) Limited' and 'A R ASSOCIATES (PRIVATE)
    LIMITED' both normalize to 'A R ASSOCIATES (PRIVATE) LIMITED'. Still
    an EXACT match on the result - this only closes the gap on
    formatting differences (case, punctuation, extra spaces), never on a
    genuinely different name."""
    if _is_blank(value):
        return ""
    text = str(value).strip().upper()
    text = _NAME_PUNCT_RE.sub("", text)
    text = _NAME_WS_RE.sub(" ", text).strip()
    return text


# ---------------------------------------------------------------------------
# Column alias matching (same pattern as every other GIMS loader)
# ---------------------------------------------------------------------------

GL_ALIASES = {
    "Payee Name": ["payee name", "payee", "vendor name", "supplier name", "party name"],
    "NTN": ["ntn", "n t n", "ntn no", "ntn number"],
    "CNIC": ["cnic", "cnic no", "cnic number", "c n i c"],
}

NTN_LOCATOR_ALIASES = {
    "Taxpayer Name": ["taxpayer name", "tax payer name", "payee name"],
    "NTN": ["ntn"],
    "CNIC": ["cnic"],
}

STATUS_ALIASES = {
    "NTN": ["taxpayer_ntn", "taxpayer ntn", "ntn"],
    "CNIC": ["taxpayer_cnic", "taxpayer cnic", "cnic"],
    "Name": ["taxpayer_name", "taxpayer name"],
    "Status": ["taxpayer_status", "taxpayer status", "status"],
    "Address": ["address"],
    "Location": ["location"],
}


def _load_generic(filepath, aliases_dict, sheet_name, report, required_any=None, required_all=None):
    """
    Shared loader body for the GL / NTN Locator / Status files - all three
    are small-to-medium Excel files with a header row that may sit a few
    banner rows down, matched via an alias list, exactly like every other
    GIMS loader (see premium_register_processor.py's own helpers, reused
    here directly rather than re-implemented).

    required_any: canonical names where AT LEAST ONE must be found (e.g.
        NTN or CNIC - a file can reasonably have only one of the two).
    required_all: canonical names that MUST all be found.
    Raises WithholdingTaxError with a clear, actionable message if a
    requirement isn't met.

    Returns (df, inv) where inv is {canonical_name: original_column_name}
    for every alias actually matched (first occurrence wins if a header
    is duplicated in the source file).
    """
    alias_lookup = _build_alias_lookup(aliases_dict)
    xls, resolved_filepath = _read_workbook_any_engine(filepath)

    if sheet_name is None:
        sheet_name, header_row = _pick_best_sheet(xls, alias_lookup)
    else:
        preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=25)
        header_row = find_header_row(preview, alias_lookup)

    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".xlsx" and _PERF_AVAILABLE:
        rows, _n = fast_read_excel_rows(resolved_filepath, sheet_name, report=report)
        columns = _dedupe_columns(list(rows[header_row]))
        data = pd.DataFrame(rows[header_row + 1:], columns=columns)
    else:
        raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        columns = _dedupe_columns(raw.iloc[header_row].tolist())
        data = raw.iloc[header_row + 1:].reset_index(drop=True)
        data.columns = columns

    data = data.dropna(how="all").reset_index(drop=True)
    data.attrs["source_sheet"] = sheet_name

    colmap = _map_columns_to_canonical(data.columns, alias_lookup)
    inv = {}
    for orig, canon in colmap.items():
        inv.setdefault(canon, orig)

    for canon in (required_all or []):
        if canon not in inv:
            raise WithholdingTaxError(
                f"Could not find a '{canon}' column in '{os.path.basename(filepath)}'. "
                f"Recognized headers were: {list(data.columns)}."
            )
    if required_any and not any(c in inv for c in required_any):
        raise WithholdingTaxError(
            f"Could not find any of {required_any} in '{os.path.basename(filepath)}'. "
            f"Recognized headers were: {list(data.columns)}."
        )
    return data, inv


def load_general_ledger(filepath, sheet_name=None, report=None):
    """Returns (gl_df, colmap). Payee Name is required; NTN/CNIC are
    optional (used only as a manual-entry fallback - see
    resolve_final_ntn_cnic)."""
    return _load_generic(filepath, GL_ALIASES, sheet_name, report, required_all=["Payee Name"])


def load_ntn_locator(filepath, sheet_name=None, report=None):
    """Returns (locator_df, colmap). Taxpayer Name is required; at least
    one of NTN/CNIC must be present."""
    return _load_generic(filepath, NTN_LOCATOR_ALIASES, sheet_name, report,
                          required_all=["Taxpayer Name"], required_any=["NTN", "CNIC"])


def load_status_file(filepath, sheet_name=None, report=None):
    """Returns (status_df, colmap). At least one of NTN/CNIC must be
    present (real data confirmed both NTN-only and CNIC-only rows)."""
    return _load_generic(filepath, STATUS_ALIASES, sheet_name, report, required_any=["NTN", "CNIC"])


# ---------------------------------------------------------------------------
# NTN Locator lookup
# ---------------------------------------------------------------------------

def build_ntn_locator_lookup(df, colmap):
    """
    {normalized_taxpayer_name: {"ntn": digits_or_None, "cnic": digits_or_None, "ambiguous": bool}}

    "ambiguous" is set True the moment the SAME normalized name is found
    on more than one Locator row with a DIFFERENT NTN/CNIC combination -
    matching that name automatically would risk silently assigning one
    real taxpayer's withholding tax to a different taxpayer's NTN, which
    is exactly the kind of wrong-but-confident-looking match this module
    must never produce (see the module docstring). An ambiguous name is
    therefore deliberately left unresolved here; resolve_final_ntn_cnic()
    is what turns that into a manual-review flag on the GL output.
    """
    name_col = colmap["Taxpayer Name"]
    ntn_col = colmap.get("NTN")
    cnic_col = colmap.get("CNIC")

    lookup = {}
    for rec in df.to_dict("records"):
        name = normalize_name(rec.get(name_col))
        if not name:
            continue
        ntn = digits_only(rec.get(ntn_col)) if ntn_col else None
        cnic = digits_only(rec.get(cnic_col)) if cnic_col else None
        if not ntn and not cnic:
            continue

        existing = lookup.get(name)
        if existing is None:
            lookup[name] = {"ntn": ntn, "cnic": cnic, "ambiguous": False}
        elif not existing["ambiguous"] and (existing["ntn"], existing["cnic"]) != (ntn, cnic):
            existing["ambiguous"] = True
    return lookup


# ---------------------------------------------------------------------------
# Status lookup
# ---------------------------------------------------------------------------

def build_status_lookup(df, colmap):
    """Returns (status_by_ntn, status_by_cnic), both
    {digits: {"status", "address", "location"}}. The first row seen for a
    given NTN/CNIC wins if the Status file happens to carry more than one
    row for the same taxpayer."""
    ntn_col = colmap.get("NTN")
    cnic_col = colmap.get("CNIC")
    status_col = colmap.get("Status")
    address_col = colmap.get("Address")
    location_col = colmap.get("Location")

    by_ntn, by_cnic = {}, {}
    for rec in df.to_dict("records"):
        info = {
            "status": rec.get(status_col) if status_col else None,
            "address": rec.get(address_col) if address_col else None,
            "location": rec.get(location_col) if location_col else None,
        }
        ntn = digits_only(rec.get(ntn_col)) if ntn_col else None
        cnic = digits_only(rec.get(cnic_col)) if cnic_col else None
        if ntn and ntn not in by_ntn:
            by_ntn[ntn] = info
        if cnic and cnic not in by_cnic:
            by_cnic[cnic] = info
    return by_ntn, by_cnic


# ---------------------------------------------------------------------------
# ATL (Active Taxpayer List) loading - the one genuinely large input.
# See the module docstring's PERFORMANCE section for why this bypasses
# pandas/openpyxl entirely and reads via python-calamine directly.
# ---------------------------------------------------------------------------

def _find_atl_header_row(rows_preview):
    """Scan a small preview (list of raw row lists) for whichever row has
    a cell that reads exactly 'NTN' (case-insensitive) - FBR's own ATL
    export header. Returns (header_row_index, ntn_column_index), or
    (None, None) if no such row is found in the preview (the caller skips
    that sheet with a warning rather than guessing a column)."""
    for i, row in enumerate(rows_preview):
        for j, v in enumerate(row):
            if v is not None and str(v).strip().lower() == "ntn":
                return i, j
    return None, None


def load_atl_sets(filepath, report=None):
    """
    Reads every sheet of the FBR Active Taxpayer List export and returns,
    per sheet, a plain Python set of every distinct NTN/CNIC found on it
    (digits-only, see digits_only()) - which is all this module ever
    needs from the ATL file: "is this taxpayer's NTN/CNIC present on this
    sheet, yes or no".

    Returns (atl_sets, sheet_order, meta):
      atl_sets - {sheet_name: frozenset(digit_strings)}
      sheet_order - sheet names in workbook order; this becomes the
          column order of the Filer/Non-Filer columns added to the GL.
      meta - {"total_rows": int, "rows_by_sheet": {sheet_name: int}}

    Raises WithholdingTaxError if python-calamine isn't installed - the
    real FBR export runs into the millions of rows (confirmed: ~8.9
    million across 23 sheets in a real 180MB file), and openpyxl's own
    reader (even in read_only streaming mode) is far too slow at that
    scale to be usable here - see the module docstring.
    """
    if not _CALAMINE_AVAILABLE:
        raise WithholdingTaxError(
            "python-calamine is not installed. The ATL (Active Taxpayer List) file is typically "
            "millions of rows, and reading it without python-calamine would take an impractically "
            "long time. Install it with: pip install python-calamine"
        )
    if not os.path.exists(filepath):
        raise WithholdingTaxError(f"File not found: {filepath}")

    wb = CalamineWorkbook.from_path(filepath)
    atl_sets = {}
    rows_by_sheet = {}
    total_rows = 0

    for sheet_name in wb.sheet_names:
        if report:
            report(f"  Reading ATL sheet '{sheet_name}'...")
        sheet = wb.get_sheet_by_name(sheet_name)
        data = sheet.to_python()

        if not data:
            atl_sets[sheet_name] = frozenset()
            rows_by_sheet[sheet_name] = 0
            continue

        header_row_idx, ntn_col_idx = _find_atl_header_row(data[:10])
        if header_row_idx is None:
            atl_sets[sheet_name] = frozenset()
            rows_by_sheet[sheet_name] = 0
            if report:
                report(f"    WARNING: no 'NTN' column header found on '{sheet_name}' - this sheet "
                       f"will be skipped (every row treated as not-on-this-list).")
            continue

        digit_set = set()
        for row in data[header_row_idx + 1:]:
            if ntn_col_idx >= len(row):
                continue
            d = digits_only(row[ntn_col_idx])
            if d:
                digit_set.add(d)

        n_rows = len(data) - header_row_idx - 1
        atl_sets[sheet_name] = frozenset(digit_set)
        rows_by_sheet[sheet_name] = n_rows
        total_rows += n_rows
        if report:
            report(f"    {sheet_name}: {n_rows:,} row(s), {len(digit_set):,} distinct NTN/CNIC value(s).")

    meta = {"total_rows": total_rows, "rows_by_sheet": rows_by_sheet}
    return atl_sets, list(wb.sheet_names), meta


# ---------------------------------------------------------------------------
# Matching - applied to the GL, in order
# ---------------------------------------------------------------------------

def match_ntn_locator(gl_df, gl_colmap, locator_lookup):
    """
    Adds "NTN (Locator Match)", "CNIC (Locator Match)", and "Locator
    Match Status" (one of "Matched", "Not Found", "Ambiguous") to a COPY
    of gl_df, matched on normalize_name(Payee Name) against
    locator_lookup (see build_ntn_locator_lookup). Exact match only -
    see the module docstring for why this is never made fuzzy.
    """
    payee_col = gl_colmap["Payee Name"]
    out = gl_df.copy()

    ntns, cnics, statuses = [], [], []
    for payee in out[payee_col]:
        name = normalize_name(payee)
        entry = locator_lookup.get(name) if name else None
        if entry is None:
            ntns.append(None); cnics.append(None); statuses.append("Not Found")
        elif entry["ambiguous"]:
            ntns.append(None); cnics.append(None); statuses.append("Ambiguous")
        else:
            ntns.append(entry["ntn"]); cnics.append(entry["cnic"]); statuses.append("Matched")

    out["NTN (Locator Match)"] = ntns
    out["CNIC (Locator Match)"] = cnics
    out["Locator Match Status"] = statuses
    return out


def resolve_final_ntn_cnic(gl_df, gl_colmap):
    """
    Adds "NTN (Final)", "CNIC (Final)", and "Manual Review Reason" to a
    COPY of gl_df (which must already carry the Locator match columns -
    see match_ntn_locator). Priority order, per instruction:

      1. The NTN Locator's own match (by Payee Name), whenever it was
         unambiguous - this is the confident, independently-derived
         answer and always wins when available.
      2. Only when the Locator step didn't produce a confident answer
         (Payee Name not found, or ambiguous): whatever NTN/CNIC the
         source GL file's OWN row already carries, if a person has
         already filled one in by hand. This never overrides a
         confident Locator match - it only fills the gap when the
         Locator has nothing to offer.
      3. If neither of the above produces an NTN or a CNIC, the row is
         left blank and "Manual Review Reason" is set, so the generated
         workbook can highlight it rather than silently carrying a gap
         through Status/ATL matching (which would otherwise show
         "Non-Filer" for a taxpayer we never actually identified).

    "Manual Review Reason" is also set (independently of whether a
    fallback NTN/CNIC was found) whenever the Locator match was
    Ambiguous, since an ambiguous name is a real problem to flag for a
    person even if the GL's own pre-filled value happens to paper over it.
    """
    orig_ntn_col = gl_colmap.get("NTN")
    orig_cnic_col = gl_colmap.get("CNIC")

    final_ntns, final_cnics, reasons = [], [], []
    for rec in gl_df.to_dict("records"):
        locator_status = rec["Locator Match Status"]
        ntn = rec["NTN (Locator Match)"]
        cnic = rec["CNIC (Locator Match)"]
        ntn = None if _is_blank(ntn) else ntn
        cnic = None if _is_blank(cnic) else cnic
        reason = None

        if locator_status != "Matched":
            fallback_ntn = digits_only(rec.get(orig_ntn_col)) if orig_ntn_col else None
            fallback_cnic = digits_only(rec.get(orig_cnic_col)) if orig_cnic_col else None
            if fallback_ntn or fallback_cnic:
                ntn, cnic = fallback_ntn, fallback_cnic
            else:
                ntn, cnic = None, None

            if locator_status == "Ambiguous":
                reason = ("Payee name matches more than one different NTN/CNIC in the NTN Locator "
                          "- ambiguous, needs manual review")
            elif not ntn and not cnic:
                reason = ("Payee name not found in the NTN Locator, and no NTN/CNIC already on file "
                          "- needs manual review")
            # else: Locator said "Not Found" but the GL's own manually-entered
            # NTN/CNIC was usable as a fallback - no review needed.

        final_ntns.append(ntn)
        final_cnics.append(cnic)
        reasons.append(reason)

    out = gl_df.copy()
    out["NTN (Final)"] = final_ntns
    out["CNIC (Final)"] = final_cnics
    out["Manual Review Reason"] = reasons
    return out


def match_status(gl_df, status_by_ntn, status_by_cnic):
    """Adds "Status", "Address", "Location" to a COPY of gl_df, matched
    against NTN (Final) first, falling back to CNIC (Final) only if the
    NTN itself isn't found in the Status file (both are checked
    regardless of which one is populated, since real Status data has
    NTN-only AND CNIC-only rows)."""
    out = gl_df.copy()
    statuses, addresses, locations = [], [], []
    for ntn, cnic in zip(out["NTN (Final)"], out["CNIC (Final)"]):
        info = None
        if not _is_blank(ntn) and ntn in status_by_ntn:
            info = status_by_ntn[ntn]
        elif not _is_blank(cnic) and cnic in status_by_cnic:
            info = status_by_cnic[cnic]
        statuses.append(info["status"] if info else None)
        addresses.append(info["address"] if info else None)
        locations.append(info["location"] if info else None)
    out["Status"] = statuses
    out["Address"] = addresses
    out["Location"] = locations
    return out


def match_atl(gl_df, atl_sets, sheet_order):
    """
    Adds one column per ATL sheet (named after the sheet, e.g.
    "Part1".."Part23") to a COPY of gl_df, each holding "Filer",
    "Non-Filer", or "Not Determined" - matched against NTN (Final)
    first, falling back to CNIC (Final).

    "Not Determined" (rather than "Non-Filer") is used for a row with
    NEITHER an NTN nor a CNIC available at all - there is a real
    difference between "checked this list and genuinely didn't find
    them" (a confident Non-Filer, which directly sets a withholding tax
    rate) and "never had an NTN/CNIC to check in the first place" (a
    data gap). Collapsing the two into "Non-Filer" would misrepresent a
    gap as a determination on a figure with real financial consequences.
    """
    out = gl_df.copy()
    ntn_series = out["NTN (Final)"]
    cnic_series = out["CNIC (Final)"]

    for sheet_name in sheet_order:
        atl_set = atl_sets.get(sheet_name, frozenset())
        col = []
        for ntn, cnic in zip(ntn_series, cnic_series):
            ntn_ok = not _is_blank(ntn)
            cnic_ok = not _is_blank(cnic)
            if not ntn_ok and not cnic_ok:
                col.append("Not Determined")
                continue
            found = (ntn_ok and ntn in atl_set) or (cnic_ok and cnic in atl_set)
            col.append("Filer" if found else "Non-Filer")
        out[sheet_name] = col
    return out


OVERALL_FILER_COL = "Overall Filer Status"


def compute_overall_filer_status(gl_df, atl_sheet_order):
    """
    Adds a single "Overall Filer Status" column to a COPY of gl_df, so a
    person reviewing the workbook doesn't have to check every individual
    ATL sheet column on every row just to see whether that payee is a
    Filer anywhere at all.

      - "Filer" if that payee is a Filer on AT LEAST ONE ATL sheet. Being
        found on any one active list is what actually matters for the
        withholding rate, so one "Filer" among the per-sheet columns is
        enough to call the whole row a Filer, regardless of what the
        other sheets say.
      - "Non-Filer" if at least one sheet could genuinely be checked
        (there was an NTN/CNIC to check against) and came back
        Non-Filer, and no sheet came back Filer.
      - "Not Determined" if every single per-sheet column reads "Not
        Determined" (no NTN/CNIC was ever available to check against any
        list). This is a data gap, not a confident Non-Filer finding,
        and is kept as its own separate value rather than folded into
        Non-Filer, for the same reason match_atl keeps the two apart at
        the per-sheet level.

    Placed as a column immediately before the first per-sheet ATL column
    (rather than left at the end, after every sheet), so the summary sits
    right next to the detail it's summarizing and reads first.
    """
    out = gl_df.copy()
    if not atl_sheet_order:
        out[OVERALL_FILER_COL] = "Not Determined"
        return out

    sub = out[atl_sheet_order]
    is_filer = sub.eq("Filer").any(axis=1)
    is_nonfiler = sub.eq("Non-Filer").any(axis=1)
    overall = pd.Series("Not Determined", index=out.index)
    overall = overall.mask(is_nonfiler, "Non-Filer")
    overall = overall.mask(is_filer, "Filer")  # Filer wins even if another sheet says Non-Filer
    out[OVERALL_FILER_COL] = overall

    cols = list(out.columns)
    cols.remove(OVERALL_FILER_COL)
    insert_at = cols.index(atl_sheet_order[0])
    cols.insert(insert_at, OVERALL_FILER_COL)
    return out[cols]


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def run_withholding_tax_processing(gl_filepath, gl_sheet_name, ntn_locator_filepath, ntn_locator_sheet_name,
                                    status_filepath, status_sheet_name, atl_filepath, report=None):
    """
    Full pipeline: load all four inputs, then match the GL against each
    in order (Locator -> resolve final NTN/CNIC -> Status -> ATL).

    Each of gl_filepath / ntn_locator_filepath / status_filepath /
    atl_filepath may point at the SAME physical workbook (e.g. one
    combined file with a "General Ledger", "NTN Locator", and "Status"
    sheet) or at three/four entirely separate files - every loader
    auto-detects its own sheet within whatever file it's given by
    scoring against its own alias list, so either way works with no
    special-casing needed.

    Returns a result dict:
      {
        "merged_df": pd.DataFrame,          # the GL with every new column appended
        "gl_colmap": {...},
        "atl_sheet_order": [str, ...],       # column order of the Filer/Non-Filer columns
        "stats": {
            "num_rows", "num_matched_locator", "num_not_found_locator",
            "num_ambiguous_locator", "num_manual_review", "num_status_matched",
        },
      }
    Raises WithholdingTaxError with a clear, user-facing message on any
    problem reading or matching a file.
    """
    if report:
        report(f"Loading General Ledger: {os.path.basename(gl_filepath)}...")
    gl_df, gl_colmap = load_general_ledger(gl_filepath, sheet_name=gl_sheet_name, report=report)
    if report:
        report(f"  {len(gl_df):,} row(s) loaded. Payee Name column = '{gl_colmap['Payee Name']}'.")

    if report:
        report(f"\nLoading NTN Locator: {os.path.basename(ntn_locator_filepath)}...")
    locator_df, locator_colmap = load_ntn_locator(ntn_locator_filepath, sheet_name=ntn_locator_sheet_name, report=report)
    locator_lookup = build_ntn_locator_lookup(locator_df, locator_colmap)
    if report:
        report(f"  {len(locator_df):,} row(s) loaded -> {len(locator_lookup):,} distinct payee name(s) indexed.")

    if report:
        report(f"\nLoading Status file: {os.path.basename(status_filepath)}...")
    status_df, status_colmap = load_status_file(status_filepath, sheet_name=status_sheet_name, report=report)
    status_by_ntn, status_by_cnic = build_status_lookup(status_df, status_colmap)
    if report:
        report(f"  {len(status_df):,} row(s) loaded -> {len(status_by_ntn):,} indexed by NTN, "
               f"{len(status_by_cnic):,} by CNIC.")

    if report:
        report(f"\nLoading ATL file: {os.path.basename(atl_filepath)}... (large file - may take a moment)")
    atl_sets, atl_sheet_order, atl_meta = load_atl_sets(atl_filepath, report=report)
    if report:
        report(f"  {atl_meta['total_rows']:,} ATL row(s) loaded across {len(atl_sheet_order)} sheet(s).")

    if report:
        report(f"\nMatching {len(gl_df):,} GL row(s)...")
    merged = match_ntn_locator(gl_df, gl_colmap, locator_lookup)
    merged = resolve_final_ntn_cnic(merged, gl_colmap)
    merged = match_status(merged, status_by_ntn, status_by_cnic)
    merged = match_atl(merged, atl_sets, atl_sheet_order)
    merged = compute_overall_filer_status(merged, atl_sheet_order)

    num_rows = len(merged)
    stats = {
        "num_rows": num_rows,
        "num_matched_locator": int((merged["Locator Match Status"] == "Matched").sum()),
        "num_not_found_locator": int((merged["Locator Match Status"] == "Not Found").sum()),
        "num_ambiguous_locator": int((merged["Locator Match Status"] == "Ambiguous").sum()),
        "num_manual_review": int(merged["Manual Review Reason"].notna().sum()),
        "num_status_matched": int(merged["Status"].notna().sum()),
        "num_overall_filer": int((merged[OVERALL_FILER_COL] == "Filer").sum()),
        "num_overall_nonfiler": int((merged[OVERALL_FILER_COL] == "Non-Filer").sum()),
        "num_overall_not_determined": int((merged[OVERALL_FILER_COL] == "Not Determined").sum()),
    }

    if report:
        report(f"\nDone: {stats['num_matched_locator']:,}/{num_rows:,} matched to an NTN/CNIC via the "
               f"Locator ({stats['num_not_found_locator']:,} not found, {stats['num_ambiguous_locator']:,} "
               f"ambiguous). {stats['num_manual_review']:,} row(s) need manual review. "
               f"{stats['num_status_matched']:,}/{num_rows:,} matched to a Status/Address/Location.")
        report(f"Overall Filer Status: {stats['num_overall_filer']:,} Filer, "
               f"{stats['num_overall_nonfiler']:,} Non-Filer, "
               f"{stats['num_overall_not_determined']:,} Not Determined (no NTN/CNIC to check at all).")

    return {
        "merged_df": merged,
        "gl_colmap": gl_colmap,
        "atl_sheet_order": atl_sheet_order,
        "stats": stats,
    }


if __name__ == "__main__":
    import sys

    gl_fp = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/GL_-_Combined-26-07-2026.xlsx"
    locator_fp = sys.argv[2] if len(sys.argv) > 2 else gl_fp
    status_fp = sys.argv[3] if len(sys.argv) > 3 else gl_fp
    atl_fp = sys.argv[4] if len(sys.argv) > 4 else "/mnt/user-data/uploads/ATL_IT.xlsx"

    result = run_withholding_tax_processing(
        gl_fp, "WORKING WHT (2)", locator_fp, None, status_fp, None, atl_fp, report=print)
    df = result["merged_df"]
    print("\nColumns:", list(df.columns))
    print("\nFirst few rows of new columns:")
    new_cols = ["Payee Name" if "Payee Name" in df.columns else df.columns[0],
                "NTN (Final)", "CNIC (Final)", "Locator Match Status", "Manual Review Reason",
                "Status", "Address", "Location"] + result["atl_sheet_order"][:3]
    print(df[new_cols].head(10).to_string())
