"""
Commission Income (Finance module).

WHAT THIS DOES: computes Reinsurance (RI) Commission Income and the
closing RI Unearned Commission for a selected 12-month period, from:

  * the Premium Registers - SIX departments (Fire, Marine, Motor,
    Miscellaneous, Engineering, Bond), TWO registers each (the previous
    year's register and the current year's register), and
  * the FAC Outward Register - one for the previous year, one for the
    current year.

TREATY COMMISSION (Premium Registers)
  Required columns per register: Policy, Treaty Description, Comm
  (commencement) Date, Treaty Premium.
  1. A "Period" is assigned to every row from its Comm Date against the
     selected 12-month period: before the period start -> "Previous",
     after the period end -> "Future Policies", otherwise the month
     name (same bucketing as UPR Automation - upr_processor.
     assign_buckets, reused directly).
  2. The system sorts out every distinct Treaty Description per
     department and the user enters a commission % for each one (one %
     for the current-year register, one for the previous-year register -
     real treaty rates do change between years). Treaty Commission =
     Treaty Premium x that row's Treaty Description %.
  3. The Treaty Premium of every description is totalled per department
     so the user can see exactly what each % is being applied to.

FAC COMMISSION (FAC Outward Register)
  The register already carries the FAC commission amount per line.
  Required columns: Policy#, Department, Comm Date, Acceptance Date,
  Comm amt. Marine lines carry no commencement date in the FAC register
  - for Marine (and ONLY Marine) the Acceptance Date is used as the
  line's date instead. If a Marine line has no Acceptance Date but does
  carry a Comm Date (the Dec 2025 FAC register is shaped this way), the
  Comm Date is used rather than dropping the line, and this is reported.

FINAL WORKING
  Per department and per period (Previous / Jan..Dec / Future Policies),
  Treaty Commission + FAC Commission from BOTH years are added together
  and then amortized with the same per-department method as UPR
  Automation: 1/24 (most departments), 1/6 (Marine, Travel), zero
  (BIMA) - upr_processor.build_unearned_table, reused directly - giving
  the closing RI Unearned Commission.

A treaty description left without a % defaults to 0% (the same fallback
as Deferred Wakala Fee's department %) and is reported as a warning so
a forgotten description is never silent.
"""
import os
import re

import pandas as pd

from premium_register_processor import _read_workbook_any_engine
from claims_register_processor import _read_full_table_with_continuations
import upr_processor as up

# The six departments that carry a Premium Register, in working order.
DEPARTMENTS = ["Fire", "Marine", "Motor", "Miscellaneous", "Engineering", "Bond"]
SLOT_PREVIOUS = "previous"
SLOT_CURRENT = "current"
SLOTS = [SLOT_PREVIOUS, SLOT_CURRENT]

BLANK_DESC_LABEL = "(No Treaty Description)"

# Re-exported so the app/generator only need to import this module.
TABLE_COLUMNS = up.TABLE_COLUMNS
MONTH_NAMES = up.MONTH_NAMES
PREVIOUS_LABEL = up.PREVIOUS_LABEL
FUTURE_LABEL = up.FUTURE_LABEL
TOTAL_LABEL = up.TOTAL_LABEL
get_available_period_options = up.get_available_period_options
compute_period_bounds = up.compute_period_bounds
row_total = up.row_total
column_totals = up.column_totals
grand_total = up.grand_total


class CommissionIncomeError(Exception):
    """A problem that should stop the run with a clear, actionable
    message (missing required column, unreadable file, period not
    covered by the uploaded registers)."""
    pass


# ---------------------------------------------------------------------------
# Header matching - same alias pattern as every other GIMS loader.
# ---------------------------------------------------------------------------
REGISTER_ALIASES = {
    "Policy": ["policy", "policy number", "policy no", "policy no.", "policy #", "policy#"],
    "Insured": ["insured"],
    "Treaty Description": ["treaty desc", "treaty description", "treaty descriptions", "treaty desc."],
    "Comm Date": ["comm date", "commencement date", "comm. date", "comm date."],
    "Treaty Premium": [
        "treaty premium", "treaty premiumium", "treaty premium/contrib.",
        "treaty premium / contrib.", "treaty premium contrib.",
    ],
}
REGISTER_REQUIRED = ["Policy", "Treaty Description", "Comm Date", "Treaty Premium"]

FAC_ALIASES = {
    "Policy": ["policy#", "policy", "policy number", "policy no"],
    "Insured": ["insured"],
    "Closing": ["closing#", "closing #", "closing no", "closing"],
    "Department": ["department", "dept", "class"],
    "Comm Date": ["comm date", "commencement date", "comm. date"],
    "Acceptance Date": ["acceptance date", "accept date", "acceptance dt"],
    "FAC Premium": ["fac premium", "fac prem", "fac premium "],
    "FAC Commission": ["comm amt", "comm. amt", "commission amount", "commission amt", "comm amount"],
}
FAC_REQUIRED = ["Policy", "Department", "Comm Date", "Acceptance Date", "FAC Commission"]


def _normalize_header(value):
    text = str(value).replace("\n", " ")
    text = re.sub(r"[()]", " ", text)
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def _alias_lookup(aliases):
    lookup = {}
    for canonical, names in aliases.items():
        for a in names:
            lookup[_normalize_header(a)] = canonical
    return lookup


_REGISTER_LOOKUP = _alias_lookup(REGISTER_ALIASES)
_FAC_LOOKUP = _alias_lookup(FAC_ALIASES)


def _find_header_row(preview_df, lookup, max_scan=25, min_score=3):
    best_row, best_score = None, 0
    for i in range(min(max_scan, len(preview_df))):
        score = 0
        seen = set()
        for v in preview_df.iloc[i].tolist():
            if pd.isna(v):
                continue
            canon = lookup.get(_normalize_header(v))
            if canon and canon not in seen:
                seen.add(canon)
                score += 1
        if score > best_score:
            best_score, best_row = score, i
    if best_score < min_score:
        return None, best_score
    return best_row, best_score


def _pick_best_sheet(xls, lookup, label, max_scan=25, min_score=3):
    best = None
    for sn in xls.sheet_names:
        preview = pd.read_excel(xls, sheet_name=sn, header=None, nrows=max_scan)
        header_row, score = _find_header_row(preview, lookup, max_scan=max_scan, min_score=min_score)
        if header_row is None:
            continue
        if best is None or score > best[0]:
            best = (score, sn, header_row)
    if best is None:
        raise CommissionIncomeError(
            f"Could not find the header row in any sheet of this {label} (no sheet had a row matching "
            f"at least {min_score} of the expected columns)."
        )
    return best[1], best[2]


def _read_table(filepath, lookup, label, report=None):
    """Open one workbook, find its real header row and read the whole
    table (merging legacy .xls continuation sheets). Returns (data,
    sheet_note)."""
    if not os.path.exists(filepath):
        raise CommissionIncomeError(f"File not found: {filepath}")
    xls, _resolved = _read_workbook_any_engine(filepath)
    sheet_name, header_row = _pick_best_sheet(xls, lookup, label)
    if report:
        report(f"Reading '{os.path.basename(filepath)}' - sheet '{sheet_name}'...")
    data, note = _read_full_table_with_continuations(filepath, xls, sheet_name, header_row, lookup, report=report)
    return data, note


def _canonical_columns(data, lookup):
    """{canonical: [original column names in sheet order]} - a canonical
    field can legitimately match more than one column (e.g. a register
    with both a blank 'Comm Date' and a populated one)."""
    out = {}
    for col in data.columns:
        canon = lookup.get(_normalize_header(col))
        if canon:
            out.setdefault(canon, []).append(col)
    return out


def _coalesce(data, columns):
    """First non-blank value across the given columns, row by row."""
    result = data[columns[0]]
    for extra in columns[1:]:
        result = result.where(result.notna(), data[extra])
    return result


def _canonical_department(name):
    text = str(name).strip()
    for canon in up.DEPARTMENT_ORDER:
        if text.upper() == canon.upper():
            return canon
    return text


def normalize_desc_key(value):
    """Matching key for a Treaty Description: case/whitespace-insensitive.
    A blank description maps to the empty key."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = re.sub(r"\s+", " ", str(value)).strip()
    if text.lower() == "nan":
        return ""
    return text.upper()


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

def load_premium_register(filepath, department, report=None):
    """Load one department's Premium Register. Returns (df, meta).

    df columns: Department, Policy, Insured (may be blank), Treaty
    Description (original text, "" if blank), Desc Key (normalized),
    Comm Date (datetime), Treaty Premium (float; blank -> 0), Source File.
    A row is excluded only when its Comm Date is missing/unparseable (it
    can't be placed in a period); excluded rows are counted and their
    Treaty Premium reported in meta so nothing disappears silently.
    """
    fname = os.path.basename(filepath)
    data, note = _read_table(filepath, _REGISTER_LOOKUP, f"{department} Premium Register '{fname}'", report=report)
    cols = _canonical_columns(data, _REGISTER_LOOKUP)
    missing = [c for c in REGISTER_REQUIRED if c not in cols]
    if missing:
        raise CommissionIncomeError(
            f"'{fname}' ({department}) is missing required column(s) {missing}. Recognized headers were: "
            f"{list(data.columns)}. Add the real header text to REGISTER_ALIASES in commission_income_processor.py."
        )

    out = pd.DataFrame(index=data.index)
    out["Department"] = department
    out["Policy"] = _coalesce(data, cols["Policy"]).astype(str).str.strip()
    out["Insured"] = _coalesce(data, cols["Insured"]) if "Insured" in cols else ""
    desc = _coalesce(data, cols["Treaty Description"])
    out["Treaty Description"] = desc.map(lambda v: "" if normalize_desc_key(v) == "" else re.sub(r"\s+", " ", str(v)).strip())
    out["Desc Key"] = out["Treaty Description"].map(normalize_desc_key)
    out["Comm Date"] = pd.to_datetime(_coalesce(data, cols["Comm Date"]), errors="coerce")
    out["Treaty Premium"] = pd.to_numeric(_coalesce(data, cols["Treaty Premium"]), errors="coerce").fillna(0.0)
    out["Source File"] = fname

    valid = out["Comm Date"].notna()
    excluded = out.loc[~valid]
    meta = {
        "sheet": note,
        "rows_read": len(out),
        "excluded_rows": int((~valid).sum()),
        "excluded_premium": float(excluded["Treaty Premium"].sum()),
    }
    return out.loc[valid].reset_index(drop=True), meta


def load_fac_register(filepath, report=None):
    """Load a FAC Outward Register. Returns (df, meta).

    df columns: Policy, Insured, Closing, Department (canonical name),
    Comm Date (the line's own date - the Comm Date for every department
    EXCEPT Marine, which uses its Acceptance Date; see the module
    docstring), FAC Premium, FAC Commission, Date Source.
    """
    fname = os.path.basename(filepath)
    data, note = _read_table(filepath, _FAC_LOOKUP, f"FAC Outward Register '{fname}'", report=report)
    cols = _canonical_columns(data, _FAC_LOOKUP)
    missing = [c for c in FAC_REQUIRED if c not in cols]
    if missing:
        raise CommissionIncomeError(
            f"'{fname}' is missing required FAC column(s) {missing}. Recognized headers were: "
            f"{list(data.columns)}. Add the real header text to FAC_ALIASES in commission_income_processor.py."
        )

    out = pd.DataFrame(index=data.index)
    out["Policy"] = _coalesce(data, cols["Policy"]).astype(str).str.strip()
    out["Insured"] = _coalesce(data, cols["Insured"]) if "Insured" in cols else ""
    out["Closing"] = _coalesce(data, cols["Closing"]) if "Closing" in cols else ""
    dept_raw = _coalesce(data, cols["Department"])
    out["Department"] = dept_raw.map(lambda v: "" if pd.isna(v) else _canonical_department(v))
    comm = pd.to_datetime(_coalesce(data, cols["Comm Date"]), errors="coerce")
    accept = pd.to_datetime(_coalesce(data, cols["Acceptance Date"]), errors="coerce")
    out["FAC Premium"] = pd.to_numeric(_coalesce(data, cols["FAC Premium"]), errors="coerce").fillna(0.0) if "FAC Premium" in cols else 0.0
    out["FAC Commission"] = pd.to_numeric(_coalesce(data, cols["FAC Commission"]), errors="coerce").fillna(0.0)

    is_marine = out["Department"].str.upper() == "MARINE"
    # Marine: Acceptance Date. Everything else: Comm Date.
    date_value = comm.where(~is_marine, accept)
    date_source = pd.Series("Comm Date", index=out.index, dtype=object)
    date_source = date_source.mask(is_marine, "Acceptance Date")
    # Marine line with no Acceptance Date but a Comm Date - keep the line.
    marine_fallback = is_marine & accept.isna() & comm.notna()
    date_value = date_value.mask(marine_fallback, comm)
    date_source = date_source.mask(marine_fallback, "Comm Date (Marine fallback)")
    out["Comm Date"] = date_value
    out["Date Source"] = date_source
    out["Source File"] = fname

    has_dept = out["Department"] != ""
    # Blank spacer rows / subtotal rows carry no department - not real lines.
    real = out.loc[has_dept]
    valid = real["Comm Date"].notna()
    meta = {
        "sheet": note,
        "rows_read": int(has_dept.sum()),
        "ignored_blank_department_rows": int((~has_dept & (out["FAC Commission"] != 0)).sum()),
        "excluded_rows": int((~valid).sum()),
        "excluded_commission": float(real.loc[~valid, "FAC Commission"].sum()),
        "marine_fallback_rows": int(marine_fallback[has_dept].sum()),
    }
    return real.loc[valid].reset_index(drop=True), meta


def load_all(register_files, fac_files, report=None):
    """
    register_files: {department: {"previous": path, "current": path}} -
        any slot may be missing/None (that register is simply not part of
        the run).
    fac_files: {"previous": path, "current": path}, same rule.

    Returns a bundle dict:
      {
        "registers": {(department, slot): df},
        "fac": {slot: df},
        "meta": {"registers": {(dept, slot): meta}, "fac": {slot: meta}},
      }
    """
    registers, fac, reg_meta, fac_meta = {}, {}, {}, {}
    for dept in DEPARTMENTS:
        for slot in SLOTS:
            path = (register_files.get(dept) or {}).get(slot)
            if not path:
                continue
            if report:
                report(f"[{dept} - {slot} year] loading premium register...")
            df, meta = load_premium_register(path, dept, report=report)
            registers[(dept, slot)] = df
            reg_meta[(dept, slot)] = meta
            if report:
                report(f"[{dept} - {slot} year] {len(df):,} row(s) usable"
                       + (f", {meta['excluded_rows']:,} excluded (no Comm Date)" if meta["excluded_rows"] else "") + ".")
    for slot in SLOTS:
        path = (fac_files or {}).get(slot)
        if not path:
            continue
        if report:
            report(f"[FAC - {slot} year] loading FAC outward register...")
        df, meta = load_fac_register(path, report=report)
        fac[slot] = df
        fac_meta[slot] = meta
        if report:
            note = f", {meta['marine_fallback_rows']:,} Marine line(s) used Comm Date (no Acceptance Date)" if meta["marine_fallback_rows"] else ""
            report(f"[FAC - {slot} year] {len(df):,} line(s) usable{note}.")
    if not registers and not fac:
        raise CommissionIncomeError("No register files were provided.")
    return {"registers": registers, "fac": fac, "meta": {"registers": reg_meta, "fac": fac_meta}}


# ---------------------------------------------------------------------------
# Treaty description scan (feeds the % entry table in the app)
# ---------------------------------------------------------------------------

def scan_treaty_descriptions(bundle):
    """Every distinct Treaty Description per department, sorted by
    department (working order) then description, with its total Treaty
    Premium for each register year.

    Returns a list of dicts: {department, key, description,
    current_premium, previous_premium}. A blank description is only
    listed when it actually carries Treaty Premium (a blank description
    with zero premium needs no %).
    """
    rows = []
    for dept in DEPARTMENTS:
        agg = {}
        for slot in SLOTS:
            df = bundle["registers"].get((dept, slot))
            if df is None or df.empty:
                continue
            grouped = df.groupby("Desc Key")["Treaty Premium"].sum()
            first_text = df.drop_duplicates("Desc Key").set_index("Desc Key")["Treaty Description"]
            for key, prem in grouped.items():
                entry = agg.setdefault(key, {"description": first_text.get(key, key), "current": 0.0, "previous": 0.0})
                if not entry["description"] and first_text.get(key):
                    entry["description"] = first_text.get(key)
                entry[slot] += float(prem)
        for key in sorted(agg, key=lambda k: (k == "", k)):
            e = agg[key]
            if key == "" and abs(e["current"]) < 0.005 and abs(e["previous"]) < 0.005:
                continue
            rows.append({
                "department": dept,
                "key": key,
                "description": e["description"] or BLANK_DESC_LABEL,
                "current_premium": e["current"],
                "previous_premium": e["previous"],
            })
    return rows


def available_period_options(bundle):
    """Period dropdown options from every register's Comm Date."""
    frames = [df[["Comm Date"]] for df in bundle["registers"].values() if len(df)]
    if not frames:
        return []
    return up.get_available_period_options(pd.concat(frames, ignore_index=True))


# ---------------------------------------------------------------------------
# Calculation
# ---------------------------------------------------------------------------

def _pct_lookup(rates):
    """{(DEPT_UPPER, desc_key): {"current": pct, "previous": pct or None}}"""
    out = {}
    for dept, per_desc in (rates or {}).items():
        for key, slot_pcts in per_desc.items():
            out[(str(dept).strip().upper(), normalize_desc_key(key))] = slot_pcts
    return out


def _rate_for(lookup, dept, key, slot):
    """Percent for this dept/description/slot. A previous-year % left
    blank (None) falls back to the current-year %; nothing entered at all
    is 0%. Returns (pct, was_entered)."""
    entry = lookup.get((dept.strip().upper(), key))
    if not entry:
        return 0.0, False
    cur = entry.get(SLOT_CURRENT)
    if slot == SLOT_CURRENT:
        return (float(cur), True) if cur is not None else (0.0, False)
    prev = entry.get(SLOT_PREVIOUS)
    if prev is not None:
        return float(prev), True
    if cur is not None:
        return float(cur), True
    return 0.0, False


def _grid(bucketed, value_col, departments):
    grouped = bucketed.groupby(["Department", "Bucket"])[value_col].sum()
    table = {}
    for dept in departments:
        table[dept] = {col: float(grouped.get((dept, col), 0.0)) for col in up.TABLE_COLUMNS}
    return table


def _add_tables(*tables, departments):
    out = {}
    for dept in departments:
        out[dept] = {col: sum(t[dept][col] for t in tables) for col in up.TABLE_COLUMNS}
    return out


def run_commission_income_calculation(bundle, rates, period_year, period_month, report=None):
    """
    bundle: from load_all().
    rates: {department: {description_key_or_text: {"current": pct,
        "previous": pct_or_None}}} - PERCENT numbers (26 = 26%).

    Returns a result dict (see generator for how it is written out):
      period_start, period_end, departments, methods, year_labels,
      detail (per register: DataFrame with Period + commission),
      fac_detail ({slot: DataFrame with Period}),
      treaty_summary (list of per-description rows),
      treaty_table/fac_table ({slot: {dept: {column: amount}}}),
      total_table, unearned_table, closing_unearned,
      income ({slot: {dept: {"treaty", "fac", "total"}}}),
      warnings, rates (echo, normalized).
    """
    period_start, period_end = up.compute_period_bounds(period_year, period_month)
    if report:
        report(f"Selected period: {period_start:%d-%b-%Y} to {period_end:%d-%b-%Y}.")

    # Coverage check over every register's Comm Date (same rule as UPR).
    date_frames = [df[["Comm Date"]] for df in bundle["registers"].values() if len(df)]
    if not date_frames:
        raise CommissionIncomeError("No premium register rows were loaded.")
    try:
        up.validate_coverage(pd.concat(date_frames, ignore_index=True), period_start, period_end)
    except up.UPRError as e:
        raise CommissionIncomeError(str(e))

    lookup = _pct_lookup(rates)
    warnings = []
    unentered = {}

    detail = {}
    treaty_tables = {s: {} for s in SLOTS}
    fac_tables = {s: {} for s in SLOTS}
    summary_rows = {}   # (dept, key) -> dict

    # ---- Treaty commission per register ----
    bucketed_by_slot = {s: [] for s in SLOTS}
    for (dept, slot), df in bundle["registers"].items():
        work = df.copy()
        keys = work["Desc Key"].unique()
        pct_by_key = {}
        for key in keys:
            pct, entered = _rate_for(lookup, dept, key, slot)
            pct_by_key[key] = pct
            has_premium = abs(float(work.loc[work["Desc Key"] == key, "Treaty Premium"].sum())) > 0.005
            if not entered and has_premium:
                unentered.setdefault((dept, key), set()).add(slot)
        work["Rate %"] = work["Desc Key"].map(pct_by_key)
        work["Treaty Commission"] = work["Treaty Premium"] * work["Rate %"] / 100.0
        work = up.assign_buckets(work, period_start, period_end)
        work = work.rename(columns={"Bucket": "Period"})
        detail[(dept, slot)] = work
        bucketed_by_slot[slot].append(work.rename(columns={"Period": "Bucket"}))

        for key, g in work.groupby("Desc Key"):
            row = summary_rows.setdefault((dept, key), {
                "department": dept, "key": key,
                "description": g["Treaty Description"].iloc[0] or BLANK_DESC_LABEL,
                "current_premium": 0.0, "previous_premium": 0.0,
                "current_pct": None, "previous_pct": None,
                "current_commission": 0.0, "previous_commission": 0.0,
            })
            row[f"{slot}_premium"] += float(g["Treaty Premium"].sum())
            row[f"{slot}_commission"] += float(g["Treaty Commission"].sum())
            row[f"{slot}_pct"] = float(g["Rate %"].iloc[0])

    # ---- FAC commission per register ----
    fac_detail = {}
    fac_bucketed = {}
    for slot, df in bundle["fac"].items():
        work = up.assign_buckets(df.copy(), period_start, period_end)
        fac_bucketed[slot] = work
        fac_detail[slot] = work.rename(columns={"Bucket": "Period"})

    # ---- Departments present anywhere ----
    dept_frames = [pd.DataFrame({"Department": [d]}) for (d, _s) in bundle["registers"].keys()]
    dept_frames += [df[["Department"]] for df in bundle["fac"].values() if len(df)]
    present_departments, dept_warnings = up.resolve_department_list(pd.concat(dept_frames, ignore_index=True))
    warnings.extend(dept_warnings)
    # Every standard department is always shown (Health, Travel and BIMA
    # carry no Premium Register here and simply show zero, as in the
    # reference workbook) followed by any non-standard department that
    # appeared in the data.
    departments = list(up.DEPARTMENT_ORDER) + [d for d in present_departments if d not in up.DEPARTMENT_ORDER]

    # ---- Grids ----
    for slot in SLOTS:
        frames = bucketed_by_slot[slot]
        if frames:
            treaty_tables[slot] = _grid(pd.concat(frames, ignore_index=True), "Treaty Commission", departments)
        else:
            treaty_tables[slot] = {d: {c: 0.0 for c in up.TABLE_COLUMNS} for d in departments}
        if slot in fac_bucketed and len(fac_bucketed[slot]):
            fac_tables[slot] = _grid(fac_bucketed[slot], "FAC Commission", departments)
        else:
            fac_tables[slot] = {d: {c: 0.0 for c in up.TABLE_COLUMNS} for d in departments}

    total_table = _add_tables(treaty_tables[SLOT_PREVIOUS], treaty_tables[SLOT_CURRENT],
                              fac_tables[SLOT_PREVIOUS], fac_tables[SLOT_CURRENT], departments=departments)
    unearned_table, methods = up.build_unearned_table(total_table, departments, period_start)
    closing_unearned = up.grand_total(unearned_table, departments)

    # ---- Income (whole register, every bucket) per slot ----
    income = {}
    for slot in SLOTS:
        income[slot] = {}
        for dept in departments:
            treaty = row_total(treaty_tables[slot][dept])
            fac = row_total(fac_tables[slot][dept])
            income[slot][dept] = {"treaty": treaty, "fac": fac, "total": treaty + fac}

    # ---- Warnings ----
    for (dept, key), slots in sorted(unentered.items()):
        label = key and next((r["description"] for (d, k), r in summary_rows.items() if d == dept and k == key), key) or BLANK_DESC_LABEL
        which = " and ".join(f"{s}-year" for s in SLOTS if s in slots)
        warnings.append(f"No commission % entered for {dept} - '{label}' ({which}); treated as 0%.")
    for (dept, slot), meta in bundle["meta"]["registers"].items():
        if meta["excluded_rows"]:
            warnings.append(f"{dept} ({slot} year): {meta['excluded_rows']:,} register row(s) with no Comm Date were excluded "
                            f"(Treaty Premium {meta['excluded_premium']:,.2f}).")
    for slot, meta in bundle["meta"]["fac"].items():
        if meta["excluded_rows"]:
            warnings.append(f"FAC ({slot} year): {meta['excluded_rows']:,} line(s) with no usable date were excluded "
                            f"(commission {meta['excluded_commission']:,.2f}).")
        if meta["marine_fallback_rows"]:
            warnings.append(f"FAC ({slot} year): {meta['marine_fallback_rows']:,} Marine line(s) had no Acceptance Date - "
                            f"their Comm Date was used instead.")

    summary_list = sorted(
        summary_rows.values(),
        key=lambda r: (DEPARTMENTS.index(r["department"]) if r["department"] in DEPARTMENTS else 99, r["key"] == "", r["key"]))

    if report:
        report(f"Treaty commission - current year: {sum(income[SLOT_CURRENT][d]['treaty'] for d in departments):,.2f}; "
               f"previous year: {sum(income[SLOT_PREVIOUS][d]['treaty'] for d in departments):,.2f}.")
        report(f"FAC commission - current year: {sum(income[SLOT_CURRENT][d]['fac'] for d in departments):,.2f}; "
               f"previous year: {sum(income[SLOT_PREVIOUS][d]['fac'] for d in departments):,.2f}.")
        report(f"Closing RI Unearned Commission: {closing_unearned:,.2f}")

    return {
        "period_start": period_start,
        "period_end": period_end,
        "year_labels": {SLOT_CURRENT: period_end.year, SLOT_PREVIOUS: period_end.year - 1},
        "departments": departments,
        "methods": methods,
        "detail": detail,
        "fac_detail": fac_detail,
        "treaty_summary": summary_list,
        "treaty_table": treaty_tables,
        "fac_table": fac_tables,
        "total_table": total_table,
        "unearned_table": unearned_table,
        "closing_unearned": closing_unearned,
        "income": income,
        "warnings": warnings,
        "rates": {k: dict(v) for k, v in (rates or {}).items()},
    }
