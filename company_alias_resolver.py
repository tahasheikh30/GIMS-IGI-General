"""
Company merge logic for the Coinsurance modules (Claims Register and
Premium/Claims Email Automation), both of which group source rows by a
"company" name field (Coinsurer Name / Company) that real source data
records inconsistently under several different literal strings for the
same real company - e.g. "ATLAS_INSURANCE_FAISALABAD" and
"ATLAS_INSURANCE_CLAIMRECOVERY" are both Atlas Insurance, just typed
differently by whoever entered them.

REBUILT FROM SCRATCH (previous version removed entirely) around a single,
simple, fully deterministic rule - no suggestion files, no confidence
scores, no manual-review workbook, no ambiguity handling:

  1. PREFIX MATCH: take the text before the first underscore (or, if
     there's no underscore, the first space) in each bucket label - e.g.
     "Atlas" out of "Atlas_Insurance_Faisalabad". Two bucket labels with
     the same leading token (case-insensitive) are merged into one file.
  2. ABBREVIATIONS: a small fixed list of known short-forms used in real
     data are mapped to their real company name first, so e.g. "AIC ..."
     and "Adamjee ..." land in the same merged bucket even though their
     own leading tokens don't literally match each other.

This is deliberately simple - there is no external canonical-company file
to load, nothing to review, and merging is entirely a POST-PROCESSING
step applied to whatever buckets a run already produced (every module's
own initial grouping, by the raw/exact company name, is completely
unaffected) - so it's always safe to toggle off and regenerate the
original, un-merged files.

Also still handles parsing the real "Co-insurance Companies List" format
(company names + email addresses) for the Email Lookup File feature -
that is a completely separate concern from merging (it supplies
recipient email addresses, never influences which bucket a row lands in)
and is unchanged here.
"""
import os
import re
import numbers

import pandas as pd

# ---------------------------------------------------------------------------
# Fixed abbreviation list - the only "knowledge" this module has about
# specific companies. Add more here as they come up; keys are matched
# case-insensitively against a bucket label's own leading token (see
# _first_token / get_merge_key below), letters only.
# ---------------------------------------------------------------------------
ABBREVIATIONS = {
    "AIC": "Adamjee Insurance",
    "JGE": "Jubilee Insurance",
    "JGI": "Jubilee Insurance",  # common alternate spelling of the same abbreviation
    "RIC": "Reliance Insurance",
    "AGI": "Askari General Insurance",
    # NOTE: "UBL" is deliberately NOT mapped to "United Bank" here (it used
    # to be). UBL Insurers and United Insurance Company are two different
    # real companies - but mapping "UBL" -> "United Bank" made
    # get_merge_key() reduce BOTH of them down to the same "UNITED" key
    # (United Bank's own leading token), which silently merged "UBL ..."
    # labels into "United Insurance ..."'s bucket/file. UBL's own leading
    # token ("UBL") is already a perfectly good, distinct merge key on its
    # own, so no abbreviation entry is needed (or wanted) for it at all.
}


def _first_token(label):
    """The text before the first underscore, or (if there's no
    underscore) before the first space - e.g. "Atlas" from
    "Atlas_Insurance_Faisalabad", "Atlas" from "Atlas Insurance Ltd"."""
    text = str(label).strip()
    if "_" in text:
        return text.split("_", 1)[0].strip()
    if " " in text:
        return text.split(" ", 1)[0].strip()
    return text


def _letters_upper(text):
    return re.sub(r"[^A-Za-z]", "", str(text)).upper()


def get_merge_key(label):
    """
    The normalized grouping key for one bucket label: its own leading
    token (before the first underscore/space), letters-only and
    uppercased - e.g. "ADAMJEE" from "Adamjee Insurance Company Limited".

    If that leading token is a known abbreviation (see ABBREVIATIONS),
    the key is instead derived the exact same way from the real company
    name it stands for - e.g. "AIC" -> "Adamjee Insurance" -> its own
    leading token "Adamjee" -> "ADAMJEE" - so an abbreviation-prefixed
    label always reduces to the SAME key a plain-text label for that same
    company would, and the two merge together correctly.
    """
    token_letters = _letters_upper(_first_token(label))
    if token_letters in ABBREVIATIONS:
        canonical = ABBREVIATIONS[token_letters]
        return _letters_upper(_first_token(canonical))
    return token_letters or str(label).strip().upper()


# ---------------------------------------------------------------------------
# Canonical output filenames for a fixed list of known companies. Used by
# the Claims Register and Receivable generators to name a bucket's output
# file/workbook after the company using this exact, requested convention,
# instead of deriving the filename mechanically from whatever raw/merged
# bucket label happened to win (which can vary - "Atlas Insurance Co Ltd"
# vs "ATLAS_INSURANCE_FAISALABAD" vs "Atlas Insurance Company Limited" all
# refer to the same real company but would otherwise produce different
# filenames depending on which raw label happened to be picked as the
# merge representative).
#
# Matching is substring-based against the label's own compacted (letters
# only, uppercased, spaces/punctuation stripped) form, checked in the
# order listed below - deliberately looser than get_merge_key()'s "first
# token only" rule, since a filename should resolve correctly regardless
# of exactly how the raw or merged bucket label happens to be worded/
# ordered (e.g. "AIC Adamjee House Branch" and "Adamjee Insurance Co Ltd"
# should both resolve to the same canonical Adamjee filename). A company
# NOT in this list falls back to the existing safe_filename_part()-based
# naming in each generator, unchanged.
#
# IMPORTANT - UBL vs United: kept as two separate, non-overlapping entries
# on purpose. "UBL" and "UNITED" share no common substring, so a UBL
# Insurers bucket can never accidentally resolve to the United Insurance
# Company filename (or vice versa) - see the note above ABBREVIATIONS for
# the related merge-key bug this was fixed alongside.
# ---------------------------------------------------------------------------
CANONICAL_COMPANY_FILENAMES = [
    ("ADAMJEE", "Adamjee_Insurance_Company_LTD"),
    ("ALFALAH", "Alfalah_Insurance_Company_LTD"),
    ("ALPHA", "Alpha_Insurance_Company_LTD"),
    ("ASKARI", "Askari_Insurance_Company_LTD"),
    ("ASIA", "Asia_Insurance_Company_LTD"),
    ("ATLAS", "Atlas_Insurance_Company_LTD"),
    ("CENTURY", "Century_Insurance_Company_LTD"),
    ("CHUBB", "Chubb_Insurance_Company_Pakistan_LTD"),
    ("EASTWEST", "East_West_Insurance_Company_LTD"),
    ("EFU", "EFU_General_Insurance_Company_LTD"),
    ("HABIB", "Habib_Insurance_Company_LTD"),
    ("IGI", "IGI_General_Insurance_Company_LTD"),
    ("JUBILEE", "Jubilee_General_Insurance_Company_LTD"),
    ("PAKQATAR", "PAK_Qatar_General_Takaful"),
    ("PREMIER", "Premier_Insurance_Company_LTD"),
    ("RELIANCE", "Reliance_Insurance_Company_LTD"),
    ("SALAAM", "Salaam_Takaful_LTD"),
    ("SAUDIPAK", "Saudi_PAK_Insurance_Company_LTD"),
    ("SECURITYGENERAL", "Security_General_Insurance_Company_LTD"),
    ("SHAHEEN", "Shaheen_Insurance_Company_LTD"),
    ("PAKISTANGENERAL", "The_Pakistan_General_Insurance_Company_LTD"),
    ("TPL", "TPL_Insurance_Company_LTD"),
    ("UBL", "UBL_Insurance_Company_LTD"),
    ("UNITED", "United_Insurance_Company_LTD"),
    ("UNIVERSAL", "Universal_Insurance_Company_LTD"),
]


def canonical_filename_for_label(label):
    """
    Return the fixed canonical filename stub (no date suffix, no
    extension) for `label` if it identifies one of the companies listed in
    CANONICAL_COMPANY_FILENAMES above, else None - callers should fall
    back to their own existing safe_filename_part()-based naming when this
    returns None, so every company not on the fixed list keeps behaving
    exactly as before.

    An abbreviation-prefixed label (e.g. "AIC ...", "JGE ...") is resolved
    to its real company name first (the same ABBREVIATIONS lookup
    get_merge_key() already uses), so it matches by real name rather than
    needing its own separate substring entry here.
    """
    token_letters = _letters_upper(_first_token(label))
    search_text = _letters_upper(ABBREVIATIONS.get(token_letters, "")) + _letters_upper(label)
    for substring, filename in CANONICAL_COMPANY_FILENAMES:
        if substring in search_text:
            return filename
    return None


def self_merge_bucket_labels(bucket_labels, weight_by=None):
    """
    Group bucket_labels (the coinsurer/company names actually present in
    one run's output) by get_merge_key(). No external file, no ambiguity
    handling - every label always maps to something, deterministically.

    weight_by: optional {label: number} - typically that bucket's own
        claim/row count in this run. Within a merged group, the
        representative name is whichever member has the HIGHEST weight
        (ties broken alphabetically) - UNLESS at least one member's own
        leading token is a known abbreviation, in which case the real
        company name it stands for (e.g. "Adamjee Insurance") is always
        used as the representative, since that's more meaningful than an
        arbitrary raw label.

    Returns merge_map: {label: representative_label} - every input label
        maps to something (itself, if it never merged with anything).
    """
    labels = list(dict.fromkeys(str(l) for l in bucket_labels if str(l).strip()))
    weight_by = weight_by or {}

    groups = {}        # merge_key -> [labels]
    abbrev_name = {}    # merge_key -> the real company name, if any member's own leading token was a known abbreviation
    for l in labels:
        key = get_merge_key(l)
        token_letters = _letters_upper(_first_token(l))
        if token_letters in ABBREVIATIONS:
            abbrev_name[key] = ABBREVIATIONS[token_letters]
        groups.setdefault(key, []).append(l)

    merge_map = {}
    for key, members in groups.items():
        if len(members) == 1:
            merge_map[members[0]] = members[0]
            continue
        if key in abbrev_name:
            representative = abbrev_name[key]
        else:
            representative = sorted(members, key=lambda m: (-weight_by.get(m, 0), m))[0]
        for m in members:
            merge_map[m] = representative

    return merge_map


def merge_buckets(groups, totals, breakdowns, merge_map, combine_group):
    """
    Combine existing output buckets according to merge_map
    ({old_label: new_label}, from self_merge_bucket_labels()) into fewer,
    merged buckets - the actual "give me one consolidated file" step.
    No data is ever dropped: every row from every original bucket being
    folded together is preserved in the merged bucket.

    combine_group(list_of_group_values, new_label) -> merged_group_value:
        caller-supplied, since Claims Register's group value is a single
        DataFrame and Receivable's is a dict of two DataFrames
        ({"premium_rows", "claims_rows"}) - this function stays agnostic
        to that shape. The caller's combine_group is expected to also
        re-tag whatever column holds the company/coinsurer name to
        new_label on every row, so the merged workbook's own data sheet
        reads consistently rather than mixing several raw names together.

    totals: numeric values are summed across every bucket folded into a
        merged label - nothing is lost or recomputed from a subset.

    breakdowns: per-insured/per-client tables are concatenated then
        re-aggregated by name (first column), so a client/insured that
        shows up under two different raw company variants folds into one
        row instead of two, with its own numeric columns summed and
        re-sorted by the first numeric column, largest first.

    Returns (merged_groups, merged_totals, merged_breakdowns) - same
    shapes as the inputs, just with merge_map's clusters combined under
    their representative label.
    """
    from collections import defaultdict
    clusters = defaultdict(list)
    for old_label, new_label in merge_map.items():
        if old_label in groups:
            clusters[new_label].append(old_label)

    merged_groups, merged_totals, merged_breakdowns = {}, {}, {}
    for new_label, old_labels in clusters.items():
        merged_groups[new_label] = combine_group([groups[o] for o in old_labels], new_label)

        combined_totals = {}
        for o in old_labels:
            for k, v in totals.get(o, {}).items():
                # isinstance(v, (int, float)) alone misses numpy scalar
                # types (numpy.int64/float64 - exactly what a pandas
                # .sum() returns, which is what every totals dict here is
                # built from) - those would silently fall through to the
                # "keep only the first bucket's value" branch below,
                # quietly dropping every other merged bucket's amount.
                # numbers.Number (which numpy registers its scalar types
                # against) catches both plain Python and numpy numerics.
                if isinstance(v, numbers.Number) and not isinstance(v, bool):
                    combined_totals[k] = combined_totals.get(k, 0) + v
                elif k not in combined_totals:
                    combined_totals[k] = v
        merged_totals[new_label] = combined_totals

        frames = [breakdowns[o] for o in old_labels if o in breakdowns and breakdowns[o] is not None and len(breakdowns[o])]
        if frames:
            combined = pd.concat(frames, ignore_index=True)
            name_col = combined.columns[0]
            numeric_cols = [c for c in combined.columns[1:] if pd.api.types.is_numeric_dtype(combined[c])]
            other_cols = [c for c in combined.columns[1:] if c not in numeric_cols]
            agg = {c: "sum" for c in numeric_cols}
            for c in other_cols:
                agg[c] = "first"
            g = combined.groupby(name_col, as_index=False).agg(agg)
            g = g[list(combined.columns)]  # restore original column order
            if numeric_cols:
                g = g.sort_values(numeric_cols[0], ascending=False).reset_index(drop=True)
            merged_breakdowns[new_label] = g
        elif old_labels and old_labels[0] in breakdowns:
            merged_breakdowns[new_label] = breakdowns[old_labels[0]]

    return merged_groups, merged_totals, merged_breakdowns


# ---------------------------------------------------------------------------
# Email Lookup File parsing - UNCHANGED, unrelated to merging. Supplies
# recipient email addresses (and an optional point-of-contact name) for a
# company, keyed by its own name, for receivable_email.py /
# claims_register_email.py to look up. Never influences which bucket a
# row lands in.
# ---------------------------------------------------------------------------

_HEADER_ALIASES = {
    "s_no": ["s. no.", "s no", "sno", "sr.", "sr", "sr no", "sr. no."],
    "company_name": ["company name", "company", "coinsurer", "coinsurer name"],
    "email": ["email address", "email", "e-mail", "e-mail address", "emails"],
}


def _find_header_and_columns(raw_df, max_scan=25):
    """Scan the first `max_scan` rows (header=None read) for the row that
    looks like the real header (banner/title rows sit above it in the
    real file - confirmed 4 blank/title rows above the real header).
    Returns (header_row_index, company_col_idx, email_col_idx)."""
    for i in range(min(max_scan, len(raw_df))):
        row = raw_df.iloc[i].tolist()
        company_col, email_col = None, None
        for j, v in enumerate(row):
            if pd.isna(v):
                continue
            norm = str(v).strip().lower().rstrip(".")
            if company_col is None and norm in _HEADER_ALIASES["company_name"]:
                company_col = j
            if email_col is None and norm in [a.rstrip(".") for a in _HEADER_ALIASES["email"]]:
                email_col = j
        if company_col is not None and email_col is not None:
            return i, company_col, email_col
    raise ValueError(
        "Could not find a 'Company Name' and 'Email address' column in this file. "
        "Add the real header text to _HEADER_ALIASES in company_alias_resolver.py."
    )


_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def _clean_email_cell(value):
    """
    One cell's worth of a real email column can be:
      - a bare address: "south.cl@adamjeeinsurance.com"
      - "Name <address>": "Zaid Haris <Zaid.Haris@adamjeeinsurance.com>"
      - wrapped with a stray trailing ">" left over from a copy-pasted
        email header: "SM.OWAIS@adamjeeinsurance.com>"
      - a stray trailing ";" from a copy-pasted list: "ranajawed@...;"
      - MULTIPLE addresses in one cell, comma or semicolon separated
      - a placeholder meaning "no email on file": "N.A.", "-", "NA"
      - genuinely blank (NaN)
    Returns a list of cleaned addresses (possibly empty).
    """
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text:
        return []
    return _EMAIL_RE.findall(text)


def load_canonical_companies_and_emails(filepath, sheet_name=0):
    """
    Parse the real Co-insurance Companies List format: Company Name is
    only present on the FIRST row of each company's block (forward-fill
    down through blank rows), with one email address per subsequent row
    until the next company's name appears or a fully blank separator row
    is hit.

    Returns (canonical_companies, email_lookup, poc_lookup):
      canonical_companies: ordered list of distinct company names, exactly
          as they appear in the file.
      email_lookup: {normalized_company_name: "email1;email2;..."} - every
          valid email found for that company, semicolon-joined.
      poc_lookup: {normalized_company_name: "First Person Name"} - the
          first "Name <email>"-style entry found for that company, if any.
    """
    raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
    header_row, company_col, email_col = _find_header_and_columns(raw)

    canonical_companies = []
    email_lookup = {}
    poc_lookup = {}
    current_company = None
    current_key = None

    name_email_re = re.compile(r"^\s*([A-Za-z][A-Za-z .'\-]*?)\s*<")

    for i in range(header_row + 1, len(raw)):
        row = raw.iloc[i].tolist()
        company_val = row[company_col] if company_col < len(row) else None
        email_val = row[email_col] if email_col < len(row) else None

        if not pd.isna(company_val) and str(company_val).strip():
            current_company = str(company_val).strip()
            current_key = current_company.upper()
            if current_company not in canonical_companies:
                canonical_companies.append(current_company)

        if current_key is None:
            continue

        cleaned_emails = _clean_email_cell(email_val)
        if cleaned_emails:
            existing = email_lookup.get(current_key, "")
            emails = existing.split(";") if existing else []
            for cleaned in cleaned_emails:
                if cleaned.lower() not in (e.lower() for e in emails):
                    emails.append(cleaned)
            email_lookup[current_key] = ";".join(emails)

            if current_key not in poc_lookup and not pd.isna(email_val):
                name_match = name_email_re.match(str(email_val))
                if name_match:
                    poc_lookup[current_key] = name_match.group(1).strip()

    return canonical_companies, email_lookup, poc_lookup


if __name__ == "__main__":
    # Quick smoke test of the merge logic
    labels = [
        "Atlas_Insurance_Faisalabad",
        "Atlas_Insurance_ClaimRecovery",
        "AIC Adamjee House Branch",
        "Adamjee Insurance Company Limited",
        "JGE Bahawalnagar Branch",
        "UBL Insurers",
        "EFU General Insurance",
    ]
    weights = {l: i for i, l in enumerate(labels)}
    merge_map = self_merge_bucket_labels(labels, weight_by=weights)
    for l in labels:
        print(f"{l!r} -> {merge_map[l]!r}")
