"""
Company Alias Resolver - for the Coinsurance modules (Claims Register and
Premium/Claims Email Automation), both of which group source rows by a
"company" name field (Coinsurer Name / Company) that real source data
records inconsistently: the SAME real company shows up under dozens of
different literal strings, because branch offices, divisions, and window
takaful units are each entered separately, e.g. (confirmed against a real
27,976-claim file - 98 distinct names for roughly 25 actual companies):

    ADAMJEE INSURANCE COMPANY LIMITED
    ADAMJEE INSURANCE CO LTD- ISLAMABAD BRANCH
    AIC Adamjee House Branch
    AIC Corp. Division Karachi
    AIC. Head Office, Karachi
    ...

Without merging, each of these becomes its OWN separate output workbook
and its own separate (usually "missing email") row on the Send Emails
tab, instead of one combined file/email per real company.

WHY THIS IS ITS OWN MODULE, NOT JUST FUZZIER STRING MATCHING: some of the
real variants above share almost no characters with the company's own
full name at all - "AIC", "JGI", "RIC" are ACRONYMS (Adamjee Insurance
Company, Jubilee General Insurance, Reliance Insurance Company), which a
generic character-similarity ratio (SequenceMatcher etc.) cannot detect,
since "AIC" and "ADAMJEE" simply don't resemble each other as strings.
Acronym matching is added as its own signal for exactly this reason - but
acronym matching is ALSO inherently ambiguous on its own (see
_generate_acronym's docstring: "AIC" is the acronym of FOUR different
real Pakistani insurance companies - Adamjee, Alfalah, Alpha, and Asia -
purely by coincidence of all four starting with "A" + "Insurance" +
"Company"). Silently picking one would risk merging (and therefore
misrouting a recovery email for) a genuinely different company's claims.

THE SAFE DEFAULT, per explicit instruction: anything the matcher can't
resolve with high confidence is left UNMAPPED - it stays exactly as its
own separate raw name/bucket/file, exactly like today's behavior, unless
a person has explicitly reviewed and filled it in. Nothing is ever
silently merged on a guess.

WORKFLOW:
  1. generate_alias_suggestions_file() - scan the raw company/coinsurer
     names actually present in a loaded source file, suggest a canonical
     company (from the email lookup file's own Company Name list) for
     every name the matcher is confident about, and write out a plain
     Excel file with every raw name, its (possibly blank) suggestion, a
     confidence score, and a plain-English reason - for a person to
     review, correct, or fill in the blanks.
  2. load_alias_mapping_file() - read that file back (after a person has
     possibly edited it) into a plain {raw_name: canonical_company} dict,
     using ONLY the rows where the Canonical Company column is filled in.
  3. apply_alias() - look up one raw name; returns the canonical name if
     mapped, otherwise returns the raw name completely unchanged (the
     safe default - "not confidently mapped" behaves identically to "no
     mapping exists at all").
"""
import os
import re
from difflib import SequenceMatcher

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Name normalization
# ---------------------------------------------------------------------------

# Generic insurance-industry words that appear in almost every company's own
# name and therefore carry no distinguishing information on their own -
# excluded when looking for a company's own DISTINCTIVE brand word (e.g.
# "ADAMJEE", "ALFALAH", "ATLAS"), but DELIBERATELY KEPT IN when generating a
# company's own ACRONYM (see _generate_acronym's docstring for why - real
# acronyms used in practice, like "AIC" and "JGI", are built from these
# words too, not just the brand name alone).
_GENERIC_WORDS = {
    "INSURANCE", "GENERAL", "COMPANY", "CO", "LTD", "LIMITED", "TAKAFUL",
    "THE", "OF", "AND", "WINDOW", "OPERATIONS", "OPERATION", "BRANCH",
    "BRANCHES", "DIVISION", "DIVISIONS", "UNIT", "UNITS", "OFFICE",
    "OFFICES", "MAIN", "HEAD", "REGIONAL", "CORPORATE", "CORP", "SALES",
    "COMMERCIAL", "CENTRAL", "GLASS", "TOWER", "NEW", "CITY", "H.O.",
    "HO", "PVT", "PRIVATE",
}

# Words that are branch/city/location qualifiers commonly appended to a
# company's own name in real data - stripping these before comparison
# avoids a location word coincidentally resembling a different company's
# own brand word.
_LOCATION_WORDS = {
    "KARACHI", "LAHORE", "ISLAMABAD", "RAWALPINDI", "FAISALABAD", "MULTAN",
    "PESHAWAR", "SIALKOT", "GULBERG", "GULLBERG", "BAHAWALNAGAR", "FSD",
    "KHI", "LHR", "MUL", "ROAD", "MARKET", "MARKAZ", "PLAZA", "CENTRE",
    "CENTER", "F-7", "MALL",
}

_WORD_RE = re.compile(r"[A-Za-z]+")


def _words(name):
    return _WORD_RE.findall(str(name).upper())


def normalize_name(name):
    """Uppercase, letters-only-per-word, generic/location words stripped -
    used as the dict key for alias lookups, so trivial punctuation/spacing
    differences (periods, commas, extra whitespace, hyphens) never cause a
    lookup miss on an otherwise-identical raw name."""
    return " ".join(w for w in _words(name))


def _distinctive_words(name):
    """A company's own BRAND word(s) - e.g. "ADAMJEE", "ALFALAH", "ATLAS",
    "JUBILEE", "RELIANCE" - with generic insurance-industry terms and
    branch/location qualifiers stripped out. Only words of 4+ letters are
    kept, since shorter words (e.g. "PAK" in "PAK QATAR") are too common
    to safely use as a standalone matching signal on their own."""
    return [w for w in _words(name)
            if len(w) >= 4 and w not in _GENERIC_WORDS and w not in _LOCATION_WORDS]


def _generate_acronym(name):
    """
    A company's own likely acronym, built from the first letter of every
    SIGNIFICANT word in its name INCLUDING generic industry terms like
    "INSURANCE"/"GENERAL"/"COMPANY" (only pure grammatical filler - "THE",
    "OF", "AND", "LTD", "LIMITED" - is excluded). This is deliberately
    different from _distinctive_words(), which excludes those same
    generic terms - because real acronyms used in practice are built FROM
    them: confirmed against real data, "AIC" = Adamjee INSURANCE COMPANY,
    "JGI" = Jubilee General Insurance (further confirmed: real entries
    literally spell out "JGI COMMERCIAL UNIT II" etc. alongside "JUBILEE
    GENERAL INSURANCE LTD" entries for the same company), "RIC" =
    Reliance Insurance Company. A brand-word-only acronym (skipping
    "INSURANCE"/"COMPANY") would produce the wrong 1-2 letter result and
    match nothing real.

    IMPORTANT - this is INHERENTLY AMBIGUOUS and must never be trusted
    alone: many different real companies share the exact same acronym by
    coincidence (e.g. "AIC" is the acronym of Adamjee, Alfalah, Alpha,
    AND Asia Insurance Company - all four happen to start with "A" +
    "Insurance" + "Company"). See suggest_canonical_company()'s own
    docstring for how this is only ever used as a confirmed-unique signal,
    never as a standalone match.
    """
    words = [w for w in _words(name) if w not in {"THE", "OF", "AND", "LTD", "LIMITED"}]
    return "".join(w[0] for w in words)


# ---------------------------------------------------------------------------
# Parsing the real Co-insurance Companies List (canonical names + emails)
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
      - MULTIPLE addresses in one cell, comma or semicolon separated (an
        older lookup-file convention some existing files still use, e.g.
        "a@x.com, b@y.com") - every address in the cell is extracted, not
        just the first, so this format keeps working unchanged
      - a placeholder meaning "no email on file": "N.A.", "-", "NA"
      - genuinely blank (NaN)
    Returns a list of cleaned addresses (possibly empty) - extracting via
    a regex findall rather than trying to enumerate every wrapper format,
    so it's robust to variants not explicitly listed above.
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
          as they appear in the file (used as the target list for
          suggest_canonical_company()).
      email_lookup: {normalized_company_name: "email1;email2;..."} - every
          valid email found for that company, semicolon-joined (Outlook's
          To field accepts this natively) - matches the shape
          receivable_email.py / claims_register_email.py already expect.
      poc_lookup: {normalized_company_name: "First Person Name"} - the
          first "Name <email>"-style entry found for that company, if any
          (used for the email greeting) - left unset for a company with
          only bare addresses, exactly like the rest of GIMS already does
          when no contact name is available.
    """
    raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
    header_row, company_col, email_col = _find_header_and_columns(raw)

    canonical_companies = []
    email_lookup = {}
    poc_lookup = {}
    current_company = None
    current_key = None  # keyed by simple .strip().upper() - matches the
    # existing convention already used by receivable_email.py's and
    # claims_register_email.py's own _normalize_key(), so this dict drops
    # in as-is wherever those modules already expect an email_lookup /
    # poc_lookup dict, with no changes needed on their side.

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


# ---------------------------------------------------------------------------
# Suggesting a canonical company for a raw name
# ---------------------------------------------------------------------------

HIGH_CONFIDENCE_THRESHOLD = 0.85


def _build_distinctive_word_index(canonical_companies):
    """{word: [company, ...]} for every distinctive word across every
    canonical company. A word is only safe to use as a matching signal
    for ONE company if that company is the ONLY one whose name contains
    it - otherwise the word isn't actually distinctive at all, it's just
    a word two or more real companies happen to share (confirmed real
    case: "PAKISTAN" appears in both "CHUBB INSURANCE COMPANY PAKISTAN
    LTD" and the separately-listed "THE PAKISTAN GENERAL INSURANCE
    COMPANY" - treating it as if it belonged to only one of them would
    misroute the other company's claims entirely)."""
    index = {}
    for c in canonical_companies:
        for w in _distinctive_words(c):
            index.setdefault(w, [])
            if c not in index[w]:
                index[w].append(c)
    return index


def suggest_canonical_company(raw_name, canonical_companies, canonical_acronyms=None, distinctive_word_index=None):
    """
    Try to confidently match one raw company/coinsurer name (as it
    literally appears in source data) against the list of canonical
    company names from the email lookup file. Returns
    (canonical_name_or_None, confidence, reason).

    Checked in order, each a stronger/weaker signal:
      1. Exact match (after normalize_name) -> confidence 1.0
      2. The raw name contains a word that is a DISTINCTIVE word (see
         _distinctive_words) of exactly ONE canonical company (checked
         via distinctive_word_index - a word shared by two or more
         canonical companies' own names is excluded here, it isn't
         actually distinguishing anything) -> confidence 0.9. This is
         what correctly merges e.g. "ADAMJEE INSURANCE CO LTD- ISLAMABAD
         BRANCH" under "ADAMJEE INSURANCE COMPANY LTD" - it literally
         contains the word ADAMJEE, and no other canonical company does.
      3. The raw name's own leading acronym-shaped token (e.g. "AIC",
         "JGI", "RIC") matches a canonical company's generated acronym -
         either exactly, or as a prefix of it (some real acronyms in use
         drop a trailing generic word, e.g. "JGI" for what this module
         would otherwise generate as "JGIC" from "JUBILEE GENERAL
         INSURANCE COMPANY") - AND that acronym prefix is UNIQUE among
         every canonical company's own acronym -> confidence 0.85. If
         ambiguous (shared by 2+ canonical companies, e.g. bare "AIC"
         alone could be Adamjee/Alfalah/Alpha/Asia), this is deliberately
         NOT used as a match - returns (None, 0.0, an explanation naming
         every ambiguous candidate) rather than guessing one of them.

         IMPORTANT SAFETY CHECK: if step 2 (distinctive word) ALSO found
         a candidate, and it DISAGREES with what step 3 would suggest,
         this is treated as a CONFLICT, not resolved silently in either
         direction - confirmed real case: "JGI UNITED MALL BR, MULTAN"
         acronym-matches Jubilee ("JGI"), but "UNITED" is also, by pure
         coincidence, Insurance Company's own distinctive word (United
         Mall is a real shopping mall name, unrelated to the insurer) -
         picking either signal alone here would have been a genuine
         misroute, so a disagreement between the two signals falls
         through to manual review instead.
      4. A generic character-similarity ratio (SequenceMatcher) against
         each canonical name -> only returned as a suggestion (with its
         own, generally lower, confidence) if it clears
         HIGH_CONFIDENCE_THRESHOLD; a good fuzzy ratio below that
         threshold is still reported back to the caller as the closest
         name found, but at a confidence too low to auto-suggest.

    canonical_acronyms / distinctive_word_index: optional precomputed
        lookups (see generate_alias_suggestions_file) - built fresh if
        not given.
    """
    if not str(raw_name).strip():
        return None, 0.0, "blank name"

    raw_norm = normalize_name(raw_name)
    raw_words = _words(raw_name)
    raw_words_set = set(raw_words)

    for c in canonical_companies:
        if normalize_name(c) == raw_norm:
            return c, 1.0, "exact match"

    if distinctive_word_index is None:
        distinctive_word_index = _build_distinctive_word_index(canonical_companies)

    word_match = None  # (word, company)
    for w in raw_words:
        owners = distinctive_word_index.get(w)
        if owners and len(owners) == 1:
            word_match = (w, owners[0])
            break

    if canonical_acronyms is None:
        canonical_acronyms = {}
        for c in canonical_companies:
            canonical_acronyms.setdefault(_generate_acronym(c), []).append(c)

    acronym_match = None  # (token, company)
    leading_token = raw_words_ordered_first_token(raw_name)
    if leading_token:
        # Exact acronym match, or the token is a genuine PREFIX of a
        # longer canonical acronym (e.g. "JGI" is a prefix of "JGIC") -
        # collect every canonical company either way could plausibly mean.
        candidates = set(canonical_acronyms.get(leading_token, []))
        for full_acronym, owners in canonical_acronyms.items():
            if len(leading_token) >= 2 and full_acronym.startswith(leading_token) and full_acronym != leading_token:
                candidates.update(owners)
        if len(candidates) == 1:
            acronym_match = (leading_token, next(iter(candidates)))
        elif len(candidates) > 1:
            acronym_match = (leading_token, None)  # ambiguous - marker, not a real match

    if word_match and acronym_match:
        _, word_company = word_match
        _, acronym_company = acronym_match
        if acronym_company is None:
            # Acronym alone is ambiguous, but the distinctive word narrows
            # it down unambiguously - safe to trust the word.
            return word_company, 0.9, f"contains distinctive word '{word_match[0]}' from '{word_company}'"
        if word_company != acronym_company:
            return None, 0.0, (f"conflicting signals - leading token '{leading_token}' suggests "
                                f"'{acronym_company}' but the word '{word_match[0]}' suggests '{word_company}' - "
                                f"needs manual review")
        return word_company, 0.9, f"contains distinctive word '{word_match[0]}' from '{word_company}' (acronym agrees)"

    if word_match:
        _, word_company = word_match
        return word_company, 0.9, f"contains distinctive word '{word_match[0]}' from '{word_company}'"

    if acronym_match:
        token, acronym_company = acronym_match
        if acronym_company is not None:
            return acronym_company, 0.85, f"acronym '{token}' uniquely matches '{acronym_company}'"
        owners = canonical_acronyms.get(token, [])
        all_owners = sorted(set(owners) | {c for full, own in canonical_acronyms.items()
                                            if len(token) >= 2 and full.startswith(token) for c in own})
        return None, 0.0, (f"acronym '{token}' is ambiguous - could be any of: "
                            f"{', '.join(all_owners)} - needs manual review")

    best_name, best_ratio = None, 0.0
    for c in canonical_companies:
        ratio = SequenceMatcher(None, raw_norm, normalize_name(c)).ratio()
        if ratio > best_ratio:
            best_name, best_ratio = c, ratio

    if best_ratio >= HIGH_CONFIDENCE_THRESHOLD:
        return best_name, round(best_ratio, 3), f"close text match to '{best_name}' ({best_ratio:.0%} similar)"
    if best_name is not None:
        return None, round(best_ratio, 3), f"closest guess was '{best_name}' ({best_ratio:.0%} similar) - not confident enough"
    return None, 0.0, "no similar canonical company found"


def raw_words_ordered_first_token(name):
    """The first ALL-CAPS-shaped alphabetic token in the raw name, if it
    looks acronym-shaped (2-5 letters) - e.g. "AIC" from "AIC Corp.
    Division Karachi", "JGI" from "JGI BAHAWALNAGAR BRANCH". Returns None
    if the first word doesn't look like an acronym (too long, or the
    source text wasn't actually leading with one)."""
    words = _words(name)
    if not words:
        return None
    first = words[0]
    return first if 2 <= len(first) <= 5 else None


# ---------------------------------------------------------------------------
# Generating and loading the editable alias suggestions file
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial")
NEEDS_REVIEW_FILL = PatternFill("solid", start_color="FDF3EA", end_color="FDF3EA")
NEEDS_REVIEW_FONT = Font(name="Arial", color="8A6100")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

ALIAS_FILE_COLUMNS = ["Raw Name (from source data)", "Canonical Company", "Confidence", "Match Reason"]


def generate_alias_suggestions_file(raw_names, canonical_companies, output_path):
    """
    raw_names: every distinct raw company/coinsurer name found in a
        loaded source file (e.g. matched_df["Coinsurer Name"].unique()
        for Claims Register, or clean_df[company_col].unique() for
        Receivable).
    canonical_companies: from load_canonical_companies_and_emails().
    output_path: where to write the editable suggestions workbook.

    Writes one row per distinct raw name, "Canonical Company" PRE-FILLED
    only for names the matcher is confident about (see
    suggest_canonical_company - confidence >= HIGH_CONFIDENCE_THRESHOLD),
    left BLANK for anything ambiguous or unrecognized, highlighted amber
    with "Confidence" and a plain-English "Match Reason" so a person can
    see WHY it wasn't auto-filled (e.g. "acronym 'AIC' is ambiguous...")
    and fill in the right answer themselves rather than guessing.

    This file is meant to be reviewed/edited, then handed back to
    load_alias_mapping_file() as the actual, approved mapping to use -
    nothing here is applied automatically just by generating it.

    Returns output_path.
    """
    canonical_acronyms = {}
    for c in canonical_companies:
        canonical_acronyms.setdefault(_generate_acronym(c), []).append(c)
    distinctive_word_index = _build_distinctive_word_index(canonical_companies)

    rows = []
    for raw_name in sorted(set(str(n).strip() for n in raw_names if str(n).strip())):
        suggestion, confidence, reason = suggest_canonical_company(
            raw_name, canonical_companies, canonical_acronyms, distinctive_word_index)
        rows.append((raw_name, suggestion or "", confidence, reason))

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Company Aliases"

    for j, col in enumerate(ALIAS_FILE_COLUMNS, start=1):
        c = ws.cell(row=1, column=j, value=col)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 20

    n_needs_review = 0
    for i, (raw_name, suggestion, confidence, reason) in enumerate(rows, start=2):
        needs_review = not suggestion
        if needs_review:
            n_needs_review += 1
        values = [raw_name, suggestion, confidence if confidence else None, reason]
        for j, val in enumerate(values, start=1):
            cell = ws.cell(row=i, column=j, value=val)
            cell.border = BORDER
            if needs_review:
                cell.font = NEEDS_REVIEW_FONT
                cell.fill = NEEDS_REVIEW_FILL
            else:
                cell.font = BODY_FONT
            if j == 3 and val is not None:
                cell.number_format = "0%"

    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 60

    note_row = len(rows) + 3
    ws.cell(row=note_row, column=1,
             value=(f"{n_needs_review} of {len(rows)} name(s) need manual review (highlighted) - fill in "
                    f"'Canonical Company' for any row you want merged under a real company. Leave blank to "
                    f"keep that name as its own separate file/email, exactly like today.")
             ).font = Font(name="Arial", size=9, italic=True, color="6C7A89")

    wb.save(output_path)
    return output_path


def load_alias_mapping_file(filepath):
    """
    Read a (possibly hand-edited) alias suggestions file back in. Only
    rows with a non-blank "Canonical Company" are included - anything
    left blank is deliberately excluded, so it falls through to the
    normal "use the raw name as its own bucket" behavior everywhere this
    mapping is applied (see apply_alias).

    Returns {normalize_name(raw_name): canonical_company}.
    """
    df = pd.read_excel(filepath, sheet_name=0)
    col_map = {c.strip().lower(): c for c in df.columns}
    raw_col = col_map.get("raw name (from source data)") or df.columns[0]
    canon_col = col_map.get("canonical company") or df.columns[1]

    mapping = {}
    for _, row in df.iterrows():
        raw_name = row.get(raw_col)
        canonical = row.get(canon_col)
        if pd.isna(raw_name) or pd.isna(canonical):
            continue
        canonical = str(canonical).strip()
        if not canonical:
            continue
        mapping[normalize_name(raw_name)] = canonical
    return mapping


def apply_alias(raw_name, mapping):
    """Look up one raw name in a loaded alias mapping. Returns the
    canonical company name if mapped, otherwise returns raw_name
    completely UNCHANGED - the safe default: "not in the mapping"
    behaves identically to "no mapping was ever built", so this is
    always safe to call even with an empty/None mapping."""
    if not mapping or pd.isna(raw_name):
        return raw_name
    return mapping.get(normalize_name(raw_name), raw_name)


if __name__ == "__main__":
    import sys
    lookup_fp = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/Co-insurance_Companies_List.xlsx"
    companies, emails, pocs = load_canonical_companies_and_emails(lookup_fp)
    print(f"Loaded {len(companies)} canonical companies, {len(emails)} with email(s), {len(pocs)} with a POC name")
    for c in companies[:5]:
        key = c.upper()
        print(f"  {c!r} -> emails={emails.get(key)!r} poc={pocs.get(key)!r}")
