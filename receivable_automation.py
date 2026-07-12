"""
Premium & Claims Receivable Email Automation
Top-level orchestration module - mirrors the Outstanding Premium
workflow (data_processor -> report_generator -> email_generator ->
outlook_sender) but for the Premium/Claims Receivable statements.

Run as a script for a quick end-to-end test:
    python receivable_automation.py <master_file.xlsx> [--sheet "Detail"]
        [--addresses-sheet "Addresses Co Insurance"] [--send]

Or call run_pipeline(...) directly from a GUI (see the "Receivable" tab
wired into app.py, following the same pattern as the existing Outstanding
Premium / Loss Ratio tabs).
"""
import os
import re
import pandas as pd

import receivable_processor as rp
import receivable_generator as rg
import receivable_email as re_
import receivable_run_log as rlog
import email_generator as eg


def load_email_lookup_from_addresses_sheet(filepath, sheet_name="Addresses Co Insurance"):
    """
    Load the recipient-email lookup from the real file's own address-book
    sheet, keyed by Company (the co-insurance partner) - e.g. "ADAMJEE
    INSURANCE COMPANY LTD" -> its Finance contact email(s). This is a
    different sheet in the SAME workbook as the transaction data, not a
    separate lookup file, matching the real sample's layout.

    Handles two real quirks of that sheet:
      - The header row isn't row 1 (a blank spacer row sits above it) -
        reuses the same header auto-detection as the transaction sheet.
      - The email column itself has no header text in the real sample -
        detected by scanning for the column with the most "@" values.

    A company can have several emails in one cell, comma/semicolon
    separated (e.g. "a@x.com, b@x.com") - these are preserved and passed
    to Outlook as a semicolon-separated To list.

    Returns (email_lookup, poc_lookup) - both dicts keyed by normalized
    (upper/stripped) company name.
    """
    raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
    # The address-book sheet typically has only one alias-recognizable
    # header ("Company Name" - things like "Contact Person" or "S. No"
    # aren't in ALIASES), so use a lower confidence threshold than the
    # transaction sheet's header detection.
    header_row = rp.find_header_row(raw, min_score=1)
    columns = raw.iloc[header_row].tolist()
    df = raw.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = columns
    df = df.dropna(how="all").reset_index(drop=True)

    colmap = rp._map_columns_to_canonical(df.columns)
    company_col = next((orig for orig, canon in colmap.items() if canon == "Company"), None)
    if company_col is None:
        raise ValueError(
            f"Could not find a Company Name column in sheet '{sheet_name}'. "
            f"Recognized headers were: {list(colmap.keys())}."
        )

    email_col = None
    for col in df.columns:
        if col == company_col:
            continue
        sample = df[col].dropna().astype(str)
        if len(sample) and sample.str.contains("@").sum() > 0:
            email_col = col
            break
    if email_col is None:
        raise ValueError(f"Could not find an email column (no column contains '@' values) in sheet '{sheet_name}'.")

    poc_col = next((c for c in df.columns
                     if c not in (company_col, email_col)
                     and any(k in str(c).lower() for k in ["poc", "contact", "attn", "attention", "recipient name", "contact person"])),
                    None)

    email_lookup, poc_lookup = {}, {}
    for _, row in df.iterrows():
        company_raw = row.get(company_col)
        email_raw = row.get(email_col)
        if pd.isna(company_raw) or pd.isna(email_raw):
            continue
        key = str(company_raw).strip().upper()
        if not key:
            continue
        emails = [e.strip() for e in re.split(r"[;,]", str(email_raw)) if "@" in e]
        if emails:
            email_lookup[key] = ";".join(emails)
        if poc_col is not None and not pd.isna(row.get(poc_col)):
            poc_name = str(row[poc_col]).strip()
            if poc_name:
                poc_lookup[key] = poc_name

    return email_lookup, poc_lookup


def run_pipeline(master_filepath, output_dir, sheet_name="Detail", addresses_sheet_name="Addresses Co Insurance",
                  load_email_lookup=True, require_ref_for_claims_only_rows=True,
                  signature=None, dry_run=True, report_date=None, logger=print):
    """
    Full pipeline: load -> validate -> group -> generate workbooks ->
    draft emails -> send (or dry-run) -> log + return a summary.

    master_filepath: the uploaded workbook. Both the transaction sheet
        (sheet_name) and the address-book sheet (addresses_sheet_name) are
        read from this SAME file, matching the real sample's layout. Set
        load_email_lookup=False to skip email lookup entirely (e.g. if a
        given file has no address-book sheet) - all drafts will then show
        "Missing email".
    output_dir: configurable folder for generated workbooks (created if
        missing).
    require_ref_for_claims_only_rows: see receivable_processor.validate_data
        - defaults to the literal spec (blank Leader Ref Number is always
        excluded). On a file shaped like the real sample this means Claims
        Receivable sheets come out mostly empty; set False to keep
        claims-only rows that just lack a Leader Ref Number.
    logger: callable(str) used for progress/status lines - defaults to
        print, but a GUI can pass in a log-widget-appending function
        instead, same pattern as app.py's self.log().

    Returns a summary dict:
        {
          "num_groups": int,
          "generated_files": [filepath, ...],
          "validation_issues": DataFrame,
          "drafts": [draft, ...],
          "send_results": [result, ...],
          "missing_email": [ (client, company), ... ],
          "errors": [str, ...],
        }
    """
    errors = []
    logger(f"Loading master file: {master_filepath} (sheet: {sheet_name})")
    try:
        raw_df = rp.load_raw_data(master_filepath, sheet_name=sheet_name)
    except Exception as e:
        logger(f"ERROR: {e}")
        return _empty_summary(str(e))

    try:
        clean_df, issues_df, dropped_columns, colmap = rp.validate_data(
            raw_df, require_ref_for_claims_only_rows=require_ref_for_claims_only_rows
        )
    except ValueError as e:
        logger(f"ERROR: {e}")
        return _empty_summary(str(e))

    logger(f"Detected layout: {colmap['mode']} mode "
           f"({'Premium Receivable + Claim Receivable amount columns' if colmap['mode'] == 'dual' else 'Type/Category + single Amount column'}).")
    logger(f"Loaded {len(raw_df)} rows, {len(clean_df)} valid after validation.")
    if dropped_columns:
        logger(f"Removed {len(dropped_columns)} completely empty column(s): {', '.join(dropped_columns)}")
    if len(issues_df) > 0:
        logger(f"{len(issues_df)} row(s) flagged (see validation_issues in the returned summary):")
        issue_counts = issues_df["issue"].value_counts()
        for issue_label, count in issue_counts.items():
            logger(f"   {count}x  {issue_label}")
        if not require_ref_for_claims_only_rows:
            logger("   (claims-only rows missing a Leader Ref Number were flagged but NOT excluded, per require_ref_for_claims_only_rows=False)")
        elif "Missing Leader Ref Number" in issue_counts and issue_counts.get("Missing Leader Ref Number", 0) > 0:
            logger("   NOTE: a large share of missing-Leader-Ref-Number rows are commonly claims-only rows in this "
                   "export format - if Claims Receivable sheets come out mostly empty, re-run with "
                   "require_ref_for_claims_only_rows=False.")

    groups = rp.build_groups(clean_df)
    totals = rp.group_totals(clean_df, groups)
    logger(f"\n{len(groups)} Client+Company group(s) identified.")

    if colmap["mode"] == "dual":
        premium_amount_col, claims_amount_col = colmap["premium_amount"], colmap["claims_amount"]
    else:
        premium_amount_col = claims_amount_col = colmap["amount"]

    os.makedirs(output_dir, exist_ok=True)
    generated_files = []
    filename_lookup = {}
    for (client_name, company_name), g in groups.items():
        try:
            fp = rg.build_receivable_workbook(
                client_name, company_name, g["premium_rows"], g["claims_rows"],
                output_dir, premium_amount_col=premium_amount_col, claims_amount_col=claims_amount_col,
                run_date=None,
            )
            generated_files.append(fp)
            filename_lookup[(client_name, company_name)] = os.path.basename(fp)
            logger(f"  Generated: {os.path.basename(fp)}")
        except Exception as e:
            msg = f"Failed to generate workbook for {client_name} / {company_name}: {e}"
            logger(f"  ERROR: {msg}")
            errors.append(msg)

    email_lookup, poc_lookup = {}, {}
    if load_email_lookup:
        try:
            email_lookup, poc_lookup = load_email_lookup_from_addresses_sheet(master_filepath, addresses_sheet_name)
            logger(f"\nLoaded {len(email_lookup)} company email mapping(s) from '{addresses_sheet_name}'.")
        except Exception as e:
            msg = f"Could not load email lookup: {e}"
            logger(f"WARNING: {msg}")
            errors.append(msg)

    if signature is None:
        signature = eg.load_signature()

    drafts = re_.build_receivable_email_drafts(
        groups, totals, email_lookup, poc_lookup=poc_lookup,
        signature=signature, report_date=report_date,
    )

    missing_email = [(d["client_name"], d["company_name"]) for d in drafts if not d.get("email")]
    if missing_email:
        logger(f"\n{len(missing_email)} group(s) have no recipient email (flagged, not skipped):")
        for client_name, company_name in missing_email[:20]:
            logger(f"   {client_name} / {company_name}")
        if len(missing_email) > 20:
            logger(f"   ...and {len(missing_email) - 20} more.")

    logger(f"\n{'[DRY RUN] ' if dry_run else ''}Sending {len(drafts)} email(s)...")
    send_results = re_.send_receivable_emails(
        drafts, output_dir, filename_lookup, dry_run=dry_run,
        progress_callback=lambda i, total, r: logger(
            f"  [{i}/{total}] {r['client_name']} / {r['company_name']}: {r['send_message']}"
        ),
    )

    num_sent = sum(1 for r in send_results if r["sent"])
    num_failed = len(send_results) - num_sent

    rlog.log_run(
        num_groups=len(groups),
        num_files_generated=len(generated_files),
        num_sent=num_sent,
        num_failed=num_failed,
        num_missing_email=len(missing_email),
        dry_run=dry_run,
        errors=errors,
    )

    logger("\n" + "=" * 60)
    logger("RUN SUMMARY")
    logger("=" * 60)
    logger(f"Groups processed:      {len(groups)}")
    logger(f"Workbooks generated:   {len(generated_files)}")
    logger(f"Emails {'previewed (dry run)' if dry_run else 'sent'}:      {num_sent}")
    logger(f"Emails failed:         {num_failed}")
    logger(f"Missing email:         {len(missing_email)}")
    logger(f"Other errors:          {len(errors)}")

    return {
        "num_groups": len(groups),
        "generated_files": generated_files,
        "validation_issues": issues_df,
        "drafts": drafts,
        "send_results": send_results,
        "missing_email": missing_email,
        "errors": errors,
    }


def _empty_summary(error_message):
    return {
        "num_groups": 0, "generated_files": [], "validation_issues": pd.DataFrame(),
        "drafts": [], "send_results": [], "missing_email": [], "errors": [error_message],
    }


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python receivable_automation.py <master_file.xlsx> [--send]")
        sys.exit(1)

    master_file = sys.argv[1]
    dry_run = "--send" not in sys.argv[2:]

    summary = run_pipeline(
        master_filepath=master_file,
        output_dir=os.path.join(os.getcwd(), "receivable_output"),
        dry_run=dry_run,
    )
