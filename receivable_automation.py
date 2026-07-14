"""
Premium & Claims Receivable Email Automation
Top-level orchestration module - mirrors the Outstanding Premium
workflow (data_processor -> report_generator -> email_generator ->
outlook_sender) but for the Premium/Claims Receivable statements.

Run as a script for a quick end-to-end test:
    python receivable_automation.py <master_file.xlsx> [--email-lookup <lookup_file.xlsx>] [--send]

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


def load_email_lookup_from_file(filepath):
    """
    Load the recipient-email lookup from a SEPARATE lookup file (.xlsx or
    .csv) that the user browses to on tab 1 - keyed by Company (the
    co-insurance partner), e.g. "ADAMJEE INSURANCE COMPANY LTD" -> its
    Finance contact email(s). This mirrors the Outstanding Premium
    module's own email-lookup-file pattern (see app.py's load_data /
    data_processor), just keyed by Company instead of Client, since the
    master transaction workbook itself only has a single sheet and no
    longer carries its own address-book sheet.

    Expects a plain header row at the top of the file: a name/company
    column (matched by "company", "client", or "participant" appearing in
    the header text - falls back to the first column if none match), an
    "Email" column (matched by "email" in the header text), and an
    optional POC/contact column (matched by "poc", "contact", "attn",
    "attention", "recipient name", or "contact person").

    A company can have several emails in one cell, comma/semicolon
    separated (e.g. "a@x.com, b@x.com") - these are preserved and passed
    to Outlook as a semicolon-separated To list.

    Returns (email_lookup, poc_lookup) - both dicts keyed by normalized
    (upper/stripped) company name, matching receivable_email.py's lookup
    convention.
    """
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(filepath)
    else:
        df = pd.read_excel(filepath)

    if len(df.columns) == 0:
        raise ValueError("The email lookup file has no columns.")

    name_col = next(
        (c for c in df.columns if any(k in str(c).lower() for k in ["company", "client", "participant"])),
        df.columns[0],
    )
    email_col = next((c for c in df.columns if "email" in str(c).lower()), None)
    if email_col is None:
        raise ValueError("Could not find an 'Email' column in the lookup file.")
    poc_col = next(
        (c for c in df.columns
         if c not in (name_col, email_col)
         and any(k in str(c).lower() for k in ["poc", "contact", "attn", "attention", "recipient name", "contact person"])),
        None,
    )

    email_lookup, poc_lookup = {}, {}
    for _, row in df.iterrows():
        company_raw = row.get(name_col)
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


def run_pipeline(master_filepath, output_dir, load_email_lookup=True, email_lookup_filepath=None,
                  signature=None, dry_run=True, report_date=None, logger=print):
    """
    Full pipeline: load -> validate -> group -> generate workbooks ->
    draft emails -> send (or dry-run) -> log + return a summary.

    master_filepath: the uploaded workbook, which has a single data sheet
        (no sheet name needs to be specified - it's read directly).
    email_lookup_filepath: optional SEPARATE lookup file (.xlsx or .csv)
        the user browses to, keyed by Company name -> email(s) (see
        load_email_lookup_from_file). Set load_email_lookup=False, or
        leave this unset, to skip email lookup entirely - all drafts will
        then show "Missing email" for manual entry.
    output_dir: configurable folder for generated workbooks (created if
        missing).
    logger: callable(str) used for progress/status lines - defaults to
        print, but a GUI can pass in a log-widget-appending function
        instead, same pattern as app.py's self.log().

    Returns a summary dict:
        {
          "num_groups": int,   # number of output files/buckets (companies + pooled OFC)
          "generated_files": [filepath, ...],
          "validation_issues": DataFrame,
          "drafts": [draft, ...],
          "send_results": [result, ...],
          "missing_email": [company_name, ...],
          "errors": [str, ...],
        }
    """
    errors = []
    logger(f"Loading master file: {master_filepath}")
    try:
        raw_df = rp.load_raw_data(master_filepath)
    except Exception as e:
        logger(f"ERROR: {e}")
        return _empty_summary(str(e))

    try:
        clean_df, issues_df, dropped_columns, colmap = rp.validate_data(raw_df)
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

    groups = rp.build_company_groups(clean_df)
    totals = rp.company_totals(clean_df, groups)
    logger(f"\n{len(groups)} output file(s) identified (one per insurance company, "
           f"with every company containing \"OFC\" in its name pooled into a single shared 'OFC' file).")

    if colmap["mode"] == "dual":
        premium_amount_col, claims_amount_col = colmap["premium_amount"], colmap["claims_amount"]
    else:
        premium_amount_col = claims_amount_col = colmap["amount"]

    os.makedirs(output_dir, exist_ok=True)
    generated_files = []
    filename_lookup = {}
    for bucket_label, g in groups.items():
        try:
            breakdown = rp.client_breakdown_for_bucket(clean_df, bucket_label)
            fp = rg.build_company_workbook(
                bucket_label, g["premium_rows"], g["claims_rows"], breakdown,
                output_dir, premium_amount_col=premium_amount_col, claims_amount_col=claims_amount_col,
                run_date=None,
            )
            generated_files.append(fp)
            filename_lookup[bucket_label] = os.path.basename(fp)
            logger(f"  Generated: {os.path.basename(fp)} ({totals[bucket_label]['num_clients']} client(s))")
        except Exception as e:
            msg = f"Failed to generate workbook for {bucket_label}: {e}"
            logger(f"  ERROR: {msg}")
            errors.append(msg)

    email_lookup, poc_lookup = {}, {}
    if load_email_lookup and email_lookup_filepath:
        try:
            email_lookup, poc_lookup = load_email_lookup_from_file(email_lookup_filepath)
            logger(f"\nLoaded {len(email_lookup)} company email mapping(s) from '{os.path.basename(email_lookup_filepath)}'.")
        except Exception as e:
            msg = f"Could not load email lookup: {e}"
            logger(f"WARNING: {msg}")
            errors.append(msg)
    elif load_email_lookup and not email_lookup_filepath:
        logger("\nNo email lookup file provided - all drafts will show 'Missing email'.")

    if signature is None:
        signature = eg.load_signature()

    drafts = re_.build_receivable_email_drafts(
        groups, totals, email_lookup, poc_lookup=poc_lookup,
        signature=signature, report_date=report_date,
    )

    missing_email = [d["company_name"] for d in drafts if not d.get("email")]
    if missing_email:
        logger(f"\n{len(missing_email)} file(s) have no recipient email (flagged, not skipped):")
        for company_name in missing_email[:20]:
            logger(f"   {company_name}")
        if len(missing_email) > 20:
            logger(f"   ...and {len(missing_email) - 20} more.")

    logger(f"\n{'[DRY RUN] ' if dry_run else ''}Sending {len(drafts)} email(s)...")
    send_results = re_.send_receivable_emails(
        drafts, output_dir, filename_lookup, dry_run=dry_run,
        progress_callback=lambda i, total, r: logger(
            f"  [{i}/{total}] {r['company_name']}: {r['send_message']}"
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
    logger(f"Files processed:       {len(groups)}")
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
        print("Usage: python receivable_automation.py <master_file.xlsx> [--email-lookup <lookup_file.xlsx>] [--send]")
        sys.exit(1)

    master_file = sys.argv[1]
    dry_run = "--send" not in sys.argv[2:]
    email_lookup_filepath = None
    if "--email-lookup" in sys.argv[2:]:
        idx = sys.argv.index("--email-lookup")
        email_lookup_filepath = sys.argv[idx + 1]

    summary = run_pipeline(
        master_filepath=master_file,
        output_dir=os.path.join(os.getcwd(), "receivable_output"),
        email_lookup_filepath=email_lookup_filepath,
        dry_run=dry_run,
    )
