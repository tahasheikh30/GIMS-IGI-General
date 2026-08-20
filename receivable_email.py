"""
Email drafting + sending for Premium & Claims Receivable Automation.

Reuses the Outstanding Premium module's own signature storage
(email_generator.load_signature/save_signature) and its Outlook-COM send
mechanism (outlook_sender.send_email) so credentials, templates, and the
signature box behave identically to the existing module.

One draft is built per output bucket (one per insurance company, or one
for the pooled "OFC" bucket - see receivable_processor.build_company_groups),
matching the one-workbook-per-company change: each email now covers every
client for that company in a single attached statement, rather than one
email per individual client.

RECIPIENT LOOKUP: the master transaction file has a single data sheet
and no longer carries its own address book, so recipient emails come
from a separate lookup file the user browses to (see
receivable_automation.load_email_lookup_from_file), keyed by Company
name - recipient email lookup is by that same bucket_label. For the
pooled "OFC" bucket, this looks up a lookup-file row keyed exactly "OFC"
- if the lookup file doesn't have one, that draft is flagged "Missing
email" for manual entry, same as any other unmatched company, rather
than guessing which of the pooled companies' contacts to use.
"""
from datetime import date

import company_alias_resolver as car
import email_generator as eg
import outlook_sender as osend

SUBJECT_TEMPLATE = "Premium & Claims Receivable Statement \u2013 {company_name} \u2013 {date}"

BODY_TEMPLATE = (
    "Dear {greeting_name},\n\n"
    "Please find attached the Premium & Claims Receivable statement for {company_name}, "
    "covering {num_clients} client{plural} as of {report_date}. It shows a total outstanding "
    "Premium Receivable of PKR {premium_total} and a total outstanding Claims Receivable of "
    "PKR {claims_total} across those clients.\n\n"
    "A full client-by-client breakdown is included in the attached workbook, along with a "
    "Summary sheet with charts.\n\n"
    "We would appreciate your review and settlement at your earliest convenience. "
    "If payment has already been made, please disregard this notice and share the "
    "payment confirmation with us.\n\n"
    "Please do not hesitate to contact us if you have any questions.\n\n"
    "Regards,\n"
    "{signature}"
)


def _normalize_key(text):
    return str(text).strip().upper()


def build_receivable_email_drafts(groups, totals, email_lookup, poc_lookup=None, signature=None, report_date=None):
    """
    groups: dict from receivable_processor.build_company_groups()
        {bucket_label: {"premium_rows": df, "claims_rows": df}}
    totals: dict from receivable_processor.company_totals()
        {bucket_label: {"premium_total": x, "claims_total": y, "num_clients": n}}
    email_lookup: dict, keyed by normalized (upper/stripped) Company name
        -> email address (semicolon-joined if the source listed several)
    poc_lookup: optional dict, same normalized Company-name keys -> contact
        person name. Carried through onto each draft's own "poc_name"
        field for reference, but never used for the greeting - every
        email always addresses the company itself ("Dear ADAMJEE
        INSURANCE COMPANY LTD,"), regardless of whether a POC name is
        on file.
    Returns list of draft dicts: {company_name, email, cc, bcc, poc_name,
        subject, body, num_clients, premium_total, claims_total, status}
    cc/bcc start blank - there's no lookup-file source for them, they're
    filled in (and saved per-draft) by hand on the Review & Send Emails
    tab, same manual-entry convention as a missing "email" value.
    """
    poc_lookup = poc_lookup or {}
    signature = signature.strip() if signature and signature.strip() else eg.DEFAULT_SIGNATURE
    report_date = report_date or date.today().isoformat()
    date_str = date.today().strftime("%d-%b-%Y")

    # See company_alias_resolver.resolve_company_email's own docstring -
    # same fix as claims_register_email.py: an exact string match between
    # the bucket's own raw Company name and the Email Lookup File's own
    # (differently-worded) Company Name essentially never holds, so this
    # falls back to the same merge-key logic the "Merge files for the
    # same company" checkbox already uses, computed once per run.
    merge_index = car.build_email_merge_index(email_lookup, poc_lookup)

    drafts = []
    for bucket_label, _rows in groups.items():
        t = totals[bucket_label]
        email, poc_name = car.resolve_company_email(bucket_label, email_lookup, poc_lookup, merge_index=merge_index)
        num_clients = t["num_clients"]

        subject = SUBJECT_TEMPLATE.format(company_name=bucket_label, date=date_str)
        body = BODY_TEMPLATE.format(
            greeting_name=bucket_label,
            company_name=bucket_label,
            num_clients=num_clients,
            plural="" if num_clients == 1 else "s",
            report_date=report_date,
            premium_total=f"{t['premium_total']:,.0f}",
            claims_total=f"{t['claims_total']:,.0f}",
            signature=signature,
        )

        drafts.append({
            "company_name": bucket_label,
            "email": email,
            "cc": "",
            "bcc": "",
            "poc_name": poc_name,
            "subject": subject,
            "body": body,
            "num_clients": num_clients,
            "premium_total": t["premium_total"],
            "claims_total": t["claims_total"],
            "status": "Ready" if email else "Missing email - needs manual entry",
        })
    return drafts


def send_receivable_emails(drafts, attachments_dir, filename_lookup, dry_run=False, progress_callback=None):
    """
    Send (or dry-run) each draft via the same Outlook mechanism as the
    Outstanding Premium module (outlook_sender.send_email - same signed-in
    Outlook account, no separate credentials).

    filename_lookup: dict bucket_label -> filename, as returned by the
        generator step, so the attachment always matches the file
        actually written for that company (or the pooled OFC bucket).
    Returns list of result dicts merged with each draft: adds sent (bool)
    and send_message (str). A draft with no email is recorded as a failed
    send (message: "No recipient email address") rather than silently
    skipped, so it shows up in the run's error report.
    """
    import os

    results = []
    total = len(drafts)
    for i, d in enumerate(drafts, start=1):
        key = d["company_name"]
        attachment_path = None
        fname = filename_lookup.get(key)
        if fname:
            candidate = os.path.join(attachments_dir, fname)
            if os.path.exists(candidate):
                attachment_path = candidate

        success, message = osend.send_email(d.get("email"), d["subject"], d["body"], attachment_path, dry_run,
                                             cc=d.get("cc"), bcc=d.get("bcc"))
        result = {**d, "sent": success, "send_message": message}
        results.append(result)
        if progress_callback:
            progress_callback(i, total, result)
    return results
