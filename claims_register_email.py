"""
Email drafting + sending for the Claims Register (Coinsurance) module.

One draft per output bucket (one per coinsurer, or one shared draft for
the pooled "OFC" bucket - see claims_register_processor.build_coinsurer_
buckets), reusing the Outstanding Premium module's own signature storage
(email_generator.load_signature/save_signature) and Outlook-COM send
mechanism (outlook_sender.send_email), same pattern as receivable_email.py.

RECIPIENT LOOKUP: same convention as receivable_email.py - a separate
lookup file the user browses to, keyed by (normalized upper/stripped)
coinsurer name. The pooled "OFC" bucket is looked up under the literal
key "OFC".
"""
from datetime import date

import company_alias_resolver as car
import email_generator as eg
import outlook_sender as osend

# Default Cc/Bcc recipients for every Claims Intimation (Claims Register)
# email, per Finance's own standing instruction - always included unless
# a person clears/edits the Cc/Bcc field by hand on the Review & Send
# Emails tab for a specific draft. Only the bare address is kept from
# each "Name/Dept/City <email>" entry Finance supplied, since that's the
# only part Outlook's own Cc/Bcc fields actually need; the name/
# department/city portion was just how the addresses were originally
# copied out of Outlook.
DEFAULT_BCC_RECIPIENTS = [
    "awais.amjad@IGI.COM.PK",
    "murtaza.feroz@IGI.COM.PK",
]
DEFAULT_BCC = ";".join(DEFAULT_BCC_RECIPIENTS)

DEFAULT_CC_RECIPIENTS = [
    "kashif.qayyum@IGI.COM.PK",
    "abdul.rauf@IGI.COM.PK",
    "majid.masood@IGI.COM.PK",
    "zahid.mehmood@IGI.COM.PK",
    "waqas.muhammad@IGI.COM.PK",
    "fazal.abbas@IGI.COM.PK",
    "Haider.Ali@IGI.COM.PK",
    "nadeem.alvi@IGI.COM.PK",
    "Muhammad.Nadeem@IGI.COM.PK",
    "adeel.ahmad@IGI.COM.PK",
]
DEFAULT_CC = ";".join(DEFAULT_CC_RECIPIENTS)

SUBJECT_TEMPLATE = "{type_prefix}Claims Intimation Register \u2013 {company_name} \u2013 {date}"

BODY_TEMPLATE = (
    "Dear {greeting_name},\n\n"
    "Please find attached the {type_prefix}Claims Intimation Register for {period_text}, for your "
    "review and necessary action.\n\n"
    "Should you have any questions or require further clarification, please do not hesitate to "
    "contact us.\n\n"
    "For any queries, please feel free to contact Abdul Rauf ({abdul_rauf_email}) or "
    "Muhammad Majid Masood ({majid_masood_email}).\n\n"
    "Regards,\n"
    "{signature}"
)

# Bare addresses reused in BODY_TEMPLATE's contact line above - kept as
# their own constants (rather than hardcoded inline in the template
# string) so they stay in sync with DEFAULT_CC_RECIPIENTS above if either
# person's address is ever updated, without needing to edit the template
# text itself.
ABDUL_RAUF_EMAIL = "abdul.rauf@IGI.COM.PK"
MAJID_MASOOD_EMAIL = "majid.masood@IGI.COM.PK"


def build_claims_register_email_drafts(groups, totals, email_lookup, poc_lookup=None, signature=None,
                                        report_date=None, period_text=None, business_type="conventional"):
    """
    groups: {bucket_label: claims_df} from claims_register_processor.build_coinsurer_buckets
    totals: {bucket_label: {"num_claims", "total_share_paid", "total_share_os"}}
    email_lookup / poc_lookup: dicts keyed by normalized (upper/stripped) coinsurer name.
        poc_lookup is carried through onto each draft's own "poc_name"
        field for reference, but never used for the greeting - every
        email always addresses the coinsurer itself ("Dear ADAMJEE
        INSURANCE COMPANY LTD,"), regardless of whether a POC name is
        on file.
    period_text: text describing the period covered (e.g. "the period from
        01 January 2026 to 31 July 2026"), dropped into the body as-is.
        Falls back to a generic phrase built from report_date if not given.
    business_type: "conventional" (default) or "takaful" - a Takaful run
        prefixes both the subject and the body's own reference to the
        Claims Intimation Register with "Takaful ", so the recipient
        never has to guess which book of business a statement belongs
        to. Conventional runs are worded exactly as before (no prefix).
    Returns list of draft dicts: {company_name, email, cc, bcc, poc_name,
        subject, body, num_claims, paid_total, os_total, status}
    cc and bcc both start pre-filled with Finance's own standing lists
    (DEFAULT_CC_RECIPIENTS / DEFAULT_BCC_RECIPIENTS) on every draft -
    still fully editable/clearable per draft on the Review & Send Emails
    tab like any other field, this is just the default each one starts at.
    """
    poc_lookup = poc_lookup or {}
    signature = signature.strip() if signature and signature.strip() else eg.DEFAULT_SIGNATURE
    report_date = report_date or date.today().isoformat()
    date_str = date.today().strftime("%d-%b-%Y")
    period_text = period_text or f"the period ending {report_date}"
    type_prefix = "Takaful " if business_type == "takaful" else ""

    # Precomputed once per run (see company_alias_resolver.resolve_company_email's
    # own docstring for why exact-match alone misses almost every real
    # company - the bucket label is the coinsurer's raw name straight off
    # the Premium Register's own "Ins" column, e.g. "AIC Corp. Division
    # Karachi", while the Email Lookup File's own Company Name column
    # spells the same real company differently, e.g. "ADAMJEE INSURANCE
    # COMPANY LTD" - the merge-key fallback is what actually bridges that).
    merge_index = car.build_email_merge_index(email_lookup, poc_lookup)

    drafts = []
    for bucket_label in groups:
        t = totals[bucket_label]
        email, poc_name = car.resolve_company_email(bucket_label, email_lookup, poc_lookup, merge_index=merge_index)
        num_claims = t["num_claims"]

        subject = SUBJECT_TEMPLATE.format(type_prefix=type_prefix, company_name=bucket_label, date=date_str)
        body = BODY_TEMPLATE.format(
            type_prefix=type_prefix,
            greeting_name=bucket_label,
            period_text=period_text,
            signature=signature,
            abdul_rauf_email=ABDUL_RAUF_EMAIL,
            majid_masood_email=MAJID_MASOOD_EMAIL,
        )

        drafts.append({
            "company_name": bucket_label,
            "email": email,
            "cc": DEFAULT_CC,
            "bcc": DEFAULT_BCC,
            "poc_name": poc_name,
            "subject": subject,
            "body": body,
            "num_claims": num_claims,
            "paid_total": t["total_share_paid"],
            "os_total": t["total_share_os"],
            "status": "Ready" if email else "Missing email - needs manual entry",
        })
    return drafts


def send_claims_register_emails(drafts, attachments_dir, filename_lookup, dry_run=False, progress_callback=None):
    """Same send mechanism/shape as receivable_email.send_receivable_emails."""
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
