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
    "Faisal.Khan@IGI.COM.PK",
    "muhammad.hisham@IGI.COM.PK",
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
]
DEFAULT_CC = ";".join(DEFAULT_CC_RECIPIENTS)

SUBJECT_TEMPLATE = "Claims Recovery Statement \u2013 {company_name} \u2013 {date}"

BODY_TEMPLATE = (
    "Dear {greeting_name},\n\n"
    "Please find attached the claims recovery statement covering {num_claims} claim{plural} on which "
    "{company_name} is a coinsurer, as of {report_date}. Your share of claims already paid by IGI totals "
    "PKR {paid_total}, and your share of claims still outstanding totals PKR {os_total}.\n\n"
    "A full claim-by-claim breakdown, including your own Share % on each policy, is included in the "
    "attached workbook, along with a Summary sheet.\n\n"
    "We would appreciate your review and settlement of the paid-claims portion at your earliest "
    "convenience. If payment has already been made against any of these claims, please disregard that "
    "line and share the payment confirmation with us.\n\n"
    "Please do not hesitate to contact us if you have any questions.\n\n"
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


def _normalize_key(text):
    return str(text).strip().upper()


def build_claims_register_email_drafts(groups, totals, email_lookup, poc_lookup=None, signature=None, report_date=None):
    """
    groups: {bucket_label: claims_df} from claims_register_processor.build_coinsurer_buckets
    totals: {bucket_label: {"num_claims", "total_share_paid", "total_share_os"}}
    email_lookup / poc_lookup: dicts keyed by normalized (upper/stripped) coinsurer name
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

        subject = SUBJECT_TEMPLATE.format(company_name=bucket_label, date=date_str)
        body = BODY_TEMPLATE.format(
            greeting_name=poc_name if poc_name else bucket_label,
            company_name=bucket_label,
            num_claims=num_claims,
            plural="" if num_claims == 1 else "s",
            report_date=report_date,
            paid_total=f"{t['total_share_paid']:,.0f}",
            os_total=f"{t['total_share_os']:,.0f}",
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
