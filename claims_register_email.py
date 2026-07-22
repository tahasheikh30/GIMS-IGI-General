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

import email_generator as eg
import outlook_sender as osend

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
    "Regards,\n"
    "{signature}"
)


def _normalize_key(text):
    return str(text).strip().upper()


def build_claims_register_email_drafts(groups, totals, email_lookup, poc_lookup=None, signature=None, report_date=None):
    """
    groups: {bucket_label: claims_df} from claims_register_processor.build_coinsurer_buckets
    totals: {bucket_label: {"num_claims", "total_share_paid", "total_share_os"}}
    email_lookup / poc_lookup: dicts keyed by normalized (upper/stripped) coinsurer name
    Returns list of draft dicts: {company_name, email, poc_name, subject, body,
        num_claims, paid_total, os_total, status}
    """
    poc_lookup = poc_lookup or {}
    signature = signature.strip() if signature and signature.strip() else eg.DEFAULT_SIGNATURE
    report_date = report_date or date.today().isoformat()
    date_str = date.today().strftime("%d-%b-%Y")

    drafts = []
    for bucket_label in groups:
        t = totals[bucket_label]
        key = _normalize_key(bucket_label)
        email = email_lookup.get(key)
        poc_name = poc_lookup.get(key)
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
        )

        drafts.append({
            "company_name": bucket_label,
            "email": email,
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

        success, message = osend.send_email(d.get("email"), d["subject"], d["body"], attachment_path, dry_run)
        result = {**d, "sent": success, "send_message": message}
        results.append(result)
        if progress_callback:
            progress_callback(i, total, result)
    return results
