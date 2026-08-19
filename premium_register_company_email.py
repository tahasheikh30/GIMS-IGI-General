"""
Email drafting + sending for the Premium Register (Coinsurance) company
split module. Same shape as claims_register_email.py / receivable_email.py:
one draft per output bucket (one per coinsurer, or one shared draft for
the pooled "OFC" bucket), reusing email_generator.load_signature and
outlook_sender.send_email.

RECIPIENT LOOKUP: same convention as the other coinsurance modules - a
separate lookup file the user browses to, keyed by (normalized upper/
stripped) company name. The pooled "OFC" bucket is looked up under the
literal key "OFC".
"""
from datetime import date

import company_alias_resolver as car
import email_generator as eg
import outlook_sender as osend

SUBJECT_TEMPLATE = "Premium Register Statement \u2013 {company_name} \u2013 {date}"

BODY_TEMPLATE = (
    "Dear {greeting_name},\n\n"
    "Please find attached the Premium Register statement covering {num_policies} "
    "polic{plural} on which {company_name} is a coinsurer, as of {report_date}. Your share "
    "of Gross Premium across these policies totals PKR {gross_total}, and your share of "
    "Net Premium totals PKR {net_total}.\n\n"
    "A full policy-by-policy breakdown, including your own Share % on each policy, is "
    "included in the attached workbook, along with a Summary sheet.\n\n"
    "Please do not hesitate to contact us if you have any questions.\n\n"
    "Regards,\n"
    "{signature}"
)


def _normalize_key(text):
    return str(text).strip().upper()


def build_premium_register_email_drafts(groups, totals, email_lookup, poc_lookup=None, signature=None, report_date=None):
    """
    groups: {bucket_label: rows_df} from premium_register_company_processor.build_company_buckets
    totals: {bucket_label: {"num_policies", "total_gross_premium", "total_net_premium"}}
    email_lookup / poc_lookup: dicts keyed by normalized (upper/stripped) company name
    Returns list of draft dicts: {company_name, email, cc, bcc, poc_name,
        subject, body, num_policies, gross_total, net_total, status}
    cc/bcc start blank - filled in (and saved per-draft) by hand on the
    Review & Send Emails tab, same convention as every other draft in GIMS.
    """
    poc_lookup = poc_lookup or {}
    signature = signature.strip() if signature and signature.strip() else eg.DEFAULT_SIGNATURE
    report_date = report_date or date.today().isoformat()
    date_str = date.today().strftime("%d-%b-%Y")

    merge_index = car.build_email_merge_index(email_lookup, poc_lookup)

    drafts = []
    for bucket_label in groups:
        t = totals[bucket_label]
        email, poc_name = car.resolve_company_email(bucket_label, email_lookup, poc_lookup, merge_index=merge_index)
        num_policies = t["num_policies"]

        subject = SUBJECT_TEMPLATE.format(company_name=bucket_label, date=date_str)
        body = BODY_TEMPLATE.format(
            greeting_name=poc_name if poc_name else bucket_label,
            company_name=bucket_label,
            num_policies=num_policies,
            plural="y" if num_policies == 1 else "ies",
            report_date=report_date,
            gross_total=f"{t['total_gross_premium']:,.0f}",
            net_total=f"{t['total_net_premium']:,.0f}",
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
            "num_policies": num_policies,
            "gross_total": t["total_gross_premium"],
            "net_total": t["total_net_premium"],
            "status": "Ready" if email else "Missing email - needs manual entry",
        })
    return drafts


def send_premium_register_emails(drafts, attachments_dir, filename_lookup, dry_run=False, progress_callback=None):
    """Same send mechanism/shape as receivable_email.send_receivable_emails
    and claims_register_email.send_claims_register_emails."""
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
