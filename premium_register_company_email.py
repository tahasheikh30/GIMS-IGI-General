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

# Default Cc/Bcc recipients for every Premium Register email, per
# Finance's own standing instruction - identical to Claims Intimation's
# own lists (see claims_register_email.py) with a few substitutions/
# additions on the Cc list: Muhammad Majid Masood is replaced by Usama
# Jahangir (Underwriting, Lahore), and Sabeen Ahmed Nomani and Jahanzaib
# Khan (this module's own three named contacts in the body below - Abdul
# Rauf and Sabeen for Karachi, Jahanzaib for Lahore) are added alongside
# him. Always included unless a person clears/edits the Cc/Bcc field by
# hand on the Review & Send Emails tab for a specific draft. Only the
# bare address is kept, same convention as claims_register_email.py's
# own DEFAULT_BCC_RECIPIENTS.
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
    "usama.jahangir@IGI.COM.PK",
    "sabeen.nomani@IGI.COM.PK",
    "jahanzaib.khan@IGI.COM.PK",
    "zahid.mehmood@IGI.COM.PK",
    "waqas.muhammad@IGI.COM.PK",
    "fazal.abbas@IGI.COM.PK",
    "Haider.Ali@IGI.COM.PK",
    "nadeem.alvi@IGI.COM.PK",
    "Muhammad.Nadeem@IGI.COM.PK",
]
DEFAULT_CC = ";".join(DEFAULT_CC_RECIPIENTS)

SUBJECT_TEMPLATE = "Premium Ceded Register \u2013 {company_name} \u2013 {date}"

BODY_TEMPLATE = (
    "Dear {greeting_name},\n\n"
    "Please find attached our premium ceded register relating to {company_name} for {period_text}, "
    "for your review and reconciliation.\n\n"
    "We would appreciate it if you could also share your premium register for the same period for "
    "our verification and reconciliation purposes.\n\n"
    "For any queries or if any additional/missing documents are required, please feel free to "
    "coordinate directly with the following representatives:\n\n"
    "1. Abdul Rauf \u2013 {abdul_rauf_email} (Karachi Region)\n"
    "2. Sabeen Nomani \u2013 {sabeen_email} (Karachi Region)\n"
    "3. Jahanzaib Khan \u2013 {jahanzaib_email} (Lahore Region)\n\n"
    "Regular sharing and reconciliation of premium registers will help both companies maintain "
    "accurate records and ensure that our books are reconciled on a timely basis.\n\n"
    "Your cooperation in this regard will be highly appreciated.\n\n"
    "Regards,\n"
    "{signature}"
)

ABDUL_RAUF_EMAIL = "abdul.rauf@IGI.COM.PK"
SABEEN_EMAIL = "sabeen.nomani@IGI.COM.PK"
# Lahore Region point of contact, added alongside the two existing Karachi
# Region contacts (Abdul Rauf, Sabeen Nomani) - named in this module's own
# body template so recipients outside Karachi have a Lahore-based contact
# to coordinate with directly, rather than defaulting to Karachi contacts
# regardless of the coinsurer's own region.
JAHANZAIB_EMAIL = "jahanzaib.khan@IGI.COM.PK"


def _normalize_key(text):
    return str(text).strip().upper()


def build_premium_register_email_drafts(groups, totals, email_lookup, poc_lookup=None, signature=None,
                                         report_date=None, period_text=None):
    """
    groups: {bucket_label: rows_df} from premium_register_company_processor.build_company_buckets
    totals: {bucket_label: {"num_policies", "total_gross_premium", "total_net_premium"}}
    email_lookup / poc_lookup: dicts keyed by normalized (upper/stripped) company name
    period_text: text describing the period covered (e.g. "the period from
        01 January 2026 to 31 July 2026"), dropped into the body as-is.
        Falls back to a generic phrase built from report_date if not given.
    Returns list of draft dicts: {company_name, email, cc, bcc, poc_name,
        subject, body, num_policies, gross_total, net_total, status}
    cc and bcc both start pre-filled with Finance's own standing lists
    (DEFAULT_CC_RECIPIENTS / DEFAULT_BCC_RECIPIENTS) on every draft - same
    as claims_register_email.py's own drafts - still fully editable/
    clearable per draft on the Review & Send Emails tab like any other
    field, this is just the default each one starts at.
    """
    poc_lookup = poc_lookup or {}
    signature = signature.strip() if signature and signature.strip() else eg.DEFAULT_SIGNATURE
    report_date = report_date or date.today().isoformat()
    date_str = date.today().strftime("%d-%b-%Y")
    period_text = period_text or f"the period ending {report_date}"

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
            period_text=period_text,
            signature=signature,
            abdul_rauf_email=ABDUL_RAUF_EMAIL,
            sabeen_email=SABEEN_EMAIL,
            jahanzaib_email=JAHANZAIB_EMAIL,
        )

        drafts.append({
            "company_name": bucket_label,
            "email": email,
            "cc": DEFAULT_CC,
            "bcc": DEFAULT_BCC,
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
