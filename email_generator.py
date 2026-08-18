"""
Email drafting: fills a template with per-client data.
Templates use {placeholder} syntax, e.g.:
    "Dear {client_name}, your outstanding balance of {amount} ..."

PERSISTENCE: signature.json and templates.json are stored in a local,
per-machine folder (see app_paths.py), not next to this script, and are
written atomically (see safe_io.py). Both together are what stop these
files from corrupting when the project itself is shared across laptops
via a synced OneDrive folder. See either module's own docstring for the
full rationale.
"""
import os

import app_paths
import safe_io

DEFAULT_SIGNATURE = "Finance Department\nIGI General Takaful"

DEFAULT_TEMPLATE = {
    "name": "Standard Reminder",
    "subject": "Outstanding Balance Notice - {client_name}",
    "body": (
        "Dear {greeting_name},\n\n"
        "As per our records, your account currently shows an outstanding balance "
        "of PKR {amount} as of {report_date}.\n\n"
        "A detailed breakdown by department and year is attached for your reference.\n\n"
        "We would appreciate your settlement at your earliest convenience. "
        "If payment has already been made, please disregard this notice and share "
        "the payment confirmation with us.\n\n"
        "Please do not hesitate to contact us if you have any questions.\n\n"
        "Regards,\n"
        "{signature}"
    ),
}

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_FILE = app_paths.resolve_state_file("templates.json", legacy_dir=_APP_DIR)
SIGNATURE_FILE = app_paths.resolve_state_file("signature.json", legacy_dir=_APP_DIR)


def load_signature():
    """Load the user's personal signature (persisted locally, per machine)."""
    data = safe_io.safe_read_json(SIGNATURE_FILE, default=None)
    if data:
        return data.get("signature", DEFAULT_SIGNATURE)
    return DEFAULT_SIGNATURE


def save_signature(signature_text):
    safe_io.atomic_write_json(SIGNATURE_FILE, {"signature": signature_text})


def load_templates():
    data = safe_io.safe_read_json(TEMPLATES_FILE, default=None)
    if data:
        return data
    return {"Standard Reminder": DEFAULT_TEMPLATE}


def save_template(template_name, subject, body):
    templates = load_templates()
    templates[template_name] = {"name": template_name, "subject": subject, "body": body}
    safe_io.atomic_write_json(TEMPLATES_FILE, templates)


def fill_template(template, fields):
    """
    template: dict with 'subject' and 'body' containing {placeholder} tokens
    fields: dict of placeholder -> value
    Missing placeholders are left as-is rather than raising, so partial
    data doesn't break the whole run (flagged separately for review).
    """
    class SafeDict(dict):
        def __missing__(self, key):
            return "{" + key + "}"

    safe_fields = SafeDict(fields)
    subject = template["subject"].format_map(safe_fields)
    body = template["body"].format_map(safe_fields)
    return subject, body


def build_email_drafts(qualifying_df, grand_totals, email_lookup, poc_lookup=None, template=None, report_date=None, signature=None):
    """
    qualifying_df: index = client names that qualify
    grand_totals: Series client -> total O/s
    email_lookup: dict client_name -> email address (or None if missing)
    poc_lookup: optional dict client_name -> point-of-contact person name.
        When present for a client, the email greeting addresses that person
        ("Dear John,") instead of the company name ("Dear ACME CORP,").
    signature: text shown after "Regards," at the end of every email. Falls
        back to DEFAULT_SIGNATURE if not provided or blank.
    Returns list of draft dicts: {client_name, email, cc, bcc, subject, body, amount, status}
    cc/bcc start blank - there's no source column for them, they're filled
    in (and saved per-draft) by hand on the Review & Send Emails tab,
    same manual-entry convention as a missing "email" value.
    """
    if template is None:
        template = DEFAULT_TEMPLATE
    poc_lookup = poc_lookup or {}
    signature = signature.strip() if signature and signature.strip() else DEFAULT_SIGNATURE
    drafts = []
    for client in qualifying_df:
        amount = grand_totals.get(client, 0)
        email = email_lookup.get(client)
        poc_name = poc_lookup.get(client)
        fields = {
            "client_name": client,
            "greeting_name": poc_name if poc_name else client,
            "amount": f"{amount:,.0f}",
            "report_date": report_date or "",
            "signature": signature,
        }
        subject, body = fill_template(template, fields)
        drafts.append({
            "client_name": client,
            "email": email,
            "cc": "",
            "bcc": "",
            "poc_name": poc_name,
            "subject": subject,
            "body": body,
            "amount": amount,
            "status": "Ready" if email else "Missing email - needs manual entry",
        })
    return drafts


if __name__ == "__main__":
    import data_processor as dp
    raw = dp.load_raw_data("/mnt/user-data/uploads/Karachi_Conventional_Outstanding.xlsx", sheet_name="Sheet1")
    clean, issues, _dropped = dp.validate_data(raw)
    df, totals = dp.build_pivot(clean)
    q = dp.qualifying_clients(df, totals, threshold=2_000_000)
    fake_email_lookup = {q.index[0]: "finance@example-client.com", q.index[1]: "ap@example-client2.com"}
    fake_poc_lookup = {q.index[0]: "Ahmed Khan"}  # only first client has a named POC
    drafts = build_email_drafts(q.index[:3], totals, fake_email_lookup, poc_lookup=fake_poc_lookup, report_date="2026-07-07")
    for d in drafts:
        print(d["client_name"], "|", d["status"], "|", d["subject"])
        print(d["body"][:80], "...\n")
