"""
Sends emails through the user's own signed-in Outlook desktop application
using COM automation (pywin32). This only works on Windows with Outlook
installed and configured - it cannot be tested in this sandbox.

Install dependency on the target Windows machine:
    pip install pywin32
"""

try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False


def send_email(to_address, subject, body, attachment_path=None, dry_run=False):
    """
    Send a single email via Outlook.
    Returns (success: bool, message: str).
    """
    if not to_address:
        return False, "No recipient email address"

    if dry_run:
        return True, f"[DRY RUN] Would send to {to_address}: '{subject}'" + \
                      (f" with attachment {attachment_path}" if attachment_path else "")

    if not OUTLOOK_AVAILABLE:
        return False, "pywin32 / Outlook not available on this machine"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = to_address
        mail.Subject = subject
        mail.Body = body
        if attachment_path:
            mail.Attachments.Add(attachment_path)
        mail.Send()
        return True, f"Sent to {to_address}"
    except Exception as e:
        return False, f"Failed to send to {to_address}: {e}"


def send_bulk(drafts, attachments_dir=None, dry_run=False, progress_callback=None):
    """
    drafts: list of dicts with keys client_name, email, subject, body
    attachments_dir: folder where per-client reports live (matched by filename)
    progress_callback: optional fn(index, total, result_dict) called after each send
    Returns list of result dicts merged with each draft.
    """
    import os
    import report_generator as rg

    results = []
    total = len(drafts)
    for i, d in enumerate(drafts, start=1):
        attachment_path = None
        if attachments_dir:
            fname = f"{rg.safe_filename(d['client_name'])}_OutstandingBalance_Report.xlsx"
            candidate = os.path.join(attachments_dir, fname)
            if os.path.exists(candidate):
                attachment_path = candidate

        success, message = send_email(d.get("email"), d["subject"], d["body"], attachment_path, dry_run)
        result = {**d, "sent": success, "send_message": message}
        results.append(result)
        if progress_callback:
            progress_callback(i, total, result)
    return results


if __name__ == "__main__":
    # Dry-run demo (safe to run anywhere, does not require Outlook)
    demo_draft = [{
        "client_name": "TEST CLIENT",
        "email": "test@example.com",
        "subject": "Outstanding Balance Notice - TEST CLIENT",
        "body": "Dear TEST CLIENT, ...",
    }]
    results = send_bulk(demo_draft, dry_run=True)
    for r in results:
        print(r["send_message"])
