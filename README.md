# Outstanding Balance Reporting & Client Email System

A desktop tool for the Finance department that replaces manual pivot-table
building and email drafting: upload the client balance sheet, set a
threshold, and it generates per-client Excel breakdowns and drafts
(and optionally sends) reminder emails through your own Outlook account.

## What's included

| File | Purpose |
|---|---|
| `app.py` | Main desktop GUI (Tkinter) — run this |
| `data_processor.py` | Loads, validates, and aggregates the raw balance sheet |
| `report_generator.py` | Builds one styled Excel file per qualifying client (Department x Year) |
| `email_generator.py` | Fills the email template per client; template stored in `templates.json` |
| `outlook_sender.py` | Sends via Outlook COM automation (your own signed-in account) |
| `run_log.py` | Writes an audit trail to `run_history.json` |

## Requirements

- Windows PC with **Microsoft Outlook desktop app** installed and signed in
  (emails send through your own mailbox — no separate login needed)
- Python 3.9+

Install dependencies:
```
pip install -r requirements.txt
```

## Running it

```
python app.py
```

### Step 1 — Upload & Filter
1. Browse to your balance data `.xlsx` file (the raw transaction sheet, e.g. "Sheet1").
2. Confirm the sheet name.
3. **Optional:** browse to an email lookup file (`.xlsx` or `.csv`) with a client-name
   column and an `Email` column. Rows are matched to clients by exact name.
   If you skip this, drafted emails will be flagged "missing email" and you can
   fill them in manually in tab 3.
   - **Optional POC/contact column:** if the lookup file also has a column named
     something like `POC`, `Contact`, `Contact Person`, or `Attention`, that
     person's name is used in the greeting ("Dear Ahmed Khan,") instead of the
     company name. Rows with no name in that column just fall back to the
     company name as before — no need for every row to have one.
4. Click **Load & Validate Data** — this flags rows with a missing client name,
   missing department, or a non-numeric balance (these are excluded from totals,
   and listed in the log so you can fix the source sheet if needed).
5. Choose which date drives the "Year" breakdown: **Issue Date** or
   **Comm. Date** (commencement date). Issue Date is the default.
6. Set your **threshold** (PKR). Optionally select one or more departments
   and/or one or more agencies/agents (Ctrl+click for multiple), and/or type
   exact company names to restrict to specific clients. Selecting an agency
   first narrows to that agency's clients, and the threshold then applies to
   each of those clients individually.
7. Click **Analyze / Apply Filters**.

### Step 2 — Qualifying Clients
Shows every client whose total outstanding balance (across all departments/years,
within any filters applied) meets or exceeds the threshold, largest first, along
with the agent(s)/agency associated with that client. Click **Generate Excel
Reports** to write one file per client (named
`ClientName_OutstandingBalance_Report.xlsx`) into an `output/` folder. Each file
now contains **two sheets**:
- **"Outstanding Balance"** — the Department x Year pivot with Grand Total row/column
- **"Raw Data"** — every original transaction row for that client from the source sheet

### Step 3 — Review & Send Emails
1. Click **Draft Emails** to generate a personalized email per qualifying client
   from the template in `email_generator.py` / `templates.json`.
2. Click any row to preview/edit its subject and body. Click **Save Edits** to
   keep your changes for that client only.
3. **Dry run** is checked by default — this lets you preview exactly what would
   be sent (and to whom) without actually sending anything.
4. Uncheck dry run and click **Send All** to send for real. Each email attaches
   that client's report from `output/` if one was generated. You'll get a
   confirmation prompt first since this action can't be undone.
5. Every run (qualifying count, sent/failed counts, filters used) is logged to
   `run_history.json` for audit purposes.

## Customizing the email template

Edit the `DEFAULT_TEMPLATE` dict in `email_generator.py`, or save additional
templates at runtime (the `save_template()` function is ready to be wired to a
"Save as new template" button if you want multiple templates — e.g. "First
Reminder" vs "Final Notice" — let me know and I'll add that UI).

Placeholders currently supported: `{client_name}`, `{greeting_name}`, `{amount}`,
`{report_date}`. `{greeting_name}` resolves to the matched POC name if one exists
for that client, otherwise the company name — use it in the "Dear ..." line.
Add more by including the field name in the template and passing it into
`fields` in `build_email_drafts()`.

## Packaging as a Windows .exe (optional)

Once you've tested it with `python app.py`, you can package it as a single
executable so non-technical staff can just double-click it:

```
pip install pyinstaller
pyinstaller --onefile --windowed app.py
```

The `.exe` will appear in `dist/app.exe`.

## Known limitations / things to confirm with you

- **Outlook COM automation only works on Windows** with Outlook installed —
  I could not test the actual send step in this environment, only the dry-run
  logic. Please test a dry run first, then try sending to yourself before
  using it on real clients.
- **Year breakdown source**: defaults to Issue Date; switch to Comm. Date in
  the dropdown on tab 1 if that better matches how Finance recognizes balances.
- Threshold currently applies to each client's **Grand Total** (summed across
  all departments and years) — as originally specified. Let me know if you'd
  instead want it applied per department (e.g. flag a client if any single
  department alone exceeds the threshold).
- Email matching is by **exact client name** between the balance sheet and the
  lookup file. If names differ slightly (e.g. "LTD" vs "LIMITED"), those rows
  will show as "missing email" for manual entry — I can add fuzzy matching if
  this turns out to be common in your data.
