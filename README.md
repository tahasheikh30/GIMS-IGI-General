# GIMS — General Insurance Management System

A desktop tool for the Finance department, built around a persistent sidebar
with three modules:
- **Outstanding Premium** — fully built: upload the client balance sheet, set
  a threshold, generate per-client Excel PivotTable reports, and draft/send
  reminder emails through your own Outlook account.
- **Loss Ratio** — fully built: calculate Loss Ratio (Claims Paid ÷ Gross
  Premium) for one company or a group of companies from their own Excel
  files, with a styled, color-flagged summary report.
- **Renewal** — placeholder, reserved for a future module.

## What's included

| File | Purpose |
|---|---|
| `app.py` | Main desktop GUI (Tkinter) — run this |
| `data_processor.py` | Loads, validates, and aggregates the raw balance sheet |
| `report_generator.py` | Writes each client's raw transaction data into its own Excel file |
| `pivot_native_generator.py` | Drives Excel via COM to build a genuine, interactive PivotTable in each report (falls back to a static formatted table if Excel isn't available) |
| `email_generator.py` | Fills the email template per client; template stored in `templates.json` |
| `outlook_sender.py` | Sends via Outlook COM automation (your own signed-in account) |
| `loss_ratio_processor.py` | Reads/validates Loss Ratio input files, normalizes column-name variations, aggregates per company |
| `loss_ratio_generator.py` | Builds the styled Loss Ratio Summary Excel output |
| `run_log.py` | Writes an audit trail to `run_history.json` |

## Requirements

- Windows PC with **Microsoft Outlook desktop app** installed and signed in
  (emails send through your own mailbox — no separate login needed)
- **Microsoft Excel desktop app installed** — needed to build genuine, interactive
  PivotTables (real field list, filters, drill-down) in each report. Without Excel,
  the app still works but falls back to a static formatted table that looks similar
  but can't be interacted with like a real pivot.
- Python 3.9+

Install dependencies:
```
pip install -r requirements.txt
```

## Running it

```
python app.py
```

On launch you'll see a fixed header with the GIMS wordmark, and a left
sidebar with three sections: **Outstanding Premium**, **Loss Ratio**, and
**Renewal**. The active section is highlighted with an accent bar and bold
text. All the functionality below (Steps 1–3) lives inside **Outstanding
Premium** — click it in the sidebar if it isn't already selected. The other
two sections are placeholders reserved for future modules.

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
contains **two sheets**:
- **"Outstanding Balance"** — a genuine, interactive **Excel PivotTable** (built
  by driving Excel itself via COM automation), with Department in rows, Year in
  columns, and Sum of O/s as the value — fully drillable, filterable, and
  re-configurable using Excel's normal PivotTable Field List, exactly as if you'd
  built it by hand via Insert → PivotTable. The header cells read "Department"
  and "Year" (rather than Excel's generic "Row Labels"/"Column Labels" default).
- **"Raw Data"** — every original transaction row for that client from the source
  sheet (this is also the PivotTable's source range, so refreshing the pivot
  in Excel re-pulls from this sheet)

If Microsoft Excel isn't installed on the machine generating the reports, this
step automatically falls back to a static formatted table that *looks* like a
pivot but has no interactive functionality — you'll see a note in the completion
popup if that happens.

### Step 3 — Review & Send Emails
1. At the top of this tab, there's a **"Your Signature"** box — whatever you
   type there appears after "Regards," at the end of every email. It's saved
   automatically the moment you click **Draft Emails** (or explicitly via
   **Save Signature**), and reloads automatically the next time you open the
   app, so you only need to set it once.
2. Click **Draft Emails** to generate a personalized email per qualifying client
   from the template in `email_generator.py` / `templates.json`.
3. Click any row to preview/edit its subject and body. Click **Save Edits** to
   keep your changes for that client only.
4. **Dry run** is checked by default — this lets you preview exactly what would
   be sent (and to whom) without actually sending anything.
5. Uncheck dry run and click **Send All** to send for real. Each email attaches
   that client's report from `output/` if one was generated. You'll get a
   confirmation prompt first since this action can't be undone.
6. Every run (qualifying count, sent/failed counts, filters used) is logged to
   `run_history.json` for audit purposes.

## Loss Ratio module

Click **Loss Ratio** in the sidebar. The page is organized into four
numbered cards: Mode & Report Name, Input Files, Threshold & Generate, and
Processing Log.

**Formula (verified against a real IGI sample file and its source report):**
```
Loss Ratio % = (Incurred Loss + Commission Accrual) / Gross Premium Written × 100
```
This matches your system's own "Loss Ratio (B+C/A)" definition, where A = Gross
Premium Written, B = Incurred Loss, C = Commission Accrual.

**Accepted input files: Excel (`.xlsx`) or PDF.**
- **Excel**: expects `Gross Premium Written`, `Incurred Loss`, and optionally
  `Commission Accrual` columns (see alias list below). Multi-year files with
  a "Client Total"/"Grand Total" row use that row directly rather than
  summing every row, to avoid double-counting.
- **PDF**: reads IGI's own "Loss Ratio Gross Premium Written" Crystal
  Reports export format directly — the same layout as your system's PDF
  print-out. It finds each client's "Client Total" line and pulls the
  Gross Premium Written / Incurred Loss / Commission Accrual totals from
  it. **A single PDF covering a client range produces one row per client**,
  so one multi-client PDF export can populate an entire Group report by
  itself — you don't need to split it into per-company files first. The
  report's own overall "Grand Total" line (the sum across every client in
  the PDF) is correctly excluded from the row list, since it isn't a company.

1. **Choose a mode:**
   - **Individual Company** — browse to a single file for one company.
   - **Group of Companies** — click **Add File(s)...** to select one or
     more files (Excel and PDF can be mixed freely in the same batch).
     Select a file in the list and click **Remove Selected** to take it
     out before processing.
2. **Group / Report Name** — text shown in the green banner at the top of
   the output (e.g. "POPULAR GROUP"). Leave blank for a generic default.
3. **Set the red highlight threshold** (default 100%) — any company's Loss
   Ratio above this is flagged red. Anything ≥70% is always flagged amber
   (this coloring is an addition on top of your existing template, which
   doesn't otherwise style individual rows).
4. Click **Calculate Loss Ratio**. The Processing Log shows each file's
   result or error, tagged with the source filename — a problem in one
   file (missing column, non-numeric value, unreadable file, unrecognized
   PDF layout) doesn't stop the others from processing. If any files
   errored, you'll be asked whether to continue with just the successful
   ones.
5. A **Save As** dialog lets you choose where to save the output `.xlsx`.

**Expected Excel column names** (case-insensitive, naming variations like
"Gross Premium Written ( A )" or "GP" are automatically recognized):
- Required: `Gross Premium Written`, `Incurred Loss`
- Optional: `Commission Accrual` (treated as 0 if absent — flagged in the
  Processing Log so it's never silently wrong), `Company Name` (falls back
  to the filename if absent)

A leading account-code prefix in a client name cell (e.g. "1120106182 SAYA
WEAVING MILLS...") is stripped automatically, in both Excel and PDF inputs.

**Output format** replicates your real "Group Summary" template exactly:
merged green group-name banner, `CLIENT | GP | CLAIM | COMMISSION | LR`
header row, and — critically — **live Excel formulas** for the LR column
(`=SUM(D:E)/C%`) and the Total row, not hardcoded numbers, so the sheet
recalculates correctly if you ever edit a value by hand. The Total row's LR
cell is highlighted yellow, matching your existing template.

I validated this two ways: reconstructing all 11 companies from your real
`Popular_Group_Loss_Ratio.xlsx` sample as separate input files (every value
matched to 6 decimal places, including the 105.408432% group total), and
parsing your real `Loss_Ratio.pdf` sample directly (extracted GP, Loss, and
Commission matched the report's own printed 10.73% total ratio exactly).

If any of your real files use column headers, or a PDF layout, not covered
here, send me an example and I'll extend the matching.

**Output:** a "Loss Ratio Summary" sheet with one row per company, a bold
Total/Group row in Group mode, and cells color-coded red/amber based on the
Loss Ratio thresholds — styled consistently with the Outstanding Premium
reports (same fonts, header colors, borders, number formats).

## Customizing the email template

Edit the `DEFAULT_TEMPLATE` dict in `email_generator.py`, or save additional
templates at runtime (the `save_template()` function is ready to be wired to a
"Save as new template" button if you want multiple templates — e.g. "First
Reminder" vs "Final Notice" — let me know and I'll add that UI).

Placeholders currently supported: `{client_name}`, `{greeting_name}`, `{amount}`,
`{report_date}`, `{signature}`. `{greeting_name}` resolves to the matched POC
name if one exists for that client, otherwise the company name — use it in the
"Dear ..." line. `{signature}` resolves to whatever is saved in the "Your
Signature" box on tab 3 (falls back to a generic Finance Department signature
if left blank). Add more by including the field name in the template and
passing it into `fields` in `build_email_drafts()`.

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
