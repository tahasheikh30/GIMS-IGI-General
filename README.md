# GIMS — General Insurance Management System

A desktop tool for the Finance department, built around a persistent sidebar
with two collapsible menus, each expanding to its own set of sections:

- **General Operations** — a collapsible menu with three sub-items:
  - **Outstanding Premium** — fully built: upload the client balance sheet,
    set a threshold, generate per-client Excel PivotTable reports, and
    draft/send reminder emails through your own Outlook account.
  - **Loss Ratio** — fully built: calculate Loss Ratio (Claims Paid ÷ Gross
    Premium) for one company or a group of companies from their own Excel
    files, with a styled, color-flagged summary report.
  - **Renewal** — placeholder, reserved for a future module.
- **Coinsurance** — a collapsible menu with three sub-items:
  - **Premium/Claims Email Automation** — fully built: reads a co-insurance
    outstanding-position export, generates one two-sheet Excel statement
    (Premium Receivable / Claims Receivable) per Client + Company group, and
    drafts/sends the notification emails through your own Outlook account.
    Structured as the same three-tab workflow as Outstanding Premium
    (Upload & Validate → Client + Company Groups → Review & Send Emails).
  - **Premium Register** — fully built: matches a co-insurance leader's own
    policy numbers against IGI's Premium Register using fuzzy digit-run
    matching, pulls in the matched IGI Policy Number, and flags any
    unmatched rows (highlighted red) for manual entry.
  - **Claims Register** — fully built: matches the Claim Intimation Register
    against IGI's own Premium Register to recover each claim's coinsurer(s)
    (a policy shared by several coinsurers fans a claim out to each of
    their files), then writes one Excel workbook per coinsurer plus a
    separate "Unmatched Claims" workbook for manual review, and drafts/
    sends a recovery-statement email per coinsurer file. Same three-tab
    structure as Premium/Claims Email Automation (Upload & Match →
    Insurance Company Files → Review & Send Emails), with an optional
    date filter to scope a run to a specific period.

## What's included

| File | Purpose |
|---|---|
| `app.py` | Main desktop GUI (Tkinter) — run this |
| `assets/gims_logo_cropped.png` | GIMS logo shown top-left in the header |
| `assets/company_logo_cropped.png` | Company logo shown top-right in the header |
| `data_processor.py` | Loads, validates, and aggregates the raw balance sheet |
| `perf_utils.py` | Background-thread runner for Tkinter; a fast Excel reader (`python-calamine` when installed, ~7x faster than openpyxl's own reader on a real 500,000-row file, falling back to openpyxl's streaming reader otherwise); `StyleIndexCache`, which makes per-cell font/border/fill/number_format application cheap by reusing openpyxl's own internal style index instead of re-hashing the same style on every cell; a visual progress bar widget. Keeps the app responsive on huge (500k+ row) source files and fast when generating Excel output. See `PERFORMANCE_NOTES.md`. |
| `report_generator.py` | Writes each client's raw transaction data into its own Excel file; also generates every client's static (non-native-pivot) report in parallel across CPU cores when Excel isn't available and the batch is large enough |
| `pivot_native_generator.py` | Drives Excel via COM to build a genuine, interactive PivotTable in each report (falls back to a static formatted table if Excel isn't available) |
| `email_generator.py` | Fills the email template per client; template stored in `templates.json` |
| `outlook_sender.py` | Sends via Outlook COM automation (your own signed-in account) |
| `loss_ratio_processor.py` | Reads/validates Loss Ratio input files, normalizes column-name variations, aggregates per company |
| `loss_ratio_generator.py` | Builds the styled Loss Ratio Summary Excel output |
| `receivable_processor.py` | Loads, validates, and groups the Premium & Claims Receivable master file (auto-detects the header row and either amount-column layout) |
| `receivable_generator.py` | Writes each insurance company's three-sheet (Summary w/ charts, Premium Receivable, Claims Receivable) Excel statement, covering every client for that company; also generates every company's workbook in parallel across CPU cores when the batch is large enough |
| `receivable_email.py` | Drafts the statement email per company file and sends via `outlook_sender` |
| `receivable_automation.py` | `load_email_lookup_from_file()` (used by the GUI, reads a separate lookup file) plus a standalone `run_pipeline()` for command-line / scripted runs of the whole flow |
| `receivable_run_log.py` | Writes the Receivable module's audit trail to `receivable_run_history.json` |
| `run_log.py` | Writes an audit trail to `run_history.json` (Outstanding Premium module) |
| `premium_register_processor.py` | Loads the leader's policy file and IGI's Premium Register (auto-detecting header rows/sheets, handling legacy `.xls` including files split across multiple 65,536-row sheets), and fuzzy-matches policy numbers between them |
| `premium_register_generator.py` | Writes the matched-output Excel file, with unmatched rows highlighted for manual entry |
| `claims_register_processor.py` | Loads the Claim Intimation Register and IGI's Premium Register (merging legacy `.xls` files split across multiple 65,536-row sheets), fuzzy-matches each claim to its policy, and fans it out to every coinsurer sharing that policy |
| `claims_register_generator.py` | Writes each coinsurer's two-sheet (Summary w/ charts, Claims Data) Excel workbook plus a shared "Unmatched Claims" workbook; generates every coinsurer's workbook in parallel across CPU cores when the batch is large enough |
| `claims_register_email.py` | Drafts the recovery-statement email per coinsurer file and sends via `outlook_sender` |
| `claims_register_run_log.py` | Writes the Claims Register module's audit trail to `claims_register_run_history.json` |

## Requirements

- Windows PC with **Microsoft Outlook desktop app** installed and signed in
  (emails send through your own mailbox — no separate login needed)
- **Microsoft Excel desktop app installed** — needed to build genuine, interactive
  PivotTables (real field list, filters, drill-down) in each report. Without Excel,
  the app still works but falls back to a static formatted table that looks similar
  but can't be interacted with like a real pivot.
- Python 3.9+
- **Pillow** (for the header logos — falls back to a plain drawn "G" badge on
  the left if it isn't installed or the `assets/` files are missing; the
  rest of the app is unaffected either way)

Install dependencies:
```
pip install -r requirements.txt
```

## Running it

```
python app.py
```

On launch you'll see a fixed header with the GIMS wordmark, and a left
sidebar with two collapsible menus: **General Operations** (click it to
expand/collapse its three sub-items: **Outstanding Premium**, **Loss
Ratio**, **Renewal**) and **Coinsurance** (click it to expand/collapse its
own three sub-items: **Premium/Claims Email Automation**, **Premium
Register**, **Claims Register**). Both menus start expanded, and each can
be collapsed independently without affecting the other. The active section
is highlighted with an accent bar and bold text. All the functionality
below (Steps 1–3) lives inside **Outstanding Premium**, under General
Operations — click it in the sidebar if it isn't already selected. Renewal
is a placeholder reserved for a future module; Claims Register is fully
built (see its own section below).

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

## Coinsurance > Premium/Claims Email Automation module

Click **Coinsurance** in the sidebar to expand it, then click
**Premium/Claims Email Automation**. The page uses the same three-tab
structure as Outstanding Premium:

- **1. Upload & Validate** — source file, sheet names, validation options,
  and an optional Client/Company filter, plus an Activity Log.
- **2. Insurance Company Files** — one row per output file (one per
  insurance company, with every "OFC"-named company pooled into a single
  shared file — see below), showing its client count and Premium/Claims
  Receivable totals, plus the "Generate Excel Reports" / "Open Output
  Folder" buttons.
- **3. Review & Send Emails** — draft, preview/edit, and send the
  statement emails (one per company file), with a shared "Your Signature"
  box, dry run toggle, and per-row Send/Save Edits buttons — identical
  interaction pattern to Outstanding Premium's tab 3.

**What it does:** reads a co-insurance outstanding-position export and
groups every transaction row by **Company** (the co-insurance partner,
e.g. "ADAMJEE INSURANCE COMPANY LTD" — the company that owes IGI its
share of a client's premium/claims), writing **one Excel workbook per
insurance company**, covering every client that company is associated
with. Every company whose name contains "OFC" (case-insensitive) is
pulled out of the normal per-company grouping and pooled together into a
single shared file named `OFC_Receivable_{date}.xlsx` instead — so if you
have several differently-named OFC branches/entities in your data, they
all land in one combined file rather than each getting their own.

Each output workbook has exactly three sheets:
- **"Summary"** (opens first) — a redesigned dashboard-style sheet: a
  navy title banner naming the company and generation date, three KPI
  cards (Premium Receivable / Claims Receivable / Total Receivable), a
  Premium-vs-Claims pie chart, and the full per-client breakdown table
  with zebra-striped rows, a bold Total row, and a bar chart comparing
  Premium and Claims Receivable across every client in that file. Every
  number — the KPI cards, the pie chart's source table, and the Total
  row — is a live Excel formula referencing the client table on the same
  sheet, not a hardcoded value, so it all recalculates correctly if you
  ever edit a number by hand. Both charts are genuine embedded Excel
  chart objects (built via openpyxl), not static pictures — you can
  click into them and adjust like any chart you built by hand.
- **"Premium Receivable"** — every row (across every client in the file)
  with a non-zero Premium Receivable amount
- **"Claims Receivable"** — every row with a non-zero Claim Receivable
  amount (still written with headers only if the file has none)

Every other original column (Class, Leader Ref. No, Code, Client Name,
Company, Date, Type, etc.) is preserved exactly as it appears in the
source file on the two data sheets — nothing is renamed or reformatted
beyond splitting rows by category and number-formatting the relevant
amount column. Client Name stays visible on every row, since a file now
covers many clients at once (and, for the pooled OFC file, several
different real companies too).

### Tab 1 — Upload & Validate
- **Master workbook** — the co-insurance export, a single-sheet `.xlsx`
  file — there's no sheet name to specify, it's read directly. The real
  header row is located automatically even though the source file has a
  few title/banner rows above it (report title, "As at ..." date, a
  blank row) — you don't need to trim those first.
- **Look up recipient emails from a lookup file** — uncheck to skip email
  lookup entirely (every draft will show "Missing email" and can still be
  reviewed/sent manually on tab 3).
- **Email Lookup File** — a *separate* `.xlsx` or `.csv` file you browse
  to (not part of the master workbook), with a Company-name column
  (matched by "company", "client", or "participant" in the header — falls
  back to the first column if none match), an **Email** column, and an
  optional POC/contact-person column. Rows are matched by exact **Company
  Name**, the same "Company" field used on the transaction data, so
  recipient emails are looked up by co-insurance company rather than by
  client. For the pooled OFC file, the lookup key is literally `OFC` —
  add a row with that exact name if you want that file's email to
  resolve automatically; otherwise it's flagged "Missing email" for
  manual entry on tab 3, same as any other unmatched company. A company
  can list several emails in one cell (comma or semicolon separated) —
  all of them are used as Outlook "To" recipients. If a company's
  contact-person column is present, that name is used in the email
  greeting instead of the company name. Leave this blank to skip email
  lookup — this mirrors how the Outstanding Premium module's own Email
  Lookup File works.
- **Output folder** — configurable; created automatically if it doesn't
  exist.
- **Leader Ref Number** is the *other* co-insurance company's own policy
  reference, so it's only required (and only checked) on **Premium
  Receivable** rows — where IGI is the follower being billed against a
  policy the other company leads. On **Claims Receivable** rows IGI is
  the primary payer being reimbursed, so there's no such reference to
  expect from the other side; a blank Leader Ref Number there is normal
  and is never flagged or excluded. (Confirmed against a real sample
  file: all 460 claims-only rows had a blank Leader Ref Number, and
  every single premium-only row had one — this is structural, not messy
  data.) This is handled automatically now — there's no checkbox for it.
- Click **Load & Validate Data** to load the file — the Activity Log shows
  the detected column layout and a per-issue-type validation count.
- **Filters** — optional Client filter and/or Company filter (comma-
  separated exact names). Click **Analyze / Build Groups** to apply them
  and populate tab 2.

### Tab 2 — Insurance Company Files
Shows one row per output file — each insurance company (plus the pooled
"OFC" row, if any OFC-named companies are present) — with its client
count and Premium/Claims Receivable totals. Click **Generate Excel
Reports** to write one workbook per row (named per the pattern below)
into the output folder set on tab 1, or **Open Output Folder** to browse
what's been generated.

### Tab 3 — Review & Send Emails
1. **Your Signature** — shared with the Outstanding Premium module (same
   `signature.json`), so it only needs to be set once across both modules.
2. Click **Draft Emails** to generate one statement email per company file.
3. Click any row to preview/edit its recipient, subject, and body. Click
   **Save Edits to Selected Email** to keep changes for that file only.
4. **Dry run** is checked by default — preview exactly what would be sent
   without actually sending.
5. Uncheck dry run and click **Send All** (or **Send This Email** for a
   single row) to send for real via Outlook. Each email attaches that
   file's generated workbook if one exists. Every run is logged to
   `receivable_run_history.json`, the same way the Outstanding Premium
   module logs to `run_history.json`.

**Email subject/body:**
```
Subject: Premium & Claims Receivable Statement – {Company Name} – {Date}
```
The body is a short cover note referencing the attached statement, how
many clients it covers, and its total Premium/Claims Receivable.

**Filename format:** `{CompanyName}_Receivable_{DDMMYYYY}.xlsx` (or
`OFC_Receivable_{DDMMYYYY}.xlsx` for the pooled file), with spaces and
special characters stripped from the name for filesystem/attachment
safety.

**Column layout flexibility:** the module recognizes two possible source
layouts automatically — (a) separate "Premium Receivable" and "Claim
Receivable" amount columns on every row (the confirmed real format), or
(b) a single "Amount" column plus a "Type"/"Category" column naming which
category each row belongs to (a fallback in case a different export
format is used later). Column headers are matched case-insensitively
against a small alias list in `receivable_processor.py` — if a future
export uses different header text than what's listed there, add it to
`ALIASES` rather than renaming columns in the source file.

**Known limitation:** like the Outstanding Premium module, sending only
works on Windows with Outlook installed and signed in — this could only be
verified here in dry-run mode. Please test a dry run first, then send to
yourself before using it on real recipients.

## Coinsurance > Premium Register module

Click **Coinsurance** in the sidebar to expand it, then click
**Premium Register**. This is a single-page workflow (no multi-tab
notebook, since there's no grouping/email step) with three numbered
cards: Input Files, Run Matching, and Processing Log.

**What it does:** matches every policy in a co-insurance **leader's** own
file (their own policy number, e.g. in a column called `NUMS`) against
IGI's own **Premium Register** (which records the leader's policy number
back to us in a column such as `Leader#`, alongside our own `Doc Ref #` /
IGI Policy Number), and pulls in the matched IGI Policy Number wherever a
confident match is found. Rows that can't be matched confidently are
flagged so they can be entered into the Premium Register manually,
instead of guessing — a wrong policy number would be worse than none,
since someone would act on it.

**Matching is fuzzy, not exact string equality**, since the leader's own
policy-number format essentially never matches our `Leader#` column
character-for-character (e.g. `1/114591/25585/11006` vs. our own
`2691025585/12/2023`). In practice the two numbers reliably share at
least one distinctive digit run — the policy/endorsement's own numeric
core — even when everything else around it (separators, prefixes,
date-like suffixes) differs. Matching works in two stages:
1. **Blocking** — every register row is indexed by its own digit-run
   substrings, so a leader value only needs to be compared against a
   small set of real candidates rather than the entire register (fast
   even against a ~24k-row register).
2. **Scoring** — among those candidates, whichever register row shares
   the longest *complete* digit run (the query's own policy number as a
   whole token, not a fragment) with the query wins; a generic
   digits-only similarity ratio is used only as a tiebreaker. A row with
   no candidates, or whose best candidate doesn't clear the acceptance
   threshold, is left **Not Matched** rather than forced onto a poor guess.

1. **Input Files**
   - **Leader's Policy File** — the co-insurance leader's own file
     (`.xlsx` or `.xls`). Only the policy-number column is identified
     (matched via an alias list, e.g. `NUMS`, `Policy No`, `Leader Ref
     No`) — every other column is left exactly as-is and carried through
     to the output untouched.
   - A leader's file is very often a **multi-sheet settlement workbook**
     (e.g. a `Summary` tab, a `Premium Payable` tab, `Claim Rec`/`Prem
     Rec` tabs, and one or more dated snapshot tabs) rather than one
     plain sheet — and more than one of those tabs can independently
     contain a `POLICY NO.`-style column, so column content alone isn't
     always enough to reliably pick the right one. The right sheet is
     found **by its own name first** (matched loosely — ignoring spacing
     and punctuation, so `"Premium Payable"`, `"PremPayable"`, and
     `"Prem. Payable - AIC"` are all recognized the same way), and only
     falls back to scanning every sheet's column content if no sheet
     name is a clear match.
   - **IGI Premium Register** — our own register (`.xlsx` or `.xls`).
     The `Leader Ref Number` column (e.g. `Leader#`) and `IGI Policy
     Number` column (e.g. `Doc Ref #`) are identified the same way. A
     Crystal Reports export quirk — some register cells carry the
     column's own caption baked into the value (e.g. `"Leader Ref. No.
     PL-0720-200501-F01-000053"`) — is automatically stripped.
   - Both files' real header row is auto-detected even if there are
     banner/title rows above it, and column-header punctuation (e.g.
     `"POLICY NO."` vs `"Policy No"`) doesn't need to match exactly —
     trailing periods are ignored automatically, since real exports from
     different companies punctuate the same header inconsistently.
   - Legacy `.xls` files are read via the `xlrd` package if installed;
     otherwise the app falls back to converting the file with a local
     LibreOffice install, if one is present on the machine.
2. **Run Matching** — click **Run Matching**. The Processing Log shows
   how many rows were loaded from each file and which columns were
   detected, then the match count and percentage. A **Save As** dialog
   then lets you choose where to write the output `.xlsx`
   (default filename: `{LeaderFileName}_PolicyMatched_{DDMMYYYY}.xlsx`).
3. **Output** — every original column from the leader's file, plus four
   new columns: `IGI Policy Number`, `Matched Leader Ref (Register)`,
   `Match Score`, and `Match Status`. Any row with `Match Status` =
   "Not Matched" has its **entire row highlighted light red**, so it's
   immediately visible which rows still need to be entered into the
   Premium Register by hand. A small note below the table summarizes the
   overall match rate.

## Coinsurance > Claims Register module

Click **Coinsurance** in the sidebar, then expand **Claims** and click
**Claims Intimation**. Same three-tab structure as Premium/Claims Email
Automation (Upload & Match → Insurance Company Files → Review & Send
Emails).

**What it does:** matches the Claim Intimation Register (Finance's own
log of claims as they come in — sparsely filled in on the coinsurer side,
since that column is manually entered) against IGI's Premium Register to
recover which coinsurer(s) share each claim's underlying policy. A policy
can be shared by several coinsurers at once (confirmed against real data:
one policy had 11 separate coinsurer rows) — matching a claim's own
`Policy/PMD #` against the register's `Base Doc No.` finds the right
policy, then every coinsurer sharing that same Base Doc No. is pulled in
at once, so a single claim can fan out into several coinsurers' own
recovery files, each showing only that coinsurer's own Share % and their
proportional cut of Claim Paid / Claim O/S.

Matching uses the same fuzzy digit-run approach as Premium Register
(blocking + longest-complete-run scoring), tuned for this module's own
reference format — see `claims_register_processor.py`'s own docstrings
for the details, including a documented, deliberate blocking-key design
that keeps matching fast even against a 200,000+ row register.

1. **Upload & Match**
   - **Claim Intimation Register** and **IGI Premium Register** — both
     `.xlsx`/`.xls`, real header row auto-detected the same way as every
     other GIMS module. Legacy `.xls` files that are split across several
     sheets at Excel's old 65,536-row-per-sheet limit are automatically
     detected and merged back into one table (confirmed against a real
     204,220-row register split across 4 sheets) — the header only needs
     to exist on the first sheet.
   - **Date Filter (optional)** — scope a run to claims whose Intimation
     Date or Loss Date falls within a chosen range, without needing to
     trim the source files first. The Premium Register itself is never
     date-filtered (it's only used as a lookup table), so it can still
     cover a much wider period than the claims being matched.
   - Click **Load & Match Claims**. The Activity Log reports how many
     claims matched to a coinsurer, how many claim-coinsurer pairs that
     produced across how many coinsurer files, and — for anything that
     couldn't be matched — a breakdown of why (no matching policy found,
     or the matched policy has no coinsurer on it at all, i.e. 100% IGI).
2. **Insurance Company Files** — one row per coinsurer (plus a shared
   pooled "OFC" row for every coinsurer whose name contains "OFC"), with
   its claim count and total Share of Claim Paid / O/S. Click **Generate
   Excel Reports** to write one workbook per coinsurer, plus a separate
   **"Unmatched Claims"** workbook for anything that needs manual review.
   Each coinsurer workbook has two sheets: **"Summary"** (a dashboard —
   KPI cards, a Paid-vs-Outstanding pie, and a per-insured bar chart, all
   driven by live Excel formulas) and **"Claims Data"** (every matched
   claim row for that coinsurer, original claim columns preserved plus
   the fan-out columns: Coinsurer Name, Coinsurer Share %, Coinsurer
   Share of Claim Paid/O/S, Matched Policy Doc No., Match Score). The
   Unmatched Claims workbook highlights every row red (the whole sheet
   needs review) with a Reason column explaining why.
3. **Review & Send Emails** — draft, preview/edit, and send a claims
   recovery-statement email per coinsurer file, with the same shared
   signature box, dry-run toggle, and per-row Send/Save Edits controls as
   every other GIMS email workflow.

**Performance note:** both matching and Excel generation for this module
were specifically profiled and optimized against real, large production
files (a 204,220-row register and a 27,976-claim register) — see
`PERFORMANCE_NOTES.md`'s Part 4 for the full detail, including the one
disclosed limitation (a single coinsurer with an extremely large claim
count is still bound by openpyxl's own per-cell cost once its data sheet
runs into the millions of cells, since that sheet's workbook also needs
real embedded charts).

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
