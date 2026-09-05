# GIMS — General Insurance Management System

A Windows desktop tool (Tkinter) for the Finance department at IGI General
Insurance, built around a persistent sidebar with three top-level collapsible
menus, each expanding into its own sections:

- **General Operations**
  - **Outstanding Premium** — upload the client balance sheet, set a
    threshold, generate per-client Excel PivotTable reports, and draft/send
    reminder emails through your own Outlook account.
  - **Loss Ratio** — calculate Loss Ratio (Incurred Loss + Commission
    Accrual) ÷ Gross Premium Written for one company or a group of
    companies, from Excel or PDF source files, with a styled, color-flagged
    summary report.
- **Finance**
  - **Management Accounts**
    - **UPR Automation** — computes Closing Unearned Premium Revenue from
      the Premium Register using the 1/24 (most departments) and 1/6
      (Marine/Travel) amortization methods for a selected 12-month period.
    - **Reinsurance Automation** — the same period/bucketing/amortization
      method as UPR, applied instead to Total Reinsurance Ceded (Treaty +
      Local + Foreign Premium) from a separate register per department, to
      produce Closing Prepaid Reinsurance Premium Ceded.
    - **Deferred Commission** — matches three commission/service-charge
      ledgers to the Premium Register by policy reference, sums them into a
      Total Commission per policy, and amortizes it the same UPR way to
      produce Closing Deferred Commission Expense.
  - **Withholding Tax**
    - **Tax Payer Check** — matches the General Ledger against an NTN Locator, a
      Status file, and FBR's own multi-sheet Active Taxpayer List (ATL) to
      resolve each payee's NTN/CNIC, tax status, and Filer/Non-Filer
      standing per ATL sheet.
- **Coinsurance**
  - **Premium/Claims Email Automation** — reads a co-insurance
    outstanding-position export, generates one Excel statement per Company
    (Premium Receivable / Claims Receivable), and drafts/sends the
    notification emails.
  - **Premium**
    - **Policy Check** — matches a co-insurance leader's own policy numbers
      against IGI's Premium Register using fuzzy digit-run + branch/class
      matching, pulls in the matched IGI Policy Number, and flags unmatched
      rows for manual entry.
    - **Premium Register** — splits IGI's own raw Premium Register export
      by coinsurer (every row where Leader = "No") and produces one
      statement workbook + draft email per company.
  - **Claims**
    - **Claims Intimation** (Claims Register) — matches the Claim
      Intimation Register against IGI's own Premium Register to recover
      each claim's coinsurer(s) — a policy shared by several coinsurers
      fans a claim out to each of their files — then writes one workbook
      per coinsurer plus a separate "Unmatched Claims" workbook, and
      drafts/sends a recovery-statement email per coinsurer.
    - **Claims Check** — reuses Policy Check's own matching engine to
      confirm whether a co-insurance leader's claimed policy actually
      exists in IGI's own Premium Register (with a second, strict Doc-No
      matcher auto-selected for purely numeric reference formats).

Every module that groups source rows by company (Claims Register, Premium
Register split, Premium/Claims Email Automation) shares one deterministic
merge/alias engine (`company_alias_resolver.py`) and a fixed canonical
filename list for known companies, so a coinsurer never ends up split across
two differently-spelled output files by accident.

Every module that produces a Takaful vs. Conventional distinction (Claims
Intimation, Premium Register split) has a **Business Type** radio button
that works "Takaful" into the output filenames, email subject, and email
body for that run — a Conventional run (the default) is worded exactly as
if the toggle didn't exist.

## What's included

### Core app / shared infrastructure
| File | Purpose |
|---|---|
| `app.py` | Main desktop GUI (Tkinter) — run this |
| `assets/gims_logo_cropped.png`, `assets/company_logo_cropped.png` | Header logos |
| `perf_utils.py` | Background-thread runner for Tkinter; a fast Excel reader (`python-calamine` when installed, ~7x faster than openpyxl's own reader, falling back to openpyxl's streaming reader otherwise); `StyleIndexCache`, which makes per-cell font/border/fill/number_format application cheap by reusing openpyxl's own internal style index instead of re-hashing the same style on every cell; a visual progress bar widget. See `PERFORMANCE_NOTES.md`. |
| `app_paths.py` | Keeps GIMS' own small per-user state (signature, templates, run-history logs) in a local, per-machine folder (`%LOCALAPPDATA%\GIMS`), never inside the OneDrive-shared project folder — this is the fix for "the app corrupts when moved between laptops." See `README_ONEDRIVE_FIX.md`. |
| `safe_io.py` | Atomic (write-to-temp-then-rename) JSON read/write helpers used by every state file, so a mid-write crash or a sync engine reading at the wrong moment can never leave a half-written file behind. |
| `company_alias_resolver.py` | Shared, fully deterministic company-merge logic (leading-token + known-abbreviation rules, no manual-review workbook), canonical output filenames for a fixed list of known companies, and the Co-insurance Companies List / Email Lookup File parser used across Claims Register, Premium Register (split), and Receivable. |

### Outstanding Premium
| File | Purpose |
|---|---|
| `data_processor.py` | Loads, validates, and aggregates the raw balance sheet |
| `report_generator.py` | Writes each client's raw transaction data + pivot into its own Excel file; parallelized across CPU cores on the non-native-pivot fallback path |
| `pivot_native_generator.py` | Drives Excel via COM to build a genuine, interactive PivotTable in each report (falls back to a static table if Excel isn't available) |
| `email_generator.py` | Fills the shared email template per client; signature/templates persisted via `app_paths`/`safe_io` |
| `outlook_sender.py` | Sends via Outlook COM automation |
| `run_log.py` | Audit trail (`run_history.json`) |

### Loss Ratio
| File | Purpose |
|---|---|
| `loss_ratio_processor.py` | Reads/validates Excel or PDF Loss Ratio inputs, normalizes header variations, aggregates per company |
| `loss_ratio_generator.py` | Builds the styled Loss Ratio Summary Excel output with live formulas |

### UPR Automation / Reinsurance Automation / Deferred Commission
| File | Purpose |
|---|---|
| `upr_processor.py` | Loads Premium Register file(s), validates period coverage, buckets by Comm Date (Previous/Jan–Dec/Future Policies), and amortizes Gross Premium (1/24 / 1/6 / zero per department) — the shared engine all three "Management Accounts" modules build on |
| `upr_generator.py` | Writes the UPR working workbook (Summary + a single two-table working sheet, live SUM/reference formulas) |
| `upr_ledger_processor.py` | Parses the "All Ledgers Balance" PDF to auto-fill UPR's own Section 1 premium-income figures |
| `reinsurance_processor.py` | Computes Total Reinsurance Ceded per row (department taken from the file name, since these registers carry no Department column) and reuses `upr_processor`'s period/bucketing/amortization logic unchanged |
| `reinsurance_generator.py` | Writes the Reinsurance working workbook, including a Treaty/FAC/Foreign-FAC breakdown table |
| `deferred_commission_processor.py` | Matches three commission/service-charge ledgers to the Premium Register by Doc Ref # (after splitting a "~"-joined reference), sums into Total Commission, and reuses `upr_processor`'s amortization logic |
| `deferred_commission_generator.py` | Writes the Deferred Commission working workbook, including a Ledger-Balance-PDF reconciliation block |
| `deferred_commission_ledger_processor.py` | Parses the same "All Ledgers Balance" PDF for the five GL accounts this module cares about |

### Withholding Tax — Tax Payer Check
| File | Purpose |
|---|---|
| `withholding_tax_processor.py` | Matches the General Ledger's Payee Name against an NTN Locator (exact match only — no fuzzy matching, since this feeds a real tax rate), resolves NTN/CNIC, matches a Status file, and checks FBR's own multi-sheet Active Taxpayer List (via `python-calamine` directly — a real ATL file can run into the millions of rows) |
| `withholding_tax_generator.py` | Writes the single working sheet, red-highlighting every row that needs manual review (no confident NTN/CNIC could be established), plus a bold Overall Filer Status column |

### Coinsurance — Policy Check / Claims Check
| File | Purpose |
|---|---|
| `premium_register_processor.py` | Loads the leader's policy file and IGI's Premium Register (auto-detecting header rows/sheets, handling legacy `.xls` split across multiple 65,536-row sheets), and fuzzy-matches policy numbers (digit-run blocking + branch/class-code safeguard) |
| `premium_register_generator.py` | Writes the matched-output Excel file, unmatched rows highlighted red |
| `claims_check_processor.py` | Reuses Policy Check's own matching engine for a leader's claims file against the same register columns; auto-switches to a strict, Doc-No-only matcher for a purely numeric reference format (where the normal branch/class safeguard would be a silent no-op) |
| `claims_check_generator.py` | Writes the Claims Check output, unmatched rows highlighted red |

### Coinsurance — Claims Intimation (Claims Register)
| File | Purpose |
|---|---|
| `claims_register_processor.py` | Loads the Claim Intimation Register and IGI's Premium Register, fuzzy-matches each claim to its policy (compound branch/class + serial blocking keys for speed on 200,000+ row registers), and fans a claim out to every coinsurer sharing that policy |
| `claims_register_generator.py` | Writes each coinsurer's two-sheet (Summary dashboard + Claims Data) workbook plus a shared Unmatched Claims workbook; parallelized across CPU cores |
| `claims_register_email.py` | Drafts the recovery-statement email per coinsurer file |
| `claims_register_run_log.py` | Audit trail (`claims_register_run_history.json`) |

### Coinsurance — Premium/Claims Email Automation & Premium Register (split)
| File | Purpose |
|---|---|
| `receivable_processor.py` | Loads, validates, and groups the Premium & Claims Receivable master file (auto-detects header row, sheet, and either dual-amount-column or legacy single-amount-column layout) |
| `receivable_generator.py` | Writes each company's three-sheet (Summary dashboard, Premium Receivable, Claims Receivable) statement; parallelized across CPU cores |
| `receivable_email.py`, `receivable_automation.py`, `receivable_run_log.py` | Drafting/sending, a scripted CLI pipeline, and the audit trail |
| `premium_register_company_processor.py` | Splits IGI's own raw Premium Register export by coinsurer (Leader = "No" rows only, IGI's own share excluded) — accepts either the original Leader-bearing register shape or a "Co Insurance Ceded Register" export shape that has no Leader column at all |
| `premium_register_company_generator.py` | Writes each company's two-sheet (Summary + Premium Data) workbook; parallelized across CPU cores |
| `premium_register_company_email.py`, `premium_register_company_run_log.py` | Drafting/sending and the audit trail |

## Requirements

- Windows PC with **Microsoft Outlook desktop app** installed and signed in
  (emails send through your own mailbox)
- **Microsoft Excel desktop app** — needed for genuine, interactive
  PivotTables in Outstanding Premium reports; without it, that module falls
  back to a static formatted table
- Python 3.9+
- **Pillow** — header logos and the progress-bar state dot (falls back to a
  plain drawn badge/oval if missing)
- **python-calamine** (strongly recommended) — the fast Excel/legacy-`.xls`
  reader every loader in the app prefers; falls back to openpyxl
  automatically if not installed. **Required** (no fallback) for the
  Withholding Tax module's ATL file, since that file can run into the
  millions of rows
- **xlrd** (optional) — reading legacy `.xls` files directly; without it,
  the app falls back to converting via a local LibreOffice install if one
  is present
- **pdfplumber** — required for any PDF input (Loss Ratio's PDF mode, the
  UPR/Deferred Commission Ledger Balance PDF auto-fill)

Install dependencies:
```
pip install -r requirements.txt
```

## Running it

```
python app.py
```

Click a top-level menu in the sidebar (**General Operations**, **Finance**,
**Coinsurance**) to expand/collapse it; every menu — and every nested
submenu inside it (e.g. Finance's own "Management Accounts", Coinsurance's
own "Premium"/"Claims") — starts collapsed, so only the section you're
actually using takes up space. The active page is highlighted with a gold
accent bar and bold white text.

## Module notes

### Outstanding Premium
Three tabs: **Upload & Filter** (balance file, sheet name, Year-basis
choice, email lookup file), **Qualifying Clients** (threshold + Department/
Agent/Branch filters, one row per client meeting the threshold, generate
reports), **Review & Send Emails** (draft/preview/edit/send, shared
signature box, dry-run toggle). Each client's report has a "Raw Data"
sheet plus a genuine, interactive Excel PivotTable (native COM automation)
or a static fallback table if Excel isn't installed.

### Loss Ratio
```
Loss Ratio % = (Incurred Loss + Commission Accrual) / Gross Premium Written × 100
```
Individual or Group mode, Excel and/or PDF input mixed freely, a
configurable red-highlight threshold (amber is always triggered at ≥70%),
and a live-formula output matching the real "Group Summary" template
exactly (merged green banner, `CLIENT | GP | CLAIM | COMMISSION | LR`).

### UPR Automation / Reinsurance Automation / Deferred Commission
All three share one three-tab shape (**Register/Files & Period** →
**Opening Balance & Calculate** → **Generate & Save**) and the same
underlying 12-month bucketing/amortization engine: a department's own
premium (or reinsurance-ceded, or commission) figure is spread across
Previous / Jan–Dec / Future Policies buckets by Comm Date, then amortized
1/24 (most departments), 1/6 (Marine/Travel), or not at all (BIMA). Each
module raises a clear error rather than silently understating a figure if
the uploaded register(s) don't actually cover the full selected 12-month
period. UPR and Deferred Commission can both auto-fill part of their
Section 1 / Summary figures from an "All Ledgers Balance" PDF upload.

### Withholding Tax — Tax Payer Check
Single-page workflow: General Ledger, NTN Locator, Status file, and ATL
file (any of the first three can be separate files or separate sheets in
the same workbook — each is auto-detected independently). Matching is
**exact only** end to end (name normalization closes formatting gaps, but
never approximates a match) — any payee that can't be confidently
identified is highlighted red for manual review rather than guessed at,
since this feeds an actual withholding tax rate.

### Coinsurance — Policy Check / Claims Check
Policy Check matches a leader's own policy numbers against IGI's Premium
Register via digit-run blocking + longest-complete-run scoring, with a
branch/class-code safeguard (`MIN_ALPHA_RATIO`) that stops two unrelated
transactions from different departments matching just because their
serial numbers happen to collide (every department restarts numbering at
1). Claims Check reuses this same engine for a leader's own claims file,
automatically switching to a stricter Doc-No-only matcher when the
reference format is purely numeric (a shape where the normal safeguard
would otherwise be a silent no-op).

### Coinsurance — Claims Intimation (Claims Register)
Matches the Claim Intimation Register's own `Policy/PMD #` against the
Premium Register's `Base Doc No.` (falling back to `Document No.`), then
fans a matched claim out to every coinsurer sharing that same policy (a
real policy can have 11+ coinsurers on it). Supports an optional Date
Filter (Intimation Date or Loss Date) and Unit/Client Name/Department
multi-select filters, both applied before matching so everything
downstream only ever reflects the scoped run. One workbook per coinsurer
(OFC-named companies pooled) plus a shared "Unmatched Claims" workbook.

### Coinsurance — Premium/Claims Email Automation & Premium Register (split)
Both group rows by company (pooling every "OFC"-named company into one
shared file) and both support the optional "Merge files for the same
company" checkbox, which folds branch-name variants of the same real
company (or a known abbreviation like AIC/JGE/JGI/UBL) into one combined
output — a fully deterministic, no-manual-review operation that never
drops or double-counts a row. Premium Register (split) additionally
accepts a "Co Insurance Ceded Register" export shape that has no Leader
column at all (every row on that export is already ceded business by
construction).

## Company merge logic & multiple email recipients

`company_alias_resolver.py` is the single shared engine behind every
company-grouping module:

- **Merging branch-name variants**: two bucket labels merge whenever they
  share the same leading token before the first underscore/space (e.g.
  "Atlas_Insurance_Faisalabad" and "Atlas_Insurance_ClaimRecovery" both
  reduce to "ATLAS"), or when one is a known abbreviation (AIC → Adamjee,
  JGE/JGI → Jubilee, RIC → Reliance, AGI → Askari General, plus a couple of
  confirmed real misspellings like "SALAM" → Salaam Takaful). This is
  fully deterministic — no confidence scores, no manual-review workbook —
  so it's always safe to toggle the merge checkbox on/off and regenerate.
- **Canonical output filenames**: a fixed, word-boundary-matched list
  (`CANONICAL_COMPANY_FILENAMES`) gives a known set of companies (Adamjee,
  Atlas, EFU, IGI, Jubilee, UBL, United, ...) a stable filename regardless
  of exactly how their raw/merged bucket label happens to be worded.
- **Email lookup**: resolves a bucket's raw company name against the Email
  Lookup File first by exact match, then by the same merge-key logic —
  this alone fixed a confirmed real bug where 0 of 11 real coinsurer
  drafts matched under plain exact-match lookup.
- **Multiple recipients per company**: the Email Lookup File can list more
  than one address per company (one per row, cleaned of `Name <email>`
  wrappers and stray trailing punctuation); all are joined into Outlook's
  "To" field with semicolons. You can also just type multiple addresses
  directly into the "To" field on any Review & Send Emails tab.

## Performance

See `PERFORMANCE_NOTES.md` for the full detail (five documented rounds of
profiling against real production-sized files — up to a 204,220-row
Premium Register and a 27,976-claim Claims Register run). In short:

- Every heavy action (Load, Analyze, Generate, Match) runs on a background
  thread via `perf_utils.run_in_background`, so the GUI never freezes.
- Excel reads go through `python-calamine` when installed (~7x faster than
  openpyxl's own reader on a real 500,000-row file).
- `perf_utils.StyleIndexCache` makes per-cell styling cheap by reusing
  openpyxl's own internal style index instead of re-hashing the same
  Font/Border combination on every cell (measured ~5x faster on a real
  output workbook).
- Bucketing/breakdown computations are O(rows), not O(items × rows) —
  fixed after a real 113-bucket Receivable run was found to be re-scanning
  the whole dataset once per output file.
- Output workbook generation is parallelized across CPU cores wherever
  each file is fully independent (Receivable, Claims Register, Premium
  Register split), falling back to sequential generation on a small batch
  or a single-core machine.
- One disclosed, un-fixable floor: a single coinsurer bucket with tens of
  thousands of matched rows still takes roughly a minute to write, since
  at that scale openpyxl's own per-cell object construction (not styling,
  not iteration) becomes the bottleneck — parallel generation is what
  keeps the *whole batch's* wall-clock time down in that case, not this
  one file's own write time.

## Multi-laptop / OneDrive safety

See `README_ONEDRIVE_FIX.md` for the full story. In short: GIMS' own small
state files (signature, templates, run-history logs) live in a local,
per-machine folder (`app_paths.py`) rather than inside the shared project
folder, and every write to them is atomic (`safe_io.py` — write to a temp
file, then rename into place) — so two laptops sharing the app folder via
OneDrive can never corrupt each other's signature or run history, and a
mid-write crash can never leave a half-written file behind. The
application code itself (every `.py` file) is read-only at runtime and
safe to keep in OneDrive as before.

## Customizing the email template (Outstanding Premium)

Edit `DEFAULT_TEMPLATE` in `email_generator.py`, or save additional
templates at runtime via `save_template()`. Supported placeholders:
`{client_name}`, `{greeting_name}` (resolves to a matched POC name if one
exists, else the company name), `{amount}`, `{report_date}`, `{signature}`.
Every other module's email templates (Claims Register, Premium/Claims
Email Automation, Premium Register split) are plain Python string
templates near the top of their own `*_email.py` file — edit those
directly (`SUBJECT_TEMPLATE` / `BODY_TEMPLATE`) if the wording needs to
change.

## Packaging as a Windows .exe

```
pip install pyinstaller
pyinstaller GIMS.spec
```
Produces `dist/GIMS/` — an `onedir` build (a folder of files, not a single
self-extracting exe) is used deliberately: a onefile exe re-unpacks itself
on every launch and both antivirus software and OneDrive's own sync
scanning tend to flag it far more aggressively than a plain folder of
files. Copy the whole `GIMS` folder into OneDrive, not just the `.exe`, and
run `GIMS.exe` from inside it. See `README_ONEDRIVE_FIX.md` for the icon
setup and the two OneDrive habits worth adopting (mark the folder "Always
keep on this device"; point each module's own output folder at a local,
non-synced location for large batch runs).

## Known limitations / things to confirm on the target machine

- **Outlook send, native Excel PivotTables/COM automation, and Excel-COM
  ScreenUpdating speedups** were built and correctness-checked but could
  only be exercised in dry-run/fallback mode in the development
  environment (no Windows/Excel available there) — confirm on a real
  Windows + Outlook + Excel machine before relying on them.
- **Parallel workbook generation** was correctness-verified (identical,
  valid output every time) but only throughput-tested on a single-CPU-core
  development sandbox — worth a timing check on a real multi-core machine
  for a large batch.
- **One disclosed performance floor**: a single coinsurer/company bucket
  with tens of thousands of rows still takes roughly a minute to write
  (see Performance, above) — not a bug, a ceiling worth knowing about.
- **Withholding Tax matching is exact-only by design** — a payee name that
  differs from the NTN Locator's own spelling beyond simple case/
  punctuation/whitespace normalization will not match, and is correctly
  flagged for manual review rather than guessed at.
- Emails are matched to companies/clients by **exact name** (after the
  merge-key fallback described above) — a spelling difference the merge
  logic doesn't already cover will still show as "missing email" for
  manual entry.
