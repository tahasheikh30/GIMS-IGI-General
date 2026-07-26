# Performance and progress bar notes

## Part 1, keeping the app responsive on huge files

### What was slow, and why

1. **GUI freezing.** Every heavy button (Load & Validate, Analyze, Generate
   Reports, in both Outstanding Premium and the Coinsurance module, plus
   Premium Register matching) ran directly on the Tkinter main thread. On a
   real 1,000,000-row export that call alone can take well over a minute,
   during which the window stops repainting entirely (Windows shows "Not
   Responding").
2. **Reading.** `pd.read_excel()` builds a full in-memory object graph for
   every cell before pandas even sees it. Fine at 10,000 rows, expensive at
   1,000,000.
3. **Writing.** Applying an openpyxl `Font`/`Border` style object to every
   individual cell (the pattern used throughout) is the single biggest
   writer cost once you're past roughly 50,000 rows, worse than the
   `iterrows()` loop it was riding on.
4. **`iterrows()`** itself is also meaningfully slower than `itertuples()`
   for large frames, on top of point 3.

### What changed

- **`perf_utils.py`**. A threading helper (`run_in_background`) that moves
  a long-running call onto a worker thread and marshals progress/result
  back to Tkinter safely via `widget.after()`, plus a streaming Excel
  reader (`fast_read_excel_rows`) built on openpyxl's `read_only=True`
  mode, and a `style_budget_ok()` gate the writers use to skip purely
  cosmetic per-cell styling above 50,000 rows (values and number formats
  are never affected, only font/border).
- **`data_processor.py` / `receivable_processor.py`**: `load_raw_data()`
  checks the sheet's row count cheaply first, and automatically switches
  to the streaming reader above 100,000 rows.
- **`report_generator.py` / `receivable_generator.py` /
  `premium_register_generator.py`**: the flat data-sheet writers do a
  single vectorized pandas pass to clean NaN/date values, then iterate
  with `itertuples()`, then apply the style budget described above.
  Premium Register's red "Not Matched" row highlight is never skipped by
  the budget, since it's functional information (what needs manual entry
  into the Premium Register), not decoration, only the green
  "matched, new column" tint and plain body font are.
- **`app.py`**: `load_data`, `generate_reports` (Outstanding Premium),
  `load_receivable_data`, `generate_receivable_reports` (Coinsurance),
  and `_run_premium_register_match` (+ its save step) all run through
  `perf_utils.run_in_background`.

### One thing to be aware of

`generate_reports` builds native Excel PivotTables via COM automation
(`pivot_native_generator.py`), and COM has to be initialized on whichever
thread actually calls it. The main thread gets this for free from
Tkinter's message loop; a plain background thread does not. The patched
`generate_reports` worker calls `pythoncom.CoInitialize()` /
`CoUninitialize()` around the Excel automation for exactly this reason.

## Part 2, the progress bar

A `ttk.Progressbar` + status label (`perf_utils.ProgressBarPanel`) now sits
directly under the "Generate Excel Reports" button on the Outstanding
Premium and Coinsurance pages, and under "Run Matching" on Premium
Register. It is entirely a passive reflection of the background thread's
own progress; nothing about it can slow down or otherwise affect the
actual generation work:

- The worker thread still runs at full, unthrottled speed. It calls
  `report.progress(current, total, text=...)` at natural points (once per
  file for the multi-file flows, throttled to a few times a second for
  the single-file Premium Register write), which just drops a tuple onto
  a thread-safe queue and returns immediately, no waiting, no lock
  contention with the UI.
- The main thread polls that queue roughly every 80ms (the same
  `widget.after()` loop already used for the text log) and updates the
  bar's value from whatever's waiting. If the worker is faster than the
  poll interval, intermediate updates are simply superseded by the latest
  one, they're never a bottleneck for the worker.
- Two progress scales are kept deliberately separate so the bar never
  jumps backwards: `ThrottledReporter` (used inside a single sheet's row
  loop) stays text-log-only when that write sits inside an outer
  per-file loop, so the per-file loop is the only thing driving the bar
  in Outstanding Premium and Coinsurance. `RowProgress` is used instead
  for the single-output flows (Premium Register matching and its Excel
  write) where there's no outer per-file loop, so row count is free to
  drive the bar directly there.
- The bar is indeterminate ("busy stripe") for Premium Register's
  matching pass, since the exact per-row timing isn't predictable before
  the leader file finishes loading, then switches to determinate once
  matching starts (leader row count is known by then). It's always
  determinate for the multi-file flows, since the file count is known
  upfront.
- The panel stays permanently visible in its slot (never appears/
  disappears, which would otherwise reflow the rest of the page) and
  simply resets to an empty, idle state before each run and shows the
  final tally after.

All of this was verified against synthetic 40,000 to 60,000 row datasets
(past the style budget, so both code paths get exercised) with a fake
Tkinter widget standing in for `after()`, confirming the text log, the
numeric progress channel, and the final result all arrive correctly and
in order.

## How to apply

Drop `perf_utils.py` into the project folder alongside the others, then
replace `app.py`, `report_generator.py`, `receivable_generator.py`,
`data_processor.py`, `receivable_processor.py`,
`premium_register_processor.py`, and `premium_register_generator.py` with
the versions here. Every change degrades gracefully: if `perf_utils.py` is
ever missing, each module falls back to its exact original behavior (a
plain `try/except ImportError` guards every import), so nothing hard-breaks
if only some files get updated.

## Part 3, eliminating O(items x rows) re-scans

Parts 1 and 2 kept the GUI responsive on a huge file by moving the work
off the main thread and speeding up the read/write paths. They didn't
touch a separate problem: several places re-filtered or re-bucketed the
**entire** dataset once per output item, inside a loop over those items -
fine on a handful of clients/companies, but scaling as O(num_items x
num_rows) rather than O(num_rows) as either number grows. On real data
this was the single biggest remaining bottleneck, worse than anything
Parts 1-2 addressed.

### What was slow, and why

- **Receivable module - the worst offender.** `client_breakdown_for_bucket()`
  re-scanned and re-bucketed the whole cleaned dataset from scratch, using
  a non-vectorized `Series.apply()`, and was called once per OUTPUT BUCKET
  inside the report-generation loop (both in `app.py`'s GUI flow and in
  `receivable_automation.run_pipeline()`'s scripted flow). A real logged
  run (`receivable_run_history.json`) had 113 output buckets - so a file
  that should be scanned once was instead being fully re-scanned 113
  times over. `_bucket_series()` itself (used by this and by
  `build_company_groups()`/`company_totals()`) was also a per-row Python
  `.apply()` rather than a vectorized string check, so every one of those
  113 passes was slower than it needed to be too.
- **Outstanding Premium - the same shape of bug, twice.** `analyze()`'s
  treeview-population loop called `client_agents()`/`client_branches()`
  once per QUALIFYING CLIENT, each re-scanning the whole filtered
  dataset. `generate_reports()`'s worker did the same thing twice over -
  `get_client_raw_rows()` and (on the non-native-Excel fallback path,
  which is the *only* path when Excel/pywin32 isn't installed at all)
  `client_pivot_table()` each re-filtered the whole dataset once per
  qualifying client.
- Neither `analyze()` (Outstanding Premium) nor `analyze_receivable()`
  (Coinsurance) ran through `perf_utils.run_in_background` at all - only
  the Load and Generate steps did. On a huge file, clicking Analyze could
  still freeze the window even after Parts 1-2.

### What changed

- **`receivable_processor.py`**: `_bucket_series()` is now a single
  vectorized `Series.str.contains()` call instead of a per-row
  `.apply()`. A new `build_client_breakdowns()` computes every bucket's
  client breakdown in one `groupby(["_rec_bucket", "_rec_client"])` pass
  regardless of how many buckets there are. A new `analyze_buckets()`
  goes further and computes groups + totals + breakdowns all from a
  single bucket assignment, for callers (the GUI, the scripted pipeline)
  that need all three. `client_breakdown_for_bucket()` and the older
  `build_company_groups()`/`company_totals()` are kept as-is for
  one-off/backward-compatible use, with their docstrings now pointing at
  the bulk versions for anything that loops over multiple buckets. The
  legacy single-amount-column fallback's `Type`/`Category` classification
  (`_classify_type`) also got a vectorized `_classify_type_bulk()`
  counterpart, for the same reason.
- **`data_processor.py`**: new `group_raw_rows_by_client()` (one groupby
  pass, O(1) lookup per client afterward) and
  `client_agents_branches_bulk()` (same idea, for the Agent/Branch
  columns) replace the per-client-loop scanning pattern. A new
  `client_pivot_table_from_rows()` builds a client's Department x Year
  pivot directly from that client's own already-filtered rows instead of
  re-filtering the full dataset - `client_pivot_table()` itself is now
  implemented on top of it and kept for one-off use. The `Year` column
  derivation in `validate_data()` was also switched from a per-row
  `.apply()` to a vectorized nullable-Int64 cast.
- **`app.py`**: `generate_reports()` and `generate_receivable_reports()`
  now consume these bulk-computed lookups instead of re-scanning per
  item. `analyze()` and `analyze_receivable()` now both run through
  `perf_utils.run_in_background` too (with their own "Analyze" button
  passed as a `busy_widget` so a second click can't queue up a duplicate
  run), computing everything the treeview needs via the new bulk
  functions before handing the result back to the main thread to render.
- **`receivable_automation.py`**: `run_pipeline()` (the scripted/CLI
  entry point, used outside the GUI) now uses `analyze_buckets()` too, so
  a scripted run gets the identical fix rather than only the GUI path.

### Verified

Every new bulk function was checked against its old per-item equivalent
on synthetic data (identical output, not just "doesn't crash") before
being wired in, and benchmarked at a scale matching real usage:

- Outstanding Premium, 300,000 rows / 400 qualifying clients: the old
  per-client raw-row scan took ~7.1s; the new one-pass grouping took
  ~0.2s (about 33x).
- Receivable, 150,000 rows / 111 output buckets (matching the 113-bucket
  real run in `receivable_run_history.json`): the old per-bucket
  breakdown loop took ~4.5s; the new single bulk pass took ~0.14s (about
  32x).

Both numbers scale further in the app's favor as the row count or item
count grows, since the old pattern was O(items x rows) and the new one is
O(rows) regardless of item count. Empty-dataframe and missing-column edge
cases (no data after filtering, no Agent Name/Branch desc column present)
were checked separately and degrade to empty results rather than raising.

## Part 4, fixing the two remaining bottlenecks Parts 1-3 didn't touch: reading huge files, and writing huge output workbooks

Parts 1-3 fixed GUI freezing, slow row-by-row reads, and O(items x rows)
re-scans. Two bottlenecks remained untouched, confirmed against real
production-sized files: matching in the Claims Register module (204,220
register rows / 27,976 claims) and generating output in both the
Receivable and Claims Register modules (a real 500,000-row Receivable
master file / 143 companies, and the same 27,976-claim run fanning out to
274,000 matched claim-coinsurer rows across 98 coinsurers).

### Claims Register matching (33 minutes -> ~10 seconds)

`claims_register_processor.py`'s own blocking (`_build_policy_index`)
indexed register rows by their bare 5+ digit serial number alone. On the
real register, low serials collide constantly (every branch/class starts
its own numbering back at 1) - some keys matched 1,500-5,600 register
rows, and every claim landing on one of them had to be fuzzy-scored
(SequenceMatcher) against the whole bloated pool.

Fix: a second, far more selective blocking key pairs the branch/class
code with the serial (e.g. `("HOFFNOP", "00025")` instead of just
`"00025"`) - see `_compound_keys()` / `MAX_BLOCK_SIZE`'s docstring. Only
the old bare-serial keys are purged once their candidate set is
unreasonably large; compound keys are never purged, since they're small
by construction. Also added `_build_siblings_index()` so pulling in a
matched claim's co-insurer rows is an O(1) lookup instead of a fresh
full-register scan per match.

Verified this doesn't trade away correctness for speed: compared the fast
(purged) result against the true unpurged result on a random 2,000-claim
sample. Every difference found was the unpurged version accepting a weak,
coincidental match (branch/class similarity barely over the acceptance
threshold, e.g. `LHRHHDP` matched to unrelated `LHREBPOT`) - a pre-existing
weakness in the acceptance thresholds that predates this fix, not
something introduced by it. Zero exact reference-number matches were lost.

### Reading Excel files (openpyxl's own reader was the actual cost)

Profiled `receivable_processor.py`'s `load_raw_data()` against a real
500,000-row file: openpyxl's `read_only` row iteration alone (no other
code running around it) took ~43 of the ~50 total seconds. Switched
`perf_utils.fast_read_excel_rows()` to try `python-calamine` (a Rust-based
reader) first, falling back to the original openpyxl path if it isn't
installed - ~7x faster on the same file (50s -> 7.4s total). Also reads
legacy `.xls` directly. `receivable_processor.load_raw_data()` now always
uses this fast path regardless of file size, rather than only above the
old 100,000-row threshold, since it's strictly faster at every size.

### Writing output workbooks (openpyxl re-hashing the same Font/Border on every cell)

Profiled `receivable_generator.py` and `claims_register_generator.py`
building real output files: ~73% of per-file time was inside openpyxl's
own style-deduplication machinery
(`styleable.py` / `indexed_list.py` / `Serialisable.__hash__`) - re-hashing
the exact same, already-reused `Font`/`Border` objects on every single
cell, even though the generators were already reusing shared constants
rather than creating new style objects per cell.

Fix: `perf_utils.StyleIndexCache` (new). The first time a given
font/border/fill/alignment/number_format combination is used, it's
applied the normal way and the resulting internal openpyxl style index is
cached; every later cell needing the identical combination gets that
index assigned directly (`cell._style = index`), skipping the repeated
hashing. Keyed on **object identity** (`id()`), not the style objects'
own values - hashing a Font/Border object to use as a dict key would be
exactly the cost being avoided, even on a cache hit. Falls back to the
always-correct per-attribute path if `_style` (a private openpyxl
attribute) is ever unavailable, so a future openpyxl version can only
make this slower, never wrong. Wired into both generators'
per-cell-styled loops (`_write_sheet` / `_write_data_sheet` /
`_write_unmatched_sheet` and each module's Summary-sheet breakdown
table), replacing the old `style_budget_ok()` row-count gate that used to
skip font/border entirely above 50,000 rows - with the cache, full
styling is cheap enough that skipping it is no longer necessary.

Measured on a real ~2,700-row company workbook: 4.18s -> 0.79s per file
(~5.3x) once keyed correctly by identity. Confirmed correct by reopening
generated files and checking font/border/number_format round-trip
correctly through the reused style index.

**A genuine floor, not fixed by this**: one real coinsurer bucket
(HABIB INSURANCE COMPANY LIMITED, 58,423 matched claim-coinsurer rows,
~4.4 million cells once every column is counted) still takes roughly a
minute to write even with styling no longer the bottleneck - at that cell
count, openpyxl's own per-cell object construction in its normal
(non-streaming) write mode is the cost, confirmed via cProfile
(`lxml_write_cell` / `Worksheet._get_cell` / `Cell.__init__`). openpyxl's
much faster `write_only` streaming mode was tested as a fix and rejected:
it cannot be safely mixed into the same workbook as the Summary sheet's
merged cells and embedded charts - confirmed directly by attempting it,
which silently dropped every row written to the write-only sheet (the
file reopened with `max_row == 1`, header only). Rather than risk
producing a corrupted file, huge single buckets are left on normal mode
and instead handled by parallelization (below).

### Parallel generation across CPU cores

Each output workbook (one per company / one per coinsurer) is fully
independent - its own DataFrame slice in, its own `.xlsx` out - so
generation is an "embarrassingly parallel" batch. Added
`build_company_workbooks_parallel()` (`receivable_generator.py`) and
`build_coinsurer_workbooks_parallel()` (`claims_register_generator.py`),
both using `concurrent.futures.ProcessPoolExecutor` to spread file
generation across cores, wired into `app.py`'s "Generate Excel Reports"
workers in place of the old one-at-a-time loop. The Claims Register
version also folds the "Unmatched Claims" workbook into the same batch as
one more task, so it runs concurrently with the coinsurer files instead
of as a separate step afterward.

Both fall back to the original sequential loop whenever parallelism
wouldn't help or isn't safely available: fewer than 2 CPU cores, a small
batch (parallel process start-up/pickling overhead would outweigh the
benefit), or a process pool that fails to start at all for any
environment-specific reason - caught broadly so a batch never fails
outright just because the parallel path couldn't start.

**Caveat, disclosed rather than glossed over**: this was built and
correctness-verified (identical, fully valid output, confirmed by
reopening generated files) in a sandbox with only 1 CPU core available,
so the actual multi-core speedup could not be measured directly in that
environment - only reasoned about from the fact that each file's build
cost is CPU-bound and fully independent of every other file's. On a real
multi-core machine, generating a batch where a handful of buckets are
huge and the rest are small should take roughly as long as those few huge
buckets alone (run concurrently), not the sum of every file in the batch.

### Net effect, measured on real files where possible

- Claims Register matching: ~33 minutes -> ~10-50 seconds (see Part 4's
  first section above for the range depending on machine load).
- Receivable, 500,000 rows / 143 companies: load+validate+analyze+generate
  ~204s -> ~74-79s on a single core (reader + styling fixes only; parallel
  generation would improve this further on a multi-core machine).
- Claims Register generation, 27,976 claims / 98 coinsurers / 274,000
  matched rows: the styling fix alone took the small/medium buckets (91
  of 98, ~22,000 total rows) from a proportional share of ~744s down to
  ~36s; the 7 large buckets (~256,000 rows combined) improved less
  dramatically in absolute terms since they're bound by openpyxl's
  per-cell construction cost at that scale, not styling - parallel
  generation is what brings the WHOLE BATCH's wall-clock time down on a
  real multi-core machine, by running those 7 large buckets concurrently
  instead of back to back.

## Part 5, Policy Check (Premium Register) and Outstanding Premium

### Policy Check - two correctness bugs found while investigating "why is it slow"

Before any performance work, loading the real Premium Register file
(the same 204,220-row file the Claims Register module uses) into this
module failed outright: `REGISTER_ALIASES` only recognizes header text
like `"Leader#"` / `"Doc Ref #"`, but the real file uses `"Document No."`
/ `"Base Doc No."` (the same header style Claims Register's own aliases
already expect). This module was evidently built against an older or
different register export format than what's actually in use now - not
a performance issue at all, but worth fixing before performance mattered.

Separately, and more seriously: **this module had no continuation-sheet
merging at all**. A legacy `.xls` register bigger than Excel's old 65,536-
row-per-sheet cap (confirmed: the real 204,220-row register is split
across 4 sheets, header only on the first) would have silently been read
as just its first ~65,523 rows - 68% of the register missing, with no
error or warning. Claims Register already had this fix
(`_find_continuation_sheets`); ported the same logic here
(`premium_register_processor._find_continuation_sheets`), verified
against a synthetic 3-sheet, 9,000-row split file that all 9,000 rows
load correctly.

### Policy Check - performance (matching 15,000 leader rows against a 200,000-row register)

Profiled with cProfile: `register_df.at[idx, leader_ref_col]` -
DataFrame-level `.at[]` access - was being called on every single
candidate check inside the matching loop. A microbenchmark of the same
"same column, many different row indices" access pattern confirmed
`Series.at[]` (i.e. access the column ONCE as its own Series, then index
into that) is roughly 5x faster than `DataFrame.at[]` for this exact
shape of repeated lookup - `DataFrame.at[]` goes through column lookup,
scalar "boxing", and a pandas `__finalize__` call that deep-copies frame
metadata on every access, none of which `Series.at[]` needs to repeat.

Separately, `_score_candidate` always ran a full `SequenceMatcher.ratio()`
call for every candidate, even though the final ranking always compares
`(run_len, ratio)` with `run_len` first - a candidate with a lower
run_len can never win regardless of its ratio. Split scoring into two
passes: a cheap run_len-only check (plain substring test) across every
candidate first, then `SequenceMatcher.ratio()` only for whichever
candidates tie for the best run_len found - provably the same result as
scoring everyone, just skipping ratio computation for candidates that
can't possibly win. **Verified this is bit-for-bit identical to the
original algorithm**: ran 1,500 deliberately fuzzy/corrupted queries
(real register references with random slashes-to-dashes swaps, digit
typos, and truncation) through both the old and new scoring logic side
by side - zero differences in status, matched policy, or score.

Also applied the same fast-reader (python-calamine) and
`StyleIndexCache` fixes as every other module.

Combined result on a real-shaped 200,000-row register / 15,000-row
leader file: load+match+write went from ~97s+ to **~14s**.

### Outstanding Premium

`data_processor.py` already called the shared `fast_read_excel_rows()`
(and so already benefited from the calamine swap in Part 4), but only
above a 100,000-row threshold - anything smaller still went through
`pd.read_excel()` directly. Removed the threshold, matching the fix
already applied to `receivable_processor.py`: a real-shaped 50,000-row
file went from 2.94s to 0.41s. `report_generator.py`'s `_write_raw_data_into`
(the "Raw Data" sheet, which scales with a client's own transaction
volume) got the same `StyleIndexCache` fix as the other three generators.

This module has two genuinely different generation paths depending on
whether Excel is installed, which matters for how far parallelization
could safely go:

- **Native PivotTable path** (Excel/pywin32 available): drives ONE shared
  Excel COM instance across every client sequentially. Added
  `ScreenUpdating = False` and manual calculation mode to that shared
  instance for the duration of the batch (restored before it quits) - a
  standard, well-documented speedup for batch Excel COM automation.
  **This could not be verified at all** - this was developed on Linux
  with no Windows/Excel available, so unlike every other fix in this
  document, this one is applied on the strength of it being a known
  technique, not a project-specific measurement. COM automation also
  isn't safe to parallelize the same way plain `openpyxl` writing is (a
  single Application object isn't built for concurrent access from
  multiple processes), so this path was deliberately left sequential
  rather than risk something that couldn't be tested here.
- **Fallback static-table path** (no Excel, or a client's native pivot
  build fails): pure `openpyxl`, fully independent per client - added
  `report_generator.write_client_reports_parallel()` using the same
  `ProcessPoolExecutor` pattern as the other three modules, with the same
  fallback conditions. Verified correct: 500/500 synthetic client reports
  generated with 0 errors, reopened and checked for valid content/fonts.

Measured on a real-shaped 600,000-row file / 500 clients (fallback path,
single CPU core in this environment): load+validate+analyze 7.76s,
generate 500 reports 69.5s - the generation step would improve further
on a real multi-core machine, same caveat as Part 4's parallel
generation work.



