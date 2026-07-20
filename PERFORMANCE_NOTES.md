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

