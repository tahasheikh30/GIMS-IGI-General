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
