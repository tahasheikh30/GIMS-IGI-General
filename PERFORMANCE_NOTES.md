# Performance pass, summary

## What was slow, and why

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

## What changed

- **`perf_utils.py`** (new file). A small threading helper
  (`run_in_background`) that moves a long-running call onto a worker
  thread and marshals progress/result back to Tkinter safely via
  `widget.after()`, plus a streaming Excel reader
  (`fast_read_excel_rows`) built on openpyxl's `read_only=True` mode, and
  a `style_budget_ok()` gate the writers use to skip purely cosmetic
  per-cell styling above 50,000 rows (values and number formats are
  never affected, only font/border).
- **`data_processor.py` / `receivable_processor.py`**: `load_raw_data()`
  now checks the sheet's row count cheaply first, and automatically
  switches to the streaming reader above 100,000 rows. Below that
  threshold, behavior is identical to before.
- **`report_generator.py` / `receivable_generator.py`**: the two flat
  data-sheet writers (`_write_raw_data_into`, `_write_sheet`) now do a
  single vectorized pandas pass to clean NaN/date values, then iterate
  with `itertuples()`, then apply the style budget described above. Both
  accept an optional `report(msg)` callback for progress.
- **`app.py`**: `load_data`, `generate_reports` (Outstanding Premium),
  `load_receivable_data`, `generate_receivable_reports` (Coinsurance),
  and `_run_premium_register_match` (+ its save step) now all run
  through `perf_utils.run_in_background`, with their existing log
  widgets receiving live progress lines instead of the app appearing
  hung. Buttons on this pattern can be extended the same way if you want
  it on `analyze` / `analyze_receivable` too (those stayed as direct
  calls since pandas groupby/filter on a clean, already-loaded frame is
  fast even at 1,000,000 rows, generally well under a second, but the
  same `run_in_background` wrapper works there unchanged if you'd rather
  be consistent everywhere).

## One thing to be aware of

`generate_reports` builds native Excel PivotTables via COM automation
(`pivot_native_generator.py`), and COM has to be initialized on whichever
thread actually calls it. The main thread gets this for free from
Tkinter's message loop; a plain background thread does not. The patched
`generate_reports` worker now calls `pythoncom.CoInitialize()` /
`CoUninitialize()` around the Excel automation for exactly this reason.
If you later move `send_all_emails`'s Outlook COM calls onto a *different*
thread than the one `outlook_sender.py` already uses, do the same there.

## How to apply

Drop `perf_utils.py` into the project folder alongside the others, then
replace `app.py`, `report_generator.py`, `receivable_generator.py`,
`data_processor.py`, and `receivable_processor.py` with the versions here.
Every change degrades gracefully: if `perf_utils.py` is ever missing, each
module falls back to its exact original behavior (a plain `try/except
ImportError` guards every import), so nothing hard-breaks if only some
files get updated.

All five patched files were verified with `py_compile` and exercised
against synthetic 40,000 to 60,000 row datasets to confirm identical
output (values, NaN handling, dates) versus the original code, just
faster and with the style budget kicking in as expected.
