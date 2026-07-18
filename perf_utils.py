"""
Performance utilities shared across GIMS modules, added to keep the app
responsive on huge source files (real production exports have been seen
up to ~1,000,000 rows) and to speed up both the read and write paths.

Two independent problems are solved here:

1. THE GUI FREEZES ON BIG FILES
   Every "Load & Validate", "Analyze", and "Generate Reports" button in
   app.py currently calls straight into pandas/openpyxl on the Tkinter
   main thread. On a million-row file that call can take well over a
   minute, during which the whole window stops repainting and looks
   hung (Windows will often show "Not Responding"). run_in_background()
   moves that work onto a worker thread and marshals progress/result
   back onto the Tkinter main thread via widget.after(), which is the
   only thread-safe way to touch Tkinter widgets.

2. READING/WRITING HUGE EXCEL FILES IS SLOW
   pandas' read_excel() for .xlsx always goes through openpyxl under the
   hood, but in its default ("normal") mode openpyxl builds a full
   in-memory object graph for every cell before pandas ever sees it -
   expensive in both time and RAM for a million-row sheet. openpyxl's
   own `read_only=True` mode streams rows directly off disk instead and
   is dramatically lighter. fast_read_excel_rows() below uses that mode
   directly, bypassing pandas' own reader entirely for the "give me all
   the raw rows" step (header-row detection, then DataFrame construction,
   stay unchanged - the alias-matching logic in each *_processor.py file
   is untouched, only the underlying row source changes).

   On the write side, applying a Font/Border/Alignment object to every
   individual cell (the pattern used throughout report_generator.py and
   receivable_generator.py) is the dominant cost once you get into the
   hundreds of thousands of rows - worse than the pandas iteration
   itself. style_budget_ok() below is a simple row-count gate the
   generators use to skip per-cell font/border on very large sheets
   (values and number formats are always kept - only the cosmetic
   per-cell font/border is skipped past the threshold, and a note is
   logged so this is never a silent behavior change).
"""
import queue
import threading
import time

import openpyxl


# ---------------------------------------------------------------------------
# 1. Background task runner for Tkinter
# ---------------------------------------------------------------------------

def run_in_background(widget, worker, on_progress=None, on_done=None, on_error=None,
                       busy_widgets=None, poll_ms=80):
    """
    Run `worker(report)` on a background thread so the Tkinter main loop
    never blocks, then marshal progress/result/errors back onto the main
    thread safely.

    worker(report): your long-running function. Call report("some status
        line") from inside it at reasonable intervals (e.g. every 5,000-
        50,000 rows) to surface progress. Return whatever result value
        you want handed to on_done.
    on_progress(msg): called on the MAIN thread for every report() call.
        Typical use: append the line to a ScrolledText log.
    on_done(result): called on the MAIN thread once worker() returns
        normally, with its return value.
    on_error(exc): called on the MAIN thread if worker() raises. If not
        given, the exception is re-raised into the on_progress channel
        as a plain string instead of crashing the polling loop.
    busy_widgets: optional list of Tkinter widgets (buttons, etc.) to
        disable for the duration and re-enable afterwards, so the user
        can't double-click "Generate Reports" while it's already running
        on a huge file.
    poll_ms: how often (ms) the main thread checks for new messages.
    """
    q = queue.Queue()
    busy_widgets = busy_widgets or []

    for w in busy_widgets:
        try:
            w.configure(state="disabled")
        except Exception:
            pass

    def _worker_wrapper():
        def report(msg):
            q.put(("progress", msg))
        try:
            result = worker(report)
            q.put(("done", result))
        except Exception as exc:  # noqa: BLE001 - surfaced to the UI, not swallowed
            q.put(("error", exc))

    def _poll():
        drained_terminal = None
        try:
            while True:
                kind, payload = q.get_nowait()
                if kind == "progress" and on_progress:
                    on_progress(payload)
                elif kind in ("done", "error"):
                    drained_terminal = (kind, payload)
        except queue.Empty:
            pass

        if drained_terminal is None:
            widget.after(poll_ms, _poll)
            return

        for w in busy_widgets:
            try:
                w.configure(state="normal")
            except Exception:
                pass

        kind, payload = drained_terminal
        if kind == "done" and on_done:
            on_done(payload)
        elif kind == "error":
            if on_error:
                on_error(payload)
            elif on_progress:
                on_progress(f"ERROR: {payload}")

    thread = threading.Thread(target=_worker_wrapper, daemon=True)
    thread.start()
    widget.after(poll_ms, _poll)


class ThrottledReporter:
    """
    Wraps a report(msg) callback so progress lines only actually get
    pushed to the queue/UI at most once every `min_interval` seconds (or
    every `min_rows` rows processed) - reporting on literally every row
    of a million-row loop would itself become a bottleneck and would
    flood the log widget with unreadable noise.
    """
    def __init__(self, report_fn, min_interval=0.3, min_rows=5000):
        self._report = report_fn
        self._min_interval = min_interval
        self._min_rows = min_rows
        self._last_time = 0.0
        self._last_row = 0

    def maybe(self, row_index, total=None, label="rows"):
        now = time.monotonic()
        if row_index - self._last_row < self._min_rows and now - self._last_time < self._min_interval:
            return
        self._last_row = row_index
        self._last_time = now
        if total:
            pct = (row_index / total) * 100 if total else 0
            self._report(f"  ... {row_index:,} / {total:,} {label} ({pct:.0f}%)")
        else:
            self._report(f"  ... {row_index:,} {label}")

    def final(self, row_index, label="rows"):
        self._report(f"  ... {row_index:,} {label} (done)")


# ---------------------------------------------------------------------------
# 2. Fast streaming Excel reads (bypasses pandas' own reader for the raw
#    row scan; header detection / column-alias logic in each processor
#    module is unchanged and works the same off the returned rows)
# ---------------------------------------------------------------------------

def fast_read_excel_rows(filepath, sheet_name, report=None):
    """
    Stream every row of one sheet as a list of lists using openpyxl's
    read_only mode, which reads directly off disk row-by-row instead of
    building pandas' full in-memory frame up front. For a ~1,000,000 row
    sheet this is typically several times faster and uses a fraction of
    the peak memory of pd.read_excel().

    Returns (rows, n_rows) where rows is a list of tuples (raw cell
    values, exactly as openpyxl would give a normal read - dates come
    back as datetime objects, blanks as None).
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    try:
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[sheet_name]
        rows = []
        reporter = ThrottledReporter(report, min_rows=50000) if report else None
        for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            rows.append(row)
            if reporter:
                reporter.maybe(i, label="rows read")
        if reporter:
            reporter.final(len(rows), label="rows read")
        return rows, len(rows)
    finally:
        wb.close()


def rows_to_dataframe(rows, header_row_index):
    """
    Build a pandas DataFrame from rows returned by fast_read_excel_rows(),
    given the (0-indexed) header row - mirrors the header/data split each
    *_processor.py module already does after pd.read_excel(header=None).
    """
    import pandas as pd
    columns = list(rows[header_row_index])
    data_rows = rows[header_row_index + 1:]
    df = pd.DataFrame(data_rows, columns=columns)
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def estimate_row_count(filepath, sheet_name=None):
    """
    Cheap row-count check (reads the sheet's own stored dimensions, not
    the data itself) used to decide whether a file is "huge" enough to
    warrant the fast streaming path and reduced per-cell styling on
    write. Falls back to 0 (treated as "unknown/small") if the file's
    dimension metadata isn't available for some reason.
    """
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True)
        try:
            ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.worksheets[0]
            return ws.max_row or 0
        finally:
            wb.close()
    except Exception:
        return 0


# ---------------------------------------------------------------------------
# 3. Write-side style budget
# ---------------------------------------------------------------------------

STYLE_ROW_THRESHOLD = 50000  # above this, skip per-cell font/border (values + number formats are always kept)


def style_budget_ok(n_rows):
    return n_rows <= STYLE_ROW_THRESHOLD
