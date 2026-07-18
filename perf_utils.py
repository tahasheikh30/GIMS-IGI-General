"""
Performance utilities shared across GIMS modules, added to keep the app
responsive on huge source files (real production exports have been seen
up to ~1,000,000 rows) and to speed up both the read and write paths.

Three independent problems are solved here:

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

3. NO VISUAL FEEDBACK WHILE GENERATING EXCEL FILES
   ProgressBarPanel is a small ttk.Progressbar + status label meant to
   sit directly under a "Generate Excel Reports" button. It is driven
   entirely off the same background thread used to solve problem 1, so
   it never affects the flow of the application - the worker thread
   still does the real work at full speed; the progress bar is purely a
   read-only reflection of how far along it is, updated a few times a
   second at most (see Reporter.progress below), never something the
   worker waits on.
"""
import queue
import threading
import time

import openpyxl


# ---------------------------------------------------------------------------
# 1. Background task runner for Tkinter
# ---------------------------------------------------------------------------

class Reporter:
    """
    Passed into every worker(report) function. Calling it directly,
    report("some status line"), behaves exactly as before (pushes a text
    line onto the log). report.progress(current, total, text=None)
    additionally drives a numeric progress bar via on_progress_value,
    completely independent of the text log - a worker can call either,
    both, or neither, and existing worker functions that only ever call
    report("...") keep working unchanged.
    """
    def __init__(self, q):
        self._q = q

    def __call__(self, msg):
        self._q.put(("progress", msg))

    def progress(self, current, total=None, text=None):
        self._q.put(("progress_value", (current, total, text)))


def run_in_background(widget, worker, on_progress=None, on_progress_value=None,
                       on_done=None, on_error=None, busy_widgets=None, poll_ms=80):
    """
    Run `worker(report)` on a background thread so the Tkinter main loop
    never blocks, then marshal progress/result/errors back onto the main
    thread safely.

    worker(report): your long-running function. report is a Reporter
        (see above) - call report("some status line") for text progress,
        and/or report.progress(i, total) to drive a progress bar. Return
        whatever result value you want handed to on_done.
    on_progress(msg): called on the MAIN thread for every report("...")
        call. Typical use: append the line to a ScrolledText log.
    on_progress_value(current, total, text): called on the MAIN thread
        for every report.progress(...) call - typical use: update a
        ProgressBarPanel. total may be None (indeterminate/"busy" mode);
        text may be None (leave the panel's current label as-is).
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
        report = Reporter(q)
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
                elif kind == "progress_value" and on_progress_value:
                    current, total, text = payload
                    on_progress_value(current, total, text)
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

    Deliberately text-only (does not also drive a numeric progress bar):
    the callers that write a single huge sheet are themselves usually
    called from inside a per-file loop in app.py that already drives the
    progress bar at file-level granularity (see ProgressBarPanel below).
    Mixing "row N of this file" and "file i of N files" onto the same
    bar would make it jump backwards every time a new file starts, so
    only one of the two scales is allowed to own the bar at a time - the
    file-level one, since that's what "Generate Excel Reports" naturally
    means from the user's side.
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


class RowProgress:
    """
    Throttled row-level progress for a SINGLE Excel write that has no
    outer per-file loop around it (e.g. Premium Register's one output
    file, or Loss Ratio's one summary workbook) - drives both the text
    log and, when the wrapped callback is a Reporter, the numeric
    progress bar directly off row count. Use ThrottledReporter instead
    when the write sits inside an outer per-file loop that already owns
    the bar at file-level granularity - see the note on ThrottledReporter
    above for why the two are kept separate.
    """
    def __init__(self, report_fn, total, min_interval=0.3, min_rows=2000, label="rows"):
        self._report = report_fn
        self._total = total or 0
        self._label = label
        self._min_interval = min_interval
        self._min_rows = min_rows
        self._last_time = 0.0
        self._last_row = 0

    def maybe(self, row_index):
        now = time.monotonic()
        if row_index - self._last_row < self._min_rows and now - self._last_time < self._min_interval:
            return
        self._last_row = row_index
        self._last_time = now
        pct = (row_index / self._total * 100) if self._total else 0
        self._report(f"  ... {row_index:,} / {self._total:,} {self._label} ({pct:.0f}%)")
        if hasattr(self._report, "progress"):
            self._report.progress(row_index, self._total,
                                   text=f"Writing {self._label}: {row_index:,} / {self._total:,}")

    def final(self):
        self._report(f"  ... {self._total:,} {self._label} (done)")
        if hasattr(self._report, "progress") and self._total:
            self._report.progress(self._total, self._total, text=f"{self._total:,} {self._label} written")


# ---------------------------------------------------------------------------
# 1b. Visual progress bar panel (Tkinter), driven off run_in_background's
#     on_progress_value channel - purely a read-only reflection of the
#     background thread's own progress, never something the worker waits
#     on, so it can never slow down or otherwise affect the actual work.
# ---------------------------------------------------------------------------

def _resolve_bg_color(widget):
    """
    Best-effort background color lookup that works for BOTH plain tk
    widgets (which support cget("bg")) and themed ttk widgets (which do
    NOT - ttk.Frame, ttk.Label, etc. raise _tkinter.TclError on
    cget("bg")/cget("background") entirely, since their background comes
    from the active ttk theme/style instead of a per-widget option).
    Returns None if no color could be determined either way, in which
    case the caller should simply omit the bg kwarg and let Tk use
    whatever the platform default is - never raises.
    """
    for option in ("bg", "background"):
        try:
            return widget.cget(option)
        except Exception:
            pass
    try:
        from tkinter import ttk
        style = ttk.Style()
        style_name = widget.cget("style") or widget.winfo_class()
        color = style.lookup(style_name, "background")
        if color:
            return color
    except Exception:
        pass
    return None


class ProgressBarPanel:
    """
    A small determinate/indeterminate progress bar + status label, meant
    to sit directly under a "Generate Excel Reports"-style button.

    Usage:
        panel = ProgressBarPanel(parent_frame)
        panel.frame.pack(fill="x", padx=12, pady=(0, 8))   # once, at build time
        ...
        perf_utils.run_in_background(
            self, worker,
            on_progress=..., on_progress_value=panel.update,
            on_done=lambda r: panel.finish("Done."),
            on_error=lambda e: panel.finish("Failed."),
            busy_widgets=[...],
        )

    The panel stays permanently packed in its slot (so it never disturbs
    the surrounding layout by appearing/disappearing) and simply resets
    to an empty, idle look between runs.
    """
    def __init__(self, parent, width=320):
        import tkinter as tk
        from tkinter import ttk

        bg = _resolve_bg_color(parent)

        # A plain tk.Frame (not ttk.Frame) is used deliberately here even
        # when the parent is a themed ttk widget: it's the only frame type
        # that reliably accepts an explicit bg color on every Tk build,
        # which keeps this panel's background matched to whatever it's
        # dropped into (a themed ttk page vs. a hand-styled tk.Frame card
        # elsewhere in the app) without needing the caller to tell it which
        # kind of parent it has.
        frame_kwargs = {"bg": bg} if bg else {}
        self.frame = tk.Frame(parent, **frame_kwargs)

        self.bar = ttk.Progressbar(self.frame, orient="horizontal", mode="determinate",
                                    length=width, maximum=100, value=0)
        self.bar.pack(side="left", padx=(0, 10), pady=2)

        self.status_var = tk.StringVar(value="")
        label_kwargs = {"bg": bg} if bg else {}
        self.status_label = tk.Label(self.frame, textvariable=self.status_var,
                                      font=("Segoe UI", 9), anchor="w",
                                      fg="#5B6B79", **label_kwargs)
        self.status_label.pack(side="left", fill="x", expand=True)

        self._indeterminate = False

    def start(self, total=None, text="Starting..."):
        """Call once right before the background task starts. total=None
        switches the bar to indeterminate ("busy stripe") mode, used for
        operations that don't have a natural item count up front (e.g.
        fuzzy-matching a whole file in one call)."""
        if self._indeterminate:
            self.bar.stop()
            self._indeterminate = False
        if total:
            self.bar.configure(mode="determinate", maximum=max(total, 1), value=0)
        else:
            self.bar.configure(mode="indeterminate")
            self.bar.start(12)
            self._indeterminate = True
        self.status_var.set(text)

    def update(self, current, total=None, text=None):
        """Matches the on_progress_value(current, total, text) signature
        run_in_background calls directly, so `on_progress_value=panel.update`
        can be passed straight through with no wrapper lambda needed."""
        if total:
            if int(self.bar["maximum"]) != int(total):
                self.bar.configure(mode="determinate", maximum=max(total, 1))
            if self._indeterminate:
                self.bar.stop()
                self._indeterminate = False
            self.bar["value"] = current
        if text:
            self.status_var.set(text)
        elif total:
            pct = (current / total * 100) if total else 0
            self.status_var.set(f"{current:,} / {total:,}  ({pct:.0f}%)")

    def finish(self, text="Done."):
        """Call from on_done/on_error. Leaves the bar full (or stopped, if
        it was indeterminate) with a final status line - it stays visible
        rather than disappearing, so the user can see the last run's
        outcome until they start the next one."""
        if self._indeterminate:
            self.bar.stop()
            self._indeterminate = False
            self.bar.configure(mode="determinate", maximum=100, value=100)
        else:
            try:
                self.bar["value"] = self.bar["maximum"]
            except Exception:
                pass
        self.status_var.set(text)

    def reset(self):
        if self._indeterminate:
            self.bar.stop()
            self._indeterminate = False
        self.bar.configure(mode="determinate", maximum=100, value=0)
        self.status_var.set("")


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
