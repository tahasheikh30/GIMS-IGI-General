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
   itself. StyleIndexCache below (see Part 4) removes that cost directly
   by reusing openpyxl's own internal style index instead of re-hashing
   the same Font/Border combination on every cell, so every generator can
   apply full styling at any row count without needing a row-count gate
   that skips it.

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

try:
    from python_calamine import CalamineWorkbook
    _CALAMINE_AVAILABLE = True
except ImportError:  # python-calamine not installed - fast_read_excel_rows()
    _CALAMINE_AVAILABLE = False  # falls back to openpyxl's read_only streaming reader below

try:
    from PIL import Image, ImageDraw, ImageTk
    _PIL_AVAILABLE = True
except ImportError:  # Pillow not installed - the progress panel's state
    _PIL_AVAILABLE = False  # chip falls back to a plain (less smooth) Canvas oval


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


# ---------------------------------------------------------------------------
# Progress bar palette - deliberately kept in sync with (but independent
# of) BalanceApp's own palette in app.py, the same way loss_ratio_generator.py
# and receivable_generator.py each keep their own small color constant
# block rather than importing app.py's (which would be a backwards,
# UI-depends-on-utility, import). Sapphire while running, muted gold
# percentage readout, success green / danger red on completion depending
# on what the caller's own finish() text says.
# ---------------------------------------------------------------------------
_PB_TRACK = "#E4E9F0"        # light neutral trough
_PB_RUNNING = "#2E6E9E"      # sapphire - matches COLOR_ACCENT
_PB_DONE = "#1E8354"         # success green - matches COLOR_SUCCESS
_PB_FAILED = "#B3402F"       # danger red - matches COLOR_DANGER
_PB_PERCENT = "#B8925A"      # muted gold - matches COLOR_GOLD
_PB_TEXT_MUTED = "#5E6E80"   # matches COLOR_MUTED
_PB_TEXT_STRONG = "#16212E"  # matches COLOR_TEXT
_PB_CARD_BG = "#FFFFFF"      # matches COLOR_PANEL
_PB_CARD_BORDER = "#DCE3ED"  # matches COLOR_BORDER


def _round_rect_points(x1, y1, x2, y2, r):
    """Corner points for a smooth-polygon rounded rectangle (a standard
    Tk technique: 12 points around the 4 corners, drawn with
    smooth=True so Tk bows each corner into a curve instead of a hard
    angle). Used for both the bar's track and its fill, which is what
    lets the bar look the same everywhere - unlike ttk.Progressbar,
    plain Canvas drawing has no dependency on the active OS theme."""
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    return [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]


_DOT_CACHE = {}


def _make_state_dot(color, size=9, supersample=4):
    """
    Render the panel's small round "state chip" as a genuinely smooth,
    anti-aliased dot, cached per color (there are only ever four: sapphire
    running / green done / red failed / muted-grey idle).

    Plain tk.Canvas.create_oval() has NO anti-aliasing at all, so at the
    small size a status chip needs to stay (a handful of pixels, so it
    doesn't compete with the status text next to it), it doesn't render
    as a circle at all - it comes out as a hard-edged, almost-square blob
    instead, since there just aren't enough pixels for Tk's line-drawing
    algorithm to trace a curve. Drawing the circle at `supersample`x the
    target resolution with Pillow and then downsampling with LANCZOS
    resampling gives every edge pixel a proper partial-opacity blend
    against the transparent background, which is what actually produces
    a round-looking dot instead of a blocky one.
    """
    key = (color, size)
    cached = _DOT_CACHE.get(key)
    if cached is not None:
        return cached
    big = size * supersample
    img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    pad = supersample  # small inset so the smoothed edge isn't clipped by the canvas bounds
    draw.ellipse([pad, pad, big - pad, big - pad], fill=color)
    img = img.resize((size, size), Image.LANCZOS)
    photo = ImageTk.PhotoImage(img)
    _DOT_CACHE[key] = photo
    return photo


class ProgressBarPanel:
    """
    A small, self-contained progress "card" meant to sit directly under a
    "Generate Excel Reports"-style button - shown ONLY while a background
    task (see run_in_background above) is actually running, and hidden
    the rest of the time so it never sits on the page as an empty bar
    with nothing to show.

    Usage is unchanged from before:
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

    `panel.frame` is still packed exactly once, permanently, by the
    caller - it's an empty outer "slot" that Tk automatically shrinks to
    zero height whenever nothing is packed inside it, so showing/hiding
    the panel never requires the caller to touch its own layout code.
    start() reveals a bordered white card inside that slot; finish()
    leaves the final result on screen just long enough to register, then
    the panel quietly collapses back to nothing on its own. The bar
    itself (track + fill, both rounded) is hand-drawn on a plain
    tk.Canvas rather than a ttk.Progressbar, so its look is identical on
    every Windows theme instead of depending on whether the active ttk
    theme happens to honor custom Progressbar colors.
    """
    def __init__(self, parent, bar_height=9, auto_hide_delay=2600):
        import tkinter as tk

        self._auto_hide_delay = auto_hide_delay
        self._hide_job = None
        self._anim_job = None
        self._indeterminate = False
        self._anim_pos = 0.0
        self._bar_height = bar_height
        self._state = "idle"
        self._fraction = 0.0
        self._visible = False

        bg = _resolve_bg_color(parent)
        outer_kwargs = {"bg": bg} if bg else {}
        # Outer slot - packed once, permanently, by the caller. Left
        # empty (no fixed size) so it collapses to 0 height whenever the
        # card below isn't packed inside it.
        self.frame = tk.Frame(parent, **outer_kwargs)

        # The visible card itself - a clean bordered white surface (the
        # same "elevated surface on a soft canvas" language used by the
        # numbered section cards elsewhere in the app) - not packed yet,
        # so the panel starts fully hidden.
        self._card = tk.Frame(self.frame, bg=_PB_CARD_BG,
                               highlightbackground=_PB_CARD_BORDER, highlightthickness=1)
        self._card_pack_kwargs = {"fill": "x"}

        inner = tk.Frame(self._card, bg=_PB_CARD_BG)
        inner.pack(fill="x", padx=14, pady=(11, 12))

        top_row = tk.Frame(inner, bg=_PB_CARD_BG)
        top_row.pack(fill="x", pady=(0, 9))

        # Small round state chip - a quiet color cue (sapphire/green/red)
        # echoing the accent language used elsewhere in the app. Rendered
        # as a genuinely smooth anti-aliased dot via Pillow (see
        # _make_state_dot) rather than a raw Canvas oval, which at this
        # size renders as a hard-edged, near-square blob instead of a
        # circle - Pillow is already a project dependency (used for the
        # header logos in app.py), so this only needs a plain Canvas-oval
        # fallback for the rare case it isn't installed.
        self._chip_uses_image = _PIL_AVAILABLE
        if self._chip_uses_image:
            self._chip_img = _make_state_dot(_PB_TEXT_MUTED)
            self._chip = tk.Label(top_row, image=self._chip_img, bg=_PB_CARD_BG, bd=0, highlightthickness=0)
            self._chip.pack(side="left", padx=(0, 9), pady=(2, 0))
        else:
            self._chip = tk.Canvas(top_row, width=8, height=8, highlightthickness=0, bg=_PB_CARD_BG)
            self._chip.pack(side="left", padx=(0, 9), pady=(2, 0))
            self._chip_id = self._chip.create_oval(0, 0, 8, 8, fill=_PB_TEXT_MUTED, outline="")

        self.status_var = tk.StringVar(value="")
        self.status_label = tk.Label(top_row, textvariable=self.status_var,
                                      font=("Segoe UI", 9), anchor="w",
                                      fg=_PB_TEXT_MUTED, bg=_PB_CARD_BG)
        self.status_label.pack(side="left", fill="x", expand=True)

        self.percent_var = tk.StringVar(value="")
        self.percent_label = tk.Label(top_row, textvariable=self.percent_var,
                                       font=("Segoe UI Semibold", 9, "bold"), anchor="e",
                                       fg=_PB_PERCENT, bg=_PB_CARD_BG)
        self.percent_label.pack(side="right")

        self._bar = tk.Canvas(inner, height=bar_height, highlightthickness=0, bg=_PB_CARD_BG)
        self._bar.pack(fill="x")
        self._bar.bind("<Configure>", self._redraw)

    # ---- visibility - the whole point of this rewrite ----
    def _show(self):
        if not self._visible:
            self._card.pack(**self._card_pack_kwargs)
            self._visible = True

    def hide(self):
        """Collapse the panel back to nothing. Safe to call any time."""
        self._cancel_hide_job()
        self._stop_animation()
        if self._visible:
            self._card.pack_forget()
            self._visible = False

    def _cancel_hide_job(self):
        if self._hide_job is not None:
            try:
                self.frame.after_cancel(self._hide_job)
            except Exception:
                pass
            self._hide_job = None

    def _schedule_hide(self):
        self._cancel_hide_job()
        if self._auto_hide_delay is not None:
            self._hide_job = self.frame.after(self._auto_hide_delay, self.hide)

    # ---- drawing ----
    def _redraw(self, event=None):
        self._bar.delete("all")
        w = self._bar.winfo_width()
        h = self._bar_height
        if w <= 2:
            return
        r = h / 2
        fill_color = {"running": _PB_RUNNING, "done": _PB_DONE,
                      "failed": _PB_FAILED, "idle": _PB_TRACK}[self._state]

        self._bar.create_polygon(_round_rect_points(0, 0, w, h, r), smooth=True, fill=_PB_TRACK, outline="")

        if self._indeterminate:
            seg_w = max(34, w * 0.22)
            x = self._anim_pos * (w + seg_w) - seg_w
            x1, x2 = max(0, x), min(w, x + seg_w)
            if x2 > x1:
                self._bar.create_polygon(_round_rect_points(x1, 0, x2, h, r), smooth=True, fill=fill_color, outline="")
        else:
            fw = max(0, min(w, w * self._fraction))
            if fw > 0:
                fr = min(r, fw / 2)  # keep the fill's own corners rounded even when very narrow
                self._bar.create_polygon(_round_rect_points(0, 0, fw, h, fr), smooth=True, fill=fill_color, outline="")

    def _stop_animation(self):
        if self._anim_job is not None:
            try:
                self.frame.after_cancel(self._anim_job)
            except Exception:
                pass
            self._anim_job = None

    def _animate(self):
        self._anim_pos += 0.02
        if self._anim_pos > 1.4:
            self._anim_pos = 0.0
        self._redraw()
        self._anim_job = self.frame.after(20, self._animate)

    # ---- state ----
    def _set_state(self, state):
        """state: 'running' | 'done' | 'failed' | 'idle' - drives both the
        bar's fill color and the small chip beside the status text."""
        self._state = state
        chip_color = {"running": _PB_RUNNING, "done": _PB_DONE,
                      "failed": _PB_FAILED, "idle": _PB_TEXT_MUTED}[state]
        if self._chip_uses_image:
            self._chip_img = _make_state_dot(chip_color)  # keep a live reference - Tk drops GC'd PhotoImages
            self._chip.configure(image=self._chip_img)
        else:
            self._chip.itemconfig(self._chip_id, fill=chip_color)

    def start(self, total=None, text="Starting..."):
        """Call once right before the background task starts. Reveals the
        panel (if it was hidden) and switches to running state. total=None
        renders an indeterminate "marching" bar, for operations without a
        natural item count up front (e.g. fuzzy-matching a whole file in
        one call)."""
        self._cancel_hide_job()
        self._show()
        self._set_state("running")
        self._fraction = 0.0
        self.status_var.set(text)
        self.status_label.configure(fg=_PB_TEXT_STRONG)
        self._stop_animation()
        if total:
            self._indeterminate = False
            self.percent_var.set("0%")
        else:
            self._indeterminate = True
            self.percent_var.set("")
            self._anim_pos = 0.0
            self._animate()
        self._redraw()

    def update(self, current, total=None, text=None):
        """Matches the on_progress_value(current, total, text) signature
        run_in_background calls directly, so `on_progress_value=panel.update`
        can be passed straight through with no wrapper lambda needed."""
        if total:
            if self._indeterminate:
                self._stop_animation()
                self._indeterminate = False
            self._fraction = max(0.0, min(1.0, current / total))
            self.percent_var.set(f"{self._fraction * 100:.0f}%")
            self._redraw()
        if text:
            self.status_var.set(text)
        elif total:
            self.status_var.set(f"{current:,} / {total:,}")

    def finish(self, text="Done."):
        """Call from on_done/on_error. Shows the final result (full bar,
        green or red depending on whether `text` itself reads as a
        failure - looks for "fail"/"error" case-insensitively) just long
        enough to register, then the panel automatically hides itself
        again a couple of seconds later."""
        self._stop_animation()
        self._indeterminate = False
        failed = any(w in text.lower() for w in ("fail", "error"))
        self._set_state("failed" if failed else "done")
        self._fraction = 1.0
        self.percent_var.set("" if failed else "100%")
        self.status_var.set(text)
        self.status_label.configure(fg=_PB_FAILED if failed else _PB_DONE)
        self._redraw()
        self._schedule_hide()

    def reset(self):
        """Hide immediately and clear all state - safe to call any time,
        e.g. right before kicking off a new run."""
        self._set_state("idle")
        self._fraction = 0.0
        self.percent_var.set("")
        self.status_var.set("")
        self.status_label.configure(fg=_PB_TEXT_MUTED)
        self.hide()


# ---------------------------------------------------------------------------
# 2. Fast streaming Excel reads (bypasses pandas' own reader for the raw
#    row scan; header detection / column-alias logic in each processor
#    module is unchanged and works the same off the returned rows)
# ---------------------------------------------------------------------------

def _read_rows_calamine(filepath, sheet_name, report=None):
    """
    Read every row of one sheet using python-calamine (a Rust-based
    spreadsheet reader) instead of openpyxl. Verified against a real
    500,000-row / 7-column production-shaped file: openpyxl's own
    read_only row iteration alone (with no other work happening around
    it) took ~43 seconds; python-calamine read the identical file in
    ~6.4 seconds - roughly 7x faster, and it was openpyxl's row
    iteration itself, not anything in this module's own code around it,
    that was the dominant cost of a "Load & Validate" click on a huge
    file. Also reads legacy .xls files directly (no xlrd / LibreOffice
    conversion needed for that path).

    Normalizes calamine's own blank-cell convention (empty string "") to
    None, matching openpyxl's read_only convention - every caller
    downstream (header detection via pd.isna(), dropna(how="all"), the
    various *_processor.py alias-matching logic) already assumes None
    for a blank cell, not "". Left un-normalized, a blank cell would
    silently stop counting as blank (e.g. a fully-empty spacer row would
    no longer get dropped by dropna(how="all")).

    Returns (rows, n_rows), same shape as the openpyxl-based reader
    below, so callers never need to know which one actually ran.
    """
    wb = CalamineWorkbook.from_path(filepath)
    if sheet_name in wb.sheet_names:
        sheet = wb.get_sheet_by_name(sheet_name)
    else:
        sheet = wb.get_sheet_by_index(sheet_name if isinstance(sheet_name, int) else 0)
    data = sheet.to_python()
    rows = [tuple((v if v != "" else None) for v in row) for row in data]
    if report:
        report(f"  ... {len(rows):,} rows read (done)")
    return rows, len(rows)


def _read_rows_openpyxl(filepath, sheet_name, report=None):
    """Original openpyxl read_only streaming reader - the fallback path
    when python-calamine isn't installed. See fast_read_excel_rows's
    docstring for the relative performance difference."""
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


def fast_read_excel_rows(filepath, sheet_name, report=None):
    """
    Stream every row of one sheet as a list of tuples, using
    python-calamine when it's installed (dramatically faster - see
    _read_rows_calamine's docstring) and transparently falling back to
    openpyxl's own read_only mode otherwise (same behavior as before
    python-calamine was added, so nothing breaks on a machine that
    doesn't have it). A calamine read error (e.g. a file format quirk
    calamine doesn't like) also falls back to the openpyxl path rather
    than failing the whole load outright.

    Returns (rows, n_rows) where rows is a list of tuples (raw cell
    values - dates come back as datetime/date objects, blanks as None,
    matching openpyxl's own convention either way).
    """
    if _CALAMINE_AVAILABLE:
        try:
            return _read_rows_calamine(filepath, sheet_name, report=report)
        except Exception:
            pass  # fall through to the openpyxl path below
    return _read_rows_openpyxl(filepath, sheet_name, report=report)


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
# 4. Fast per-cell style application
# ---------------------------------------------------------------------------

class StyleIndexCache:
    """
    Speeds up writing a large number of cells that only ever use a SMALL,
    fixed set of distinct (font, border, fill, alignment, number_format)
    combinations - the normal case for every GIMS generator's data sheets
    (e.g. "every body cell uses BODY_FONT + BORDER, plus NUM_FORMAT on
    amount columns and DATE_FORMAT on date columns" - a handful of
    combinations reused across tens of thousands of cells).

    WHY THIS EXISTS: setting cell.font = X / cell.border = Y / etc. the
    normal way is NOT cheap to repeat, even when X and Y are the exact
    same reused Python objects every time. openpyxl deduplicates styles
    internally by re-hashing the full Font/Border/Fill/Alignment/
    number-format combination on every single assignment, to find (or
    register) its slot in the workbook's shared style table - that
    hashing, not the cell-value write itself, was confirmed (via
    cProfile against a real ~2,700-row company workbook) to be the
    dominant cost of generating an output file: ~73% of total build time
    was spent inside openpyxl's own style-hashing machinery
    (styleable.py / indexed_list.py / Serialisable.__hash__), not in this
    project's own code.

    THE FIX: the first time a given style combination is used, apply it
    the normal way (so openpyxl registers it through its own public API,
    exactly as before - this is what keeps the resulting file byte-for-
    byte the kind of file openpyxl would always produce) and remember the
    resulting internal style index. Every later cell that needs the
    IDENTICAL combination gets that index assigned directly
    (`cell._style = index`), skipping the repeated hashing entirely.
    Verified correct (font/border/alignment/number_format all round-trip
    correctly through the reused index) and verified fast: a 50,000-cell
    microbenchmark of this exact pattern dropped from ~5.8s to ~0.6s.

    `_style` is an internal openpyxl attribute, not part of its public
    API - if it's ever renamed/removed by a future openpyxl version,
    every method below catches that and falls back to the always-correct
    "set each attribute the normal way" path, so this can only ever
    become slower on a future openpyxl, never wrong.

    Usage:
        cache = StyleIndexCache()
        cell = ws.cell(row=r, column=c, value=v)
        cache.apply(cell, font=BODY_FONT, border=BORDER, number_format=NUM_FORMAT)
    """
    def __init__(self):
        self._cache = {}
        self._broken = False  # set True the first time the fast path fails, so we stop retrying it

    def apply(self, cell, font=None, border=None, fill=None, alignment=None, number_format=None):
        if self._broken:
            self._apply_normal(cell, font, border, fill, alignment, number_format)
            return

        # Keyed on OBJECT IDENTITY (id()), not the Font/Border/Alignment
        # objects themselves. This matters: those are openpyxl
        # Serialisable objects with an expensive custom __hash__ (it
        # recursively hashes every sub-attribute) - using them directly
        # as dict-key components would mean Python re-hashes them on
        # EVERY apply() call just to do the cache lookup, which is
        # exactly the cost this cache exists to avoid, even on a cache
        # hit. id() is a cheap integer identity check instead. This is
        # safe as long as callers pass the same reused object instances
        # every time (which every GIMS generator already does - shared
        # module-level Font/Border constants) rather than constructing a
        # fresh-but-equal object on each call.
        key = (id(font), id(border), id(fill), id(alignment), number_format)
        style_id = self._cache.get(key)
        if style_id is not None:
            try:
                cell._style = style_id
                return
            except Exception:
                self._broken = True  # give up on the fast path for the rest of this cache's life
                self._apply_normal(cell, font, border, fill, alignment, number_format)
                return

        self._apply_normal(cell, font, border, fill, alignment, number_format)
        try:
            self._cache[key] = cell._style
        except Exception:
            self._broken = True

    @staticmethod
    def _apply_normal(cell, font, border, fill, alignment, number_format):
        if font is not None:
            cell.font = font
        if border is not None:
            cell.border = border
        if fill is not None:
            cell.fill = fill
        if alignment is not None:
            cell.alignment = alignment
        if number_format is not None:
            cell.number_format = number_format
