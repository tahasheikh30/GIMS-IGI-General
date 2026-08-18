"""
Crash-safe, OneDrive-sync-safe file writing helpers.

WHY THIS EXISTS
----------------
Every small persistence file GIMS writes on its own (signature.json,
templates.json, run_history.json, receivable_run_history.json,
claims_register_run_history.json) used to be written with a plain
open(path, "w") followed by json.dump(). That pattern has two real
corruption risks once a file like this can be touched by a sync engine
or opened from more than one place:

1. A plain open(path, "w") truncates the file to zero bytes before a
   single byte of the new content is written back. If anything reads
   the file in that window (OneDrive's own sync watcher, or the app
   itself on a second laptop), it can see, and sync out, a half-written
   or completely empty file. The next launch, on any laptop, then fails
   to parse the JSON, or silently loses history, even though nothing
   about that particular run actually failed.
2. A process killed mid-write (a crash, a forced shutdown, Windows
   Update rebooting the machine) leaves the file in whatever partial
   state it was in at that instant, since there is no way to know a
   plain write ever finished.

THE FIX
-------
Write the new content to a temporary file in the SAME directory, flush
and fsync it to disk, then atomically rename it over the real target
with os.replace(). os.replace() is atomic on both Windows and POSIX, any
reader (including OneDrive's own sync watcher) either sees the old,
complete file or the new, complete file, never a partially written one.

Paired with app_paths.py (which keeps these files in a local, per-machine
folder rather than inside the synced project folder in the first place),
this removes both the root cause and the failure mode.
"""
import json
import os
import tempfile


def atomic_write_text(path, text, encoding="utf-8"):
    """Write `text` to `path` without ever leaving a truncated or
    partial file on disk, even if the process is killed mid-write or a
    sync engine reads the file at the wrong moment."""
    directory = os.path.dirname(os.path.abspath(path)) or "."
    os.makedirs(directory, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix=".tmp_", suffix=".part", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding=encoding) as f:
            f.write(text)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, path)  # atomic on Windows and POSIX
    except Exception:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise


def atomic_write_json(path, data, indent=2):
    atomic_write_text(path, json.dumps(data, indent=indent))


def safe_read_json(path, default=None):
    """Read JSON, tolerating a leftover corrupted file from before this
    fix was applied (e.g. a half-written file synced down from another
    laptop before atomic writes existed). Returns `default` instead of
    crashing the app on launch."""
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError, UnicodeDecodeError, ValueError):
        return default
