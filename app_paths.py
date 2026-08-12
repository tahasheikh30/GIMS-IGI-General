"""
Where GIMS keeps its own small, per-user runtime state: your saved
signature, saved email templates, and the four run-history logs. This is
deliberately kept OUTSIDE the folder the application code itself lives
in, so it is never inside the OneDrive folder that gets synced across
laptops.

WHY THIS MATTERS FOR MULTI-LAPTOP ONEDRIVE USE
------------------------------------------------
If GIMS is shared across laptops by putting the whole project folder in
OneDrive, every laptop that opens the app also opens, and re-saves, the
same signature.json, run_history.json, and similar files. OneDrive syncs
a changed file the moment it is saved. It does not merge two laptops'
concurrent changes to the same JSON file, it just keeps whichever save
"won" the race, or, if both laptops were offline and reconnect at
different times, creates a confusing second file such as
"run_history-LAPTOPNAME.json" next to the original. Either way, this is
the real mechanism behind "the app corrupts when moved between
laptops". It is not the code failing, it is two machines racing over
the same small file through a sync engine that was never told these
files are meant to be per-machine.

The fix here has two parts:
  1. Each laptop gets its OWN copy of this state, stored under its own
     local, non-synced, %LOCALAPPDATA%\\GIMS folder, so there is nothing
     left for two laptops to race over. This is the same pattern almost
     every desktop app uses for its own local settings.
  2. Every write to any of these files goes through safe_io.py's atomic
     write helper (write to a temp file, then rename into place) as a
     second, independent layer of protection, so even a mid-save crash
     or a forced shutdown can never leave a half-written, unreadable
     file behind.

The application CODE itself (app.py and every other *.py module) is
still perfectly safe to keep in OneDrive and share across laptops that
way. It is read-only at runtime, nothing about running GIMS ever writes
back into its own install folder, so there is no write-race risk there
at all. Only the small state files this module manages needed to move.
"""
import os
import shutil


def get_local_data_dir():
    """Per-machine folder for GIMS' own settings/history, created on
    first use. %LOCALAPPDATA%\\GIMS on Windows, ~/GIMS as a fallback on
    any other OS (e.g. if this is ever run outside Windows for testing)."""
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "GIMS")
    os.makedirs(path, exist_ok=True)
    return path


def resolve_state_file(filename, legacy_dir=None):
    """
    Return the full path to a state file (e.g. "signature.json") inside
    the local, per-machine data folder, and, the FIRST time this file is
    asked for on a given laptop, transparently copy over whatever
    same-named file already exists next to the application code
    (legacy_dir, normally the folder app.py itself lives in), if any.

    This is a one-time migration so nobody's already-saved signature or
    run history silently disappears the first time they run a version of
    GIMS built with this fix. It is copied into the new per-machine
    location once, then that laptop uses its own local copy from then
    on. The original, shared-via-OneDrive copy is left untouched, in
    case it is still needed as a reference, but is never written to
    again by this laptop.
    """
    target_dir = get_local_data_dir()
    target_path = os.path.join(target_dir, filename)

    if not os.path.exists(target_path) and legacy_dir:
        legacy_path = os.path.join(legacy_dir, filename)
        if os.path.exists(legacy_path):
            try:
                shutil.copyfile(legacy_path, target_path)
            except OSError:
                pass  # best-effort migration only, a fresh file is still fine

    return target_path
