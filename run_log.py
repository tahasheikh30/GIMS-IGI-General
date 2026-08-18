"""Simple JSON-based audit log of past runs (Outstanding Premium module).

Stored per-machine (see app_paths.py) and written atomically (see
safe_io.py), so this file can never be corrupted by OneDrive syncing a
half-written save, and two laptops never race over the same file.
"""
import os
from datetime import datetime

import app_paths
import safe_io

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = app_paths.resolve_state_file("run_history.json", legacy_dir=_APP_DIR)


def load_history():
    return safe_io.safe_read_json(LOG_FILE, default=[])


def log_run(threshold, company_filter, department_filter, num_qualifying, num_sent, num_failed, dry_run):
    history = load_history()
    history.append({
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "threshold": threshold,
        "company_filter": company_filter,
        "department_filter": department_filter,
        "num_qualifying": num_qualifying,
        "num_sent": num_sent,
        "num_failed": num_failed,
        "dry_run": dry_run,
    })
    safe_io.atomic_write_json(LOG_FILE, history)
