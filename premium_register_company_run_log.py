"""Simple JSON-based audit log for Premium Register (Coinsurance) company-
split runs, mirroring claims_register_run_log.py's / receivable_run_log.py's
pattern but kept in its own file so it never collides with either module's
run history.

Stored per-machine (see app_paths.py) and written atomically (see
safe_io.py), so this file can never be corrupted by OneDrive syncing a
half-written save, and two laptops never race over the same file.
"""
import os

import app_paths
import safe_io
from datetime import datetime

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = app_paths.resolve_state_file("premium_register_company_run_history.json", legacy_dir=_APP_DIR)


def load_history():
    return safe_io.safe_read_json(LOG_FILE, default=[])


def log_run(num_rows_loaded, num_coinsurer_rows, num_groups, num_files_generated,
            num_sent, num_failed, num_missing_email, dry_run, errors=None):
    history = load_history()
    history.append({
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "num_rows_loaded": num_rows_loaded,
        "num_coinsurer_rows": num_coinsurer_rows,
        "num_groups": num_groups,
        "num_files_generated": num_files_generated,
        "num_sent": num_sent,
        "num_failed": num_failed,
        "num_missing_email": num_missing_email,
        "dry_run": dry_run,
        "errors": errors or [],
    })
    safe_io.atomic_write_json(LOG_FILE, history)
