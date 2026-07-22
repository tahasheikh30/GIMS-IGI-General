"""Simple JSON-based audit log for Claims Register runs, mirroring
receivable_run_log.py's / run_log.py's pattern but kept in its own file
so it never collides with the other modules' run history."""
import json
import os
from datetime import datetime

LOG_FILE = os.path.join(os.path.dirname(__file__), "claims_register_run_history.json")


def load_history():
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, "r") as f:
            return json.load(f)
    return []


def log_run(num_claims_loaded, num_matched_claims, num_unmatched_claims, num_groups,
            num_files_generated, num_sent, num_failed, num_missing_email, dry_run, errors=None):
    history = load_history()
    history.append({
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "num_claims_loaded": num_claims_loaded,
        "num_matched_claims": num_matched_claims,
        "num_unmatched_claims": num_unmatched_claims,
        "num_groups": num_groups,
        "num_files_generated": num_files_generated,
        "num_sent": num_sent,
        "num_failed": num_failed,
        "num_missing_email": num_missing_email,
        "dry_run": dry_run,
        "errors": errors or [],
    })
    with open(LOG_FILE, "w") as f:
        json.dump(history, f, indent=2)
