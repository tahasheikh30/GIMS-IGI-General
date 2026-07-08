"""Simple JSON-based audit log of past runs."""
import json
import os
from datetime import datetime

LOG_FILE = os.path.join(os.path.dirname(__file__), "run_history.json")


def load_history():
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, "r") as f:
            return json.load(f)
    return []


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
    with open(LOG_FILE, "w") as f:
        json.dump(history, f, indent=2)
