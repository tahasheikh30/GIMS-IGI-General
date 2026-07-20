"""
Expected Items Register + auto-reconciliation, for the Bank Position
module's "Major Receivables" / "Major Payments" sections.

WHY THIS EXISTS (see the design discussion this was built from): a flat
manual list only handles the "add" side - someone still has to remember
to remove an entry once it clears, or notice it's gone overdue. This
module keeps manual entry as the source of truth for what's EXPECTED
(nothing in a bank statement can tell you that "Al Mumtaz Electronics
owes us 3,000,000" before it arrives), but automates the two things that
actually rot when left manual:
  1. CLEARING - each morning, every still-open item is checked against
     that day's parsed statement transactions (see
     bank_statement_processor.py) for a matching credit/debit. A
     confident match auto-marks the item Received/Paid; an ambiguous or
     no-match result leaves it open rather than guessing - same
     never-force-a-bad-match principle as
     premium_register_processor.match_one().
  2. AGING - anything still open past its expected date is flagged as
     overdue automatically instead of looking identical to a fresh entry.
As a bonus, statement credits/debits that don't match anything in the
register are surfaced separately as "unexplained" - money that moved
without anyone having logged it as expected.

Storage is a flat JSON file (same pattern as receivable_run_log.py /
run_log.py) - fine for a register that's realistically dozens, not
millions, of rows.
"""
import json
import os
import uuid
from datetime import date, datetime

import pandas as pd

STORE_FILE = os.path.join(os.path.dirname(__file__), "bank_position_expected_items.json")

CATEGORY_RECEIVABLE = "Receivable"
CATEGORY_PAYMENT = "Payment"

STATUS_OPEN = "Open"
STATUS_CLEARED = "Cleared"          # auto-matched against a statement
STATUS_CONFIRMED_MANUAL = "Confirmed (manual)"  # a human resolved it by hand instead

# Matching thresholds - deliberately tight, since a wrong auto-match on a
# CFO-facing report is worse than one that stays open an extra day for a
# human to clear by hand. Both the percentage and the absolute floor are
# applied (whichever is looser), so a very small expected amount doesn't
# become impossible to match on percentage alone.
DEFAULT_AMOUNT_TOLERANCE_PCT = 0.005   # 0.5%
DEFAULT_AMOUNT_TOLERANCE_ABS = 50.0    # PKR - covers bank-charge-sized rounding
DEFAULT_DATE_WINDOW_DAYS = 10          # how far from the expected date a transaction can still count
DEFAULT_UNEXPLAINED_THRESHOLD = 500000.0  # PKR - credits/debits above this get surfaced if unmatched


# ---------------------------------------------------------------------------
# Storage
# ---------------------------------------------------------------------------

def _load_raw():
    if os.path.exists(STORE_FILE):
        with open(STORE_FILE, "r") as f:
            return json.load(f)
    return []


def _save_raw(items):
    with open(STORE_FILE, "w") as f:
        json.dump(items, f, indent=2, default=str)


def _serialize_date(d):
    if d is None:
        return None
    if isinstance(d, (date, datetime)):
        return d.isoformat()
    return str(d)


def _parse_date(s):
    if not s:
        return None
    if isinstance(s, (date, datetime)):
        return s if isinstance(s, date) and not isinstance(s, datetime) else s.date()
    return datetime.fromisoformat(str(s)).date()


# ---------------------------------------------------------------------------
# CRUD
# ---------------------------------------------------------------------------

def add_item(category, counterparty, expected_amount, expected_date, bank_key=None, note=""):
    """
    category: CATEGORY_RECEIVABLE or CATEGORY_PAYMENT
    bank_key: optional (e.g. "js") - narrows matching to that bank's
        statement only. Leave None if you don't know which bank it'll
        land in yet; matching then searches every parsed statement.
    Returns the new item's id.
    """
    if category not in (CATEGORY_RECEIVABLE, CATEGORY_PAYMENT):
        raise ValueError(f"category must be '{CATEGORY_RECEIVABLE}' or '{CATEGORY_PAYMENT}', got {category!r}")

    items = _load_raw()
    item_id = str(uuid.uuid4())[:8]
    items.append({
        "id": item_id,
        "category": category,
        "counterparty": counterparty,
        "bank_key": bank_key,
        "expected_amount": float(expected_amount),
        "expected_date": _serialize_date(expected_date),
        "note": note,
        "status": STATUS_OPEN,
        "matched_date": None,
        "matched_amount": None,
        "matched_bank_key": None,
        "matched_description": None,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    })
    _save_raw(items)
    return item_id


def remove_item(item_id):
    items = _load_raw()
    items = [i for i in items if i["id"] != item_id]
    _save_raw(items)


def confirm_manually(item_id, matched_date=None, note_suffix=None):
    """For an item a human resolved by hand (e.g. confirmed by phone with
    the bank) rather than one the auto-matcher found - keeps the register
    accurate even when reconciliation genuinely can't find it (different
    narration, wrong amount was originally logged, etc.)."""
    items = _load_raw()
    for i in items:
        if i["id"] == item_id:
            i["status"] = STATUS_CONFIRMED_MANUAL
            i["matched_date"] = _serialize_date(matched_date or date.today())
            if note_suffix:
                i["note"] = f"{i['note']} | {note_suffix}".strip(" |")
    _save_raw(items)


def list_items(category=None, include_resolved=True):
    """Returns a list of dicts, each with an added `days_overdue` (None if
    not open or not yet due) computed fresh against today's date rather
    than stored, so it's always current."""
    items = _load_raw()
    today = date.today()
    out = []
    for i in items:
        if category and i["category"] != category:
            continue
        if not include_resolved and i["status"] != STATUS_OPEN:
            continue
        item = dict(i)
        exp_date = _parse_date(i["expected_date"])
        if i["status"] == STATUS_OPEN and exp_date and exp_date < today:
            item["days_overdue"] = (today - exp_date).days
        else:
            item["days_overdue"] = None
        out.append(item)
    return out


# ---------------------------------------------------------------------------
# Reconciliation
# ---------------------------------------------------------------------------

def _candidate_transactions(statement_results, category, bank_key_filter=None):
    """Yield (bank_key, row_index, date, amount, description) for every
    transaction row that's the relevant side (credit for a Receivable,
    debit for a Payment) of every given statement - optionally narrowed
    to one bank."""
    amount_col = "credit" if category == CATEGORY_RECEIVABLE else "debit"
    for result in statement_results:
        if bank_key_filter and result.bank_key != bank_key_filter:
            continue
        txns = result.transactions
        for idx, row in txns.iterrows():
            amt = row[amount_col]
            if amt is None or pd.isna(amt) or amt == 0:
                continue
            yield result.bank_key, idx, row["date"], amt, row["description"]


def reconcile(statement_results,
              amount_tolerance_pct=DEFAULT_AMOUNT_TOLERANCE_PCT,
              amount_tolerance_abs=DEFAULT_AMOUNT_TOLERANCE_ABS,
              date_window_days=DEFAULT_DATE_WINDOW_DAYS,
              unexplained_threshold=DEFAULT_UNEXPLAINED_THRESHOLD):
    """
    Runs every open register item against the given parsed statements
    (StatementResult objects from bank_statement_processor.py).

    For each open item: gathers every candidate transaction (right bank if
    the item specifies one, otherwise every statement) within the amount
    tolerance AND the date window of the expected date. Auto-clears the
    item ONLY if there's exactly one candidate - two or more equally
    plausible candidates are left open and flagged as ambiguous, same
    never-force-a-guess principle as premium_register_processor.match_one().

    Returns a dict:
      {
        "matched": [{"item_id", "counterparty", "matched_amount",
                      "matched_date", "matched_bank_key",
                      "matched_description"}, ...],
        "ambiguous": [{"item_id", "counterparty", "candidates": [...]}, ...],
        "still_open": [item dicts with no candidates found at all],
        "unexplained_credits": [{"bank_key", "date", "amount",
                                   "description"}, ...] - large statement
            credits/debits that matched no register item at all, category
            "Receivable" ones only (see note below),
      }

    unexplained_credits only scans the CREDIT side (money coming in) -
    the intent is "here's cash that arrived without anyone expecting it",
    which is only meaningful for receivables; a large unexpected DEBIT is
    a different problem (an unauthorized/unlogged payment going out) and
    is deliberately out of scope here to avoid conflating two different
    kinds of alert under one label.
    """
    open_items = [i for i in _load_raw() if i["status"] == STATUS_OPEN]
    matched_ids = set()
    matched_out, ambiguous_out, still_open_out = [], [], []
    consumed_txn_keys = set()  # (bank_key, row_index) already claimed by a match

    for item in open_items:
        exp_date = _parse_date(item["expected_date"])
        expected_amount = item["expected_amount"]
        tolerance = max(amount_tolerance_abs, amount_tolerance_pct * abs(expected_amount))

        candidates = []
        for bank_key, idx, txn_date, amount, description in _candidate_transactions(
                statement_results, item["category"], bank_key_filter=item.get("bank_key")):
            if abs(amount - expected_amount) > tolerance:
                continue
            if exp_date is not None and txn_date is not None and pd.notna(txn_date):
                txn_date_only = txn_date.date() if hasattr(txn_date, "date") else txn_date
                if abs((txn_date_only - exp_date).days) > date_window_days:
                    continue
            candidates.append({
                "bank_key": bank_key, "row_index": idx,
                "date": txn_date.date() if hasattr(txn_date, "date") else txn_date,
                "amount": amount, "description": description,
            })

        if len(candidates) == 1:
            c = candidates[0]
            _mark_cleared(item["id"], c["date"], c["amount"], c["bank_key"], c["description"])
            matched_ids.add(item["id"])
            consumed_txn_keys.add((c["bank_key"], c["row_index"]))
            matched_out.append({
                "item_id": item["id"], "counterparty": item["counterparty"],
                "matched_amount": c["amount"], "matched_date": c["date"],
                "matched_bank_key": c["bank_key"], "matched_description": c["description"],
            })
        elif len(candidates) > 1:
            ambiguous_out.append({
                "item_id": item["id"], "counterparty": item["counterparty"],
                "expected_amount": expected_amount, "candidates": candidates,
            })
        else:
            still_open_out.append(item)

    unexplained_credits = []
    for bank_key, idx, txn_date, amount, description in _candidate_transactions(
            statement_results, CATEGORY_RECEIVABLE):
        if (bank_key, idx) in consumed_txn_keys:
            continue
        if amount < unexplained_threshold:
            continue
        unexplained_credits.append({
            "bank_key": bank_key,
            "date": txn_date.date() if hasattr(txn_date, "date") else txn_date,
            "amount": amount, "description": description,
        })
    unexplained_credits.sort(key=lambda x: x["amount"], reverse=True)

    return {
        "matched": matched_out,
        "ambiguous": ambiguous_out,
        "still_open": still_open_out,
        "unexplained_credits": unexplained_credits,
    }


def _mark_cleared(item_id, matched_date, matched_amount, matched_bank_key, matched_description):
    items = _load_raw()
    for i in items:
        if i["id"] == item_id:
            i["status"] = STATUS_CLEARED
            i["matched_date"] = _serialize_date(matched_date)
            i["matched_amount"] = matched_amount
            i["matched_bank_key"] = matched_bank_key
            i["matched_description"] = matched_description
    _save_raw(items)


if __name__ == "__main__":
    import sys
    sys.path.insert(0, os.path.dirname(__file__))
    from bank_statement_processor import parse_batch

    # Reset the store for a clean demo run
    if os.path.exists(STORE_FILE):
        os.remove(STORE_FILE)

    samples = {
        "hbl": "/mnt/user-data/uploads/HBL.xls",
        "mcb": "/mnt/user-data/uploads/MCB_July_5655.xls",
        "faysal": "/mnt/user-data/uploads/Faysal_6018-.xls",
        "bafl": "/mnt/user-data/uploads/BAFL_July_till_date.csv",
        "js": "/mnt/user-data/uploads/JS_107791.xlsx",
    }
    results, errors = parse_batch([(fp, bank) for bank, fp in samples.items()])
    print(f"Parsed {len(results)} statements, {len(errors)} errors\n")

    # 1) A real, exact-match case (JS Bank, Bayer Pakistan inward RTGS, 08-Jul)
    id1 = add_item(CATEGORY_RECEIVABLE, "Bayer Pakistan", 1890702.00, date(2026, 7, 8), bank_key="js",
                    note="Premium payment expected per collections team")

    # 2) A real match but where we intentionally logged the wrong bank -
    #    should stay open/ambiguous-free with zero candidates, proving the
    #    matcher doesn't go hunting outside a bank_key filter.
    id2 = add_item(CATEGORY_RECEIVABLE, "Bayer Pakistan (wrong bank test)", 1890702.00, date(2026, 7, 8), bank_key="hbl")

    # 3) Same amount, no bank_key specified - should find it via a full scan
    id3 = add_item(CATEGORY_RECEIVABLE, "Bayer Pakistan (no bank specified)", 1890702.00, date(2026, 7, 8))

    # 4) Something that will never clear (nothing like it in any statement)
    id4 = add_item(CATEGORY_RECEIVABLE, "Al Mumtaz Electronics", 3000000.00, date(2026, 7, 9), bank_key="hbl",
                    note="Per Summary sheet sample")

    result = reconcile(results)

    print("=== MATCHED ===")
    for m in result["matched"]:
        print(f"  {m['counterparty']}: {m['matched_amount']:,.2f} on {m['matched_date']} ({m['matched_bank_key']}) - {m['matched_description'][:60]}")

    print("\n=== AMBIGUOUS ===")
    for a in result["ambiguous"]:
        print(f"  {a['counterparty']}: {len(a['candidates'])} candidates")

    print("\n=== STILL OPEN (no candidates) ===")
    for i in result["still_open"]:
        print(f"  {i['counterparty']}: expected {i['expected_amount']:,.2f} on {i['expected_date']}")

    print(f"\n=== UNEXPLAINED LARGE CREDITS (>{DEFAULT_UNEXPLAINED_THRESHOLD:,.0f}, top 5) ===")
    for u in result["unexplained_credits"][:5]:
        print(f"  {u['bank_key']}: {u['amount']:,.2f} on {u['date']} - {u['description'][:70]}")

    print("\n=== Register state after reconcile ===")
    for i in list_items():
        overdue = f" ({i['days_overdue']} days overdue)" if i["days_overdue"] else ""
        print(f"  [{i['status']}]{overdue} {i['counterparty']}: {i['expected_amount']:,.2f}")

    os.remove(STORE_FILE)  # clean up demo artifact
