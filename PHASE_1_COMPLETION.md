# GIMS — Phase 1 Completion

**Status:** Complete, released for UAT
**Date:** August 2026

This document records what shipped in Phase 1, so it's clear what UAT should
be testing against and what's explicitly out of scope until Phase 2.

---

## Modules delivered

| Module | Status | Notes |
|---|---|---|
| Outstanding Premium | ✅ Complete | Load/validate, threshold + Department/Agent/Branch filters, native Excel PivotTable reports (falls back to a static table if Excel/pywin32 isn't installed), email draft/send via Outlook |
| Loss Ratio | ✅ Complete | Excel + PDF input, Individual/Group modes, styled color-flagged summary output with live formulas |
| UPR Automation | ✅ Complete | Multi-register loading (consolidated or per-year files), period selection with coverage validation, Ledger Balance PDF auto-fill for Section 1, two-table (Gross Premium / Unearned Premium Working) output |
| Coinsurance → Premium/Claims Email Automation | ✅ Complete | Multi-sheet source detection, dual-amount-column and legacy single-amount-column layouts, one workbook per insurance company (OFC pooling), email automation |
| Coinsurance → Policy Check (Premium Register) | ✅ Complete | Fuzzy digit-run + branch/class-code matching, legacy `.xls` continuation-sheet merging (Excel's old 65,536-row-per-sheet split) |
| Coinsurance → Claims Intimation (Claims Register) | ✅ Complete | Coinsurer fan-out matching (one claim → several coinsurer files), date + Unit/Client/Department filters, per-coinsurer workbooks plus a separate Unmatched Claims workbook, email automation |
| Coinsurance → Claims Check | ✅ Complete | Reuses Policy Check's own matching engine against a leader's claims file |
| Company Aliases (merge-by-company) | ✅ Complete | Shared logic between Claims Register and Receivable modules; deterministic prefix/abbreviation-based merging, no manual review step |
| Performance work (Parts 1–5, see `PERFORMANCE_NOTES.md`) | ✅ Complete | Background threading so the GUI never freezes, fast Excel reads (`python-calamine`), `StyleIndexCache` for cheap per-cell styling, O(rows) bucketing instead of O(items × rows), parallel workbook generation across CPU cores |

The Renewal placeholder (sidebar entry + empty-state page) has been removed
from `app.py` — General Operations now has two sub-items (Outstanding
Premium, Loss Ratio) instead of three. It was never a working module, just
an unused stub, so there's nothing to carry into UAT for it.

---

## Final review before hand-off

- **Cross-module column contracts verified**: `NEW_COLUMNS` in each
  `*_processor.py` matches its corresponding `*_generator.py` exactly
  (Claims Register, Claims Check, Policy Check) — no column-order drift
  between what's computed and what's written.
- **Graceful-degradation shims verified**: every module that imports from
  `perf_utils` has a matching fallback stub if `perf_utils.py` is ever
  missing from a checkout, so a partial deployment doesn't hard-crash.
- **Filename/merge logic verified**: `company_alias_resolver.py` is the
  single source of truth for merge-key derivation and canonical filenames;
  Receivable and Claims Register both call it identically.
- **Signature storage verified shared** correctly across all four
  email-sending surfaces (Outstanding Premium, Receivable, Claims
  Register — all reading/writing the same `signature.json`).

---

## Known limitations carried into UAT

These are disclosed in the code/docs already — not regressions, but worth
confirming on the actual UAT machine:

1. **Outlook send and native Excel PivotTable/COM automation** were built
   and correctness-checked, but could only be verified in dry-run /
   fallback mode in the development environment (no Windows/Excel
   available there). These need real-machine confirmation during UAT.
2. **Parallel workbook generation** was correctness-verified (identical,
   valid output) but not throughput-verified on real multi-core hardware
   (development sandbox had 1 CPU core). Worth a timing check on a large
   batch during UAT.
3. **One disclosed performance floor**: a single coinsurer bucket with
   tens of thousands of matched rows still takes roughly a minute to
   write, since at that scale openpyxl's own per-cell object construction
   becomes the bottleneck (not something styling/threading fixes can
   remove). Not a bug — just a ceiling worth knowing about if UAT includes
   an extreme outlier file.

---

## Phase 2

Scope to be defined separately (Renewal module build-out and/or further
automation). This document marks the Phase 1 boundary for reference.
