"""
Loss Ratio input processing - matches IGI's real Loss Ratio formula and
report structure (verified against a real "before" source PDF and "after"
Group Summary Excel output):

    Loss Ratio (%) = (Incurred Loss + Commission Accrual) / Gross Premium Written * 100
                    = (B + C) / A * 100

Source files (per-client system exports) commonly list one row per policy
year plus "Client Total" / "Grand Total" rows summarizing the full period.
To avoid double-counting, this module detects a Grand/Client Total row when
present and uses it directly; otherwise it sums all data rows.

Column normalization: matching is case-insensitive and whitespace-trimmed,
and strips any parenthetical suffix (e.g. "Gross Premium Written ( A )"
still matches "gross premium written").
"""
import os
import re
import pandas as pd

REQUIRED_CANONICAL = ["Gross Premium Written", "Incurred Loss"]
OPTIONAL_CANONICAL = ["Commission Accrual"]

# Known header aliases -> canonical column name. Add more here as real
# sample files reveal other naming variations.
ALIASES = {
    "Company Name": [
        "company name", "company", "client", "client name", "participant",
        "insured", "insured name", "client/participant", "policyholder",
        "policy holder",
    ],
    "Gross Premium Written": [
        "gross premium written", "gross premium written a", "gp",
        "gross premium", "grosspremium", "gross prem", "g. premium",
        "gross written premium", "gwp",
    ],
    "Incurred Loss": [
        "incurred loss", "incurred loss b", "claim", "claims", "claim paid",
        "claims paid", "paid claims", "paid claim", "incurred claims",
        "total loss",
    ],
    "Commission Accrual": [
        "commission accrual", "commission accrual c", "commission",
        "total comm", "comm", "commission paid",
    ],
}

_ALIAS_LOOKUP = {}
for canonical, aliases in ALIASES.items():
    for a in aliases:
        _ALIAS_LOOKUP[a] = canonical

# Row labels that mark an already-aggregated total row in a multi-year
# source export (e.g. "Client Total", "Grand Total") - matched as a
# case-insensitive substring.
TOTAL_ROW_MARKERS = ["grand total", "client total"]


def _normalize_header(header):
    """Lowercase, strip, collapse whitespace, and drop parentheses/symbols
    so 'Gross Premium Written ( A )' still matches 'gross premium written'."""
    h = str(header)
    h = re.sub(r"[()]", " ", h)              # "( A )" -> " A "
    h = re.sub(r"\bA\b|\bB\b|\bC\b", " ", h)  # drop lone letter labels like "A"/"B"/"C"
    h = re.sub(r"\s+", " ", h).strip().lower()
    return h


def _map_columns_to_canonical(columns):
    """Return {original_column_name: canonical_name} for every column that
    matches a known alias. Unrecognized columns are simply omitted."""
    mapping = {}
    for col in columns:
        norm = _normalize_header(col)
        if norm in _ALIAS_LOOKUP:
            mapping[col] = _ALIAS_LOOKUP[norm]
    return mapping


def _clean_company_name(name):
    """Some source exports prefix the client name with an account/client
    code in the same cell (e.g. '1120106182 SAYA WEAVING MILLS (PVT)
    LIMITED'). Strip a long leading numeric code if present."""
    name = str(name).strip()
    stripped = re.sub(r"^\d{5,}\s+", "", name)
    return stripped if stripped else name


class LossRatioInputError(Exception):
    """Raised for a specific file so the caller can show a clear,
    file-scoped error message instead of crashing the whole batch."""
    def __init__(self, filepath, message):
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
        self.message = message
        super().__init__(f"{self.filename}: {message}")


def _find_total_row_mask(df, label_col):
    """Return a boolean mask marking rows whose label cell contains a
    Grand Total / Client Total marker."""
    if label_col not in df.columns:
        return pd.Series(False, index=df.index)
    labels = df[label_col].astype(str).str.lower()
    mask = pd.Series(False, index=df.index)
    for marker in TOTAL_ROW_MARKERS:
        mask = mask | labels.str.contains(marker, na=False)
    return mask


def read_and_aggregate_file(filepath):
    """
    Read one Excel file and collapse it to a single company row:
      {Company Name, Gross Premium Written, Incurred Loss, Commission Accrual,
       loss_ratio (fraction, or None if Gross Premium Written is 0),
       warning, source_file}

    Prefers an explicit "Grand Total" / "Client Total" row if the file has
    one (common in multi-year per-client exports) to avoid double-counting
    per-year rows; otherwise sums all data rows.

    Raises LossRatioInputError with a clear filename + reason on any
    problem (missing file, unreadable file, missing required column,
    non-numeric values) rather than failing silently or crashing.
    """
    if not os.path.exists(filepath):
        raise LossRatioInputError(filepath, "File not found")

    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        raise LossRatioInputError(filepath, f"Could not read Excel file ({e})")

    if df.empty:
        raise LossRatioInputError(filepath, "File has no data rows")

    colmap = _map_columns_to_canonical(df.columns)
    df = df.rename(columns=colmap)

    missing = [c for c in REQUIRED_CANONICAL if c not in df.columns]
    if missing:
        raise LossRatioInputError(
            filepath, f"Missing required column(s): {', '.join(missing)}"
        )

    has_commission_col = "Commission Accrual" in df.columns
    if not has_commission_col:
        df["Commission Accrual"] = 0

    numeric_cols_to_check = REQUIRED_CANONICAL + (["Commission Accrual"] if has_commission_col else [])
    for col in numeric_cols_to_check:
        original_non_null = df[col].notna()
        numeric = pd.to_numeric(df[col], errors="coerce")
        bad_count = int((original_non_null & numeric.isna()).sum())
        if bad_count > 0:
            raise LossRatioInputError(
                filepath, f"Column '{col}' contains {bad_count} non-numeric value(s)"
            )
        df[col] = numeric.fillna(0)

    label_col = "Company Name" if "Company Name" in df.columns else df.columns[0]
    total_mask = _find_total_row_mask(df, label_col)

    warnings = []
    if not has_commission_col:
        warnings.append("No Commission Accrual column found - treated as 0")

    if total_mask.any():
        candidates = df[total_mask]
        grand = candidates[candidates[label_col].astype(str).str.lower().str.contains("grand total")]
        total_row = (grand.iloc[-1] if len(grand) else candidates.iloc[-1])
        gross_premium = float(total_row["Gross Premium Written"])
        incurred_loss = float(total_row["Incurred Loss"])
        commission = float(total_row["Commission Accrual"])
        warnings.append("Used the file's Grand/Client Total row rather than summing all rows")
    else:
        gross_premium = float(df["Gross Premium Written"].sum())
        incurred_loss = float(df["Incurred Loss"].sum())
        commission = float(df["Commission Accrual"].sum())

    if "Company Name" in df.columns:
        data_rows = df[~total_mask] if total_mask.any() else df
        names = data_rows["Company Name"].dropna().astype(str).map(_clean_company_name)
        names = names[names != ""]
        unique_names = names.unique()
        if len(unique_names) == 1:
            company_name = unique_names[0]
        else:
            company_name = os.path.splitext(os.path.basename(filepath))[0]
    else:
        company_name = os.path.splitext(os.path.basename(filepath))[0]

    row = {
        "Company Name": company_name,
        "Gross Premium Written": gross_premium,
        "Incurred Loss": incurred_loss,
        "Commission Accrual": commission,
        "source_file": os.path.basename(filepath),
    }

    if gross_premium == 0:
        row["loss_ratio"] = None
        warnings.append("Gross Premium Written is zero - Loss Ratio not computable")
    else:
        row["loss_ratio"] = (incurred_loss + commission) / gross_premium

    row["warning"] = "; ".join(warnings) if warnings else None
    return row


def process_files(filepaths):
    """
    Process a batch of files. Returns (rows, errors):
      rows: list of successfully aggregated company rows
      errors: list of LossRatioInputError for files that failed - a
              failure in one file doesn't stop the others from processing.
    """
    rows = []
    errors = []
    for fp in filepaths:
        try:
            rows.append(read_and_aggregate_file(fp))
        except LossRatioInputError as e:
            errors.append(e)
        except Exception as e:
            errors.append(LossRatioInputError(fp, f"Unexpected error: {e}"))
    return rows, errors


if __name__ == "__main__":
    import tempfile

    tmp_dir = tempfile.mkdtemp()

    # Case 1: multi-year file with a Grand Total row (mirrors the real PDF
    # structure) - must use the Grand Total row, not sum all rows.
    df1 = pd.DataFrame({
        "Client/Participant": [
            "1120106182 SAYA WEAVING MILLS (PVT) LIMITED",
            "2024", "2025", "2026", "Client Total", "Grand Total",
        ],
        "Gross Premium Written ( A )": [None, 179985, 491114, 3589331, 4260430, 4260430],
        "Incurred Loss ( B )": [None, 0, 0, 251535, 251535, 251535],
        "Commission Accrual ( C )": [None, 19801, 56086, 381281, 457168, 457168],
    })
    fp1 = os.path.join(tmp_dir, "SAYA.xlsx")
    df1.to_excel(fp1, index=False)

    # Case 2: simple single-row file, alias headers, no commission column
    df2 = pd.DataFrame({
        "Company Name": ["POPULAR FOILS AND METALS"],
        "GP": [836035],
        "Claim": [0],
    })
    fp2 = os.path.join(tmp_dir, "POPULAR_FOILS.xlsx")
    df2.to_excel(fp2, index=False)

    rows, errors = process_files([fp1, fp2])
    print(f"{len(rows)} rows OK, {len(errors)} errors\n")
    for r in rows:
        lr = f"{r['loss_ratio']*100:.4f}%" if r["loss_ratio"] is not None else "N/A"
        print(f"  {r['Company Name']:35} GP={r['Gross Premium Written']:>12,.0f}  Loss={r['Incurred Loss']:>10,.0f}  Comm={r['Commission Accrual']:>10,.0f}  LR={lr}")
        if r["warning"]:
            print(f"      warning: {r['warning']}")
    for e in errors:
        print("  ERROR:", e)
