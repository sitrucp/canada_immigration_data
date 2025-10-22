#!/usr/bin/env python3
"""
Extract & clean: Temporary Residents - Study, by Province/Territory and Study level

What it does
------------
- Reads the first worksheet of the raw Excel with multi-row headers (Year / Quarter / Month).
- Builds a single header row of YYYY-MM month columns, keeping only monthly columns.
- Drops quarterly and yearly subtotal columns.
- Keeps two hierarchy columns: province_territory (col A) and study_level (col B).
- Backfills province_territory within groups (merged-cell style), mirroring IMP/TFW.
- Detects structural totals and emits total_flag, with the special anomaly retained as non-total.
- Cleans numeric cells, converting "--" -> 2 and blanks -> 0.
- Writes the cleaned, unpivoted data to CSV with columns:
  province_territory, study_level, total_flag, year, month, value

Usage
-----
python extract_study.py                                # default input/output
python extract_study.py input.xlsx output.csv          # custom paths

Notes
-----
- Uses identical data cleaning and total flag logic as the other extractors.
- Special-case row: "Province/territory not stated Total" must be retained with total_flag=False.
"""

import argparse
import re
from typing import List, Tuple

import pandas as pd


# ----------------------------
# Header detection utilities
# ----------------------------

_MONTH_MAP = {
    "Jan": "01", "Feb": "02", "Mar": "03",
    "Apr": "04", "May": "05", "Jun": "06",
    "Jul": "07", "Aug": "08", "Sep": "09",
    "Oct": "10", "Nov": "11", "Dec": "12",
}


def _is_year_like(x: object) -> bool:
    try:
        y = int(str(x).strip())
        return 1900 <= y <= 2100
    except Exception:
        return False


def _is_month_abbrev(x: object) -> bool:
    if pd.isna(x):
        return False
    return str(x).strip() in _MONTH_MAP


def detect_year_and_month_rows(raw: pd.DataFrame, max_scan_rows: int = 10) -> Tuple[int, int]:
    """
    Heuristically detect the header rows for year and month.
    Returns (year_row_idx, month_row_idx).
    Assumes month row appears after the year row.
    """
    nrows = min(max_scan_rows, raw.shape[0])

    year_candidates: List[int] = []
    for r in range(nrows):
        row = raw.iloc[r, :]
        num_year_like = sum(_is_year_like(v) for v in row)
        # require at least a few year-like entries
        if num_year_like >= 4:
            year_candidates.append(r)

    if not year_candidates:
        raise ValueError("Could not detect a header row with years in the first 10 rows.")

    # pick the earliest plausible year row
    year_row = min(year_candidates)

    # find the first row after year_row that looks like months
    month_row = None
    for r in range(year_row + 1, nrows):
        row = raw.iloc[r, :]
        num_month_like = sum(_is_month_abbrev(v) for v in row)
        if num_month_like >= 6:  # at least half a year's worth
            month_row = r
            break

    if month_row is None:
        raise ValueError("Could not detect a header row with months following the year row.")

    return year_row, month_row


# ----------------------------
# Province/Territory normalization
# ----------------------------

def _clean_province_territory_value(x: object) -> object:
    """
    Trim trailing ' Total' (case-insensitive) from province/territory names,
    but preserve the standalone 'Total' row exactly.
    Also preserve the specific anomaly label unchanged:
        'Province/territory not stated Total'
    """
    if pd.isna(x):
        return x
    s = str(x).strip()
    # Preserve the anomaly exactly as-is
    if s.strip().lower() == "province/territory not stated total":
        return "Province/territory not stated Total"
    if s.lower() == "total":
        return "Total"
    s2 = re.sub(r"\s+total$", "", s, flags=re.IGNORECASE)
    return s2.strip()


def _trim_at_bottom_total(df: pd.DataFrame, label_col: str) -> pd.DataFrame:
    s = df[label_col].astype(str).str.strip().str.lower()
    idx = s[s.eq("total")].index
    if len(idx) > 0:
        last_idx = int(idx.max())
        return df.loc[: last_idx].reset_index(drop=True)
    return df


# ----------------------------
# Hierarchy totals and unpivot
# ----------------------------

def transform_hierarchical(df: pd.DataFrame, hierarchy_cols: List[str]) -> pd.DataFrame:
    """
    Generalized infill + total detection for hierarchy depth of 2 (province_territory, study_level).
    Mirrors extract_imp_tfw.py behavior.
    """
    df = df.copy()

    if "province_territory" not in df.columns:
        raise ValueError("Expected 'province_territory' as the first hierarchy column.")

    # Backfill province_territory (merged-cell style)
    df["province_territory"] = df["province_territory"].bfill()

    depth = len(hierarchy_cols) - 1  # deepest index
    # Intermediate levels bfill (exclude deepest). For depth==1, no-op.
    for level in range(1, depth):
        col = hierarchy_cols[level]
        parent_key = hierarchy_cols[:level]
        df[col] = (
            df.groupby(parent_key, dropna=False, group_keys=False)[col]
              .apply(lambda s: s.bfill())
        )

    # Candidate totals at levels 0..depth-1 (exclude leaf)
    def candidate_level(row):
        for k in range(depth - 1, -1, -1):
            if pd.notna(row[hierarchy_cols[k]]) and all(pd.isna(row[hierarchy_cols[m]]) for m in range(k + 1, depth + 1)):
                return k
        return pd.NA

    df["_candidate_level"] = [candidate_level(r) for _, r in df[hierarchy_cols].iterrows()]

    totals = []
    for i, row in df.iterrows():
        k = row["_candidate_level"]
        if pd.isna(k):
            totals.append(False)
            continue
        k = int(k)

        key_vals = tuple(row[hierarchy_cols[: k + 1]].tolist())

        def same_key(r):
            return tuple(r[hierarchy_cols[: k + 1]].tolist()) == key_vals

        def has_detail(r):
            return (k + 1) <= depth and pd.notna(r[hierarchy_cols[k + 1]])

        j = i - 1
        saw_detail = False
        while j >= 0:
            rj = df.loc[j]
            if not same_key(rj):
                break
            if has_detail(rj):
                saw_detail = True
            j -= 1
        totals.append(bool(saw_detail))

    df["total_flag"] = totals
    df = df.drop(columns=["_candidate_level"])

    # Ensure the very last 'Total' row is flagged TRUE
    pt_norm = df["province_territory"].astype(str).str.strip().str.lower()
    total_rows = pt_norm.eq("total")
    if total_rows.any():
        last_total_idx = total_rows[total_rows].index.max()
        df.loc[last_total_idx, "total_flag"] = True

    # Special-case anomaly: keep it, ensure total_flag is False,
    # and set study_level to a normalized label
    mask_anomaly = df["province_territory"].astype(str).str.strip().str.lower().eq(
        "province/territory not stated total"
    )
    if mask_anomaly.any():
        df.loc[mask_anomaly, "total_flag"] = False
        if "study_level" in df.columns:
            df.loc[mask_anomaly, "study_level"] = "Education level not stated"

    return df


def unpivot_monthly(df: pd.DataFrame, hierarchy_cols: List[str], ym_cols: List[str]) -> pd.DataFrame:
    id_vars = hierarchy_cols + ["total_flag"]
    unpivoted = pd.melt(
        df,
        id_vars=id_vars,
        value_vars=ym_cols,
        var_name="year_month",
        value_name="value",
    )
    unpivoted[["year", "month"]] = unpivoted["year_month"].str.split("-", expand=True)
    unpivoted["year"] = unpivoted["year"].astype(int)
    unpivoted["month"] = unpivoted["month"].astype(int)

    # Reorder columns: place year and month after year_month
    cols = list(unpivoted.columns)
    cols.remove("year")
    cols.remove("month")
    ym_idx = cols.index("year_month")
    cols.insert(ym_idx + 1, "year")
    cols.insert(ym_idx + 2, "month")
    unpivoted = unpivoted[cols]
    return unpivoted


# ----------------------------
# Main processing
# ----------------------------

def build_monthly_dataframe(input_path: str, sheet: object = 0) -> Tuple[pd.DataFrame, List[str], List[str]]:
    """
    Reads the Excel, constructs monthly columns, and returns:
      (dataframe, hierarchy_cols, year_month_columns)
    """
    raw = pd.read_excel(input_path, sheet_name=sheet, header=None)

    year_row, month_row = detect_year_and_month_rows(raw)

    # First two columns are hierarchy labels
    FIRST_LABEL_COLS = [0, 1]
    hierarchy_cols = ["province_territory", "study_level"]

    # Year values and months
    year_vals = raw.iloc[year_row, :].copy().ffill()
    month_vals = raw.iloc[month_row, :].copy()

    keep_cols: List[int] = []
    headers: List[str] = []

    # Keep hierarchy columns
    keep_cols.extend(FIRST_LABEL_COLS)
    headers.extend(hierarchy_cols)

    # Keep only monthly columns (exclude quarters and totals)
    for col_idx in range(2, raw.shape[1]):
        month_abbrev = str(month_vals.iloc[col_idx]) if not pd.isna(month_vals.iloc[col_idx]) else None
        month_num = _MONTH_MAP.get(str(month_abbrev).strip()) if month_abbrev else None
        if month_num is None:
            continue

        yv = year_vals.iloc[col_idx]
        if pd.isna(yv):
            continue
        try:
            y = int(str(yv).strip())
        except Exception:
            m = re.match(r"^(\d{4})", str(yv))
            if not m:
                continue
            y = int(m.group(1))

        keep_cols.append(col_idx)
        headers.append(f"{y:04d}-{month_num}")

    # Data begins after the month row
    FIRST_DATA_ROW = month_row + 1
    data = raw.iloc[FIRST_DATA_ROW:, keep_cols].copy()
    data.columns = headers

    # Clean province/territory values and trim at bottom Total
    data["province_territory"] = data["province_territory"].map(_clean_province_territory_value)
    data = _trim_at_bottom_total(data, label_col="province_territory")

    # Clean numeric monthly columns
    for c in headers[2:]:
        data[c] = (
            data[c]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("\u202f", "", regex=False)
            .str.replace("\xa0", "", regex=False)
            .str.strip()
            .replace({"--": "2", "": "0", "nan": "0", "None": "0"})
        )
        data[c] = pd.to_numeric(data[c], errors="coerce").fillna(0)

    # Preserve rows with blank province_territory; they will be backfilled later
    data = data.reset_index(drop=True)

    return data, hierarchy_cols, headers[2:]


def process_excel(input_path: str, output_path: str) -> None:
    print(f"Reading {input_path} ...")
    df, hierarchy_cols, ym_cols = build_monthly_dataframe(input_path, sheet=0)
    print(f"Detected hierarchy columns: {hierarchy_cols}")
    print(f"Monthly columns: {ym_cols[0]}..{ym_cols[-1]} (total {len(ym_cols)})")

    transformed = transform_hierarchical(df, hierarchy_cols)

    export_cols = hierarchy_cols + ["total_flag"] + ym_cols
    transformed = transformed[export_cols]

    print("Creating unpivoted version...")
    unp = unpivot_monthly(transformed, hierarchy_cols, ym_cols)

    # Save as CSV
    csv_output_path = output_path.replace('.xlsx', '.csv').replace('.xlsm', '.csv')
    if not csv_output_path.endswith('.csv'):
        csv_output_path += '.csv'
    # Ensure output directory exists
    import os
    out_dir = os.path.dirname(csv_output_path) or "."
    os.makedirs(out_dir, exist_ok=True)
    unp.to_csv(csv_output_path, index=False)
    print(f"Saved CSV: {csv_output_path}")

    print(f"Original data: {len(transformed)} rows")
    print(f"Unpivoted data: {len(unp)} rows")


def main():
    parser = argparse.ArgumentParser(description="Clean Study level Excel into unpivoted CSV format.")
    parser.add_argument("input_path", nargs="?", default="goc_data_source/EN_ODP-TR-Study-IS_PT_study_level_sign.xlsx", help="Path to the input .xlsx file")
    parser.add_argument("output_path", nargs="?", default="goc_data_processed/extracted_study.csv", help="Output CSV path")
    args = parser.parse_args()

    process_excel(args.input_path, args.output_path)


if __name__ == "__main__":
    main()


