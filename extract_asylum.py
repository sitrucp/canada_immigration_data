#!/usr/bin/env python3
"""
Extract & clean: Canada - Asylum claims by Claim Office Type and Province/Territory.

What it does
------------
- Reads the first worksheet from the raw Excel with two header rows (Year / Month).
- Detects hierarchical label columns in A and B:
  - column A -> claim_office_type (merged header region)
  - column B -> province_territory
- Builds a single header row of YYYY-MM month columns (keeps only monthly columns; excludes yearly totals).
- Backfills only claim_office_type (column A). province_territory is not backfilled.
- Trims anything below the last exact "Total" found in claim_office_type.
- Cleans numeric cells, converting "--" -> 2 and blanks -> 0.
- Adds total_flag for claim_office_type subtotal rows (claim_office_type present and province_territory blank),
  and ensures the very last claim_office_type == "Total" row is flagged.
- Writes the cleaned data to unpivoted CSV format with columns:
  province_territory, claim_office_type, total_flag, year_month, year, month, value.

Usage
-----
python extract_asylum.py "EN_ODP-Asylum-OfficeType_Prov.xlsx" "extracted_asylum.csv"
# or
python extract_asylum.py input.xlsx output.csv
"""

import argparse
import re
from typing import Optional, Tuple, List
import pandas as pd


def _clean_province_territory_value(x: str) -> str:
    """
    Trim trailing ' Total' (case-insensitive) from province/territory names,
    but preserve the standalone 'Total' row exactly.
    """
    if pd.isna(x):
        return x
    s = str(x).strip()
    if s.lower() == "total":
        return "Total"
    s2 = re.sub(r"\s+total$", "", s, flags=re.IGNORECASE)
    return s2.strip()


def _normalize_month(x: object) -> Optional[str]:
    month_map = {
        "Jan": "01", "Feb": "02", "Mar": "03",
        "Apr": "04", "May": "05", "Jun": "06",
        "Jul": "07", "Aug": "08", "Sep": "09",
        "Oct": "10", "Nov": "11", "Dec": "12",
    }
    if pd.isna(x):
        return None
    x_str = str(x).strip()
    return month_map.get(x_str)


def _detect_header_structure_asylum(raw: pd.DataFrame, max_scan_rows: int = 12) -> Tuple[int, int, int]:
    """
    Detect the year row, month row, and the starting column index of the first year block.
    For this dataset, the month row is exactly 1 row below the year row (no quarters row).
    Returns (year_row, month_row, first_data_col).
    """
    def is_year(x):
        try:
            xi = int(x)
            return 1900 <= xi <= 2100
        except (ValueError, TypeError):
            return False

    nrows, ncols = raw.shape
    for i in range(min(max_scan_rows, nrows)):
        row = raw.iloc[i]
        first_year_col = None
        for col_idx in range(ncols):
            if is_year(row.iloc[col_idx]):
                first_year_col = col_idx
                break
        if first_year_col is None:
            continue

        # Count how many year tokens exist in this row (to confirm it is the year row)
        year_count = sum(1 for col_idx in range(ncols) if is_year(row.iloc[col_idx]))
        if year_count >= 3:  # looser than PR; asylum might contain fewer contiguous years
            year_row = i
            month_row = i + 1  # month row is directly below year row
            first_data_col = first_year_col
            return year_row, month_row, first_data_col

    raise ValueError("Could not detect a header structure with year and month rows.")


def parse_asylum_data(input_path: str) -> Tuple[pd.DataFrame, List[str], List[str]]:
    """
    Parse the Asylum Excel file with two hierarchy columns and year/month columns.
    Returns cleaned dataframe in wide format and metadata.
    """
    # Auto-detect first sheet
    xl = pd.ExcelFile(input_path)
    sheet_name = xl.sheet_names[0]
    raw = pd.read_excel(input_path, sheet_name=sheet_name, header=None)

    year_row, month_row, first_data_col = _detect_header_structure_asylum(raw)
    first_data_row = month_row + 1

    # Extract header rows
    year_vals = raw.iloc[year_row, :].copy()
    month_vals = raw.iloc[month_row, :].copy()
    # Forward-fill years across merged cells
    year_vals_ffill = year_vals.ffill()

    # Build hierarchy column names: A: claim_office_type, B: province_territory
    if first_data_col < 2:
        # We expect two hierarchy columns; if we detect fewer, force 2 as per spec
        first_data_col = 2
    hierarchy_cols = ["claim_office_type", "province_territory"]

    # Decide columns to keep: 2 hierarchy columns + only monthly columns
    keep_cols: List[int] = []
    new_headers: List[str] = []

    # Add hierarchy columns (0 and 1)
    keep_cols.extend([0, 1])
    new_headers.extend(hierarchy_cols)

    # Add monthly columns (skip yearly totals)
    for col_idx in range(first_data_col, raw.shape[1]):
        month_num = _normalize_month(month_vals.iloc[col_idx])
        if month_num is None:
            continue  # skip non-month columns

        year_val = year_vals_ffill.iloc[col_idx]
        if pd.isna(year_val):
            continue
        try:
            y = int(str(year_val).strip())
        except Exception:
            m = re.match(r"^(\d{4})", str(year_val))
            if not m:
                continue
            y = int(m.group(1))

        keep_cols.append(col_idx)
        new_headers.append(f"{y:04d}-{month_num}")

    # Slice data rows and apply headers
    data = raw.iloc[first_data_row:, keep_cols].copy()
    data.columns = new_headers

    # Normalize province_territory values
    data["province_territory"] = data["province_territory"].map(_clean_province_territory_value)

    # Backfill only claim_office_type, but remove trailing " - Total" ONLY on the rows that were filled
    _orig_cot = data["claim_office_type"].copy()
    _was_blank = _orig_cot.isna() | (_orig_cot.astype(str).str.strip() == "")
    _bfilled = _orig_cot.bfill()
    _bfilled_clean = _bfilled.astype(str).str.replace(r"\s*-\s*Total$", "", regex=True)
    data["claim_office_type"] = _orig_cot.where(~_was_blank, _bfilled_clean)

    # Trim below the last exact 'Total' in claim_office_type
    cot_norm = data["claim_office_type"].astype(str).str.strip().str.lower()
    total_rows = cot_norm.eq("total")
    if total_rows.any():
        last_total_idx = total_rows[total_rows].index.max()
        data = data.loc[:last_total_idx]

    # Remove stray empties/headers
    data = data[~data["claim_office_type"].astype(str).isin(["nan", "None", "", "NaN"])]

    # Identify month columns
    month_cols = [c for c in data.columns if c not in hierarchy_cols]

    # Clean numeric month columns
    for c in month_cols:
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

    # Compute total_flag for claim_office_type subtotals
    # True when claim_office_type present and province_territory is NaN/blank
    prov_blank = data["province_territory"].isna() | (data["province_territory"].astype(str).str.strip() == "")
    cot_present = data["claim_office_type"].notna() & (data["claim_office_type"].astype(str).str.strip() != "")
    # Exclude the special single-row 'Other Offices' group from subtotal flagging
    cot_norm2 = data["claim_office_type"].astype(str).str.strip().str.lower()
    is_other_offices = cot_norm2.eq("other offices")
    data["total_flag"] = (cot_present & prov_blank) & (~is_other_offices)

    # Ensure the very last claim_office_type == 'Total' row is flagged TRUE
    if total_rows.any():
        last_total_idx = total_rows[total_rows].index.max()
        data.loc[last_total_idx, "total_flag"] = True

    # Reset index
    data = data.reset_index(drop=True)

    return data, hierarchy_cols, month_cols


def unpivot_asylum(df: pd.DataFrame, hierarchy_cols: List[str], month_cols: List[str]) -> pd.DataFrame:
    """
    Unpivot monthly columns into long format with year_month, year, month and value.
    Keeps claim_office_type, province_territory and total_flag as id_vars.
    """
    id_vars = hierarchy_cols + ["total_flag"]
    unpivoted = pd.melt(
        df,
        id_vars=id_vars,
        value_vars=month_cols,
        var_name="year_month",
        value_name="value",
    )

    # Split year_month into year and month
    unpivoted[["year", "month"]] = unpivoted["year_month"].str.split("-", expand=True)
    unpivoted["year"] = unpivoted["year"].astype(int)
    unpivoted["month"] = unpivoted["month"].astype(int)

    # Reorder columns: province_territory, claim_office_type, total_flag, year_month, year, month, value
    cols = [
        "province_territory",
        "claim_office_type",
        "total_flag",
        "year_month",
        "year",
        "month",
        "value",
    ]
    # Guarantee ordering if any columns are missing due to naming issues
    cols = [c for c in cols if c in unpivoted.columns] + [c for c in unpivoted.columns if c not in cols]
    unpivoted = unpivoted[cols]

    return unpivoted


def main():
    parser = argparse.ArgumentParser(description="Clean Asylum Claim Office Type/Province Excel into unpivoted CSV format.")
    parser.add_argument("input_path", help="Path to the input .xlsx file (first sheet will be used)")
    parser.add_argument("output_path", help="Output CSV path (will be converted from .xlsx if needed)")
    args = parser.parse_args()

    print(f"Reading {args.input_path} ...")
    df, hierarchy_cols, month_cols = parse_asylum_data(args.input_path)

    print(f"Detected hierarchy columns: {hierarchy_cols}")
    print(f"Month columns: {month_cols[0]}..{month_cols[-1]} (total {len(month_cols)})")

    print("Creating unpivoted version...")
    unpivoted_df = unpivot_asylum(df, hierarchy_cols, month_cols)
    # Ensure final export column order
    _first = [
        "province_territory",
        "claim_office_type",
        "total_flag",
        "year_month",
        "year",
        "month",
        "value",
    ]
    _present_first = [c for c in _first if c in unpivoted_df.columns]
    _rest = [c for c in unpivoted_df.columns if c not in _present_first]
    unpivoted_df = unpivoted_df[_present_first + _rest]

    # Generate CSV output path (replace .xlsx with .csv)
    csv_output_path = args.output_path.replace('.xlsx', '.csv').replace('.xlsm', '.csv')
    if not csv_output_path.endswith('.csv'):
        csv_output_path += '.csv'

    print("Saving unpivoted CSV output...")
    unpivoted_df.to_csv(csv_output_path, index=False)
    print(f"Saved CSV: {csv_output_path}")

    print(f"Original data: {len(df):,} rows, {len(df.columns):,} columns")
    print(f"Unpivoted data: {len(unpivoted_df):,} rows, {len(unpivoted_df.columns):,} columns")
    print("Done.")


if __name__ == "__main__":
    main()


