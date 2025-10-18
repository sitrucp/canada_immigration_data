#!/usr/bin/env python3
"""
Extract & clean: Canada - Permanent Residents by Province and Immigration Category.

What it does
------------
- Reads the raw Excel with multi-row headers (Year / Quarter / Month).
- Detects hierarchical category columns (province_territory + category levels).
- Builds a single header row of YYYY-MM month columns.
- Drops quarterly and yearly subtotal columns.
- Removes the final "Total" row and any notes below.
- Cleans numeric cells, converting "--" -> 2 and blanks -> 0.
- Adds total_flag for hierarchical subtotals.
- Writes the cleaned data to unpivoted CSV format.

Usage
-----
python extract_pr.py "EN_ODP-PR-ProvImmCat.xlsx" "extracted_pr.csv"
# or
python extract_pr.py input.xlsx output.csv

Optional args
-------------
--sheet "Sheet1"     Worksheet name (will auto-detect if not provided)
"""

import argparse
import re
from typing import Optional, Tuple, List
import pandas as pd
import numpy as np


def detect_header_structure(raw: pd.DataFrame, max_scan_rows: int = 12) -> Tuple[int, int, int]:
    """
    Detect where the year row, month row, and data start.
    Returns (year_row, month_row, first_data_col) where first_data_col is number of hierarchy columns.
    
    For the PR dataset, we look for:
    - A row with year labels (2015, 2016, etc.)
    - The month row is 2 rows below the year row
    - Years typically start at column 4 (after 4 hierarchy columns)
    """
    # Look for year pattern (4-digit year like 2015, 2016, etc.)
    def is_year(x):
        try:
            xi = int(x)
            return 1900 <= xi <= 2100
        except (ValueError, TypeError):
            return False
    
    nrows, ncols = raw.shape
    
    # Scan for a row with years
    for i in range(min(max_scan_rows, nrows)):
        row = raw.iloc[i]
        
        # Find the first column that contains a year value
        first_year_col = None
        for col_idx in range(ncols):
            if is_year(row.iloc[col_idx]):
                first_year_col = col_idx
                break
        
        if first_year_col is None:
            continue
            
        # Count total years in this row to confirm it's the year row
        year_count = sum(1 for col_idx in range(ncols) if is_year(row.iloc[col_idx]))
        
        # If we found at least 5 years, this is our year row
        if year_count >= 5:
            year_row = i
            month_row = i + 2  # Month row is 2 rows below (quarter row in between)
            first_data_col = first_year_col  # This is where data columns start
            return year_row, month_row, first_data_col
    
    raise ValueError("Could not detect header structure with year and month rows.")


def normalize_month(x: object) -> Optional[str]:
    """Convert month abbreviation to 2-digit string."""
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


def _clean_province_territory_value(x: str) -> str:
    """
    Trim trailing ' Total' or ' -' (case-insensitive) from province/territory names,
    but preserve the standalone 'Total' row exactly.
    """
    if pd.isna(x):
        return x
    s = str(x).strip()
    if s.lower() == "total":
        return "Total"
    # remove any trailing whitespace + 'total'
    s2 = re.sub(r"\s+total$", "", s, flags=re.IGNORECASE)
    # remove any trailing ' -'
    s2 = re.sub(r"\s*-\s*$", "", s2)
    return s2.strip()


def parse_pr_data(input_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """
    Parse the PR Excel file with both hierarchical categories and year/month columns.
    Returns cleaned dataframe in wide format.
    """
    # Load raw sheet without interpreting headers
    if sheet_name is None:
        # Auto-detect first sheet
        xl = pd.ExcelFile(input_path)
        sheet_name = xl.sheet_names[0]
    
    raw = pd.read_excel(input_path, sheet_name=sheet_name, header=None)
    
    # Detect header structure
    year_row, month_row, first_data_col = detect_header_structure(raw)
    first_data_row = month_row + 1
    
    # Extract header rows
    year_vals = raw.iloc[year_row, :].copy()
    month_vals = raw.iloc[month_row, :].copy()
    # Forward-fill years across merged cells
    year_vals_ffill = year_vals.ffill()
    
    # Build hierarchy column names
    hier_n = first_data_col
    base = ["province_territory", "category_1", "category_2", "category_3"]
    if hier_n < 1 or hier_n > 4:
        raise ValueError(f"Unexpected number of hierarchy columns: {hier_n}")
    hierarchy_cols = base[:hier_n]
    
    # Decide columns to keep: hierarchy columns + only monthly columns
    keep_cols, new_headers = [], []
    
    # Add hierarchy columns
    for col_idx in range(first_data_col):
        keep_cols.append(col_idx)
    new_headers.extend(hierarchy_cols)
    
    # Add monthly columns
    for col_idx in range(first_data_col, raw.shape[1]):
        month_num = normalize_month(month_vals.iloc[col_idx])
        if month_num is None:
            continue  # skip non-month columns (quarters, totals)
        
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
    
    # Clean province_territory by removing trailing ' Total' (but keep standalone 'Total')
    data["province_territory"] = data["province_territory"].map(_clean_province_territory_value)
    
    # Find and remove the last "Total" row and everything below it
    pt_norm = data["province_territory"].astype(str).str.strip().str.lower()
    total_rows = pt_norm.eq("total")
    if total_rows.any():
        last_total_idx = total_rows[total_rows].index.max()
        data = data.loc[:last_total_idx]
    
    # Remove stray empties/headers
    data = data[~data["province_territory"].isin(["nan", "None", "", "NaN"])]
    data = data[~pt_norm.str.contains(r"^(province|territory)$", case=False, na=False)]
    
    # Clean numeric month columns:
    # - remove thousands separators/whitespace
    # - map "--" -> 2, blanks -> 0
    month_cols = [c for c in data.columns if c not in hierarchy_cols]
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
    
    # Reset index for a clean output
    data = data.reset_index(drop=True)
    
    return data, hierarchy_cols, month_cols


def transform_hierarchical(df: pd.DataFrame, hierarchy_cols: List[str]) -> pd.DataFrame:
    """
    Generalized infill + total detection for variable hierarchy depth.
    - Infill up to the second-last hierarchy level (deepest category remains as-is).
    - total_flag TRUE for structural totals; we also ensure the final 'Total' row is TRUE.
    """
    df = df.copy()
    
    # Infill
    if "province_territory" not in df.columns:
        raise ValueError("Expected 'province_territory' as the first hierarchy column.")
    df["province_territory"] = df["province_territory"].bfill()
    
    depth = len(hierarchy_cols) - 1  # deepest index
    # bfill intermediate levels only (exclude deepest)
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
        return np.nan
    df["_candidate_level"] = [candidate_level(r) for _, r in df[hierarchy_cols].iterrows()]
    
    # Confirm totals by scanning upward within group for any deeper-detail row
    totals = []
    for i, row in df.iterrows():
        k = row["_candidate_level"]
        if pd.isna(k):
            totals.append(False)
            continue
        k = int(k)
        
        key_vals = tuple(row[hierarchy_cols[:k + 1]].tolist())
        def same_key(r): return tuple(r[hierarchy_cols[:k + 1]].tolist()) == key_vals
        def has_detail(r): return (k + 1) <= depth and pd.notna(r[hierarchy_cols[k + 1]])
        
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
    
    return df


def unpivot_data(df: pd.DataFrame, hierarchy_cols: List[str], month_cols: List[str]) -> pd.DataFrame:
    """
    Unpivot the monthly columns into year_month, year, month and value columns.
    This transforms the wide format (one row per category with multiple month columns)
    into long format (multiple rows per category with year_month, year, month and value columns).
    """
    # Keep the hierarchy columns and total_flag as id_vars
    id_vars = hierarchy_cols + ["total_flag"]
    
    # Unpivot the month columns
    unpivoted = pd.melt(
        df, 
        id_vars=id_vars,
        value_vars=month_cols,
        var_name='year_month',
        value_name='value'
    )
    
    # Split year_month into separate year and month columns
    unpivoted[['year', 'month']] = unpivoted['year_month'].str.split('-', expand=True)
    unpivoted['year'] = unpivoted['year'].astype(int)
    unpivoted['month'] = unpivoted['month'].astype(int)
    
    # Reorder columns to have year and month after year_month
    cols = list(unpivoted.columns)
    # Remove year and month from their current positions
    cols.remove('year')
    cols.remove('month')
    # Insert year and month after year_month
    year_month_idx = cols.index('year_month')
    cols.insert(year_month_idx + 1, 'year')
    cols.insert(year_month_idx + 2, 'month')
    unpivoted = unpivoted[cols]
    
    # Normalize "Not stated" province handling for downstream use
    if "province_territory" in unpivoted.columns and "category_1" in unpivoted.columns:
        _mask_ns = unpivoted["province_territory"].astype(str).str.contains("not stated", case=False, na=False)
        unpivoted.loc[_mask_ns, "category_1"] = "Not stated"
    
    return unpivoted


def main():
    parser = argparse.ArgumentParser(description="Clean PR Province/Immigration Category Excel into unpivoted CSV format.")
    parser.add_argument("input_path", help="Path to the input .xlsx file")
    parser.add_argument("output_path", help="Output CSV path (will be converted from .xlsx if needed)")
    parser.add_argument("--sheet", default=None, help='Worksheet name (default: auto-detect first sheet)')
    args = parser.parse_args()
    
    print(f"Reading {args.input_path} ...")
    df, hierarchy_cols, month_cols = parse_pr_data(args.input_path, sheet_name=args.sheet)
    
    print(f"Detected hierarchy columns: {hierarchy_cols}")
    print(f"Month columns: {month_cols[0]}..{month_cols[-1]} (total {len(month_cols)})")
    
    # Apply hierarchical transformations
    print("Applying hierarchical transformations...")
    result = transform_hierarchical(df, hierarchy_cols)
    
    # Reorder columns so total_flag appears right before the month columns
    export_cols = hierarchy_cols + ["total_flag"] + month_cols
    result = result[export_cols]
    
    # Create and save unpivoted version (CSV only)
    print("Creating unpivoted version...")
    unpivoted_df = unpivot_data(result, hierarchy_cols, month_cols)
    
    # Generate CSV output path (replace .xlsx with .csv)
    csv_output_path = args.output_path.replace('.xlsx', '.csv').replace('.xlsm', '.csv')
    if not csv_output_path.endswith('.csv'):
        csv_output_path += '.csv'
    
    print("Saving unpivoted CSV output...")
    unpivoted_df.to_csv(csv_output_path, index=False)
    print(f"Saved CSV: {csv_output_path}")
    
    # Print confirmation
    print(f"Original data: {len(df):,} rows, {len(df.columns):,} columns")
    print(f"Unpivoted data: {len(unpivoted_df):,} rows, {len(unpivoted_df.columns):,} columns")
    print("Done.")


if __name__ == "__main__":
    main()

