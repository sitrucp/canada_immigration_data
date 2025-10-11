#!/usr/bin/env python3
"""
Extract & clean: Canada - Work permit holders for Humanitarian & Compassionate (HC) – by citizenship.

What it does
------------
- Reads the raw Excel with multi-row headers (Year / Quarter / Month).
- Builds a single header row of YYYY-MM month columns.
- Drops quarterly and yearly subtotal columns.
- Removes the final "Total" row and any notes below.
- Cleans numeric cells, converting "--" -> 2 and blanks -> 0.
- Writes the cleaned data to unpivoted CSV format.

Usage
-----
python extract_hc.py "EN_ODP-TR-Work-HC_citizenship_sign.xlsx" "extracted_hc.csv"
# or
python extract_hc.py input.xlsx output.csv

Optional args
-------------
--sheet "TR - HC CITZ"     Worksheet name (default shown)
"""

import argparse
import re
from typing import Optional
import pandas as pd


def clean_citizenship_xlsx(input_path: str, sheet_name: str = "TR - HC CITZ") -> pd.DataFrame:
    # Load raw sheet without interpreting headers
    raw = pd.read_excel(input_path, sheet_name=sheet_name, header=None)

    # Header rows (0-based)
    YEAR_ROW, MONTH_ROW = 2, 4
    FIRST_DATA_ROW = MONTH_ROW + 1
    FIRST_DATA_COL = 0

    # Extract header rows
    year_vals = raw.iloc[YEAR_ROW, :].copy()
    month_vals = raw.iloc[MONTH_ROW, :].copy()
    # Forward-fill years across merged cells
    year_vals_ffill = year_vals.ffill()

    month_map = {
        "Jan": "01", "Feb": "02", "Mar": "03",
        "Apr": "04", "May": "05", "Jun": "06",
        "Jul": "07", "Aug": "08", "Sep": "09",
        "Oct": "10", "Nov": "11", "Dec": "12",
    }

    def normalize_month(x: object) -> Optional[str]:
        if pd.isna(x):
            return None
        x_str = str(x).strip()
        return month_map.get(x_str)

    # Decide columns to keep: first column + only monthly columns
    keep_cols, new_headers = [], []
    for col_idx in range(raw.shape[1]):
        if col_idx == FIRST_DATA_COL:
            keep_cols.append(col_idx)
            new_headers.append("country_citizenship")
            continue

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
    data = raw.iloc[FIRST_DATA_ROW:, keep_cols].copy()
    data.columns = new_headers

    # Clean the first column and drop the 'Total' row plus any notes
    first_col = data.columns[0]
    data[first_col] = data[first_col].astype(str).str.strip()

    # Find and remove the last "Total" row (exact match, case-insensitive)
    matches = data.index[data[first_col].str.fullmatch(r"(?i)total")]
    if len(matches) > 0:
        data = data.loc[: matches.max() - 1]

    # Remove stray empties/headers
    data = data[~data[first_col].isin(["nan", "None", "", "NaN"])]
    data = data[~data[first_col].str.contains(r"^Country of Citizenship$", case=False, na=False)]

    # Clean numeric month columns:
    # - remove thousands separators/whitespace
    # - map "--" -> 2, blanks -> 0
    for c in data.columns[1:]:
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
    return data


def unpivot_data(df):
    """
    Unpivot the monthly columns into year_month and value columns.
    This transforms the wide format (one row per country with multiple month columns)
    into long format (multiple rows per country with year_month and value columns).
    Also splits year_month into separate year and month columns.
    """
    # Get the first column (country_citizenship) and all monthly columns
    id_vars = [df.columns[0]]  # country_citizenship
    value_vars = df.columns[1:]  # all monthly columns
    
    # Unpivot the monthly columns
    unpivoted = pd.melt(
        df, 
        id_vars=id_vars,
        value_vars=value_vars,
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
    
    return unpivoted


def main():
    parser = argparse.ArgumentParser(description="Clean HC citizenship Excel into unpivoted CSV format.")
    parser.add_argument("input_path", help="Path to the input .xlsx file")
    parser.add_argument("output_path", help="Output CSV path (will be converted from .xlsx if needed)")
    parser.add_argument("--sheet", default="TR - HC CITZ", help='Worksheet name (default: "TR - HC CITZ")')
    args = parser.parse_args()

    df = clean_citizenship_xlsx(args.input_path, sheet_name=args.sheet)

    # Create and save unpivoted version (CSV only)
    print("Creating unpivoted version...")
    unpivoted_df = unpivot_data(df)
    
    # Generate CSV output path (replace .xlsx with .csv)
    csv_output_path = args.output_path.replace('.xlsx', '.csv').replace('.xlsm', '.csv')
    if not csv_output_path.endswith('.csv'):
        csv_output_path += '.csv'
    
    unpivoted_df.to_csv(csv_output_path, index=False)
    print(f"Saved unpivoted CSV: {csv_output_path}")

    # Print confirmation
    print(f"Original data: {len(df):,} rows, {len(df.columns):,} columns")
    print(f"Unpivoted data: {len(unpivoted_df):,} rows, {len(unpivoted_df.columns):,} columns")
    print("Done.")

if __name__ == "__main__":
    main()
