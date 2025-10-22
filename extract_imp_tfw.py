import pandas as pd
import numpy as np
import re

def detect_header_and_year_start(df_raw, max_scan_rows=12, max_first_cols=8):
    def is_year(x):
        try:
            xi = int(x)
            return 1900 <= xi <= 2100
        except (ValueError, TypeError):
            return False
    nrows, ncols = df_raw.shape
    for i in range(min(max_scan_rows, nrows)):
        row = df_raw.iloc[i]
        for start in range(2, min(max_first_cols, ncols - 5)):
            tail = row.iloc[start:]
            years = []
            for v in tail:
                if is_year(v):
                    years.append(int(v))
                else:
                    break
            if len(years) >= 5:
                return i, start, years
    raise ValueError("Could not detect a header row with contiguous year labels.")

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
    # remove any trailing whitespace + 'total'
    s2 = re.sub(r"\s+total$", "", s, flags=re.IGNORECASE)
    return s2.strip()

def trim_at_bottom_total(df, label_col="province_territory"):
    """
    Trim anything below the LAST exact 'Total' row in the first hierarchy column.
    Keeps up to and including that row.
    """
    s = df[label_col].astype(str).str.strip().str.lower()
    idx = s[s.eq("total")].index
    if len(idx) > 0:
        last_idx = int(idx.max())
        return df.loc[:last_idx].reset_index(drop=True)
    return df

def parse_input_generic(path, trim_bottom=True):
    """
    Parse the worksheet with variable hierarchy depth (2 or 3 category cols).
    Returns (data, hierarchy_cols, years).
    """
    raw = pd.read_excel(path, header=None)
    hdr_row, year_start, years = detect_header_and_year_start(raw)

    hier_n = year_start
    base = ["province_territory", "category_1", "category_2", "category_3"]
    if hier_n < 1 or hier_n > 4:
        raise ValueError(f"Unexpected number of hierarchy columns: {hier_n}")
    hierarchy_cols = base[:hier_n]
    all_cols = hierarchy_cols + years

    data = raw.iloc[hdr_row + 1:, :len(all_cols)].reset_index(drop=True)
    data.columns = all_cols

    # Clean province_territory by removing trailing ' Total' (but keep standalone 'Total')
    data["province_territory"] = data["province_territory"].map(_clean_province_territory_value)

    # Optional: trim footers/notes below the final 'Total' row
    if trim_bottom:
        data = trim_at_bottom_total(data, label_col="province_territory")

    # Clean year columns: map "--" -> 2, strip thousands separators/whitespace, coerce to numeric, fill blanks -> 0
    for y in years:
        data[y] = (
            data[y]
            .replace({"--": 2})
        )
        data[y] = (
            data[y]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("\u202f", "", regex=False)
            .str.replace("\xa0", "", regex=False)
            .str.strip()
        )
        data[y] = pd.to_numeric(data[y], errors="coerce").fillna(0)

    return data, hierarchy_cols, years

def transform_hierarchical(df, hierarchy_cols):
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

    # Handle collisions between category_1 and category_2 by adding " (subcategory)" to category_2
    if "category_1" in df.columns and "category_2" in df.columns:
        # Find rows where category_1 and category_2 have the same value and total_flag=False
        collision_mask = (df["category_1"] == df["category_2"]) & df["category_1"].notna() & df["category_2"].notna() & (~df["total_flag"])
        if collision_mask.any():
            df.loc[collision_mask, "category_2"] = df.loc[collision_mask, "category_2"] + " (subcategory)"

    return df

def unpivot_data(df, hierarchy_cols, years):
    """
    Unpivot the year columns into year and value columns.
    This transforms the wide format (one row per category with multiple year columns)
    into long format (multiple rows per category with year and value columns).
    """
    # Keep the hierarchy columns and total_flag as id_vars
    id_vars = hierarchy_cols + ["total_flag"]
    
    # Unpivot the year columns
    unpivoted = pd.melt(
        df, 
        id_vars=id_vars,
        value_vars=years,
        var_name='year',
        value_name='value'
    )
    
    # Convert year column to integer
    unpivoted['year'] = unpivoted['year'].astype(int)
    
    # Remove rows where value is NaN (optional - you might want to keep them)
    # unpivoted = unpivoted.dropna(subset=['value'])
    
    return unpivoted

def process_excel(input_path, output_path, trim_bottom=True):
    print(f"Reading {input_path} ...")
    df, hierarchy_cols, years = parse_input_generic(input_path, trim_bottom=trim_bottom)
    print(f"Detected hierarchy columns: {hierarchy_cols}")
    print(f"Year columns: {years[0]}..{years[-1]} (total {len(years)})")

    result = transform_hierarchical(df, hierarchy_cols)

    # Reorder columns so total_flag appears right before the year columns
    export_cols = hierarchy_cols + ["total_flag"] + years
    result = result[export_cols]

    # Create unpivoted version for CSV
    print("Creating unpivoted version...")
    unpivoted_result = unpivot_data(result, hierarchy_cols, years)
    # Normalize "Not stated" province handling for downstream use
    if "province_territory" in unpivoted_result.columns and "category_1" in unpivoted_result.columns:
        _mask_ns = unpivoted_result["province_territory"].astype(str).str.contains("not stated", case=False, na=False)
        unpivoted_result.loc[_mask_ns, "category_1"] = "Not stated"
    
    # Generate CSV output path (replace .xlsx with .csv)
    csv_output_path = output_path.replace('.xlsx', '.csv')
    # Ensure output directory exists
    import os
    out_dir = os.path.dirname(csv_output_path) or "."
    os.makedirs(out_dir, exist_ok=True)
    print("Saving unpivoted CSV output...")
    unpivoted_result.to_csv(csv_output_path, index=False)
    print(f"Saved CSV: {csv_output_path}")
    
    print(f"Original data: {len(result)} rows")
    print(f"Unpivoted data: {len(unpivoted_result)} rows")

if __name__ == "__main__":
    import sys
    import os
    
    # Default processing for both IMP and TFW files
    if len(sys.argv) == 1:
        # Process both IMP and TFW files by default
        imp_input = "goc_data_source/EN_ODP_annual-TR-work-IMP_PT_program_year_end.xlsx"
        tfw_input = "goc_data_source/EN_ODP_annual-TR-work-TFW_PT_program_year_end.xlsx"
        
        if os.path.exists(imp_input):
            print("Processing IMP data...")
            process_excel(imp_input, "goc_data_processed/extracted_imp.csv", trim_bottom=True)
        else:
            print(f"Warning: {imp_input} not found")
            
        if os.path.exists(tfw_input):
            print("\nProcessing TFW data...")
            process_excel(tfw_input, "goc_data_processed/extracted_tfw.csv", trim_bottom=True)
        else:
            print(f"Warning: {tfw_input} not found")
            
    elif len(sys.argv) in (3, 4):
        # Custom file processing
        inp, outp = sys.argv[1], sys.argv[2]
        trim = True
        if len(sys.argv) == 4 and sys.argv[3].lower() == "no-trim":
            trim = False
        process_excel(inp, outp, trim_bottom=trim)
    else:
        print("Usage:")
        print("  python extract_imp_tfw.py                                    # Process both IMP and TFW files")
        print("  python extract_imp_tfw.py <input.xlsx> <output.csv> [no-trim]  # Process custom file")
        print("Add 'no-trim' as 3rd arg to disable trimming at bottom 'Total' row.")
        sys.exit(1)
