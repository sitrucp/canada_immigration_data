#!/usr/bin/env python3
import pandas as pd
import numpy as np
import json
from typing import List, Set

def _validate_aggregates(long_df: pd.DataFrame, l01: pd.DataFrame, l12: pd.DataFrame, l23: pd.DataFrame) -> None:
    """Validate that totals at each node level match CSV aggregates across all provinces and years.

    Raises AssertionError with details if mismatches are found.
    """
    # Level 1: category_1
    csv_agg_l1 = (long_df.dropna(subset=["category_1"])  # ensure only rows with category_1
                          .groupby(["category_1"], as_index=False)["value"].sum())
    links_agg_l1 = (l01.groupby(["category_1"], as_index=False)["value"].sum())

    merged_l1 = csv_agg_l1.merge(links_agg_l1, on=["category_1"], how="outer", suffixes=("_csv","_links")).fillna(0)
    l1_mismatch = merged_l1.loc[~np.isclose(merged_l1["value_csv"], merged_l1["value_links"])].copy()
    if not l1_mismatch.empty:
        raise AssertionError(f"Level 1 totals mismatch for category_1: {l1_mismatch.to_dict(orient='records')}")

    # Level 2: category_2
    csv_agg_l2 = (long_df.dropna(subset=["category_2"])  # rows that have category_2
                          .groupby(["category_2"], as_index=False)["value"].sum())
    links_agg_l2 = (l12.groupby(["category_2"], as_index=False)["value"].sum())

    merged_l2 = csv_agg_l2.merge(links_agg_l2, on=["category_2"], how="outer", suffixes=("_csv","_links")).fillna(0)
    l2_mismatch = merged_l2.loc[~np.isclose(merged_l2["value_csv"], merged_l2["value_links"])].copy()
    if not l2_mismatch.empty:
        raise AssertionError(f"Level 2 totals mismatch for category_2: {l2_mismatch.to_dict(orient='records')}")

    # Level 3: category_3 (only rows with category_3)
    csv_agg_l3 = (long_df.dropna(subset=["category_3"])  # rows that have category_3
                          .groupby(["category_3"], as_index=False)["value"].sum())
    links_agg_l3 = (l23.groupby(["category_3"], as_index=False)["value"].sum())

    merged_l3 = csv_agg_l3.merge(links_agg_l3, on=["category_3"], how="outer", suffixes=("_csv","_links")).fillna(0)
    l3_mismatch = merged_l3.loc[~np.isclose(merged_l3["value_csv"], merged_l3["value_links"])].copy()
    if not l3_mismatch.empty:
        raise AssertionError(f"Level 3 totals mismatch for category_3: {l3_mismatch.to_dict(orient='records')}")

    # Overall total
    total_csv = float(long_df["value"].sum())
    total_links_l1 = float(l01["value"].sum())
    if not np.isclose(total_csv, total_links_l1):
        raise AssertionError(f"Overall total mismatch: csv={total_csv} vs links_l1_sum={total_links_l1}")


def build_echarts_data(df: pd.DataFrame, drop_totals: bool = True):
    # Normalize text columns
    for c in ["province_territory","category_1","category_2","category_3"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
            df.loc[df[c].isin(["", "nan", "NaN"]), c] = np.nan

    # Optionally remove totals
    if drop_totals and "total_flag" in df.columns:
        df = df.loc[~df["total_flag"].astype(bool)].copy()

    # Check if data is already unpivoted (has 'year' and 'value' columns)
    if "year" in df.columns and "value" in df.columns:
        # Data is already unpivoted
        long_df = df.copy()
        long_df["year"] = long_df["year"].astype(str)
        long_df["value"] = pd.to_numeric(
            long_df["value"].astype(str).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)
        long_df = long_df.loc[long_df["value"] != 0].copy()
    else:
        # Legacy format - detect year columns and unpivot
        year_cols: List[str] = [c for c in df.columns if c.isdigit() and len(c) == 4]
        if not year_cols:
            raise ValueError("No year columns found (expected 'YYYY' headers or 'year' column).")
        year_cols = sorted(year_cols)

        # Unpivot
        id_vars = ["province_territory","category_1","category_2","category_3"]
        long_df = df.melt(id_vars=id_vars, value_vars=year_cols, var_name="year", value_name="value")
        long_df["year"] = long_df["year"].astype(str)
        long_df["value"] = pd.to_numeric(
            long_df["value"].astype(str).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)
        long_df = long_df.loc[long_df["value"] != 0].copy()
    
    # Hardcoded fix: Set category_1 = "Not stated" for province/territory not stated records
    not_stated_mask = long_df["province_territory"].str.contains("not stated", case=False, na=False)
    long_df.loc[not_stated_mask, "category_1"] = "Not stated"

    # Build node labels by levels
    level0 = ["PR"]  # Top-level node
    level1 = sorted(set(long_df["category_1"].dropna().tolist()))
    level2 = sorted(set(long_df["category_2"].dropna().tolist()))
    level3 = sorted(set(long_df["category_3"].dropna().tolist()))
    
    labels = level0 + level1 + level2 + level3
    levels = ([0]*len(level0)) + ([1]*len(level1)) + ([2]*len(level2)) + ([3]*len(level3))
    node_id_map = {label: idx for idx, label in enumerate(labels)}

    # ECharts nodes format: list of node names
    nodes = labels

    # Links
    # Level 0 to Level 1: PR to all category_1 nodes (including "not stated" province)
    l01 = (long_df.dropna(subset=["category_1"])
                 .groupby(["province_territory","year","category_1"], as_index=False)["value"].sum())
    l01["source"] = l01["category_1"].map(lambda x: node_id_map["PR"])  # All from PR node
    l01["target"] = l01["category_1"].map(node_id_map)
    
    # Level 1 to Level 2: category_1 to category_2
    l12 = (long_df.dropna(subset=["category_1","category_2"])
                 .groupby(["province_territory","year","category_1","category_2"], as_index=False)["value"].sum())
    l12["source"] = l12["category_1"].map(node_id_map)
    l12["target"] = l12["category_2"].map(node_id_map)
    
    # Level 2 to Level 3: category_2 to category_3
    l23 = (long_df.dropna(subset=["category_2","category_3"])
                 .groupby(["province_territory","year","category_2","category_3"], as_index=False)["value"].sum())
    l23["source"] = l23["category_2"].map(node_id_map)
    l23["target"] = l23["category_3"].map(node_id_map)

    links = pd.concat([
        l01[["source","target","value","province_territory","year"]],
        l12[["source","target","value","province_territory","year"]],
        l23[["source","target","value","province_territory","year"]],
    ], ignore_index=True).sort_values(["year","province_territory","source","target"]).reset_index(drop=True)

    # Validate aggregates over all years and provinces for each level
    _validate_aggregates(long_df, l01.assign(category_1=l01["category_1"]), l12.assign(category_2=l12["category_2"]), l23.assign(category_3=l23["category_3"]))

    # ECharts links format: list of {source: node_name, target: node_name, value: number}
    echarts_links = []
    for _, row in links.iterrows():
        echarts_links.append({
            "source": labels[row["source"]],
            "target": labels[row["target"]],
            "value": int(row["value"]),
            "province_territory": row["province_territory"],
            "year": row["year"]
        })

    return nodes, echarts_links

def create_color_schema(df: pd.DataFrame) -> dict:
    """Create a color schema mapping for top-level nodes (direct children of PR).
    
    Colors are assigned based on the total value of each category, ensuring
    consistent visual hierarchy across different datasets.
    """
    # Define the color palette (same as in HTML)
    color_palette = [
        '#e74c3c',  # Red
        '#3498db',  # Blue
        '#2ecc71',  # Green
        '#f39c12',  # Orange
        '#9b59b6',  # Purple
        '#1abc9c',  # Turquoise
        '#e67e22',  # Dark Orange
        '#34495e',  # Dark Blue
        '#e91e63',  # Pink
        '#00bcd4'   # Cyan
    ]
    
    # Filter out totals before calculating color schema
    if 'total_flag' in df.columns:
        df = df.loc[~df['total_flag'].astype(bool)]
    
    # Calculate total values for each top-level category
    category_values = df.groupby('category_1')['value'].sum().sort_values(ascending=False)
    
    # Create color mapping based on value ranking (largest gets first color)
    color_schema = {}
    for i, (category, value) in enumerate(category_values.items()):
        color_schema[category] = color_palette[i % len(color_palette)]
    
    return color_schema

def main():
    input_csv = "goc_data_processed/extracted_pr.csv"
    template_html = "sankey_pr_template.html"
    output_html = "sankey_pr.html"

    df = pd.read_csv(input_csv)
    # Always exclude totals
    nodes, links = build_echarts_data(df, drop_totals=True)
    
    # Create color schema for top-level nodes
    color_schema = create_color_schema(df)

    # Load template and inject JSON
    with open(template_html, "r", encoding="utf-8") as f:
        tpl = f.read()

    nodes_json = json.dumps(nodes, ensure_ascii=False)
    links_json = json.dumps(links, ensure_ascii=False)
    color_schema_json = json.dumps(color_schema, ensure_ascii=False)

    html = (tpl.replace("/*__NODES_JSON__*/ []", nodes_json)
                .replace("/*__LINKS_JSON__*/ []", links_json)
                .replace("/*__COLOR_SCHEMA_JSON__*/ {}", color_schema_json))

    # Write out
    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Wrote: {output_html}")

if __name__ == "__main__":
    main()