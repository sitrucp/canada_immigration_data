#!/usr/bin/env python3
import pandas as pd
import numpy as np
import json


def _validate_aggregates(long_df: pd.DataFrame, l01: pd.DataFrame) -> None:
    """Validate that totals at each node level match CSV aggregates across all provinces and years.

    Raises AssertionError with details if mismatches are found.
    """
    # Level 1: claim_office_type
    csv_agg_l1 = (long_df.dropna(subset=["claim_office_type"])  # ensure only rows with claim_office_type
                          .groupby(["claim_office_type"], as_index=False)["value"].sum())
    links_agg_l1 = (l01.groupby(["claim_office_type"], as_index=False)["value"].sum())

    merged_l1 = csv_agg_l1.merge(links_agg_l1, on=["claim_office_type"], how="outer", suffixes=("_csv","_links")).fillna(0)
    l1_mismatch = merged_l1.loc[~np.isclose(merged_l1["value_csv"], merged_l1["value_links"])].copy()
    if not l1_mismatch.empty:
        raise AssertionError(f"Level 1 totals mismatch for claim_office_type: {l1_mismatch.to_dict(orient='records')}")

    # Overall total (restrict to rows that form links: have claim_office_type)
    total_csv = float(long_df.dropna(subset=["claim_office_type"]) ["value"].sum())
    total_links_l1 = float(l01["value"].sum())
    if not np.isclose(total_csv, total_links_l1):
        raise AssertionError(f"Overall total mismatch: csv={total_csv} vs links_l1_sum={total_links_l1}")


def build_nodes_links(df: pd.DataFrame, drop_totals: bool = True):
    # Normalize text columns
    for c in ["province_territory","claim_office_type"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
            df.loc[df[c].isin(["", "nan", "NaN"]), c] = np.nan

    # Optionally remove totals
    if drop_totals and "total_flag" in df.columns:
        df = df.loc[~df["total_flag"].astype(bool)].copy()

    # Expect long format (has 'year' and 'value' columns) from extracted_asylum.csv
    if "year" in df.columns and "value" in df.columns:
        long_df = df.copy()
        long_df["year"] = long_df["year"].astype(str)
        long_df["value"] = pd.to_numeric(
            long_df["value"].astype(str).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)
        long_df = long_df.loc[long_df["value"] != 0].copy()
    else:
        # Legacy format (not expected here), try month/year columns
        raise ValueError("Expected unpivoted asylum CSV with 'year' and 'value' columns.")

    # Ensure provinces have explicit label for not stated
    long_df["province_territory"] = (
        long_df["province_territory"].fillna("Province/Territory of claim not stated")
    )

    # Nodes - Add top-level Asylum node
    level0 = ["Asylum"]  # Top-level node
    level1 = sorted(set(long_df["claim_office_type"].dropna().tolist()))

    labels = level0 + level1
    levels = ([0]*len(level0)) + ([1]*len(level1))
    node_id_map = {label: idx for idx, label in enumerate(labels)}

    nodes = [{"id": i, "label": lab, "level": lvl} for i, (lab, lvl) in enumerate(zip(labels, levels))]

    # Links
    # Level 0 to Level 1: Asylum to all claim_office_type nodes
    l01 = (long_df.dropna(subset=["claim_office_type"])  
                 .groupby(["province_territory","year","claim_office_type"], as_index=False)["value"].sum())
    l01["source"] = l01["claim_office_type"].map(lambda x: node_id_map["Asylum"])  # All from Asylum root
    l01["target"] = l01["claim_office_type"].map(node_id_map)

    links = l01[["source","target","value","province_territory","year"]].sort_values(["year","province_territory","source","target"]).reset_index(drop=True)

    # Validate aggregates over all years and provinces for each level
    _validate_aggregates(long_df, l01.assign(claim_office_type=l01["claim_office_type"]))

    return nodes, links.to_dict(orient="records")


def main():
    input_csv = "extracted_asylum.csv"
    template_html = "sankey_asylum_template.html"
    output_html = "sankey_asylum.html"

    df = pd.read_csv(input_csv)
    # Always exclude totals
    nodes, links = build_nodes_links(df, drop_totals=True)

    # Load template and inject JSON
    with open(template_html, "r", encoding="utf-8") as f:
        tpl = f.read()

    nodes_json = json.dumps(nodes, ensure_ascii=False)
    links_json = json.dumps(links, ensure_ascii=False)

    html = tpl.replace("/*__NODES_JSON__*/ []", nodes_json).replace("/*__LINKS_JSON__*/ []", links_json)

    # Write out
    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Wrote: {output_html}")


if __name__ == "__main__":
    main()


