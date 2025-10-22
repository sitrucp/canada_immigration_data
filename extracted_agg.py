#!/usr/bin/env python3
"""
Python script that replicates the Power Query logic for combining and summarizing CSV data.

This script:
1. Loads 6 CSV files from the goc_data_processed directory
2. Selects only the desired columns: stream, year, value, total_flag
3. Adds a 'stream' column to identify the source dataset
4. Combines all tables
5. Filters rows where total_flag = "False"
6. Groups by stream and year, summing the values
"""

import pandas as pd
import os

def main():
    # Define the data directory
    data_dir = "goc_data_processed"
    
    # Define the list of columns you want to keep in the final appended table
    desired_columns = ["stream", "year", "value", "total_flag"]
    
    # Define the source files and their corresponding stream names
    source_files = {
        "extracted_asylum.csv": "asylum",
        "extracted_hc.csv": "hc", 
        "extracted_tfw.csv": "tfw",
        "extracted_imp.csv": "imp",
        "extracted_pr.csv": "pr",
        "extracted_study.csv": "study"
    }
    
    # List to store all subset tables
    subset_tables = []
    
    print("Loading and processing CSV files...")
    
    # Process each source file
    for filename, stream_name in source_files.items():
        filepath = os.path.join(data_dir, filename)
        
        if not os.path.exists(filepath):
            print(f"Warning: File {filepath} not found, skipping...")
            continue
            
        print(f"Processing {filename}...")
        
        # Load the CSV file
        df = pd.read_csv(filepath)
        
        # Add the stream column
        df['stream'] = stream_name
        
        # Select only the desired columns (if they exist)
        available_columns = [col for col in desired_columns if col in df.columns]
        if len(available_columns) < len(desired_columns):
            missing_cols = [col for col in desired_columns if col not in df.columns]
            print(f"  Warning: Missing columns in {filename}: {missing_cols}")
        
        # Select the available desired columns
        df_subset = df[available_columns].copy()
        
        # Ensure all desired columns exist (fill missing ones with appropriate defaults)
        for col in desired_columns:
            if col not in df_subset.columns:
                if col == "total_flag":
                    df_subset[col] = False  # Default to False if not present
                elif col == "value":
                    df_subset[col] = 0  # Default to 0 if not present
                elif col == "year":
                    # Try to extract year from other columns if available
                    if "year_month" in df.columns:
                        df_subset[col] = df["year_month"].str[:4].astype(int)
                    else:
                        df_subset[col] = 2023  # Default year
                elif col == "stream":
                    df_subset[col] = stream_name
        
        # Handle missing total_flag values by filling with False
        if "total_flag" in df_subset.columns:
            df_subset["total_flag"] = df_subset["total_flag"].fillna(False)
        
        # Reorder columns to match desired_columns
        df_subset = df_subset[desired_columns]
        
        subset_tables.append(df_subset)
        print(f"  Loaded {len(df_subset)} rows from {filename}")
    
    if not subset_tables:
        print("No data files found. Please check the goc_data_processed directory.")
        return
    
    # Combine all subset tables
    print("\nCombining all tables...")
    combined_table = pd.concat(subset_tables, ignore_index=True)
    print(f"Combined table has {len(combined_table)} rows")
    
    # Filter rows where total_flag = False (boolean, not string)
    print("\nFiltering rows where total_flag = False...")
    filtered_rows = combined_table[~combined_table['total_flag']].copy()
    print(f"After filtering: {len(filtered_rows)} rows")
    
    # Group by stream and year, summing the values
    print("\nGrouping by stream and year, summing values...")
    grouped_rows = filtered_rows.groupby(['stream', 'year'])['value'].sum().reset_index()
    grouped_rows = grouped_rows.sort_values(['stream', 'year'])
    
    print(f"Final result: {len(grouped_rows)} rows")
    
    # Display the results
    print("\nSummary of results:")
    print("=" * 50)
    print(grouped_rows.to_string(index=False))
    
    # Save the results to a CSV file
    output_file = os.path.join(data_dir, "extracted_agg.csv")
    grouped_rows.to_csv(output_file, index=False)
    print(f"\nResults saved to: {output_file}")
    
    # Display summary statistics
    print("\nSummary statistics:")
    print("=" * 50)
    print(f"Total streams: {grouped_rows['stream'].nunique()}")
    print(f"Year range: {grouped_rows['year'].min()} - {grouped_rows['year'].max()}")
    print(f"Total value across all streams and years: {grouped_rows['value'].sum():,}")
    
    # Show breakdown by stream
    print("\nBreakdown by stream:")
    print("=" * 30)
    stream_totals = grouped_rows.groupby('stream')['value'].sum().sort_values(ascending=False)
    for stream, total in stream_totals.items():
        print(f"{stream:10}: {total:>10,}")

if __name__ == "__main__":
    main()
