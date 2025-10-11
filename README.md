# Canada Immigration Data Analysis

This repository contains data extraction scripts for Canadian temporary foreign worker programs, focusing on Temporary Foreign Worker Program (TFWP), International Mobility Program (IMP), and Humanitarian & Compassionate (H&C) work permit data.

## Project Overview

This project extracts and processes Canadian immigration data to create clean outputs for analysis of temporary foreign workers across different programs and provinces. The scope here focuses on extraction and processing only.

## Data Sources

All data is sourced from the Government of Canada's open data portal:

**Main Dataset**: [Temporary Residents: Temporary Foreign Worker Program (TFWP) and International Mobility Program (IMP) Work Permit Holders – Monthly IRCC Updates](https://open.canada.ca/data/en/dataset/360024f2-17e9-4558-bfc1-3616485d65b9)

### 1. TFWP Data
- **Source**: [Canada - Temporary Foreign Worker Program work permit holders on December 31st by province/territory of intended destination and program](https://open.canada.ca/data/en/dataset/360024f2-17e9-4558-bfc1-3616485d65b9/resource/3acc1b2c-8da0-405e-b54d-a1c4bcc2bd5f)
- **File**: `EN_ODP_annual-TR-work-TFW_PT_program_year_end.xlsx`

### 2. IMP Data
- **Source**: [Canada - International Mobility Program work permit holders on December 31st by province/territory and program](https://open.canada.ca/data/en/dataset/360024f2-17e9-4558-bfc1-3616485d65b9/resource/8d4c4240-88ea-421d-b80d-6cc6a3d28044)
- **File**: `EN_ODP_annual-TR-work-IMP_PT_program_year_end.xlsx`

### 3. H&C Data
- **Source**: [Canada - Work permit holders for Humanitarian & Compassionate purposes by country of citizenship and year in which permit(s) became effective](https://open.canada.ca/data/en/dataset/360024f2-17e9-4558-bfc1-3616485d65b9/resource/7257ea58-a5f0-4e58-901a-9a8785878710)
- **File**: `EN_ODP-TR-Work-HC_citizenship_sign.xlsx`

## Data Notes

- Values between 0 and 5 are shown as "--" to prevent individual identification in the data we are replacing the "--" with a value 2.
- All other values are rounded to the closest multiple of 5 for privacy protection
- Data are preliminary estimates and subject to change
- Total unique counts may not equal the sum of permit holders as individuals may hold multiple permit types
- Reporting methodology revised as of June 20, 2014 TFWP overhaul

## File Structure

### Data Files
- `EN_ODP_annual-TR-work-TFW_PT_program_year_end.xlsx` - TFWP data by province/territory
- `EN_ODP_annual-TR-work-IMP_PT_program_year_end.xlsx` - IMP data by province/territory  
- `EN_ODP-TR-Work-HC_citizenship_sign.xlsx` - H&C data by country of citizenship
- `TR_IMP_TFWP.pbix` - Power BI dashboard file

### Processed Data
- `extracted_hc.csv` - Processed H&C data
- `extracted_imp.csv` - Processed IMP data
- `extracted_tfw.csv` - Processed TFWP data

### Processing Scripts
- `extract_hc.py` - Script to extract and process H&C data
- `extract_imp_tfw.py` - Script to extract and process IMP and TFWP data

## Data Processing Workflow

1. **Data Extraction**: Python scripts extract data from Excel files
2. **Data Cleaning**: Remove privacy-protected values ("--") and handle rounding
3. **Data Aggregation**: Combine data across programs and time periods
4. **Export**: Write processed datasets to CSV files for downstream analysis

## Data Transformations and Edge Cases

The extraction scripts normalize the Excel sources into clean, analysis-ready long-format CSVs. Key steps and edge cases:

- Header normalization (HC): forward-fill years across merged cells; select monthly columns only; drop quarterly/yearly subtotals and the final Total row.
- Header detection (IMP/TFW): auto-detect the header row and the first year column; support variable hierarchy depth before the year columns.
- Filling hierarchical labels (IMP/TFW): backfill parent labels downwards for `province_territory` and intermediate `category_*` columns so every row has full context.
- Placeholder handling: values of "--" are replaced with 2; blanks and non-numeric are coerced to 0; thousands separators and non-breaking spaces are removed.
- Unpivot: convert wide year (or year-month) columns into long format with one observation per row.
- Province "Not stated" alignment (IMP/TFW): when `province_territory` contains "Not stated", the primary category is normalized to `Not stated` for consistent grouping.
- Trimming footers (IMP/TFW): by default, rows below the final exact `Total` row are removed to drop footnotes and notes. You can disable this with the `no-trim` flag.

## Output CSV Schemas

### H&C output: `extracted_hc.csv`

- Columns:
  - `country_citizenship`: country name as text
  - `year_month`: string in YYYY-MM (e.g., 2021-07)
  - `year`: integer year (e.g., 2021)
  - `month`: integer month (1-12)
  - `value`: integer count; "--" from source becomes 2; blanks become 0

- Example row:
  - `country_citizenship="Mexico"`, `year_month="2021-07"`, `year=2021`, `month=7`, `value=125`

### IMP/TFW outputs: `extracted_imp.csv` and `extracted_tfw.csv`

Depending on the source sheet’s structure, there can be 1 to 4 hierarchy columns before the year columns. The output schema is consistent:

- Columns:
  - `province_territory` (always present)
  - Optional: `category_1`, `category_2`, `category_3` (present when the source has multiple hierarchy levels)
  - `total_flag`: boolean; TRUE for structural subtotal rows, FALSE for detail/leaf rows
  - `year`: integer year (e.g., 2018)
  - `value`: integer count; "--" from source becomes 2; blanks become 0

- Example rows:
  - Two-level example: `province_territory="Ontario"`, `category_1="Program A"`, `total_flag=False`, `year=2019`, `value=15340`
  - Three-level example: `province_territory="Quebec"`, `category_1="Program B"`, `category_2="Subcategory X"`, `total_flag=False`, `year=2020`, `value=8420`

### About `total_flag`

- `total_flag=True` marks structural subtotal rows created by the source hierarchy (e.g., province × program totals). These rows aggregate the detail rows beneath them in the hierarchy.
- The very last `Total` row (across all provinces) is also flagged TRUE when present.
- Use cases:
  - To avoid double counting, filter `total_flag=False` when summing detail values.
  - To analyze rollups, filter `total_flag=True` and group by the hierarchy columns.

## Usage

### Running Data Extraction Scripts

```bash
# HC with explicit input/output (default sheet)
python extract_hc.py "EN_ODP-TR-Work-HC_citizenship_sign.xlsx" "extracted_hc.csv"

# HC with custom sheet name
python extract_hc.py "EN_ODP-TR-Work-HC_citizenship_sign.xlsx" "extracted_hc.csv" --sheet "TR - HC CITZ"

# IMP/TFW: process both default files in the current folder
python extract_imp_tfw.py

# IMP/TFW: process a specific file into a specific CSV
python extract_imp_tfw.py "EN_ODP_annual-TR-work-IMP_PT_program_year_end.xlsx" "extracted_imp.csv"

# IMP/TFW: process a specific file and keep rows below the final Total row
python extract_imp_tfw.py "EN_ODP_annual-TR-work-TFW_PT_program_year_end.xlsx" "extracted_tfw.csv" no-trim
```

## Dependencies

- Python 3.x
- pandas
- openpyxl (for Excel file processing)
 

## Data Privacy

This project processes publicly available government data that has already been anonymized and rounded for privacy protection. No personal information is included in the datasets.

## License

This project uses open government data from Canada. Please refer to the original data sources for licensing information.

## Contributing

This is a personal data analysis project. For questions or suggestions, please refer to the original data sources or government documentation.
