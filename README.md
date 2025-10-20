# Canada Immigration Data Processing Scripts

This repository contains python code to extract analysis ready data from the Canadian Immigration, Refugees and Citizenship Canada (IRCC) Canada Open Data Excel files for the following programs: 

1) Temporary Foreign Worker Program (TFWP)
2) International Mobility Program (IMP)
3) Humanitarian & Compassionate (H&C) work permit data

The current scope here focuses on extraction and processing only. 

The PDF file `power_bi_report.pdf` contains tables and charts created using this data.

## Data Sources

All data is sourced via the Government of Canada's open data portal. The datasets used are listed below:

**Main List of Available Dataset**: [Temporary Residents: Temporary Foreign Worker Program (TFWP) and International Mobility Program (IMP) Work Permit Holders – Monthly IRCC Updates](https://open.canada.ca/data/en/dataset/360024f2-17e9-4558-bfc1-3616485d65b9)

### 1. TFWP Data
- **File**: `EN_ODP_annual-TR-work-TFW_PT_program_year_end.xlsx`
- **Source**: [Canada - Temporary Foreign Worker Program work permit holders on December 31st by province/territory of intended destination and program](https://open.canada.ca/data/en/dataset/360024f2-17e9-4558-bfc1-3616485d65b9/resource/3acc1b2c-8da0-405e-b54d-a1c4bcc2bd5f)

### 2. IMP Data
- **File**: `EN_ODP_annual-TR-work-IMP_PT_program_year_end.xlsx`
- **Source**: [Canada - International Mobility Program work permit holders on December 31st by province/territory and program](https://open.canada.ca/data/en/dataset/360024f2-17e9-4558-bfc1-3616485d65b9/resource/8d4c4240-88ea-421d-b80d-6cc6a3d28044)

### 3. H&C Data
- **File**: `EN_ODP-TR-Work-HC_citizenship_sign.xlsx`
- **Source**: [Canada - Work permit holders for Humanitarian & Compassionate purposes by country of citizenship and year in which permit(s) became effective](https://open.canada.ca/data/en/dataset/360024f2-17e9-4558-bfc1-3616485d65b9/resource/7257ea58-a5f0-4e58-901a-9a8785878710)

### 4. Permanent Residents (PR) Data
- **File**: `EN_ODP-PR-ProvImmCat.xlsx`
- **Source**: Canada - Permanent Residents by Province/Territory and Immigration Category
- **Link**: https://open.canada.ca/data/en/dataset/f7e5498e-0ad8-4417-85c9-9b8aff9b9eda

### 5. Asylum Claimants Data
- **File**: `EN_ODP-Asylum-OfficeType_Prov.xlsx`
- **Source**: Canada - Asylum Claimants by Claim Office Type, Province / Territory of Claim and Claim Month
- **Link**: https://open.canada.ca/data/en/dataset/b6cbcf4d-f763-4924-a2fb-8cc4a06e3de4

### 6. Study Permit Holders Data
- **File**: `EN_ODP-TR-Study-IS_PT_study_level_sign.xlsx`
- **Source**: Canada – Study permit holders by study level, province/territory and year in which permit(s) became effective
- **Link**: https://open.canada.ca/data/en/dataset/90115b00-f9b8-49e8-afa3-b4cff8facaee

## Source Data Notes

- Values between 0 and 5 are shown as "--" to prevent individual identification. Note: As part of the processing these "--" are replaced with a value 2 (just to provide an actual value eg between 0 and 5).
- All other values are rounded to the closest multiple of 5 for privacy protection
- Data are preliminary estimates and subject to change
- Total unique counts may not equal the sum of permit holders as individuals may hold multiple permit types
- Reporting methodology revised as of June 20, 2014 TFWP overhaul

## Repo File Structure

### Source Data Files (provided in the repo but you should get latest files from the source links above)
- `EN_ODP_annual-TR-work-TFW_PT_program_year_end.xlsx` - TFWP data by province/territory
- `EN_ODP_annual-TR-work-IMP_PT_program_year_end.xlsx` - IMP data by province/territory  
- `EN_ODP-TR-Work-HC_citizenship_sign.xlsx` - H&C data by country of citizenship
 - `EN_ODP-PR-ProvImmCat.xlsx` - PR data by province/territory and immigration category
 - `EN_ODP-Asylum-OfficeType_Prov.xlsx` - Asylum claims by office type and province/territory
 - `EN_ODP-TR-Study-IS_PT_study_level_sign.xlsx` - Study permit holders by study level and province/territory

### Processed Data - Use for Analytical Purposes (run the python code on latest files to get latest processed data)
- `extracted_hc.csv` - Processed H&C data
- `extracted_imp.csv` - Processed IMP data
- `extracted_tfw.csv` - Processed TFWP data
 - `extracted_pr.csv` - Processed PR data
 - `extracted_asylum.csv` - Processed Asylum data
 - `extracted_study.csv` - Processed Study data

### Python Processing Scripts
- `extract_hc.py` - Script to extract and process H&C data
- `extract_imp_tfw.py` - Script to extract and process IMP and TFWP data
 - `extract_pr.py` - Script to extract and process PR data
 - `extract_asylum.py` - Script to extract and process Asylum data
 - `extract_study.py` - Script to extract and process Study data

### Sankey Visualization Files
- `d3-sankey.js` - D3.js Sankey plugin library (v0.12.3) for rendering flow diagrams
- `sankey_imp_data.py` - Python script to generate IMP Sankey visualization from CSV data
- `sankey_imp_template.html` - HTML template for IMP Sankey chart (dynamically populated with data)
- `sankey_imp.html` - Generated IMP Sankey visualization (interactive D3.js chart)
- `sankey_tfw_data.py` - Python script to generate TFW Sankey visualization from CSV data
- `sankey_tfw_template.html` - HTML template for TFW Sankey chart (dynamically populated with data)
- `sankey_tfw.html` - Generated TFW Sankey visualization (interactive D3.js chart)
 - `sankey_pr_data.py` - Python script to generate PR Sankey visualization from CSV data
 - `sankey_pr_template.html` - HTML template for PR Sankey chart (dynamically populated with data)
 - `sankey_pr.html` - Generated PR Sankey visualization (interactive D3.js chart)
 - `sankey_asylum_data.py` - Python script to generate Asylum Sankey visualization from CSV data
 - `sankey_asylum_template.html` - HTML template for Asylum Sankey chart (dynamically populated with data)
 - `sankey_asylum.html` - Generated Asylum Sankey visualization (interactive D3.js chart)
 - `sankey_study_data.py` - Python script to generate Study Sankey visualization from CSV data
 - `sankey_study_template.html` - HTML template for Study Sankey chart (dynamically populated with data)
 - `sankey_study.html` - Generated Study Sankey visualization (interactive D3.js chart)

**Sankey Chart Features:**
- Interactive flow diagrams showing hierarchical data relationships
- Province and year filtering capabilities
- Color-coded families (all children inherit their Category 1 parent's color)
- Hierarchical grouping with value-based sorting
- Dynamic node sizing based on flow values
- Tooltips showing detailed flow information
- Legend showing zero-flow nodes for the selected filter

## Data Processing Workflow

1. **Data Extraction**: Python scripts extract data from Excel files
2. **Data Cleaning**: Remove privacy-protected values ("--") and handle rounding
3. **Data Aggregation**: Combine data across programs and time periods
4. **Output CSV files**: Write processed datasets to CSV files for downstream analysis
5. **Sankey Chart Generation**: Python scripts read the processed CSV data, build node and link structures, validate data aggregates, and inject the data as JSON into HTML templates to create interactive D3.js Sankey visualizations

## Data Processing Notes

The extraction scripts normalize the Excel data sources into clean, analysis-ready long-format CSV format data. Key data processing steps and edge cases include:

- Header normalization (HC): forward-fill years across merged cells; select monthly columns only; drop quarterly/yearly subtotals and the final Total row.
- Header detection (IMP/TFW): auto-detect the header row and the first year column; support variable hierarchy depth before the year columns.
- Filling hierarchical labels (IMP/TFW): backfill parent labels downwards for `province_territory` and intermediate `category_*` columns so every row has full context.
- Placeholder handling: values of "--" are replaced with 2; blanks and non-numeric are coerced to 0; thousands separators and non-breaking spaces are removed.
- Unpivot month/year columns: convert wide year (or year-month) columns into long format with one observation per row.
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
  - `country_citizenship="Mexico"`
  - `year_month="2021-07"`
  - `year=2021`
  - `month=7`
  - `value=125`

### IMP/TFW outputs: `extracted_imp.csv` and `extracted_tfw.csv`

The output schema is consistent:

- Columns:
  - `province_territory` (always present)
  - `categories` => IMP source data has 3 hierarcy columns: `category_1`, `category_2`, `category_3`. TFW has only 2 hierarcy columns: `category_1`, `category_2`. 
  - `total_flag`: boolean; TRUE for structural subtotal rows, FALSE for detail/leaf rows
  - `year`: integer year (e.g., 2018)
  - `value`: integer count; "--" from source becomes 2; blanks become 0

- Example row:
  - `province_territory="Quebec"`
  - `category_1="Canadian Interests"`
  - `category_2="Reciprocal Employment"`
  - `category_3="Exchange Professors, Visiting Lecturers"` (note: IMP only)
  - `total_flag=False`
  - `year=2023`
  - `value=10`


### PR output: `extracted_pr.csv`

- Columns:
  - `province_territory`
  - `category_1`, `category_2`, `category_3` (hierarchy depth may vary by year/category)
  - `total_flag`: boolean; TRUE for structural subtotal rows, FALSE for detail rows
  - `year_month`: string in YYYY-MM
  - `year`: integer year
  - `month`: integer month (1-12)
  - `value`: integer; "--" becomes 2; blanks become 0

- Notes:
  - Province "Not stated" rows normalize `category_1` to `Not stated` for consistency in rollups.

### Asylum output: `extracted_asylum.csv`

- Columns:
  - `province_territory`
  - `claim_office_type`
  - `total_flag`: boolean; TRUE for structural subtotal rows, FALSE for detail rows
  - `year_month`: string in YYYY-MM
  - `year`: integer year
  - `month`: integer month (1-12)
  - `value`: integer; "--" becomes 2; blanks become 0

### Study output: `extracted_study.csv`

- Columns:
  - `province_territory`
  - `study_level`
  - `total_flag`: boolean; TRUE for structural subtotal rows, FALSE for detail rows
  - `year_month`: string in YYYY-MM
  - `year`: integer year
  - `month`: integer month (1-12)
  - `value`: integer; "--" becomes 2; blanks become 0


### What is the `total_flag`?

The `total_flag=False` should be set to avoid double counting of detail and total data.

- `total_flag=True` marks source data subtotal rows in the original source data (e.g., province totals, program totals, etc). These rows aggregate their related hierarchical detail rows.
- The very last `Total` row (across all provinces) is also flagged TRUE when present.
- The `total_flag` can be used to exclude the total rows:
  - To avoid double counting, filter `total_flag=False` when summing detail values.
  - To analyze rollups, filter `total_flag=True` and group by the hierarchy columns.

## Python Code Usage

### Running Source Data Extraction Scripts

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

# PR: process Province/Immigration Category file into unpivoted CSV
python extract_pr.py "EN_ODP-PR-ProvImmCat.xlsx" "extracted_pr.csv"

# Asylum: process Claim Office Type/Province file into unpivoted CSV
python extract_asylum.py "EN_ODP-Asylum-OfficeType_Prov.xlsx" "extracted_asylum.csv"

# Study: process Study level by Province/Territory into unpivoted CSV (defaults shown)
python extract_study.py
python extract_study.py "EN_ODP-TR-Study-IS_PT_study_level_sign.xlsx" "extracted_study.csv"
```

### Generating Sankey Visualizations

After extracting data to CSV files, run the Sankey generation scripts to create interactive visualizations:

```bash
# Generate IMP Sankey chart
# Reads: extracted_imp.csv and sankey_imp_template.html
# Writes: sankey_imp.html
python sankey_imp_data.py

# Generate TFW Sankey chart
# Reads: extracted_tfw.csv and sankey_tfw_template.html
# Writes: sankey_tfw.html
python sankey_tfw_data.py

# Generate PR Sankey chart
# Reads: extracted_pr.csv and sankey_pr_template.html
# Writes: sankey_pr.html
python sankey_pr_data.py

# Generate Asylum Sankey chart
# Reads: extracted_asylum.csv and sankey_asylum_template.html
# Writes: sankey_asylum.html
python sankey_asylum_data.py

# Generate Study Sankey chart
# Reads: extracted_study.csv and sankey_study_template.html
# Writes: sankey_study.html
python sankey_study_data.py
```

**How it works:**
1. The Python script reads the CSV data and builds a hierarchical node/link structure
2. IMP has 3 category levels (category_1 → category_2 → category_3), TFW has 2 (category_1 → category_2)
3. The script validates data aggregates to ensure consistency across hierarchy levels
4. Node and link data is serialized to JSON and injected into the HTML template
5. The generated HTML file contains the complete interactive D3.js Sankey visualization
6. Open the generated `.html` file in a web browser to view and interact with the chart

**Note:** The `.html` files are generated artifacts. If you update the CSV data or modify the template, re-run the appropriate Python script to regenerate the visualization.

## CSV File Data Analytical Usage

Use the csv files for analytical purposes. 

- `extracted_hc.csv` - Processed H&C data
- `extracted_imp.csv` - Processed IMP data
- `extracted_tfw.csv` - Processed TFWP data
 - `extracted_pr.csv` - Processed PR data
 - `extracted_asylum.csv` - Processed Asylum data
 - `extracted_study.csv` - Processed Study data

These have been created specifically to be in a format to be used with common data analysis tools such as Excel Pivot Tables, MS Power BI, Tableau, etc.

Recommend that you set a filter `total_flag`=False to get just detail rows and the analytical tool you use can create the totals by grouping by categories and years etc.

The PDF file `power_bi_report.pdf` contains tables and charts created using this data.

## Dependencies

- Python 3.x
- pandas
- d3-sankey.js v0.12.3

## Data Privacy

This project processes publicly available government data that has already been anonymized and rounded for privacy protection. No personal information is included in the datasets.

## License

This project uses open government data from Canada. Please refer to the original data sources for licensing information.

## Contributing

This is a personal data analysis project. For questions or suggestions, please refer to the original data sources or government documentation.
