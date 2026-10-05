FEDERAL SPENDING AND FINANCIAL DATA: FOLDER GUIDE
Cleaned data from five public sources, with tools to filter, extract and analyze it.

FOLDERS
00_Documentation   This guide, data_dictionary.csv, qa_report.csv, data_manifest.csv
01_Raw_Data        Original downloads. Do not edit.
02_Clean_Data      analytics.duckdb and parquet/ (same tables)
03_Pipeline        build_clean.py, verify_clean.py, export_workbook_inputs.py, build_filter_workbook.py, extract.py
04_Extracts        Files written by extract.py
05_Workbook        Fed Spending and Data Cross Source Dataset.xlsx
99_Archive_Audit_Framing   Earlier audit work. Kept, not used.

TABLES (02_Clean_Data)
unified_records            11.6M rows. All five sources, same columns: source, record_id, record_date, calendar_year,
                           calendar_month, fiscal_year, state, county, category, subcategory, industry_code, entity_name,
                           entity_name_std, entity_id, amount, amount_measure, data_flag
usaspending_transactions   42K     contract transactions (NAICS 236, 237, 238; FY2021 to FY2025)
usaspending_awards         20K     contract awards
fema_pa_projects           850K    Public Assistance projects
fac_reports                416K    single audit reports
fac_awards                 7.3M    federal programs audited
fac_findings               683K    audit findings
sec_filings                26K     SEC filings, 2025 Q1 to Q4
sec_tags                   330K    SEC tag dictionary
sec_facts                  5.4M    USD values, consolidated level
sba_ppp_loans              969K    PPP loans of $150,000 or more

HOW TO USE
1. Excel: open 05_Workbook/Fed Spending and Data Cross Source Dataset.xlsx. Set filters on Explorer; Extract_Request gives the command.
2. Extract:  python3 03_Pipeline/extract.py --source FAC --year-from 2019 --year-to 2023 --state TX --min-amount 10000
   Summary:  python3 03_Pipeline/extract.py --source USAspending --summarize-by state
   Help:     python3 03_Pipeline/extract.py --help   (needs duckdb; xlsx output also needs pandas and openpyxl)
3. SQL: open 02_Clean_Data/analytics.duckdb read-only.
   SELECT state, sum(amount) FROM unified_records WHERE source = 'OpenFEMA' GROUP BY 1 ORDER BY 2 DESC;

CLEANING
- Text converted to dates and numbers. Blanks set to null. Whitespace trimmed.
- State codes checked (50 states, DC, territories); others set to null.
- entity_name_std: upper case, no punctuation, for matching names across sources.
- fiscal_year: federal fiscal year (October to September).
- SBA PPP business types: a stray character replaced with a dash.
- FAC records migrated from the older system are kept and flagged (data_flag = migration_record).
- OpenFEMA records with no first obligation date use the declaration date (data_flag = date_from_declaration).
- Left out: FAC contacts, emails and phones; SBA PPP addresses, race, ethnicity and gender; non-USD, segment and
  co-registrant SEC values; SEC presentation files.
- Row counts and dollar totals tie to the raw files (qa_report.csv). Keys are unique in every table.

NOTES
- SEC amounts are separate tags. Pick one tag; do not sum across tags.
- The sources measure different dollars (obligations, expenditures, loan approvals). Compare counts freely, dollars with care.
- The database is built in local temp space and copied in (this folder blocks deletes). Rebuild: python3 03_Pipeline/build_clean.py (about 90 seconds).
- FAC agency names come from a hand-kept prefix map in build_clean.py; others show as Other federal agency (prefix NN).
