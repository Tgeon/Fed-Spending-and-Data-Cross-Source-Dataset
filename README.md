# Cross-Source Dataset: Federal Spending and Financial Data, FY1998 to FY2026

Five cleaned public datasets, with tools to filter, extract and analyze them.

| Source | Coverage |
|---|---|
| USAspending | Federal construction contracts (NAICS 236, 237, 238), FY2021 to FY2025 |
| OpenFEMA Public Assistance | Disaster recovery projects funded by FEMA, FY1998 to FY2026 |
| Federal Audit Clearinghouse (FAC) | Single audits of federal award recipients, FY2016 to FY2026 |
| SEC financial statements | Company filings made in 2025 Q1 to Q4 (reporting periods FY2019 to FY2026) |
| SBA PPP | Paycheck Protection Program loans of $150,000 or more, approved FY2020 to FY2021 |

## Files

- `03_Pipeline/` build, verify and extract scripts
- `05_Workbook/Fed Spending and Data Cross Source Dataset.xlsx` filters, charts, profile and extract builder
- `00_Documentation/` data dictionary, QA report, raw file list

Data is not included (about 6 GB). Rebuild it:

```
pip install -r requirements.txt
python3 03_Pipeline/build_clean.py     # raw files in 01_Raw_Data/ (see 00_Documentation/data_manifest.csv)
python3 03_Pipeline/verify_clean.py
```

## Use

```
python3 03_Pipeline/extract.py --source FAC --year-from 2019 --year-to 2023 --state TX --min-amount 10000
python3 03_Pipeline/extract.py --source USAspending --summarize-by state
```

```sql
SELECT state, sum(amount) FROM unified_records WHERE source = 'OpenFEMA' GROUP BY 1 ORDER BY 2 DESC;
```

Details: `00_Documentation/README.txt`
