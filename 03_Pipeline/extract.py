#!/usr/bin/env python3
"""Pull records out of the cleaned database.

Examples (run from the project folder):
  python3 03_Pipeline/extract.py --source FAC --year-from 2019 --year-to 2023 --state TX --category "Department of Education" --min-amount 10000
  python3 03_Pipeline/extract.py --table fema_pa_projects --year-from 2020 --where "state = 'LA' AND total_obligated > 1000000" --format xlsx
  python3 03_Pipeline/extract.py --source USAspending --summarize-by state
  python3 03_Pipeline/extract.py --source "SBA PPP" --state NY --count-only

Needs: pip install duckdb   (add pandas openpyxl only for --format xlsx)
Tables: unified (default; filters below all work), usaspending_transactions, usaspending_awards, fema_pa_projects,
        fac_reports, fac_awards, fac_findings, sec_filings, sec_tags, sec_facts, sba_ppp_loans.
        For any table other than unified, use --where for column filters (see the Data_Dictionary sheet); --year-from/--year-to
        work where the table has a calendar_year column.
"""
import argparse, os, sys, time, re
import duckdb

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TABLES = ["unified", "usaspending_transactions", "usaspending_awards", "fema_pa_projects", "fac_reports", "fac_awards", "fac_findings",
          "sec_filings", "sec_tags", "sec_facts", "sba_ppp_loans"]
GROUPS = ["source", "year", "month", "state", "category", "subcategory", "fiscal_year"]
GROUP_COL = {"year": "calendar_year", "month": "calendar_month", "fiscal_year": "fiscal_year"}

ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
ap.add_argument("--table", default="unified", choices=TABLES)
ap.add_argument("--source", help="USAspending, OpenFEMA, FAC, SEC or 'SBA PPP' (unified only)")
ap.add_argument("--year-from", type=int); ap.add_argument("--year-to", type=int)
ap.add_argument("--state", help="Two-letter code, or Unknown for records with no valid state (unified only)")
ap.add_argument("--category", help="Exact category name, e.g. 'Department of Education' (unified only)")
ap.add_argument("--category-contains", help="Text the category must contain (unified only)")
ap.add_argument("--entity", help="Text the standardized entity name must contain, e.g. 'TURNER CONSTRUCTION' (unified only)")
ap.add_argument("--min-amount", type=float); ap.add_argument("--max-amount", type=float)
ap.add_argument("--where", help="Extra SQL condition on the chosen table's columns")
ap.add_argument("--columns", help="Comma-separated columns to keep (default all)")
ap.add_argument("--limit", type=int, help="Maximum rows")
ap.add_argument("--summarize-by", choices=GROUPS, help="Return counts and totals by this field instead of records (unified only)")
ap.add_argument("--count-only", action="store_true", help="Only print how many records match")
ap.add_argument("--format", default="csv", choices=["csv", "xlsx", "parquet"]); ap.add_argument("--out", help="Output file (default 04_Extracts/extract_<time>.<format>)")
ap.add_argument("--db", default=os.path.join(ROOT, "02_Clean_Data", "analytics.duckdb"))
a = ap.parse_args()

tname = "unified_records" if a.table == "unified" else a.table
con = duckdb.connect(a.db, read_only=True)
cols = [r[0] for r in con.execute(f"DESCRIBE {tname}").fetchall()]
cond, par = [], []
def need_unified(flag):
    if a.table != "unified": sys.exit(f"{flag} works with --table unified only. For {a.table} use --where.")
if a.source: need_unified("--source"); cond.append("lower(source) = lower(?)"); par.append(a.source)
if a.year_from or a.year_to:
    if "calendar_year" not in cols: sys.exit(f"{a.table} has no calendar_year column; use --where on its own date columns.")
    if a.year_from: cond.append("calendar_year >= ?"); par.append(a.year_from)
    if a.year_to: cond.append("calendar_year <= ?"); par.append(a.year_to)
if a.state:
    need_unified("--state")
    if a.state.lower() == "unknown": cond.append("state IS NULL")
    else: cond.append("state = ?"); par.append(a.state.upper())
if a.category: need_unified("--category"); cond.append("lower(category) = lower(?)"); par.append(a.category)
if a.category_contains: need_unified("--category-contains"); cond.append("lower(category) LIKE ?"); par.append("%" + a.category_contains.lower() + "%")
if a.entity: need_unified("--entity"); cond.append("entity_name_std LIKE ?"); par.append("%" + a.entity.upper() + "%")
if a.min_amount is not None: need_unified("--min-amount"); cond.append("amount >= ?"); par.append(a.min_amount)
if a.max_amount is not None: need_unified("--max-amount"); cond.append("amount <= ?"); par.append(a.max_amount)
if a.where: cond.append("(" + a.where + ")")
where = (" WHERE " + " AND ".join(cond)) if cond else ""

n = con.execute(f"SELECT count(*) FROM {tname}{where}", par).fetchone()[0]
print(f"{n:,} matching records in {tname}")
if a.count_only: sys.exit(0)
if n == 0: sys.exit("Nothing matched; nothing written.")

if a.summarize_by:
    need_unified("--summarize-by"); g = GROUP_COL.get(a.summarize_by, a.summarize_by)
    sec_warn = (not a.category and not a.category_contains and (not a.source or a.source.lower() == "sec"))
    amt = "sum(amount) FILTER (WHERE source <> 'SEC')" if sec_warn else "sum(amount)"
    if sec_warn: print("Note: SEC values are separate accounting tags, so SEC is left out of the dollar totals (choose a tag with --category to include it).")
    sql = f"SELECT {g} AS {a.summarize_by}, count(*) AS records, {amt} AS total_amount, avg(amount) FILTER (WHERE source <> 'SEC' OR {str(not sec_warn).upper()}) AS average_amount FROM {tname}{where} GROUP BY 1 ORDER BY 1"
else:
    keep = a.columns if a.columns else "*"
    if a.columns:
        bad = [c for c in a.columns.split(",") if c.strip() not in cols]
        if bad: sys.exit(f"Unknown columns: {bad}. Available: {cols}")
    sql = f"SELECT {keep} FROM {tname}{where}" + (f" LIMIT {a.limit}" if a.limit else "")

os.makedirs(os.path.join(ROOT, "04_Extracts"), exist_ok=True)
out = a.out or os.path.join(ROOT, "04_Extracts", time.strftime("extract_%Y%m%d_%H%M%S") + "." + a.format)
if not os.path.isabs(out): out = os.path.join(os.getcwd(), out)
def lit(v): return "'" + v.replace("'", "''") + "'" if isinstance(v, str) else repr(v)
for v in par: sql = sql.replace("?", lit(v), 1)   # parameters are only ours (checked values); COPY cannot take bound parameters
if a.format == "xlsx":
    if min(n, a.limit or n) > 1048575 and not a.summarize_by: sys.exit("Too many rows for an Excel sheet (limit 1,048,575). Use --format csv or parquet, or narrow the filters.")
    con.execute(sql).df().to_excel(out, index=False)
elif a.format == "parquet":
    con.execute(f"COPY ({sql}) TO '{out}' (FORMAT PARQUET, COMPRESSION ZSTD)")
else:
    con.execute(f"COPY ({sql}) TO '{out}' (FORMAT CSV, HEADER)")
print("Wrote", out)
