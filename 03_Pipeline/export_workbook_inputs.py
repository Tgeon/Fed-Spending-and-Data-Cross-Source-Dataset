#!/usr/bin/env python3
"""Export summary cubes, profile and data dictionary used by the Excel filter workbook."""
import duckdb, os, re, json
ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
OUT = f"{ROOT}/05_Workbook/build_inputs"; os.makedirs(OUT, exist_ok=True)
db = duckdb.connect(f"{ROOT}/02_Clean_Data/analytics.duckdb", read_only=True)
db.execute("SET memory_limit='1800MB'; SET threads=3; SET temp_directory='/tmp/duck'")
TOPN = 14
BANDS = ["No amount", "Negative", "Zero", "Under $10K", "$10K to $100K", "$100K to $1M", "$1M to $10M", "$10M to $100M", "$100M and over"]
db.execute(f"""CREATE TEMP TABLE topc AS SELECT source, category FROM (
   SELECT source, category, row_number() OVER (PARTITION BY source ORDER BY count(*) DESC) rn FROM unified_records WHERE category IS NOT NULL GROUP BY 1,2) WHERE rn <= {TOPN}""")
db.execute("""CREATE TEMP TABLE u AS SELECT u.source, u.calendar_year AS y, u.calendar_month AS m, coalesce(u.state,'Unknown') AS state,
   u.source||' | '||CASE WHEN u.category IS NULL THEN 'Not stated' WHEN t.category IS NULL THEN 'Other' ELSE u.category END AS cat, u.amount,
   CASE WHEN u.amount IS NULL THEN 0 WHEN u.amount<0 THEN 1 WHEN u.amount=0 THEN 2 WHEN u.amount<10000 THEN 3 WHEN u.amount<100000 THEN 4
        WHEN u.amount<1000000 THEN 5 WHEN u.amount<10000000 THEN 6 WHEN u.amount<100000000 THEN 7 ELSE 8 END AS band
   FROM unified_records u LEFT JOIN topc t ON u.source=t.source AND u.category=t.category""")
db.execute(f"COPY (SELECT source, y AS year, state, cat AS category, band, count(*) AS n, coalesce(sum(amount),0) AS amount FROM u GROUP BY ALL ORDER BY ALL) TO '{OUT}/cube_main.csv' (HEADER)")
db.execute(f"COPY (SELECT source, y AS year, m AS month, state, cat AS category, count(*) AS n, coalesce(sum(amount),0) AS amount FROM u GROUP BY ALL ORDER BY ALL) TO '{OUT}/cube_month.csv' (HEADER)")
lists = dict(
    sources=[r[0] for r in db.sql("select distinct source from u order by 1").fetchall()],
    states=[r[0] for r in db.sql("select distinct state from u order by state='Unknown', 1").fetchall()],
    categories=[r[0] for r in db.sql("select distinct cat from u order by 1").fetchall()],
    years=[r[0] for r in db.sql("select distinct y from u where y is not null order by 1").fetchall()],
    bands=BANDS)
json.dump(lists, open(f"{OUT}/lists.json", "w"), indent=1)

# ---- profile
prof = {"tables": [], "by_source": []}
for src in lists["sources"]:
    r = db.sql(f"""select count(*), count(record_date), min(record_date), max(record_date), count(state), count(category), count(entity_name), count(amount),
        sum((amount<0)::int), sum((amount=0)::int), min(amount), quantile_cont(amount,0.5), quantile_cont(amount,0.9), max(amount), avg(amount), sum(amount)
        from unified_records where source='{src}'""").fetchone()
    k = ["rows","dated","date_min","date_max","with_state","with_category","with_entity","with_amount","negative","zero","amt_min","amt_median","amt_p90","amt_max","amt_mean","amt_sum"]
    d = dict(zip(k, r)); d["source"] = src; d["date_min"] = str(d["date_min"]); d["date_max"] = str(d["date_max"]); prof["by_source"].append(d)
tabs = [r[0] for r in db.sql("select table_name from duckdb_tables() where not temporary order by 1").fetchall()]
# ---- dictionary with derivation parsed from the build script
src = open(f"{ROOT}/03_Pipeline/build_clean.py").read()
deriv = {}
for body, tname in re.findall(r'out\(f?"""(.*?)""",\s*"(\w+)"\)', src, re.S):
    if tname == "unified_records": continue
    i = body.find("FROM {csv(")
    if i < 0: i = body.find("FROM {pq(")
    j = body.rfind("SELECT", 0, i)
    sel = body[j + 6:i]
    items, depth, cur, q = [], 0, "", False
    for ch in sel:
        if ch == "'" : q = not q
        if not q:
            if ch == "(": depth += 1
            elif ch == ")": depth -= 1
            elif ch == "," and depth == 0: items.append(cur); cur = ""; continue
        cur += ch
    items.append(cur)
    for it in items:
        it = " ".join(it.split())
        m = re.search(r"^(.*)\s+AS\s+(\w+)$", it)
        if m: deriv[(tname, m.group(2))] = m.group(1).strip()
NOTES = {
 "record_date": "Primary date of the record (see table notes)", "calendar_year": "Year of record_date", "calendar_month": "Month of record_date (1 to 12)",
 "fiscal_year": "U.S. federal fiscal year of record_date (October to September; October 2024 is fiscal 2025)",
 "record_date_basis": "Which date was used for record_date", "recipient_name_std": "Upper-case, punctuation removed, spaces collapsed; use to match names across sources",
 "auditee_name_std": "Standardized name for matching", "company_name_std": "Standardized name for matching", "borrower_name_std": "Standardized name for matching",
 "is_migration_record": "TRUE for records migrated from the older Federal Audit Clearinghouse system; fewer fields are reliable",
 "assistance_listing": "Assistance Listing number (agency prefix.extension), formerly CFDA number", "federal_agency": "Agency name mapped from the two-digit prefix (maintained in build_clean.py)",
}
dic = []
for t in tabs:
    n = db.sql(f"select count(*) from {t}").fetchone()[0]
    cols = db.sql(f"describe {t}").fetchall()
    exprs = ", ".join(f"count({c[0]}) AS c{i}, approx_count_distinct({c[0]}) AS d{i}" for i, c in enumerate(cols))
    row = db.sql(f"select {exprs} from {t}").fetchone()
    prof["tables"].append(dict(table=t, rows=n, columns=len(cols)))
    for i, c in enumerate(cols):
        ex = db.sql(f'select cast(min("{c[0]}") as varchar) from {t}').fetchone()[0]
        dic.append(dict(table=t, column=c[0], type=c[1], null_pct=round(1 - row[2*i] / n, 4) if n else None, distinct_approx=row[2*i+1],
                        example=(ex or "")[:60], derivation=deriv.get((t, c[0]), ""), note=NOTES.get(c[0], "")))
json.dump(prof, open(f"{OUT}/profile.json", "w"), indent=1, default=str)
json.dump(dic, open(f"{OUT}/dictionary.json", "w"), indent=1, default=str)
print("cube rows:", db.sql("select count(*) from (select source,y,state,cat,band from u group by all)").fetchone()[0],
      db.sql("select count(*) from (select source,y,m,state,cat from u group by all)").fetchone()[0])
print("categories:", len(lists["categories"]), "states:", len(lists["states"]), "years:", lists["years"][0], lists["years"][-1])
for s in lists["sources"]: print(s, [c for c in lists["categories"] if c.startswith(s + " |")][:6])
