#!/usr/bin/env python3
"""Independent QA of the cleaned layer against the raw files. Writes 00_Documentation/qa_report.csv"""
import duckdb, os, csv
ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")); RAW = f"{ROOT}/01_Raw_Data"
db = duckdb.connect(f"{ROOT}/02_Clean_Data/analytics.duckdb", read_only=True)
raw = duckdb.connect(); raw.execute("SET memory_limit='1800MB'; SET threads=3; SET temp_directory='/tmp/duck'")
def R(p, extra=""): return f"read_csv('{RAW}/{p}', all_varchar=true, header=true, parallel=true{extra})"
rows = []
def chk(name, raw_v, clean_v, tol=0.01):
    ok = abs((raw_v or 0) - (clean_v or 0)) <= tol
    rows.append((name, raw_v, clean_v, "PASS" if ok else "FAIL")); print(("PASS " if ok else "FAIL ") + f"{name}: raw={raw_v:,.2f} clean={clean_v:,.2f}")
def one(c, s): return c.sql(s).fetchone()[0]
SEC = [f"04_SEC_Financial_Statements/{q}/sec_{q}_sub.txt" for q in ["2025q1","2025q2","2025q3","2025q4"]]
T = ", delim='\t', quote='', ignore_errors=true"
# ---- row counts and amount totals, raw vs clean
U = "01_USAspending_Contracts/usaspending_contract_transactions.csv"
chk("USAspending transactions: row count", one(raw, f"select count(*) from {R(U)}"), one(db, "select count(*) from usaspending_transactions"))
chk("USAspending transactions: sum federal_action_obligation", one(raw, f"select sum(try_cast(federal_action_obligation as double)) from {R(U)}"), one(db, "select sum(federal_action_obligation) from usaspending_transactions"))
chk("USAspending awards: row count", one(raw, f"select count(*) from {R('01_USAspending_Contracts/usaspending_contract_award_summaries.csv')}"), one(db, "select count(*) from usaspending_awards"))
chk("USAspending awards: sum total_obligated", one(raw, f"select sum(try_cast(total_obligated_amount as double)) from {R('01_USAspending_Contracts/usaspending_contract_award_summaries.csv')}"), one(db, "select sum(total_obligated) from usaspending_awards"))
F = "02_OpenFEMA_Public_Assistance/openfema_pa_funded_projects.csv"
chk("OpenFEMA: row count", one(raw, f"select count(*) from {R(F)}"), one(db, "select count(*) from fema_pa_projects"))
chk("OpenFEMA: sum totalObligated", one(raw, f"select sum(try_cast(totalObligated as double)) from {R(F)}"), one(db, "select sum(total_obligated) from fema_pa_projects"))
chk("OpenFEMA: sum projectAmount", one(raw, f"select sum(try_cast(projectAmount as double)) from {R(F)}"), one(db, "select sum(project_amount) from fema_pa_projects"))
G = "03_FAC_Single_Audit/fac_general.csv"; A = "03_FAC_Single_Audit/fac_federal_awards.csv"; FI = "03_FAC_Single_Audit/fac_findings.csv"
chk("FAC reports: row count", one(raw, f"select count(*) from {R(G)}"), one(db, "select count(*) from fac_reports"))
chk("FAC reports: sum total_amount_expended", one(raw, f"select sum(try_cast(total_amount_expended as double)) from {R(G)}"), one(db, "select sum(total_amount_expended) from fac_reports"))
chk("FAC awards: row count", one(raw, f"select count(*) from {R(A)}"), one(db, "select count(*) from fac_awards"))
chk("FAC awards: sum amount_expended", one(raw, f"select sum(try_cast(amount_expended as double)) from {R(A)}"), one(db, "select sum(amount_expended) from fac_awards"), tol=1)
chk("FAC findings: row count", one(raw, f"select count(*) from {R(FI)}"), one(db, "select count(*) from fac_findings"))
B = "05_SBA_PPP_Loans/sba_ppp_loans_150k_plus_240930.csv"; L = ", encoding='latin-1'"
chk("SBA PPP: row count", one(raw, f"select count(*) from {R(B, L)}"), one(db, "select count(*) from sba_ppp_loans"))
chk("SBA PPP: sum InitialApprovalAmount", one(raw, f"select sum(try_cast(InitialApprovalAmount as double)) from {R(B, L)}"), one(db, "select sum(initial_approval_amount) from sba_ppp_loans"))
chk("SBA PPP: sum ForgivenessAmount", one(raw, f"select sum(try_cast(ForgivenessAmount as double)) from {R(B, L)}"), one(db, "select sum(forgiveness_amount) from sba_ppp_loans"))
subs = "[" + ",".join(f"'{RAW}/{s}'" for s in SEC) + "]"
chk("SEC filings: row count (distinct adsh)", one(raw, f"select count(distinct adsh) from read_csv({subs}, all_varchar=true, header=true, delim='\t', quote='', ignore_errors=true)"), one(db, "select count(*) from sec_filings"))
# ---- key uniqueness
for t, k in [("usaspending_transactions","transaction_key"),("usaspending_awards","award_key"),("fema_pa_projects","project_key"),("fac_reports","report_id"),("sba_ppp_loans","loan_number"),("sec_filings","adsh"),("unified_records","source||'|'||record_id")]:
    n, d = db.sql(f"select count(*), count(distinct {k}) from {t}").fetchone()
    rows.append((f"Unique key {t}.{k}", n, d, "PASS" if n == d else "FAIL")); print(("PASS " if n == d else "FAIL ") + f"unique {t}.{k}: rows={n:,} distinct={d:,}")
# ---- unified ties to source tables
for src, q in [("USAspending","select sum(federal_action_obligation) from usaspending_transactions"),("OpenFEMA","select sum(total_obligated) from fema_pa_projects"),
               ("FAC","select sum(amount_expended) from fac_awards"),("SBA PPP","select sum(initial_approval_amount) from sba_ppp_loans")]:
    chk(f"Unified {src} amount equals source table", one(db, q), one(db, f"select sum(amount) from unified_records where source='{src}'"), tol=1)
# ---- invalid state values dropped to NULL (count of raw non-null states that are not valid)
bad = one(raw, f"select count(*) from {R(F)} where nullif(trim(stateAbbreviation),'') is not null") - one(db, "select count(state) from fema_pa_projects")
rows.append(("OpenFEMA: raw state values rejected as invalid", bad, 0, "INFO")); print("OpenFEMA rejected states:", bad)
bad = one(raw, f"select count(*) from {R(B, L)} where nullif(trim(BorrowerState),'') is not null") - one(db, "select count(borrower_state) from sba_ppp_loans")
rows.append(("SBA PPP: raw BorrowerState values rejected as invalid", bad, 0, "INFO")); print("SBA rejected states:", bad)
bad = one(raw, f"select count(*) from {R(G)} where nullif(trim(auditee_state),'') is not null") - one(db, "select count(auditee_state) from fac_reports")
rows.append(("FAC: raw auditee_state values rejected as invalid", bad, 0, "INFO")); print("FAC rejected states:", bad)
# ---- date coverage
for t, c in [("usaspending_transactions","record_date"),("fema_pa_projects","record_date"),("fac_reports","record_date"),("sba_ppp_loans","record_date"),("sec_filings","record_date")]:
    n, nn, mn, mx = db.sql(f"select count(*), count({c}), min({c}), max({c}) from {t}").fetchone()
    rows.append((f"Date coverage {t}", n, nn, f"{mn} to {mx}")); print(f"{t}: {nn:,}/{n:,} dated, {mn} to {mx}")
with open(f"{ROOT}/00_Documentation/qa_report.csv", "w", newline="") as f:
    w = csv.writer(f); w.writerow(["check", "raw_or_expected", "clean_or_actual", "result"]); w.writerows(rows)
print("FAILS:", sum(1 for r in rows if r[3] == "FAIL"))
