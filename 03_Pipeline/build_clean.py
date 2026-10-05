#!/usr/bin/env python3
"""Build the cleaned data layer from the raw files (raw files are only read).
Outputs: 02_Clean_Data/parquet/*.parquet and 02_Clean_Data/analytics.duckdb
Run:  python3 build_clean.py
"""
import duckdb, os, time

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
RAW = f"{ROOT}/01_Raw_Data"
OUT = f"{ROOT}/02_Clean_Data"
PQ = f"{OUT}/parquet"
os.makedirs(PQ, exist_ok=True); os.makedirs("/tmp/duck", exist_ok=True)
con = duckdb.connect()
con.execute("SET memory_limit='1800MB'; SET threads=3; SET temp_directory='/tmp/duck'; SET preserve_insertion_order=false;")
T0 = time.time()
def log(m): print(f"[{time.time()-T0:6.0f}s] {m}", flush=True)

# ---------------------------------------------------------------- cleaning macros
con.execute("CREATE MACRO nz(x) AS NULLIF(trim(x), '')")
con.execute("CREATE MACRO to_dt(x) AS CAST(COALESCE(try_strptime(substr(trim(x),1,10),'%Y-%m-%d'), try_strptime(trim(x),'%m/%d/%Y'), try_strptime(trim(x),'%Y%m%d')) AS DATE)")
con.execute("CREATE MACRO to_num(x) AS TRY_CAST(NULLIF(trim(x),'') AS DOUBLE)")
con.execute("CREATE MACRO to_int(x) AS TRY_CAST(NULLIF(trim(x),'') AS INTEGER)")
con.execute("CREATE MACRO to_bool(x) AS CASE upper(trim(x)) WHEN 'Y' THEN TRUE WHEN 'YES' THEN TRUE WHEN 'TRUE' THEN TRUE WHEN 'N' THEN FALSE WHEN 'NO' THEN FALSE WHEN 'FALSE' THEN FALSE ELSE NULL END")
con.execute("""CREATE MACRO std_name(x) AS NULLIF(trim(regexp_replace(regexp_replace(replace(upper(x),'&',' AND '),'[^A-Z0-9 ]',' ','g'),'\\s+',' ','g')),'')""")
STATES = ("AL AK AZ AR CA CO CT DE DC FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY "
          "AS GU MP PR VI").split()
con.execute("CREATE MACRO st(x) AS CASE WHEN upper(trim(x)) IN (" + ",".join(f"'{s}'" for s in STATES) + ") THEN upper(trim(x)) END")
con.execute("CREATE MACRO fy_of(d) AS CASE WHEN d IS NULL THEN NULL WHEN month(d) >= 10 THEN year(d) + 1 ELSE year(d) END")

def csv(path, tsv=False, enc=None):
    p = "[" + ",".join(f"'{x}'" for x in path) + "]" if isinstance(path, list) else f"'{path}'"
    o = "all_varchar=true, header=true, parallel=true"
    if tsv: o += ", delim='\t', quote='', ignore_errors=true"
    if enc: o += f", encoding='{enc}'"
    return f"read_csv({p}, {o})"

def out(sql, name):
    con.execute(f"COPY ({sql}) TO '{PQ}/{name}.parquet' (FORMAT PARQUET, COMPRESSION ZSTD)")
    n = con.execute(f"SELECT count(*) FROM read_parquet('{PQ}/{name}.parquet')").fetchone()[0]
    log(f"{name}: {n:,} rows"); return n
def pq(name): return f"read_parquet('{PQ}/{name}.parquet')"
def dates(col):  # standard derived date columns for a fact table
    return f"{col} AS record_date, year({col}) AS calendar_year, month({col}) AS calendar_month, fy_of({col}) AS fiscal_year"

# ================================================================ USAspending
log("USAspending")
U = RAW + "/01_USAspending_Contracts"
out(f"""SELECT * EXCLUDE(d), {dates('d')} FROM (SELECT
    nz(contract_transaction_unique_key) AS transaction_key, nz(contract_award_unique_key) AS award_key,
    nz(award_id_piid) AS piid, nz(modification_number) AS modification_number, nz(parent_award_id_piid) AS parent_piid,
    to_dt(action_date) AS d, to_int(action_date_fiscal_year) AS action_fiscal_year,
    to_num(federal_action_obligation) AS federal_action_obligation,
    to_num(total_dollars_obligated) AS award_total_obligated, to_num(current_total_value_of_award) AS award_current_value,
    to_num(potential_total_value_of_award) AS award_potential_value, to_num(base_and_all_options_value) AS base_and_all_options_value,
    nz(awarding_agency_code) AS awarding_agency_code, nz(awarding_agency_name) AS awarding_agency,
    nz(awarding_sub_agency_name) AS awarding_sub_agency, nz(awarding_office_name) AS awarding_office,
    nz(funding_agency_name) AS funding_agency, nz(funding_sub_agency_name) AS funding_sub_agency,
    nz(recipient_uei) AS recipient_uei, nz(recipient_name) AS recipient_name, std_name(recipient_name) AS recipient_name_std,
    nz(recipient_parent_name) AS recipient_parent_name, nz(recipient_city_name) AS recipient_city,
    nz(recipient_county_name) AS recipient_county, st(recipient_state_code) AS recipient_state, nz(recipient_zip_4_code) AS recipient_zip,
    nz(primary_place_of_performance_city_name) AS pop_city, nz(primary_place_of_performance_county_name) AS pop_county,
    st(primary_place_of_performance_state_code) AS pop_state, nz(primary_place_of_performance_zip_4) AS pop_zip,
    nz(award_type) AS award_type, nz(type_of_contract_pricing) AS contract_pricing, nz(action_type) AS action_type,
    nz(transaction_description) AS description, nz(naics_code) AS naics_code, nz(naics_description) AS naics_description,
    nz(product_or_service_code) AS psc_code, nz(product_or_service_code_description) AS psc_description,
    nz(extent_competed) AS extent_competed, nz(type_of_set_aside) AS set_aside, to_int(number_of_offers_received) AS offers_received,
    to_dt(period_of_performance_start_date) AS pop_start_date, to_dt(period_of_performance_current_end_date) AS pop_end_date
  FROM {csv(U + '/usaspending_contract_transactions.csv')})""", "usaspending_transactions")
out(f"""SELECT nz(contract_award_unique_key) AS award_key, nz(award_id_piid) AS piid, nz(parent_award_id_piid) AS parent_piid,
    to_num(total_obligated_amount) AS total_obligated, to_num(current_total_value_of_award) AS current_value,
    to_num(potential_total_value_of_award) AS potential_value,
    to_dt(award_base_action_date) AS base_action_date, to_int(award_base_action_date_fiscal_year) AS base_fiscal_year,
    to_dt(award_latest_action_date) AS latest_action_date, to_int(award_latest_action_date_fiscal_year) AS latest_fiscal_year,
    to_dt(period_of_performance_start_date) AS pop_start_date, to_dt(period_of_performance_current_end_date) AS pop_end_date,
    nz(awarding_agency_name) AS awarding_agency, nz(awarding_sub_agency_name) AS awarding_sub_agency, nz(awarding_office_name) AS awarding_office,
    nz(funding_agency_name) AS funding_agency,
    nz(recipient_uei) AS recipient_uei, nz(recipient_name) AS recipient_name, std_name(recipient_name) AS recipient_name_std,
    nz(recipient_parent_name) AS recipient_parent_name, nz(recipient_city_name) AS recipient_city, st(recipient_state_code) AS recipient_state,
    nz(recipient_zip_4_code) AS recipient_zip,
    nz(primary_place_of_performance_city_name) AS pop_city, nz(primary_place_of_performance_county_name) AS pop_county,
    st(primary_place_of_performance_state_code) AS pop_state,
    nz(award_type) AS award_type, nz(type_of_contract_pricing) AS contract_pricing, nz(prime_award_base_transaction_description) AS description,
    nz(naics_code) AS naics_code, nz(naics_description) AS naics_description, nz(product_or_service_code) AS psc_code,
    nz(product_or_service_code_description) AS psc_description, nz(extent_competed) AS extent_competed, nz(type_of_set_aside) AS set_aside,
    to_int(number_of_offers_received) AS offers_received, to_int(number_of_actions) AS number_of_actions
  FROM {csv(U + '/usaspending_contract_award_summaries.csv')}""", "usaspending_awards")

# ================================================================ OpenFEMA
log("OpenFEMA")
out(f"""SELECT * EXCLUDE(d), {dates('d')} FROM (SELECT
    disasterNumber||'|'||pwNumber||'|'||applicantId AS project_key, to_int(disasterNumber) AS disaster_number, nz(pwNumber) AS pw_number,
    nz(applicantId) AS applicant_id, nz(applicationTitle) AS project_title, nz(incidentType) AS incident_type,
    to_dt(declarationDate) AS declaration_date, to_dt(firstObligationDate) AS first_obligation_date, to_dt(lastObligationDate) AS last_obligation_date,
    COALESCE(to_dt(firstObligationDate), to_dt(declarationDate)) AS d,
    CASE WHEN to_dt(firstObligationDate) IS NOT NULL THEN 'first_obligation_date' WHEN to_dt(declarationDate) IS NOT NULL THEN 'declaration_date' END AS record_date_basis,
    nz(damageCategoryCode) AS damage_category_code, nz(damageCategoryDescrip) AS damage_category, nz(projectStatus) AS project_status,
    nz(projectProcessStep) AS process_step, nz(projectSize) AS project_size, nz(county) AS county, nz(countyCode) AS county_fips_part,
    st(stateAbbreviation) AS state, nz(stateNumberCode) AS state_fips,
    to_num(projectAmount) AS project_amount, to_num(federalShareObligated) AS federal_share_obligated, to_num(totalObligated) AS total_obligated,
    to_num(mitigationAmount) AS mitigation_amount, to_dt(lastRefresh) AS source_last_refresh
  FROM {csv(RAW + '/02_OpenFEMA_Public_Assistance/openfema_pa_funded_projects.csv')})""", "fema_pa_projects")

# ================================================================ FAC
log("FAC")
F = RAW + "/03_FAC_Single_Audit"
out(f"""SELECT * EXCLUDE(d), {dates('d')} FROM (SELECT
    nz(report_id) AS report_id, to_int(audit_year) AS audit_year, nz(auditee_uei) AS auditee_uei, nz(auditee_ein) AS auditee_ein,
    nz(auditee_name) AS auditee_name, std_name(auditee_name) AS auditee_name_std, nz(auditee_city) AS auditee_city,
    st(auditee_state) AS auditee_state, nz(auditee_zip) AS auditee_zip, nz(entity_type) AS entity_type,
    nz(auditor_firm_name) AS auditor_firm, st(auditor_state) AS auditor_state, nz(auditor_ein) AS auditor_ein,
    nz(cognizant_agency) AS cognizant_agency, nz(oversight_agency) AS oversight_agency,
    nz(audit_type) AS audit_type, (nz(audit_type) = 'GSA_MIGRATION') AS is_migration_record,
    nz(gaap_results) AS gaap_results, to_bool(is_going_concern_included) AS going_concern_included,
    to_bool(is_internal_control_deficiency_disclosed) AS internal_control_deficiency,
    to_bool(is_internal_control_material_weakness_disclosed) AS material_weakness,
    to_bool(is_material_noncompliance_disclosed) AS material_noncompliance, to_bool(is_low_risk_auditee) AS low_risk_auditee,
    to_dt(fy_start_date) AS fy_start_date, to_dt(fy_end_date) AS fy_end_date, to_dt(submitted_date) AS submitted_date,
    to_dt(fac_accepted_date) AS d, to_num(total_amount_expended) AS total_amount_expended, to_int(number_months) AS number_months
  FROM {csv(F + '/fac_general.csv')})""", "fac_reports")
out(f"""SELECT nz(report_id) AS report_id, to_int(audit_year) AS audit_year, to_dt(fac_accepted_date) AS fac_accepted_date,
    nz(award_reference) AS award_reference, nz(federal_agency_prefix) AS agency_prefix, nz(federal_award_extension) AS award_extension,
    nz(federal_agency_prefix)||'.'||nz(federal_award_extension) AS assistance_listing,
    CASE nz(federal_agency_prefix) WHEN '10' THEN 'Department of Agriculture' WHEN '11' THEN 'Department of Commerce' WHEN '12' THEN 'Department of Defense'
      WHEN '14' THEN 'Department of Housing and Urban Development' WHEN '15' THEN 'Department of the Interior' WHEN '16' THEN 'Department of Justice'
      WHEN '17' THEN 'Department of Labor' WHEN '19' THEN 'Department of State' WHEN '20' THEN 'Department of Transportation'
      WHEN '21' THEN 'Department of the Treasury' WHEN '23' THEN 'Appalachian Regional Commission' WHEN '43' THEN 'NASA'
      WHEN '45' THEN 'Arts, Humanities and Library Services (NEA, NEH, IMLS)' WHEN '47' THEN 'National Science Foundation' WHEN '59' THEN 'Small Business Administration'
      WHEN '64' THEN 'Department of Veterans Affairs' WHEN '66' THEN 'Environmental Protection Agency' WHEN '81' THEN 'Department of Energy'
      WHEN '84' THEN 'Department of Education' WHEN '89' THEN 'National Archives' WHEN '90' THEN 'Administration for Community Living'
      WHEN '93' THEN 'Department of Health and Human Services' WHEN '94' THEN 'AmeriCorps' WHEN '96' THEN 'Social Security Administration'
      WHEN '97' THEN 'Department of Homeland Security' WHEN '98' THEN 'USAID' ELSE 'Other federal agency (prefix ' || coalesce(nz(federal_agency_prefix),'?') || ')' END AS federal_agency,
    nz(federal_program_name) AS program_name, to_num(amount_expended) AS amount_expended, nz(cluster_name) AS cluster_name,
    to_num(federal_program_total) AS program_total, to_bool(is_major) AS is_major_program, to_bool(is_loan) AS is_loan,
    to_num(loan_balance) AS loan_balance, to_bool(is_direct) AS is_direct, nz(audit_report_type) AS audit_report_type,
    to_int(findings_count) AS findings_count, to_bool(is_passthrough_award) AS is_passthrough, to_num(passthrough_amount) AS passthrough_amount
  FROM {csv(F + '/fac_federal_awards.csv')}""", "fac_awards")
out(f"""SELECT nz(report_id) AS report_id, to_int(audit_year) AS audit_year, to_dt(fac_accepted_date) AS fac_accepted_date,
    nz(award_reference) AS award_reference, nz(reference_number) AS finding_reference, to_bool(is_material_weakness) AS material_weakness,
    to_bool(is_modified_opinion) AS modified_opinion, to_bool(is_other_findings) AS other_findings, to_bool(is_other_matters) AS other_matters,
    to_bool(is_questioned_costs) AS questioned_costs, to_bool(is_repeat_finding) AS repeat_finding,
    to_bool(is_significant_deficiency) AS significant_deficiency, nz(type_requirement) AS type_requirement
  FROM {csv(F + '/fac_findings.csv')}""", "fac_findings")

# ================================================================ SEC
log("SEC")
S = RAW + "/04_SEC_Financial_Statements"; QS = ["2025q1", "2025q2", "2025q3", "2025q4"]
out(f"""SELECT * EXCLUDE(d), {dates('d')} FROM (SELECT
    nz(adsh) AS adsh, to_int(cik) AS cik, nz(name) AS company_name, std_name(name) AS company_name_std, nz(sic) AS sic,
    nz(form) AS form, to_dt(period) AS period_end, to_int(fy) AS fiscal_year_reported, nz(fp) AS fiscal_period, nz(afs) AS filer_category,
    nz(fye) AS fiscal_year_end_mmdd, nz(countryba) AS business_country, st(stprba) AS business_state, nz(cityba) AS business_city,
    nz(countryinc) AS incorporation_country, st(stprinc) AS incorporation_state, nz(ein) AS ein, to_bool(wksi) AS well_known_seasoned_issuer,
    to_dt(filed) AS d, try_strptime(substr(accepted,1,19), '%Y-%m-%d %H:%M:%S') AS accepted_ts,
    to_bool(prevrpt) AS previous_report, to_bool(detail) AS has_detail
  FROM {csv([f'{S}/{q}/sec_{q}_sub.txt' for q in QS], tsv=True)}) QUALIFY row_number() OVER (PARTITION BY adsh) = 1""", "sec_filings")
out(f"""SELECT DISTINCT nz(tag) AS tag, nz(version) AS version, to_bool(custom) AS is_custom, to_bool(abstract) AS is_abstract, nz(datatype) AS datatype,
    nz(iord) AS period_type, nz(crdr) AS balance_type, nz(tlabel) AS label
  FROM {csv([f'{S}/{q}/sec_{q}_tag.txt' for q in QS], tsv=True)}""", "sec_tags")
out(f"""SELECT nz(n.adsh) AS adsh, s.cik, nz(n.tag) AS tag, nz(n.version) AS version, to_dt(n.ddate) AS value_date, to_int(n.qtrs) AS quarters_covered,
    to_num(n.value) AS value, (to_dt(n.ddate) = s.period_end) AS is_current_period
  FROM {csv([f'{S}/{q}/sec_{q}_num.txt' for q in QS], tsv=True)} n JOIN {pq('sec_filings')} s ON nz(n.adsh) = s.adsh
  WHERE n.uom = 'USD' AND coalesce(trim(n.segments),'') = '' AND coalesce(trim(n.coreg),'') = '' AND to_num(n.value) IS NOT NULL""", "sec_facts")

# ================================================================ SBA PPP
log("SBA PPP")
out(f"""SELECT * EXCLUDE(d), {dates('d')} FROM (SELECT
    nz(LoanNumber) AS loan_number, to_dt(DateApproved) AS d, nz(BorrowerName) AS borrower_name, std_name(BorrowerName) AS borrower_name_std,
    nz(BorrowerCity) AS borrower_city, st(BorrowerState) AS borrower_state, nz(BorrowerZip) AS borrower_zip,
    nz(ProcessingMethod) AS processing_method, nz(LoanStatus) AS loan_status, to_dt(LoanStatusDate) AS loan_status_date, to_int(Term) AS term_months,
    to_num(InitialApprovalAmount) AS initial_approval_amount, to_num(CurrentApprovalAmount) AS current_approval_amount,
    to_num(UndisbursedAmount) AS undisbursed_amount, to_num(ForgivenessAmount) AS forgiveness_amount, to_dt(ForgivenessDate) AS forgiveness_date,
    nz(NAICSCode) AS naics_code, substr(nz(NAICSCode),1,2) AS naics_sector, NULLIF(trim(regexp_replace(replace(BusinessType,'¿','-'),'\\s+',' ','g')),'') AS business_type, nz(BusinessAgeDescription) AS business_age,
    to_int(JobsReported) AS jobs_reported, nz(ProjectCountyName) AS project_county, st(ProjectState) AS project_state,
    nz(RuralUrbanIndicator) AS rural_urban, nz(HubzoneIndicator) AS hubzone, nz(LMIIndicator) AS lmi, nz(NonProfit) AS non_profit, nz(Veteran) AS veteran,
    to_num(UTILITIES_PROCEED) AS proceeds_utilities, to_num(PAYROLL_PROCEED) AS proceeds_payroll, to_num(MORTGAGE_INTEREST_PROCEED) AS proceeds_mortgage_interest,
    to_num(RENT_PROCEED) AS proceeds_rent, to_num(HEALTH_CARE_PROCEED) AS proceeds_health_care, to_num(DEBT_INTEREST_PROCEED) AS proceeds_debt_interest,
    nz(OriginatingLender) AS originating_lender, st(OriginatingLenderState) AS originating_lender_state,
    nz(ServicingLenderName) AS servicing_lender, nz(FranchiseName) AS franchise_name
  FROM {csv(RAW + '/05_SBA_PPP_Loans/sba_ppp_loans_150k_plus_240930.csv', enc='latin-1')})""", "sba_ppp_loans")

# ================================================================ Unified table (one row per monetary record, all five sources)
log("Unified")
out(f"""
SELECT 'USAspending' AS source, transaction_key AS record_id, record_date, calendar_year, calendar_month, fiscal_year,
    pop_state AS state, pop_county AS county, awarding_agency AS category, naics_code||' '||coalesce(naics_description,'') AS subcategory, naics_code AS industry_code,
    recipient_name AS entity_name, recipient_name_std AS entity_name_std, recipient_uei AS entity_id,
    federal_action_obligation AS amount, 'Federal action obligation' AS amount_measure, NULL AS data_flag
  FROM {pq('usaspending_transactions')}
UNION ALL
SELECT 'OpenFEMA', project_key, record_date, calendar_year, calendar_month, fiscal_year,
    state, county, damage_category, incident_type, NULL, NULL, NULL, applicant_id,
    total_obligated, 'Total obligated', CASE WHEN record_date_basis = 'declaration_date' THEN 'date_from_declaration' END
  FROM {pq('fema_pa_projects')}
UNION ALL
SELECT 'FAC', a.report_id||'|'||coalesce(a.award_reference,'') ||'|'||coalesce(a.assistance_listing,''), r.record_date, r.calendar_year, r.calendar_month, r.fiscal_year,
    r.auditee_state, NULL, a.federal_agency, a.assistance_listing||' '||coalesce(a.program_name,''), NULL,
    r.auditee_name, r.auditee_name_std, r.auditee_uei,
    a.amount_expended, 'Federal award expended', CASE WHEN r.is_migration_record THEN 'migration_record' END
  FROM {pq('fac_awards')} a LEFT JOIN {pq('fac_reports')} r USING (report_id)
UNION ALL
SELECT 'SBA PPP', loan_number, record_date, calendar_year, calendar_month, fiscal_year,
    borrower_state, project_county, business_type, loan_status, naics_code, borrower_name, borrower_name_std, NULL,
    initial_approval_amount, 'Initial approval amount', NULL
  FROM {pq('sba_ppp_loans')}
UNION ALL
SELECT 'SEC', f.adsh||'|'||f.tag||'|'||f.version||'|'||coalesce(f.quarters_covered::VARCHAR,''), f.value_date, year(f.value_date), month(f.value_date), fy_of(f.value_date),
    s.business_state, NULL, f.tag, s.form, s.sic, s.company_name, s.company_name_std, s.cik::VARCHAR,
    f.value, 'Reported USD value (current period)', CASE WHEN f.quarters_covered IS NOT NULL THEN 'qtrs_'||f.quarters_covered::VARCHAR END
  FROM {pq('sec_facts')} f JOIN {pq('sec_filings')} s USING (adsh) WHERE f.is_current_period
""", "unified_records")

# ================================================================ DuckDB database
log("DuckDB")
import shutil
dbp = "/tmp/duck/analytics.duckdb"   # built on local disk, then copied (the project folder does not allow deletes, which DuckDB needs for its WAL)
for ext in ("", ".wal"):
    if os.path.exists(dbp + ext): os.remove(dbp + ext)
db = duckdb.connect(dbp)
names = ["unified_records", "usaspending_transactions", "usaspending_awards", "fema_pa_projects", "fac_reports", "fac_awards", "fac_findings",
         "sec_filings", "sec_tags", "sec_facts", "sba_ppp_loans"]
for nme in names:
    db.execute(f"CREATE TABLE {nme} AS SELECT * FROM read_parquet('{PQ}/{nme}.parquet')")
    log(f"table {nme}")
db.close()
shutil.copyfile(dbp, f"{OUT}/analytics.duckdb"); log("copied database to 02_Clean_Data")
open(f"{OUT}/_DONE_build", "w").write("done")
log("ALL DONE")
