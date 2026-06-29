import os
import pandas as pd
import numpy as np
import re
import datetime
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# File Input Environment Definitions
MASTER_INPUT = "Cayuse-Oracle award True Up - jth6-Copy.xlsx"
PRIME_ASST_FILE = "Assistance_PrimeAwardSummaries_2026-06-25_H19M12S13_1.csv"
PRIME_CONT_FILE = "Contracts_PrimeAwardSummaries_2026-06-25_H19M12S08_1.csv"
SUB_ASST_FILE = "Assistance_Subawards_2026-06-25_H19M12S41_1.csv"
SUB_CONT_FILE = "Contracts_Subawards_2026-06-25_H19M12S29_1.csv"

# Current data baseline extraction execution runtime anchor date
PIPELINE_RUN_DATE = pd.to_datetime("2026-06-26").date()

print("🚀 Launching Comprehensive Active & Expired Award Triage Engine (Version 6.1)...")

if not os.path.exists(MASTER_INPUT):
    print(f"❌ Error: Cannot find master workbook '{MASTER_INPUT}' in this folder.")
    exit()

def advanced_clean_id(text):
    if pd.isna(text): return ""
    text = str(text).upper().strip()
    text = re.sub(r'^(NIH|NSF|DOD|DOHA|ONR|DOE|USDA|DHHS)\s+', '', text)
    text = re.sub(r'[^A-Z0-9-]', '', text)
    text = re.sub(r'-[A-Z0-9]{1,2}$', '', text) 
    text = re.sub(r'[^A-Z0-9]', '', text)
    if re.match(r'^[A-Z]+\d{7}$', text): text = re.sub(r'^[A-Z]+', '', text)
    if re.match(r'^\d[A-Z]\d{2}', text): text = text[1:]
    return text

def tokenize_string(s):
    if pd.isna(s): return []
    tokens = re.findall(r'[A-Z0-9\-]{4,}', str(s).upper())
    noise = {'RICE', 'WEST', 'GRANT', 'FEDERAL', 'CONTR', 'CONTRACT', 'AMENDMENT', 'MODIFICATION', 'BLANK', 'NONE', 'FOUNDATION', 'UNIVERSITY'}
    return [t for t in tokens if t not in noise]

def parse_mixed_date(val):
    if pd.isna(val): return None
    val_str = str(val).strip()
    if val_str.replace('.0', '').isdigit():
        clean_int = int(val_str.replace('.0', ''))
        return pd.to_datetime(clean_int, unit='D', origin='1899-12-30').date()
    return pd.to_datetime(val_str, errors='coerce').date()

def save_as_structured_table(writer, dataframe, name_of_sheet):
    dataframe.to_excel(writer, sheet_name=name_of_sheet, index=False)
    worksheet = writer.sheets[name_of_sheet]
    total_rows = len(dataframe) + 1
    total_cols = len(dataframe.columns) if len(dataframe) > 0 else 1
    column_letter = get_column_letter(total_cols)
    cell_range = f"A1:{column_letter}{total_rows}"
    table_id = "".join(c for c in name_of_sheet if c.isalnum())
    excel_table = Table(displayName=table_id, ref=cell_range)
    excel_table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    worksheet.add_table(excel_table)

# Ingest Data Sources
print("📥 Ingesting local portfolio files and federal logs...")
master_df = pd.read_excel(MASTER_INPUT, sheet_name="Intermediate DATA")
gsum_df = pd.read_excel(MASTER_INPUT, sheet_name="Grant Summary")

gsum_df['clean_id'] = gsum_df['AWARD_NUMBER_SCRUB'].astype(str).str.upper().str.strip()
type_map = dict(zip(gsum_df['clean_id'], gsum_df['AWARD_TYPE_NAME']))
mech_map = dict(zip(gsum_df['clean_id'], gsum_df['FUNDING_MECHANISM']))
status_map = dict(zip(gsum_df['clean_id'], gsum_df['AWARD_STATUS']))
pi_map = dict(zip(gsum_df['clean_id'], gsum_df['PI_NAME']))
flow_through_map = dict(zip(gsum_df['clean_id'], gsum_df['FLOW_THROUGH_SPONSOR']))
header_limit_map = dict(zip(gsum_df['clean_id'], gsum_df['HARD_LIMIT_AMOUNT']))
project_budget_map = dict(zip(gsum_df['clean_id'], gsum_df['BUDGET_BRDND_COST']))
start_date_map = dict(zip(gsum_df['clean_id'], gsum_df['AWARD_START_DATE']))

pa_df = pd.read_csv(PRIME_ASST_FILE, low_memory=False) if os.path.exists(PRIME_ASST_FILE) else pd.DataFrame(columns=['award_id_fain'])
pc_df = pd.read_csv(PRIME_CONT_FILE, low_memory=False) if os.path.exists(PRIME_CONT_FILE) else pd.DataFrame(columns=['award_id_piid'])
sa_df = pd.read_csv(SUB_ASST_FILE, low_memory=False) if os.path.exists(SUB_ASST_FILE) else pd.DataFrame(columns=['prime_award_fain'])
sc_df = pd.read_csv(SUB_CONT_FILE, low_memory=False) if os.path.exists(SUB_CONT_FILE) else pd.DataFrame(columns=['prime_award_piid'])

pa_lookup = {advanced_clean_id(r['award_id_fain']): (r['total_obligated_amount'], r['total_funding_amount'], r['period_of_performance_current_end_date'], r['award_id_fain'], "Prime Assistance") for _, r in pa_df.iterrows() if pd.notna(r['award_id_fain'])}
pc_lookup = {advanced_clean_id(r['award_id_piid']): (r['total_obligated_amount'], r['potential_total_value_of_award'], r['period_of_performance_current_end_date'], r['award_id_piid'], "Prime Contract") for _, r in pc_df.iterrows() if pd.notna(r['award_id_piid'])}

sa_lookup = {advanced_clean_id(r['prime_award_fain']): (r['subaward_amount'], r['subaward_amount'], r['prime_award_period_of_performance_current_end_date'], r['prime_award_fain'], "Subaward Assistance") for _, r in sa_df.iterrows() if pd.notna(r['prime_award_fain'])}
sc_lookup = {advanced_clean_id(r['prime_award_piid']): (r['subaward_amount'], r['subaward_amount'], r['prime_award_period_of_performance_current_end_date'], r['prime_award_piid'], "Subaward Contract") for _, r in sc_df.iterrows() if pd.notna(r['prime_award_piid'])}

triaged_records = []

print("🔗 Cross-referencing accounts and building grouped tracking tables...")
for idx, row in master_df.iterrows():
    fed_id, fed_end_date, fed_obligated, fed_potential, fed_cat = np.nan, np.nan, 0.0, 0.0, "UNMATCHED"
    clean_award_number = str(row['Award Number']).replace('.0', '').upper().strip()
    
    source_status = status_map.get(clean_award_number, row['STATUS'] if 'STATUS' in master_df.columns else "Active")
    status_clean = str(source_status).upper().strip()
    
    # Exclude entirely completed records from tracking
    if status_clean == "CLOSED": continue  

    t_verdict, f_verdict = "IN_SYNC", "FINANCIALS_IN_SYNC"
    t_trace, f_trace = "STEADY_STATE: SYSTEM_IN_SYNC", "STEADY_STATE: SYSTEM_IN_SYNC"
    t_unit, f_unit = "N/A", "N/A"

    proj_cayuse_val = str(row['Proj CAYUSE']).strip() if ('Proj CAYUSE' in master_df.columns and pd.notna(row['Proj CAYUSE'])) else ""

    primary_id_key = advanced_clean_id(row['Award Number'])
    name_tokens = tokenize_string(row['Award Name'])
    candidates = [primary_id_key] if primary_id_key else []
    candidates.extend([advanced_clean_id(t) for t in name_tokens if t])
    candidates = [c for c in candidates if c]
    
    for k in candidates:
        if k in pa_lookup: fed_obligated, fed_potential, fed_end_date, fed_id, fed_cat = pa_lookup[k]; break
        elif k in pc_lookup: fed_obligated, fed_potential, fed_end_date, fed_id, fed_cat = pc_lookup[k]; break
        elif k in sa_lookup: fed_obligated, fed_potential, fed_end_date, fed_id, fed_cat = sa_lookup[k]; break
        elif k in sc_lookup: fed_obligated, fed_potential, fed_end_date, fed_id, fed_cat = sc_lookup[k]; break

    o_end = parse_mixed_date(row['OR END DATE'])
    c_end = parse_mixed_date(row['CAY END DATE'])
    f_end = parse_mixed_date(fed_end_date)
    
    c_funding = float(row['CAYUSE BUDGET']) if pd.notna(row['CAYUSE BUDGET']) else 0.0
    o_header_limit = float(header_limit_map.get(clean_award_number, c_funding))
    o_project_budget = float(project_budget_map.get(clean_award_number, float(row['OR Allocated Funding']) if pd.notna(row['OR Allocated Funding']) else 0.0))

    source_type = type_map.get(clean_award_number, row['Award Type'] if 'Award Type' in master_df.columns else "Federal")
    source_mechanism = str(mech_map.get(clean_award_number, "Grant")).strip()
    source_flow = flow_through_map.get(clean_award_number, np.nan)

    type_str_upper = str(source_type).upper()
    is_direct_sponsor_fed = "FED" in type_str_upper or "FEDERAL" in type_str_upper
    is_flow_through_fed = pd.notna(source_flow) and any(w in str(source_flow).upper() for w in ['NIH', 'NSF', 'FED', 'FEDERAL', 'HEALTH', 'SCIENCE', 'DEFENSE', 'DEPARTMENT', 'AGENCY', 'FOUNDATION'])
    is_federal_funding = is_direct_sponsor_fed or is_flow_through_fed or (fed_cat != "UNMATCHED")

    days_cayuse_minus_fed = (c_end - f_end).days if (c_end and f_end) else np.nan
    delta_oracle_limit_vs_cayuse = o_header_limit - c_funding

    target_end_date = o_end
    target_hard_limit = o_header_limit
    target_project_budget = o_project_budget

    # ----------------==========================
    # TIMELINE OVERWRITE EVALUATION ENGINE
    # ----------------==========================
    if is_federal_funding:
        if f_end and c_end and f_end > c_end and pd.isna(source_flow):
            t_verdict = "OSP_AMENDMENT_REVIEW"
            t_trace = "RULE_1: FED_END > CAY_END (OSP MUST CONFIRM SPONSOR NOTICE AND UPDATE CAYUSE)"
            t_unit = "OSP"
            target_end_date = c_end 
        elif c_end and o_end and c_end < o_end:
            t_verdict = "UNAUTHORIZED_EXTENSION"
            t_trace = "RULE_2: ORACLE_END > CAYUSE_END (CRITICAL RECONCILIATION: AUDIT RCA EMAIL RECORDS AND NOTIFY OSP)"
            t_unit = "BOTH"
            target_end_date = c_end 
        elif c_end and f_end and c_end > f_end:
            # 🧠 FIX: Opened the 45-day window completely for any record marked as EXPIRED to evaluate historical drift
            is_within_45_day_active_window = f_end >= (PIPELINE_RUN_DATE - datetime.timedelta(days=45)) or "EXPIRED" in status_clean
            if days_cayuse_minus_fed <= 365 and is_within_45_day_active_window:
                t_verdict = "EXPANDED_AUTHORITIES_VALID"
                t_trace = f"RULE_3A: CAY_END > FED_END BY {days_cayuse_minus_fed} DAYS (HARMONIZE ORACLE TO MATCH AUTHORIZED CAYUSE DATE)"
                t_unit = "RCA"
                target_end_date = c_end
            else:
                t_verdict = "EXPANDED_AUTHORITIES_BREACH"
                t_trace = f"RULE_3B: CAY_END > FED_END BY {days_cayuse_minus_fed} DAYS OR BREACHES EXPIRATION WINDOW STANDARDS"
                t_unit = "BOTH"
                target_end_date = f_end + datetime.timedelta(days=365) 
    else:
        if o_end != c_end:
            t_verdict = "NON_FED_SPONSOR_DRIFT"
            t_trace = "RULE_5_6: NON_FEDERAL TIMELINE DRIFT (VERIFY TRUE SIGNED SPONSOR AGREEMENT DATE AND REALIGN)"
            t_unit = "BOTH"
            target_end_date = c_end if c_end else o_end

    # ----------------==========================
    # FINANCIAL OVERWRITE EVALUATION ENGINE
    # ----------------==========================
    if "CONTRACT" in source_mechanism.upper() or fed_cat in ["Prime Contract", "Subaward Contract"]:
        if abs(o_project_budget - fed_potential) > 1.00 or abs(o_header_limit - fed_potential) > 1.00:
            f_verdict = "CONTRACT_CEILING_MISMATCH"
            f_trace = "RULE_8: CONTRACT DRIFT FROM VALUE CEILING (STOP AUTO ADJUSTMENTS: TRANSMIT TO CONTRACT OFFICER)"
            f_unit = "OSP"
    else:
        if delta_oracle_limit_vs_cayuse < -1.00:
            f_verdict = "ORACLE_UNDER_ALLOCATED"
            f_trace = f"RULE_7_9_A: CAYUSE > ORACLE BY ${abs(delta_oracle_limit_vs_cayuse):,.2f} (SAFE RE-ALLOCATION ADJUSTMENT)"
            f_unit = "RCA"
            target_hard_limit = c_funding
        elif delta_oracle_limit_vs_cayuse > 1.00:
            f_verdict = "CRITICAL_CONTROL_FAILURE"
            f_trace = f"RULE_7_9_B: ORACLE > CAYUSE BY ${delta_oracle_limit_vs_cayuse:,.2f} (STOP: RCA INVESTIGATION REQUIRED BEFORE CHANGE)"
            f_unit = "RCA"
            target_hard_limit = o_header_limit 
            
        if fed_cat != "UNMATCHED" and "GRANT" in source_mechanism.upper():
            if o_project_budget < fed_obligated:
                cay_in_sync_with_fed = abs(c_funding - fed_obligated) <= 1.00
                if cay_in_sync_with_fed:
                    f_verdict = "ORACLE_UNDER_FUNDED"
                    f_trace = "RULE_11A: ORACLE < FED_OBLIGATED AND CAYUSE IN SYNC (SAFE AUTOMATED RELATIVE INCREASE)"
                    f_unit = "RCA"
                    target_project_budget = fed_obligated
                else:
                    f_verdict = "OSP_INVESTIGATE_WITH_SPONSOR"
                    f_trace = "RULE_11A_HOLD: ORACLE < FED_OBLIGATED BUT CAYUSE MISALIGNED (OSP MUST CONFIRM DATA WITH SPONSOR)"
                    f_unit = "OSP"
                    target_project_budget = o_project_budget 
            elif o_project_budget > fed_obligated:
                f_verdict = "ORACLE_OVER_FUNDED_RISK"
                f_trace = f"RULE_11B: ORACLE > FED_OBLIGATED BY ${abs(o_project_budget - fed_obligated):,.2f} (STOP: COORDINATED PI AND OSP REMEDIATION REQUIRED)"
                f_unit = "BOTH"
                target_project_budget = o_project_budget 

    # Resolve Unified Responsible Unit Assignment Priority Flag
    if "BOTH" in [t_unit, f_unit]: responsible_unit = "BOTH"
    elif t_unit != "N/A" and f_unit != "N/A" and t_unit != f_unit: responsible_unit = "BOTH"
    elif t_unit != "N/A": responsible_unit = t_unit
    elif f_unit != "N/A": responsible_unit = f_unit
    else: responsible_unit = "N/A"

    if t_verdict in ["UNAUTHORIZED_EXTENSION", "EXPANDED_AUTHORITIES_BREACH"] or f_verdict in ["CRITICAL_CONTROL_FAILURE", "CONTRACT_CEILING_MISMATCH", "ORACLE_OVER_FUNDED_RISK", "OSP_INVESTIGATE_WITH_SPONSOR"]:
        route_class = "MANUAL_OSP_RCA_RECONCILIATION_REQUIRED"
    elif t_verdict in ["OSP_AMENDMENT_REVIEW", "NON_FED_SPONSOR_DRIFT"] or f_verdict in ["ORACLE_UNDER_ALLOCATED", "ORACLE_UNDER_FUNDED"]:
        route_class = "AUTOMATED_WORKFLOW_CHANGE_READY"
    else:
        route_class = "SYSTEMS_FULLY_ALIGNED"

    h_increase = target_hard_limit - o_header_limit if (target_hard_limit > o_header_limit) else 0.0
    h_decrease = o_header_limit - target_hard_limit if (target_hard_limit < o_header_limit) else 0.0
    p_increase = target_project_budget - o_project_budget if (target_project_budget > o_project_budget) else 0.0
    p_decrease = o_project_budget - target_project_budget if (target_project_budget < o_project_budget) else 0.0

    triaged_records.append({
        # 1. Identifiers & Route Metadata
        'AWARD_NUMBER': clean_award_number,
        'CAYUSE_PROJECT_NUMBER': proj_cayuse_val if (proj_cayuse_val.upper() != "NOT FOUND" and proj_cayuse_val.upper() != "NAN") else "N/A",
        'FEDERAL_AWARD_IDENTIFIER': fed_id if fed_cat != "UNMATCHED" else "N/A",
        'SPONSOR': row['Funding Source Name'],
        'AWARD_STATUS': source_status,
        'RECORD_RECONCILIATION_ROUTE': route_class,
        'RESPONSIBLE_UNIT': responsible_unit,
        'TIMELINE_RULE_TRACE': t_trace,
        'FINANCIAL_RULE_TRACE': f_trace,
        
        # 2. Calculated Intermediary Metrics
        'DAYS_CAYUSE_MINUS_FED': days_cayuse_minus_fed,
        'DELTA_ORACLE_LIMIT_VS_CAYUSE': delta_oracle_limit_vs_cayuse,
        
        # 3. Timeline Group (Absolute Dates)
        'EXISTING_ORACLE_END_DATE': o_end,
        'EXISTING_CAYUSE_END_DATE': c_end,
        'FEDERAL_END_DATE': f_end,
        'PROPOSED_END_DATE': target_end_date,
        
        # 4. Financial Group (Relative Adjustments)
        'EXISTING_HARD_LIMIT_AMOUNT': o_header_limit,
        'EXISTING_BUDGET_BRDND_COST': o_project_budget,
        'EXISTING_CAYUSE_BUDGET': c_funding,
        'FEDERAL_OBLIGATED_AMOUNT': fed_obligated if fed_cat != "UNMATCHED" else 0.0,
        'FEDERAL_POTENTIAL_AMOUNT': fed_potential if fed_cat != "UNMATCHED" else 0.0,
        'PROPOSED_HARD_LIMIT_AMOUNT': target_hard_limit,
        'PROPOSED_BUDGET_BRDND_COST': target_project_budget,
        'RELATIVE_HARD_LIMIT_INCREASE': h_increase if h_increase > 1.00 else np.nan,
        'RELATIVE_HARD_LIMIT_DECREASE': h_decrease if h_decrease > 1.00 else np.nan,
        'RELATIVE_PROJECT_BUDGET_INCREASE': p_increase if p_increase > 1.00 else np.nan,
        'RELATIVE_PROJECT_BUDGET_DECREASE': p_decrease if p_decrease > 1.00 else np.nan
    })

full_df = pd.DataFrame(triaged_records)

# Policy Manifest Reference Glossary Tab
manifest_matrix = [
    ["RULE_1", "OSP_AMENDMENT_REVIEW", "USAspending end date is after Cayuse end date. Tracks potential pre-award processing log backlog.", "OSP must confirm the formal funding amendment document from the sponsor and update Cayuse. RCA holds Oracle dates until Cayuse updates."],
    ["RULE_2", "UNAUTHORIZED_EXTENSION", "Federal award where the active Oracle end date exceeds the Cayuse baseline date without a matching control record.", "Joint Audit Required. RCA must audit email/system authorization records; OSP verifies whether an approved amendment bypassed sync rules."],
    ["RULE_3A", "EXPANDED_AUTHORITIES_VALID", "Cayuse exceeds federal end date by <= 365 days while the federal record falls within the active 45-day threshold window.", "Valid standard 12-month extension profile. RCA processes automated absolute update to bring Oracle into perfect alignment with Cayuse."],
    ["RULE_3B", "EXPANDED_AUTHORITIES_BREACH", "Cayuse exceeds federal limits by > 365 days, or the baseline federal date is older than today minus 45 days.", "Compliance Hold. Apply automated absolute safety cap forcing the Oracle project deadline to stay no more than 365 days past the federal record date."],
    ["RULE_5_6", "NON_FED_SPONSOR_DRIFT", "Non-federal project timeline data discrepancy between post-award setup records and pre-award logs.", "Both Units Review. Conduct manual audit to pull down and verify the true sponsor-agreed deadline explicitly stated on the signed contract agreement."],
    ["RULE_7_9_A", "ORACLE_UNDER_ALLOCATED", "Pre-award Cayuse tracked budget thresholds exceed the active Oracle header hard limit amount.", "Safe re-allocation push. RCA executes relative increase journal entry to release the unallocated funding baseline into Oracle ERP."],
    ["RULE_7_9_B", "CRITICAL_CONTROL_FAILURE", "Active Oracle header hard limits exceed total pre-award tracked funding parameters inside Cayuse.", "Halt Automation. RCA initiates manual operational lookup to investigate why funding thresholds were raised outside Cayuse checkpoints."],
    ["RULE_8", "CONTRACT_CEILING_MISMATCH", "Federal Contract file fields deviate from total potential value award limits logged on the public federal source ledger.", "Manual Procurement Review. Suspend all automated budgetary delta increases and hand the record directly to the tracking contract officer."],
    ["RULE_11A", "ORACLE_UNDER_FUNDED", "Oracle project distributed budget falls short of true legal federal obligated funding totals.", "Conditional Action. If Cayuse perfectly matches federal records, RCA pushes relative increase delta. If Cayuse deviates, OSP investigates data with sponsor."],
    ["RULE_11B", "ORACLE_OVER_FUNDED_RISK", "Oracle distributed budget exceeds total legal federal obligation balances on public direct ledgers.", "Critical compliance risk hold. Block all automated scripts. RCA, OSP, and the PI must manually clear the spending overdraft or secure the missing mod."]
]
manifest_df = pd.DataFrame(manifest_matrix, columns=["RULE_ID", "TRIAGE_CONDITION_VERDICT", "SYSTEMIC_TRIGGER_DEFINITION", "TARGETED_MANAGEMENT_ACTION_DIRECTIVE"])

# Export Standardized Tables
with pd.ExcelWriter("Ecosystem_True_Up_Audit.xlsx", engine="openpyxl") as writer:
    
    # Tab 1: Full Master
    exec_headers = [
        'AWARD_NUMBER', 'CAYUSE_PROJECT_NUMBER', 'FEDERAL_AWARD_IDENTIFIER', 'SPONSOR', 'AWARD_STATUS', 'RECORD_RECONCILIATION_ROUTE', 'RESPONSIBLE_UNIT',
        'TIMELINE_RULE_TRACE', 'EXISTING_ORACLE_END_DATE', 'EXISTING_CAYUSE_END_DATE', 'FEDERAL_END_DATE', 'PROPOSED_END_DATE',
        'FINANCIAL_RULE_TRACE', 'EXISTING_HARD_LIMIT_AMOUNT', 'EXISTING_BUDGET_BRDND_COST', 'EXISTING_CAYUSE_BUDGET', 
        'FEDERAL_OBLIGATED_AMOUNT', 'FEDERAL_POTENTIAL_AMOUNT', 'PROPOSED_HARD_LIMIT_AMOUNT', 'PROPOSED_BUDGET_BRDND_COST'
    ]
    save_as_structured_table(writer, full_df[exec_headers], "Full Triage Master")
    
    # Tab 2: Timeline Focus
    time_headers = [
        'AWARD_NUMBER', 'CAYUSE_PROJECT_NUMBER', 'FEDERAL_AWARD_IDENTIFIER', 'SPONSOR', 'RECORD_RECONCILIATION_ROUTE', 'RESPONSIBLE_UNIT', 'TIMELINE_RULE_TRACE', 'DAYS_CAYUSE_MINUS_FED',
        'EXISTING_ORACLE_END_DATE', 'EXISTING_CAYUSE_END_DATE', 'FEDERAL_END_DATE', 'PROPOSED_END_DATE'
    ]
    time_mask = full_df['TIMELINE_RULE_TRACE'].str.contains('RULE_')
    save_as_structured_table(writer, full_df[time_mask][time_headers], "Chronological Triage")
    
    # Tab 3: Financial Focus
    fin_headers = [
        'AWARD_NUMBER', 'CAYUSE_PROJECT_NUMBER', 'FEDERAL_AWARD_IDENTIFIER', 'SPONSOR', 'RECORD_RECONCILIATION_ROUTE', 'RESPONSIBLE_UNIT', 'FINANCIAL_RULE_TRACE', 'DELTA_ORACLE_LIMIT_VS_CAYUSE',
        'EXISTING_HARD_LIMIT_AMOUNT', 'EXISTING_BUDGET_BRDND_COST', 'EXISTING_CAYUSE_BUDGET', 'FEDERAL_OBLIGATED_AMOUNT', 'FEDERAL_POTENTIAL_AMOUNT',
        'PROPOSED_HARD_LIMIT_AMOUNT', 'PROPOSED_BUDGET_BRDND_COST',
        'RELATIVE_HARD_LIMIT_INCREASE', 'RELATIVE_HARD_LIMIT_DECREASE', 'RELATIVE_PROJECT_BUDGET_INCREASE', 'RELATIVE_PROJECT_BUDGET_DECREASE'
    ]
    fin_mask = full_df['FINANCIAL_RULE_TRACE'].str.contains('RULE_')
    save_as_structured_table(writer, full_df[fin_mask][fin_headers], "Financial Triage")
    
    # Tab 4: Critical Audit Hub
    audit_headers = [
        'AWARD_NUMBER', 'CAYUSE_PROJECT_NUMBER', 'FEDERAL_AWARD_IDENTIFIER', 'SPONSOR', 'RECORD_RECONCILIATION_ROUTE', 'RESPONSIBLE_UNIT',
        'EXISTING_ORACLE_END_DATE', 'EXISTING_CAYUSE_END_DATE', 'FEDERAL_END_DATE', 'PROPOSED_END_DATE',
        'EXISTING_HARD_LIMIT_AMOUNT', 'EXISTING_BUDGET_BRDND_COST', 'EXISTING_CAYUSE_BUDGET', 'FEDERAL_OBLIGATED_AMOUNT', 'FEDERAL_POTENTIAL_AMOUNT',
        'PROPOSED_HARD_LIMIT_AMOUNT', 'PROPOSED_BUDGET_BRDND_COST'
    ]
    audit_mask = full_df['RECORD_RECONCILIATION_ROUTE'].str.contains('RECONCILIATION')
    save_as_structured_table(writer, full_df[audit_mask][audit_headers], "Audit and Tracking Hub")
    
    # Tab 5: Reference Dictionary
    save_as_structured_table(writer, manifest_df, "Rule_Definitions_Registry")

# Print automated metrics validation run outputs
print("\n📊 FINAL COMPREHENSIVE ACTIVE & EXPIRED METRICS:")
print(f"Total Combined Portfolio Items Processed: {len(full_df)} awards")
print(full_df['RESPONSIBLE_UNIT'].value_counts())