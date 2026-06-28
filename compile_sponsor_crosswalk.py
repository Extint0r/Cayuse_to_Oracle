import os
import re
import sqlite3
import pandas as pd
import numpy as np
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# ==========================================
# SYSTEM AUTHENTICATION & ENVIRONMENT ANCHORS
# ==========================================
MASTER_INPUT = "Cayuse-Oracle award True Up - jth6-Copy.xlsx"
OUTPUT_FILE = "Master_Sponsor_Crosswalk_Index.xlsx"
DB_FILE = "rice_sponsor_mdm.db"

print("📊 Launching Deep-Hierarchy 1-to-1 Master Sponsor Identity Engine...")

if not os.path.exists(MASTER_INPUT):
    print(f"❌ Error: Cannot find master workbook '{MASTER_INPUT}' in this folder.")
    exit()

# ==========================================
# 1. INITIALIZE RELATIONAL SYSTEM STORAGE
# ==========================================
conn = sqlite3.connect(DB_FILE)
cursor = conn.cursor()

cursor.execute("DROP TABLE IF EXISTS staging_sponsor_inputs;")
cursor.execute("""
    CREATE TABLE staging_sponsor_inputs (
        string_id INTEGER PRIMARY KEY AUTOINCREMENT,
        oracle_sponsor_string TEXT UNIQUE,
        cayuse_sponsor_string TEXT,
        usaspending_top_tier TEXT DEFAULT 'N/A',
        usaspending_sub_tier TEXT DEFAULT 'N/A',
        sam_gov_certified_name TEXT,
        sponsor_taxonomy_class TEXT
    );
""")
conn.commit()

# ==========================================
# 2. INGEST SPREADSHEET LAYERS & META MAPS
# ==========================================
print("📥 Ingesting manual pre-award and post-award spreadsheet data layers...")
master_df = pd.read_excel(MASTER_INPUT, sheet_name="Intermediate DATA")
gsum_df = pd.read_excel(MASTER_INPUT, sheet_name="Grant Summary")

gsum_df['clean_id'] = gsum_df['AWARD_NUMBER_SCRUB'].astype(str).str.upper().str.strip()
master_df['clean_id'] = master_df['Award Number'].astype(str).str.replace('.0', '', regex=False).str.upper().str.strip()

# Build direct maps to reference structural project variables across tabs
term_name_map = dict(zip(gsum_df['clean_id'], gsum_df['AWARD_TERM_NAME']))
flow_through_map = dict(zip(gsum_df['clean_id'], gsum_df['FLOW_THROUGH_SPONSOR']))

master_df['AWARD_TERM_NAME'] = master_df['clean_id'].map(term_name_map).fillna('')
master_df['FLOW_THROUGH_SPONSOR'] = master_df['clean_id'].map(flow_through_map).fillna('')

raw_oracle_names = master_df['Funding Source Name'].dropna().unique()
raw_cayuse_names = gsum_df['FLOW_THROUGH_SPONSOR'].dropna().unique()

all_unique_strings = set()
for name in raw_oracle_names: all_unique_strings.add(str(name).strip())
for name in raw_cayuse_names: all_unique_strings.add(str(name).strip())

# ==========================================
# 3. COMPREHENSIVE REGISTRY REFERENCE MATRIX
# ==========================================
def resolve_authoritative_hierarchy(org_string, is_federal_context):
    name_up = str(org_string).upper().strip()
    
    # RULE 1: Filter out state and municipal public bodies from Federal tracking metrics
    if any(k in name_up for k in ["TEXAS", "HOUSTON", "HARRIS COUNTY", "CITY OF", "MUNICIPAL", "STATE OF"]):
        if not any(fed in name_up for fed in ["NIH", "NSF", "NASA", "DEPARTMENT OF DEFENSE", "ARPA-H"]):
            return "N/A", "N/A", org_string.upper(), "NON_FEDERAL_RECIPIENT_ENTITY"
            
    # RULE 2: Force explicit non-federal tracking if the system term name is not Federal
    if not is_federal_context:
        return "N/A", "N/A", org_string.upper(), "NON_FEDERAL_RECIPIENT_ENTITY"

    # RULE 3: Filter out commercial contractors, universities, and non-profits handling sub-awards
    # If the name string contains corporate entity tags and lacks an explicit federal command keyword,
    # it is a pass-through entity. Force agency columns to N/A.
    has_corporate_suffix = any(sfx in name_up for sfx in [" INC", " LLC", " CORP", " CO", " LTD", "UNIVERSITY", "FOUNDATION", "COLLEGE", "INSTITUTE OF TECHNOLOGY"])
    has_federal_keyword = any(fed in name_up for fed in ["DEPARTMENT", "AGENCY", "COMMISSION", "COMMAND", "ADMINISTRATION", "BUREAU", "OFFICE", "LABORATORY", "NIH", "NSF", "NASA"])
    
    if has_corporate_suffix and not has_federal_keyword:
        return "N/A", "N/A", org_string.upper(), "NON_FEDERAL_RECIPIENT_ENTITY"

    # ----------------------------------------------------
    # TRUE FEDERAL HIERARCHY DIRECTORY ROUTING
    # ----------------------------------------------------
    # 1. Health and Human Services / NIH / ARPA-H / CDC
    if any(k in name_up for k in ["HEALTH AND HUMAN SERVICES", "NATIONAL INSTITUTES OF HEALTH", "NIH", "DISEASE CONTROL", "CDC", "ARPA-H", "ADVANCED RESEARCH PROJECTS AGENCY FOR HEALTH"]):
        sub_tier = "NATIONAL INSTITUTES OF HEALTH (NIH)"
        if "ARPA-H" in name_up or "ADVANCED RESEARCH" in name_up:
            sub_tier = "ADVANCED RESEARCH PROJECTS AGENCY FOR HEALTH (ARPA-H)"
        elif "DISEASE CONTROL" in name_up or "CDC" in name_up:
            sub_tier = "CENTERS FOR DISEASE CONTROL AND PREVENTION (CDC)"
        return "DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)", sub_tier, sub_tier, "FEDERAL_FUNDING_AGENCY"
    
    # 2. National Science Foundation
    if "NATIONAL SCIENCE FOUNDATION" in name_up or name_up == "NSF":
        return "NATIONAL SCIENCE FOUNDATION (NSF)", "NATIONAL SCIENCE FOUNDATION (NSF)", "NATIONAL SCIENCE FOUNDATION (NSF)", "FEDERAL_FUNDING_AGENCY"
    
    # 3. NASA Operating Centers
    if "NATIONAL AERONAUTICS AND SPACE" in name_up or "NASA" in name_up or "SPACE TELESCOPE" in name_up:
        sub_tier = "NASA HEADQUARTERS"
        if "GODDARD" in name_up: sub_tier = "NASA GODDARD SPACE FLIGHT CENTER"
        elif "LANGLEY" in name_up: sub_tier = "NASA LANGLEY RESEARCH CENTER"
        elif "JOHNSON" in name_up or "JSC" in name_up: sub_tier = "NASA JOHNSON SPACE CENTER (JSC)"
        elif "JET PROPULSION" in name_up or "JPL" in name_up: sub_tier = "JET PROPULSION LABORATORY (JPL)"
        return "NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)", sub_tier, sub_tier, "FEDERAL_FUNDING_AGENCY"
    
    # 4. Department of Defense Commands (ARO, ARL, ARI, ONR, AFOSR, DARPA, DTRA)
    if any(k in name_up for k in ["DEFENSE", "DOD", "ARMY", "NAVY", "AIR FORCE", "ONR", "DARPA", "DTRA", "THREAT REDUCTION", "MILITARY"]):
        sub = "DEPARTMENT OF DEFENSE (DOD)"
        if "ARMY RESEARCH LABORATORY" in name_up or "ARL" in name_up: sub = "US ARMY RESEARCH LABORATORY (ARL)"
        elif "ARMY RESEARCH OFFICE" in name_up or "ARO" in name_up: sub = "US ARMY RESEARCH OFFICE (ARO)"
        elif "ARMY RESEARCH INSTITUTE" in name_up or "ARI" in name_up: sub = "US ARMY RESEARCH INSTITUTE (ARI)"
        elif "MEDICAL RESEARCH" in name_up or "USAMRAA" in name_up: sub = "US ARMY MEDICAL RESEARCH ACQUISITION ACTIVITY (USAMRAA)"
        elif "CORPS OF ENGINEERS" in name_up: sub = "US ARMY CORPS OF ENGINEERS"
        elif "CONTRACTING COMMAND" in name_up: sub = "US ARMY CONTRACTING COMMAND"
        elif "OFFICE OF NAVAL RESEARCH" in name_up or "ONR" in name_up: sub = "OFFICE OF NAVAL RESEARCH (ONR)"
        elif "NAVAL RESEARCH LABORATORY" in name_up: sub = "NAVAL RESEARCH LABORATORY (NRL)"
        elif "AIR FORCE OFFICE OF SCIENTIFIC" in name_up or "AFOSR" in name_up: sub = "AIR FORCE OFFICE OF SCIENTIFIC RESEARCH (AFOSR)"
        elif "AIR FORCE RESEARCH LABORATORY" in name_up or "AFRL" in name_up: sub = "AIR FORCE RESEARCH LABORATORY (AFRL)"
        elif "DARPA" in name_up or "DEFENSE ADVANCED RESEARCH" in name_up: sub = "DEFENSE ADVANCED RESEARCH PROJECTS AGENCY (DARPA)"
        elif "THREAT REDUCTION" in name_up or "DTRA" in name_up: sub = "DEFENSE THREAT REDUCTION AGENCY (DTRA)"
        elif "STRATEGIC ENVIRONMENTAL RESEARCH" in name_up or "SERDP" in name_up: sub = "STRATEGIC ENVIRONMENTAL RESEARCH & DEVELOPMENT PROGRAM"
        elif "MARYLAND PROCUREMENT" in name_up or "MPO" in name_up: sub = "MARYLAND PROCUREMENT OFFICE (MPO)"
        elif "ARMY" in name_up: sub = "DEPARTMENT OF THE ARMY"
        elif "NAVY" in name_up: sub = "DEPARTMENT OF THE NAVY"
        elif "AIR FORCE" in name_up: sub = "DEPARTMENT OF THE AIR FORCE"
        return "DEPARTMENT OF DEFENSE (DOD)", sub, sub, "FEDERAL_FUNDING_AGENCY"
        
    # 5. Department of Energy / National Laboratories (FFRDCs)
    if any(k in name_up for k in ["ENERGY", "DOE", "OAK RIDGE", "SANDIA", "ARGONNE", "LAWRENCE LIVERMORE", "LOS ALAMOS", "BROOKHAVEN", "FERMI", "SLAC", "BATTELLE"]):
        sub_tier = "DEPARTMENT OF ENERGY (DOE)"
        if "OAK RIDGE" in name_up: sub_tier = "OAK RIDGE NATIONAL LABORATORY"
        elif "SANDIA" in name_up: sub_tier = "SANDIA NATIONAL LABORATORIES"
        elif "ARGONNE" in name_up: sub_tier = "ARGONNE NATIONAL LABORATORY"
        elif "LAWRENCE LIVERMORE" in name_up: sub_tier = "LAWRENCE LIVERMORE NATIONAL LABORATORY"
        elif "LOS ALAMOS" in name_up: sub_tier = "LOS ALAMOS NATIONAL LABORATORY"
        elif "BROOKHAVEN" in name_up: sub_tier = "BROOKHAVEN NATIONAL LABORATORY"
        elif "FERMI" in name_up: sub_tier = "FERMI NATIONAL ACCELERATOR LABORATORY"
        elif "SLAC" in name_up: sub_tier = "SLAC NATIONAL ACCELERATOR LABORATORY"
        elif "RENEWABLE ENERGY" in name_up or "NREL" in name_up: sub_tier = "NATIONAL RENEWABLE ENERGY LABORATORY (NREL)"
        elif "OFFICE OF SCIENCE" in name_up: sub_tier = "DOE OFFICE OF SCIENCE"
        return "DEPARTMENT OF ENERGY (DOE)", sub_tier, sub_tier, "FEDERAL_FUNDING_AGENCY"
    
    # 6. Other Explicit Federal Departments
    if "DEPARTMENT OF EDUCATION" in name_up:
        return "DEPARTMENT OF EDUCATION (ED)", "DEPARTMENT OF EDUCATION (ED)", "DEPARTMENT OF EDUCATION (ED)", "FEDERAL_FUNDING_AGENCY"
    if "DEPARTMENT OF COMMERCE" in name_up or "NIST" in name_up or "NOAA" in name_up or "MBDA" in name_up or "CENSUS" in name_up or "OCEANIC ATMOSPHERIC" in name_up:
        sub_tier = "DEPARTMENT OF COMMERCE (DOC)"
        if "NIST" in name_up or "STANDARDS AND TECHNOLOGY" in name_up: sub_tier = "NATIONAL INSTITUTE OF STANDARDS AND TECHNOLOGY (NIST)"
        elif "NOAA" in name_up or "OCEANIC" in name_up: sub_tier = "NATIONAL OCEANIC AND ATMOSPHERIC ADMINISTRATION (NOAA)"
        elif "MBDA" in name_up or "MINORITY BUSINESS" in name_up: sub_tier = "MINORITY BUSINESS DEVELOPMENT AGENCY (MBDA)"
        elif "CENSUS" in name_up: sub_tier = "BUREAU OF THE CENSUS"
        return "DEPARTMENT OF COMMERCE (DOC)", sub_tier, sub_tier, "FEDERAL_FUNDING_AGENCY"
    if "DEPARTMENT OF TRANSPORTATION" in name_up or "FAA" in name_up or "AVIATION" in name_up or "ECONOMIC DEVELOPMENT" in name_up:
        sub_tier = "DEPARTMENT OF TRANSPORTATION (DOT)"
        if "FAA" in name_up or "AVIATION" in name_up: sub_tier = "FEDERAL AVIATION ADMINISTRATION (FAA)"
        elif "ECONOMIC DEVELOPMENT" in name_up: return "ECONOMIC DEVELOPMENT ADMINISTRATION (EDA)", "ECONOMIC DEVELOPMENT ADMINISTRATION (EDA)", "ECONOMIC DEVELOPMENT ADMINISTRATION (EDA)", "FEDERAL_FUNDING_AGENCY"
        return "DEPARTMENT OF TRANSPORTATION (DOT)", sub_tier, sub_tier, "FEDERAL_FUNDING_AGENCY"
    if "USAID" in name_up or "INTERNATIONAL DEVELOPMENT" in name_up:
        return "UNITED STATES AGENCY FOR INTERNATIONAL DEVELOPMENT (USAID)", "UNITED STATES AGENCY FOR INTERNATIONAL DEVELOPMENT (USAID)", "UNITED STATES AGENCY FOR INTERNATIONAL DEVELOPMENT (USAID)", "FEDERAL_FUNDING_AGENCY"
    if "DEPARTMENT OF STATE" in name_up:
        return "DEPARTMENT OF STATE (DOS)", "DEPARTMENT OF STATE (DOS)", "DEPARTMENT OF STATE (DOS)", "FEDERAL_FUNDING_AGENCY"
    if "DEPARTMENT OF INTERIOR" in name_up or "GEOLOGICAL SURVEY" in name_up or "USGS" in name_up:
        return "DEPARTMENT OF INTERIOR (DOI)", "DEPARTMENT OF INTERIOR (DOI)", "DEPARTMENT OF INTERIOR (DOI)", "FEDERAL_FUNDING_AGENCY"
    if "ENVIRONMENTAL PROTECTION" in name_up or "EPA" in name_up:
        return "ENVIRONMENTAL PROTECTION AGENCY (EPA)", "ENVIRONMENTAL PROTECTION AGENCY (EPA)", "ENVIRONMENTAL PROTECTION AGENCY (EPA)", "FEDERAL_FUNDING_AGENCY"
    if "ENDOWMENT FOR THE" in name_up:
        return "NATIONAL ENDOWMENT FOR THE HUMANITIES/ARTS", "NATIONAL ENDOWMENT FOR THE HUMANITIES/ARTS", "NATIONAL ENDOWMENT FOR THE HUMANITIES/ARTS", "FEDERAL_FUNDING_AGENCY"
    if "AGRICULTURE" in name_up or "USDA" in name_up:
        return "DEPARTMENT OF AGRICULTURE (USDA)", "DEPARTMENT OF AGRICULTURE (USDA)", "DEPARTMENT OF AGRICULTURE (USDA)", "FEDERAL_FUNDING_AGENCY"

    # Dynamic regex fallback for any other unmapped federal department structures
    match_dept = re.search(r'DEPARTMENT OF\s+([A-Z\s&]+)', name_up)
    if match_dept:
        extracted_dept = f"DEPARTMENT OF {match_dept.group(1).strip()}"
        extracted_dept = re.sub(r'\s+(LLC|INC|CORP|CO|LTD|NON\s+LOC|1306\d+|NIST).*$', '', extracted_dept).strip()
        return extracted_dept, "UNMAPPED OPERATIONAL COMPONENT", org_string.upper(), "FEDERAL_FUNDING_AGENCY"

    # Secondary gate: check if any explicit government tokens exist in the text string
    if has_federal_keyword and "FOUNDATION" not in name_up:
        return "FEDERAL_DEPARTMENT", "UNMAPPED_FEDERAL_AGENCY", org_string.upper(), "FEDERAL_FUNDING_AGENCY"
    
    # Final catch-all for private companies, universities, and non-profit entities
    return "N/A", "N/A", org_string.upper(), "NON_FEDERAL_RECIPIENT_ENTITY"

# ==========================================
# 4. EXECUTE NORMALIZATION AND SEED STORAGE
# ==========================================
print("🔗 Running layered token validations across all 270 federal context records...")
for string_item in sorted(all_unique_strings):
    if string_item.upper() == 'NAN' or not string_item.strip(): continue
        
    oracle_row_matches = master_df[master_df['Funding Source Name'] == string_item]
    cayuse_row_matches = gsum_df[gsum_df['FLOW_THROUGH_SPONSOR'] == string_item]
    
    has_fed_term_oracle = oracle_row_matches['AWARD_TERM_NAME'].astype(str).str.upper().str.contains('FEDERAL').any()
    has_fed_term_cayuse = cayuse_row_matches['AWARD_TERM_NAME'].astype(str).str.upper().str.contains('FEDERAL').any()
    is_federal_context = has_fed_term_oracle or has_fed_term_cayuse
    
    # Execute the layered hierarchy resolver
    top, sub, sam_name, tax = resolve_authoritative_hierarchy(string_item, is_federal_context)
    
    oracle_str = string_item if string_item in raw_oracle_names else "N/A"
    cayuse_str = string_item if string_item in raw_cayuse_names or string_item in raw_oracle_names else "N/A"

    cursor.execute("""
        INSERT OR IGNORE INTO staging_sponsor_inputs (oracle_sponsor_string, cayuse_sponsor_string, usaspending_top_tier, usaspending_sub_tier, sam_gov_certified_name, sponsor_taxonomy_class)
        VALUES (?, ?, ?, ?, ?, ?);
    """, (oracle_str, cayuse_str, top, sub, sam_name, tax))

conn.commit()

# ==========================================
# 5. EXPORT FINALIZED LOOKUP REFERENCE VIEW 
# ==========================================
final_query = """
    SELECT 
        oracle_sponsor_string AS ORACLE_SPONSOR_STRING,
        cayuse_sponsor_string AS CAYUSE_SPONSOR_STRING,
        usaspending_top_tier AS USASPENDING_TOP_TIER_AGENCY,
        usaspending_sub_tier AS USASPENDING_SUB_TIER_AGENCY,
        sam_gov_certified_name AS SAM_GOV_CERTIFIED_NAME,
        sponsor_taxonomy_class AS SPONSOR_TAXONOMY_CLASS
    FROM staging_sponsor_inputs;
"""
master_crosswalk_df = pd.read_sql_query(final_query, conn)
conn.close()

try:
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as excel_writer:
        master_crosswalk_df.to_excel(excel_writer, sheet_name="Master Sponsor Crosswalk", index=False)
        worksheet = excel_writer.sheets["Master Sponsor Crosswalk"]
        cell_range = f"A1:F{len(master_crosswalk_df) + 1}"
        excel_table = Table(displayName="MasterSponsorCrosswalk", ref=cell_range)
        excel_table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        worksheet.add_table(excel_table)
    print(f"\n🎉 1-to-1 master naming crosswalk matrix successfully built inside '{OUTPUT_FILE}'!")
except PermissionError:
    print(f"\n❌ Execution Blocked: Close '{OUTPUT_FILE}' in Excel before running this script.")