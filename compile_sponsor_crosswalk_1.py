import os
import re
import json
import time 
import sqlite3
import pandas as pd
import numpy as np
import requests
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# ==========================================
# ENVIRONMENT & AUTOMATED CONFIGURATION
# ==========================================
MASTER_INPUT = "Cayuse-Oracle award True Up - jth6-Copy.xlsx"
OUTPUT_FILE = "Master_Sponsor_Crosswalk_Index.xlsx"
DB_FILE = "rice_sponsor_mdm.db"
SECRETS_FILE = "secrets.json"

# Authoritative GSA Federal Hierarchy API Endpoint
SAM_FH_API_URL = "https://api.sam.gov/prod/federalorganizations/v1/orgs"

print("📊 Launching Relational SAM.gov Live API Deep-Hierarchy Pipeline...")

if not os.path.exists(MASTER_INPUT):
    print(f"❌ Error: Cannot find master workbook '{MASTER_INPUT}' in this folder.")
    exit()

# Securely load credentials from the protected secrets manifest file
SAM_API_KEY = "DEMO_KEY"
if os.path.exists(SECRETS_FILE):
    try:
        with open(SECRETS_FILE, "r") as f:
            config_data = json.load(f)
            SAM_API_KEY = config_data.get("SAM_API_KEY", "DEMO_KEY")
        print("🔑 Secure API credentials loaded successfully from secrets.json.")
    except Exception as e:
        print(f"⚠️ Error reading configuration file secrets.json: {e}. Defaulting to safe mode.")
else:
    print("ℹ️ secrets.json missing. Running in safe fallback mode using local cache rules.")

# ==========================================
# 1. INITIALIZE RELATIONAL SYSTEM STORAGE
# ==========================================
conn = sqlite3.connect(DB_FILE)
cursor = conn.cursor()

# Table A: Permanent local repository caching verified federal agency paths
cursor.execute("""
    CREATE TABLE IF NOT EXISTS ref_sam_federal_hierarchy (
        agency_string_key TEXT PRIMARY KEY,
        top_tier_agency TEXT NOT NULL,
        sub_tier_agency TEXT NOT NULL,
        cgac_code TEXT DEFAULT 'N/A'
    );
""")

# Table B: Workspace staging table tracking unique organizational row maps
cursor.execute("DROP TABLE IF EXISTS staging_sponsor_inputs;")
cursor.execute("""
    CREATE TABLE staging_sponsor_inputs (
        string_id INTEGER PRIMARY KEY AUTOINCREMENT,
        oracle_sponsor_string TEXT UNIQUE,
        cayuse_sponsor_string TEXT,
        usaspending_top_tier TEXT DEFAULT 'N/A',
        usaspending_sub_tier TEXT DEFAULT 'N/A',
        sam_gov_certified_name TEXT,
        sponsor_taxonomy_class TEXT,
        match_derivation_method TEXT
    );
""")
conn.commit()

# ==============================================================================
# 2. SEED THE PROTOCOL WITH THE COMPLETE IMMUTABLE PORTFOLIO AUTHORITY REFERENCE MAP
# ==============================================================================
print("🗂️ Seeding 100% Portfolio-Complete Master Authority Naming Grid into database...")
authority_reference_matrix = [
    # United States Department of Agriculture (USDA) Portfolio
    ('USDA NATIONAL INSTITUTE OF FOOD AND AGRICULTURE', 'DEPARTMENT OF AGRICULTURE (USDA)', 'NATIONAL INSTITUTE OF FOOD AND AGRICULTURE (NIFA)', '012'),
    ('USDA FOREST SERVICE', 'DEPARTMENT OF AGRICULTURE (USDA)', 'FOREST SERVICE (FS)', '012'),
    ('UNITED STATES DEPARTMENT OF AGRICULTURE', 'DEPARTMENT OF AGRICULTURE (USDA)', 'DEPARTMENT OF AGRICULTURE (USDA)', '012'),

    # Health and Human Services / NIH / CDC Portfolio
    ('ADVANCED RESEARCH PROJECTS AGENCY FOR HEALTH (ARPA-H)', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'ADVANCED RESEARCH PROJECTS AGENCY FOR HEALTH (ARPA-H)', '075'),
    ('NATIONAL INSTITUTES OF HEALTH', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'NATIONAL INSTITUTES OF HEALTH (NIH)', '075'),
    ('NATIONAL INSTITUTES OF HEALTH NIH DO NOT USE', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'NATIONAL INSTITUTES OF HEALTH (NIH)', '075'),
    ('DEPARTMENT OF HEALTH AND HUMAN SERVICES CENTERS FOR DISEASE CONTROL', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'CENTERS FOR DISEASE CONTROL AND PREVENTION (CDC)', '075'),
    ('DEPARTMENT OF HEALTH AND HUMAN RESOURCES HEALTH RESOURCES AND SERVICES ADMINISTRATION', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'HEALTH RESOURCES AND SERVICES ADMINISTRATION (HRSA)', '075'),
    ('DHHS HEALTH RESOURCES AND SERVICES ADMINISTRATION', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'HEALTH RESOURCES AND SERVICES ADMINISTRATION (HRSA)', '075'),
    ('DEPARTMENT OF HEALTH AND HUMAN SERVICES ADMINISTRATION FOR CHILDREN AND FAMILIES', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'ADMINISTRATION FOR CHILDREN AND FAMILIES (ACF)', '075'),
    ('DEPARTMENT OF HEALTH AND HUMAN SERVICES ADMINISTRATION FOR COMMUNITY LIVING', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'ADMINISTRATION FOR COMMUNITY LIVING (ACL)', '075'),

    # National Science Foundation Portfolio
    ('NATIONAL SCIENCE FOUNDATION', 'NATIONAL SCIENCE FOUNDATION (NSF)', 'NATIONAL SCIENCE FOUNDATION (NSF)', '049'),
    
    # NASA Operating Center Portfolio
    ('NASA HEADQUARTERS', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'NASA HEADQUARTERS', '080'),
    ('NASA JOHNSON SPACE CENTER (JSC)', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'NASA JOHNSON SPACE CENTER (JSC)', '080'),
    ('NASA LANGLEY RESEARCH CENTER', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'NASA LANGLEY RESEARCH CENTER', '080'),
    ('NATIONAL AERONAUTICS AND SPACE ADMINISTRATION GODDARD', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'NASA GODDARD SPACE FLIGHT CENTER', '080'),
    ('NASA SHARED SERVICES CENTER NSSC', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'NASA SHARED SERVICES CENTER (NSSC)', '080'),
    ('JET PROPULSION LABORATORY, CALIFORNIA INSTITUTE OF TECHNOLOGY', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'JET PROPULSION LABORATORY (JPL)', '080'),

    # Deep-Hierarchy Department of Defense (DOD) Commands
    ('DEPARTMENT OF DEFENSE AIR FORCE OFFICE OF SCIENTIFIC RESEARCH', 'DEPARTMENT OF DEFENSE (DOD)', 'AIR FORCE OFFICE OF SCIENTIFIC RESEARCH (AFOSR)', '097'),
    ('DEPARTMENT OF DEFENSE AIR FORCE RESEARCH LABORATORY', 'DEPARTMENT OF DEFENSE (DOD)', 'AIR FORCE RESEARCH LABORATORY (AFRL)', '097'),
    ('DEPARTMENT OF DEFENSE ARMY RESEARCH INSTITUTE BEHAVIORAL AND SOC SCI ARI', 'DEPARTMENT OF DEFENSE (DOD)', 'US ARMY RESEARCH INSTITUTE (ARI)', '097'),
    ('DEPARTMENT OF DEFENSE ARMY RESEARCH LABORATORY', 'DEPARTMENT OF DEFENSE (DOD)', 'US ARMY RESEARCH LABORATORY (ARL)', '097'),
    ('DEPARTMENT OF DEFENSE ARMY RESEARCH OFFICE', 'DEPARTMENT OF DEFENSE (DOD)', 'US ARMY RESEARCH OFFICE (ARO)', '097'),
    ('DEPARTMENT OF DEFENSE CONGRESSIONALLY DIRECTED MEDICAL RESEARCH PROGRAMS', 'DEPARTMENT OF DEFENSE (DOD)', 'US ARMY MEDICAL RESEARCH ACQUISITION ACTIVITY (USAMRAA)', '097'),
    ('DEPARTMENT OF DEFENSE DEFENSE ADVANCED RESEARCH PROJECTS AGENCY', 'DEPARTMENT OF DEFENSE (DOD)', 'DEFENSE ADVANCED RESEARCH PROJECTS AGENCY (DARPA)', '097'),
    ('DEPARTMENT OF DEFENSE DEFENSE THREAT REDUCTION AGENCY', 'DEPARTMENT OF DEFENSE (DOD)', 'DEFENSE THREAT REDUCTION AGENCY (DTRA)', '097'),
    ('DEPARTMENT OF DEFENSE NAVAL RESEARCH LABORATORY', 'DEPARTMENT OF DEFENSE (DOD)', 'NAVAL RESEARCH LABORATORY (NRL)', '097'),
    ('DEPARTMENT OF DEFENSE OFFICE OF NAVAL RESEARCH', 'DEPARTMENT OF DEFENSE (DOD)', 'OFFICE OF NAVAL RESEARCH (ONR)', '097'),
    ('DEPARTMENT OF DEFENSE STRATEGIC ENVIRONMENTAL RESEARCH AND DEVELOPMENT PROGRAM', 'DEPARTMENT OF DEFENSE (DOD)', 'STRATEGIC ENVIRONMENTAL RESEARCH & DEVELOPMENT PROGRAM (SERDP)', '097'),
    ('DEPARTMENT OF DEFENSE US ARMY', 'DEPARTMENT OF DEFENSE (DOD)', 'DEPARTMENT OF THE ARMY', '097'),
    ('DEPARTMENT OF DEFENSE US ARMY CORPS OF ENGINEERS', 'DEPARTMENT OF DEFENSE (DOD)', 'US ARMY CORPS OF ENGINEERS', '097'),
    ('DEPARTMENT OF DEFENSE US ARMY MEDICAL RESEARCH ACQUISITION ACTIVITY', 'DEPARTMENT OF DEFENSE (DOD)', 'US ARMY MEDICAL RESEARCH ACQUISITION ACTIVITY (USAMRAA)', '097'),
    ('DEPARTMENT OF DEFENSE OTHER AGENCY', 'DEPARTMENT OF DEFENSE (DOD)', 'DEPARTMENT OF DEFENSE (DOD)', '097'),
    ('DEPARTMENT OF DEFENSE US ARMY COMBAT CAPABILITIES DEVELOPMENT COMMAND ARMAMENT', 'DEPARTMENT OF DEFENSE (DOD)', 'US ARMY COMBAT CAPABILITIES DEVELOPMENT COMMAND', '097'),
    ('DEPARTMENT OF DEFENSE DEFENSE TECHNICAL INFORMATION CENTER DTIC', 'DEPARTMENT OF DEFENSE (DOD)', 'DEFENSE TECHNICAL INFORMATION CENTER (DTIC)', '097'),
    ('DEPARTMENT OF DEFENSE INTELLIGENCE ADVANCED RESEARCH PROJECTS ACTIVITY', 'DEPARTMENT OF DEFENSE (DOD)', 'INTELLIGENCE ADVANCED RESEARCH PROJECTS ACTIVITY (IARPA)', '097'),
    ('DEPARTMENT OF DEFENSE SMALL BUSINESS INNOVATION RESEARCH SBIR STTR', 'DEPARTMENT OF DEFENSE (DOD)', 'DEPARTMENT OF DEFENSE (DOD)', '097'),
    ('US ARMY CONTRACTING COMMAND NEW JERSEY', 'DEPARTMENT OF DEFENSE (DOD)', 'US ARMY CONTRACTING COMMAND', '097'),
    ('US ARMY MEDICAL RESEARCH AND DEVELOPMENT COMMAND USAMRDC', 'DEPARTMENT OF DEFENSE (DOD)', 'US ARMY MEDICAL RESEARCH AND DEVELOPMENT COMMAND', '097'),
    ('MARYLAND PROCUREMENT OFFICE', 'DEPARTMENT OF DEFENSE (DOD)', 'MARYLAND PROCUREMENT OFFICE (MPO)', '097'),

    # Granular Department of Energy (DOE) & National Laboratories (FFRDCs)
    ('ARGONNE NATIONAL LABORATORY', 'DEPARTMENT OF ENERGY (DOE)', 'ARGONNE NATIONAL LABORATORY', '089'),
    ('BROOKHAVEN NATIONAL LABORATORY', 'DEPARTMENT OF ENERGY (DOE)', 'BROOKHAVEN NATIONAL LABORATORY', '089'),
    ('DEPARTMENT OF ENERGY NATIONAL NUCLEAR SECURITY ADMINISTRATION', 'DEPARTMENT OF ENERGY (DOE)', 'NATIONAL NUCLEAR SECURITY ADMINISTRATION (NNSA)', '089'),
    ('DEPARTMENT OF ENERGY NATIONAL RENEWABLE ENERGY LABORATORY', 'DEPARTMENT OF ENERGY (DOE)', 'NATIONAL RENEWABLE ENERGY LABORATORY (NREL)', '089'),
    ('DEPARTMENT OF ENERGY NON LOC', 'DEPARTMENT OF ENERGY (DOE)', 'DEPARTMENT OF ENERGY (DOE)', '089'),
    ('DEPARTMENT OF ENERGY OFFICE OF SCIENCE', 'DEPARTMENT OF ENERGY (DOE)', 'DOE OFFICE OF SCIENCE', '089'),
    ('DEPARTMENT OF ENERGY VIPERS', 'DEPARTMENT OF ENERGY (DOE)', 'DEPARTMENT OF ENERGY (DOE)', '089'),
    ('LAWRENCE LIVERMORE NATIONAL LABORATORY', 'DEPARTMENT OF ENERGY (DOE)', 'LAWRENCE LIVERMORE NATIONAL LABORATORY', '089'),
    ('LOS ALAMOS NATIONAL LABORATORY', 'DEPARTMENT OF ENERGY (DOE)', 'LOS ALAMOS NATIONAL LABORATORY', '089'),
    ('OAK RIDGE NATIONAL LABORATORY', 'DEPARTMENT OF ENERGY (DOE)', 'OAK RIDGE NATIONAL LABORATORY', '089'),
    ('PACIFIC NORTHWEST NATIONAL LABORATORY', 'DEPARTMENT OF ENERGY (DOE)', 'PACIFIC NORTHWEST NATIONAL LABORATORY', '089'),
    ('SLAC NATIONAL ACCELERATOR LABORATORY', 'DEPARTMENT OF ENERGY (DOE)', 'SLAC NATIONAL ACCELERATOR LABORATORY', '089'),
    ('FERMI NATIONAL ACCELERATOR LAB', 'DEPARTMENT OF ENERGY (DOE)', 'FERMI NATIONAL ACCELERATOR LABORATORY', '089'),
    ('SANDIA NATIONAL LABORATORIES', 'DEPARTMENT OF ENERGY (DOE)', 'SANDIA NATIONAL LABORATORIES', '089'),
    ('OFFICE OF ENERGY EFFICIENCY AND RENEWABLE ENERGY EERE', 'DEPARTMENT OF ENERGY (DOE)', 'OFFICE OF ENERGY EFFICIENCY AND RENEWABLE ENERGY (EERE)', '089'),
    ('UNIVERSITY OF CALIFORNIA ERNEST ORLANDO LAWRENCE BERKELEY NATIONAL LABORATORY', 'DEPARTMENT OF ENERGY (DOE)', 'LAWRENCE BERKELEY NATIONAL LABORATORY', '089'),

    # Department of Commerce Portfolio
    ('DEPARTMENT OF COMMERCE MBDA', 'DEPARTMENT OF COMMERCE (DOC)', 'MINORITY BUSINESS DEVELOPMENT AGENCY (MBDA)', '013'),
    ('DEPARTMENT OF COMMERCE NIST 13060001 01', 'DEPARTMENT OF COMMERCE (DOC)', 'NATIONAL INSTITUTE OF STANDARDS AND TECHNOLOGY (NIST)', '013'),
    ('DEPARTMENT OF COMMERCE NIST NON LOC', 'DEPARTMENT OF COMMERCE (DOC)', 'NATIONAL INSTITUTE OF STANDARDS AND TECHNOLOGY (NIST)', '013'),
    ('DEPARTMENT OF COMMERCE US CENSUS BUREAU', 'DEPARTMENT OF COMMERCE (DOC)', 'BUREAU OF THE CENSUS', '013'),
    ('NATIONAL OCEANIC ATMOSPHERIC ADMINISTRATION', 'DEPARTMENT OF COMMERCE (DOC)', 'NATIONAL OCEANIC AND ATMOSPHERIC ADMINISTRATION (NOAA)', '013'),
    ('US ECONOMIC DEVELOPMENT ADMINISTRATION', 'DEPARTMENT OF COMMERCE (DOC)', 'ECONOMIC DEVELOPMENT ADMINISTRATION (EDA)', '013'),

    # Independent Federal Departments & National Endowments
    ('DEPARTMENT OF EDUCATION', 'DEPARTMENT OF EDUCATION (ED)', 'DEPARTMENT OF EDUCATION (ED)', '091'),
    ('DEPARTMENT OF INTERIOR', 'DEPARTMENT OF INTERIOR (DOI)', 'DEPARTMENT OF INTERIOR (DOI)', '014'),
    ('DEPARTMENT OF STATE US AGENCY FOR INTERNATIONAL DEVELOPMENT', 'UNITED STATES AGENCY FOR INTERNATIONAL DEVELOPMENT (USAID)', 'UNITED STATES AGENCY FOR INTERNATIONAL DEVELOPMENT (USAID)', '072'),
    ('ENVIRONMENTAL PROTECTION AGENCY EPA', 'ENVIRONMENTAL PROTECTION AGENCY (EPA)', 'ENVIRONMENTAL PROTECTION AGENCY (EPA)', '068'),
    ('US DEPARTMENT OF STATE', 'DEPARTMENT OF STATE (DOS)', 'DEPARTMENT OF STATE (DOS)', '019'),
    ('US DEPARTMENT OF TRANSPORTATION', 'DEPARTMENT OF TRANSPORTATION (DOT)', 'DEPARTMENT OF TRANSPORTATION (DOT)', '069'),
    ('DEPARTMENT OF TREASURY', 'DEPARTMENT OF THE TREASURY', 'DEPARTMENT OF THE TREASURY', '020'),
    ('DEPARTMENT OF THE STATE OFFICE OF THE UNDER SECRETARY FOR PUBLIC DIPLOMACY AND PUBLIC AFFAIRS', 'DEPARTMENT OF STATE (DOS)', 'DEPARTMENT OF STATE (DOS)', '019'),
    ('US DEPT OF HOMELAND SECURITY', 'DEPARTMENT OF HOMELAND SECURITY (DHS)', 'DEPARTMENT OF HOMELAND SECURITY (DHS)', '070'),
    ('US FISH AND WILDLIFE SERVICE', 'DEPARTMENT OF INTERIOR (DOI)', 'UNITED STATES FISH AND WILDLIFE SERVICE', '014'),
    ('US SOCIAL SECURITY ADMINISTRATION', 'SOCIAL SECURITY ADMINISTRATION (SSA)', 'SOCIAL SECURITY ADMINISTRATION (SSA)', '028'),
    ('NATIONAL ENDOWMENT FOR THE ARTS', 'NATIONAL FOUNDATION ON THE ARTS AND THE HUMANITIES', 'NATIONAL ENDOWMENT FOR THE ARTS', '413'),
    ('NATIONAL ENDOWMENT FOR THE HUMANITIES', 'NATIONAL FOUNDATION ON THE ARTS AND THE HUMANITIES', 'NATIONAL ENDOWMENT FOR THE HUMANITIES', '417'),

    # Explicit Higher Ed, Foundations, and Non-Profit Overrides Forced to N/A
    ('US ENDOWMENT FOR FORESTRY AND COMMUNITIES', 'N/A', 'N/A', 'FORCE_NON_FED'),
    ('BCM TRANSLATIONAL RESEARCH INSTITUTE FOR SPACE HEALTH', 'N/A', 'N/A', 'FORCE_NON_FED'),
    ('THE TAMU SYSTEM HEALTH SCIENCE CENTER', 'N/A', 'N/A', 'FORCE_NON_FED'),
    ('UNT HEALTH FORT WORTH', 'N/A', 'N/A', 'FORCE_NON_FED'),
    ('HARRIS COUNTY COMMUNITY SERVICES DEPARTMENT', 'N/A', 'N/A', 'FORCE_NON_FED'),
    ('MARINE BIOLOGICAL LABORATORY', 'N/A', 'N/A', 'FORCE_NON_FED'),
    ('TEXAS COMMISSION ON THE ARTS', 'N/A', 'N/A', 'FORCE_NON_FED'),
    ('TEXAS WORKFORCE COMMISSION', 'N/A', 'N/A', 'FORCE_NON_FED'),
    ('THE JOHN B PIERCE LABORATORY INC', 'N/A', 'N/A', 'FORCE_NON_FED')
]

cursor.executemany("INSERT OR REPLACE INTO ref_sam_federal_hierarchy VALUES (?,?,?,?);", authority_reference_matrix)
conn.commit()

# ==========================================
# 3. INGEST MANUAL SYSTEM EXCEL TABLES
# ==========================================
print("📥 Ingesting manual pre-award and post-award spreadsheet data layers...")
master_df = pd.read_excel(MASTER_INPUT, sheet_name="Intermediate DATA")
gsum_df = pd.read_excel(MASTER_INPUT, sheet_name="Grant Summary")

gsum_df['clean_id'] = gsum_df['AWARD_NUMBER_SCRUB'].astype(str).str.upper().str.strip()
master_df['clean_id'] = master_df['Award Number'].astype(str).str.replace('.0', '', regex=False).str.upper().str.strip()

term_name_map = dict(zip(gsum_df['clean_id'], gsum_df['AWARD_TERM_NAME']))
flow_through_map = dict(zip(gsum_df['clean_id'], gsum_df['FLOW_THROUGH_SPONSOR']))

master_df['AWARD_TERM_NAME'] = master_df['clean_id'].map(term_name_map).fillna('')
master_df['FLOW_THROUGH_SPONSOR'] = master_df['clean_id'].map(flow_through_map).fillna('')

raw_oracle_names = master_df['Funding Source Name'].dropna().unique()
raw_cayuse_names = gsum_df['FLOW_THROUGH_SPONSOR'].dropna().unique()

all_unique_strings = sorted(list(set(list(raw_oracle_names) + list(raw_cayuse_names))))
filtered_strings = [s for s in all_unique_strings if str(s).upper() != 'NAN' and str(s).strip()]

# ==========================================
# 4. LIVE SAM.GOV API NETWORK GATEWAY
# ==========================================
def fetch_live_sam_hierarchy(query_string):
    if SAM_API_KEY == "DEMO_KEY" or not SAM_API_KEY:
        return None
    params = {"api_key": SAM_API_KEY, "search_text": str(query_string).strip(), "status": "Active"}
    try:
        response = requests.get(SAM_FH_API_URL, params=params, timeout=4)
        if response.status_code == 200:
            data = response.json()
            orgs = data.get('organizations', data.get('getFederalOrganizationsList', []))
            if orgs and isinstance(orgs, list):
                target_org = orgs[0]
                top_tier = str(target_org.get('parentOrganization', {}).get('name', '')).upper().strip()
                sub_tier = str(target_org.get('name', '')).upper().strip()
                cgac = str(target_org.get('cgacCode', 'N/A')).strip()
                if not top_tier or top_tier == "NONE": top_tier = sub_tier
                return top_tier, sub_tier, cgac
        return None
    except Exception:
        return None

# ==========================================
# 5. EXECUTE IDENTITY CHECK & VERIFICATION
# ==========================================
total_records = len(filtered_strings)
print(f"🔗 Compiled {total_records} unique organization strings. Processing data gates...\n")

for idx, string_item in enumerate(filtered_strings, 1):
    org_upper = string_item.upper().strip()
    
    oracle_row_matches = master_df[master_df['Funding Source Name'] == string_item]
    cayuse_row_matches = gsum_df[gsum_df['FLOW_THROUGH_SPONSOR'] == string_item]
    
    is_federal_context = (
        oracle_row_matches['AWARD_TERM_NAME'].astype(str).str.upper().str.contains('FEDERAL').any() or
        cayuse_row_matches['AWARD_TERM_NAME'].astype(str).str.upper().str.contains('FEDERAL').any()
    )

    if not is_federal_context:
        top_agency, sub_agency, sam_name, taxonomy = "N/A", "N/A", org_upper, "NON_FEDERAL_RECIPIENT_ENTITY"
        derivation_method = "SYSTEM_NAME_GATE"
    else:
        # Check our relational cache map first where portfolio-wide authority elements are locked
        cursor.execute("SELECT top_tier_agency, sub_tier_agency, cgac_code FROM ref_sam_federal_hierarchy WHERE agency_string_key = ?;", (org_upper,))
        cache_match = cursor.fetchone()
        
        if cache_match:
            cached_top, cached_sub, cached_code = cache_match
            if cached_code == 'FORCE_NON_FED' or cached_top == 'N/A':
                top_agency, sub_agency, sam_name, taxonomy = "N/A", "N/A", org_upper, "NON_FEDERAL_RECIPIENT_ENTITY"
                derivation_method = "EXACT_AUTHORITY_SEED"
            elif 'FALLBACK' in str(cached_code):
                top_agency, sub_agency, sam_name, taxonomy = cached_top, cached_sub, cached_sub, "FEDERAL_FUNDING_AGENCY"
                derivation_method = "RESCUE_FALLBACK"
            else:
                top_agency, sub_agency, sam_name, taxonomy = cached_top, cached_sub, cached_sub, "FEDERAL_FUNDING_AGENCY"
                derivation_method = "EXACT_AUTHORITY_SEED"
        else:
            # IDENTITY GATEKEEPER 2: Enhanced Text Screen for Private Higher Ed & Non-Profit Entities
            has_corporate_suffix = any(sfx in org_upper for sfx in [" INC", " LLC", " CORP", " CO", " LTD", "UNIVERSITY", "FOUNDATION", "COLLEGE", "INSTITUTE", "CENTER", "SYSTEM", "HOSPITAL", "CLINIC"])
            has_federal_keyword = any(fed in org_upper for fed in ["DEPARTMENT", "AGENCY", "COMMISSION", "COMMAND", "ADMINISTRATION", "BUREAU", "OFFICE", "LABORATORY", "NIH", "NSF", "NASA", "USDA", "AGRICULTURE", "ENDOWMENT", "SERVICE"])
            
            if has_corporate_suffix and not has_federal_keyword:
                top_agency, sub_agency, sam_name, taxonomy = "N/A", "N/A", org_upper, "NON_FEDERAL_RECIPIENT_ENTITY"
                derivation_method = "SYSTEM_NAME_GATE"
            else:
                # Gateway fallback loop addressing brand-new, unmapped lines
                print(f" 🌐 [{idx} / {total_records}] Evaluating newly detected line via SAM.gov API: '{string_item}'...")
                api_result = fetch_live_sam_hierarchy(string_item)
                time.sleep(0.2)
                
                if api_result:
                    top_agency, sub_agency, cgac_code = api_result
                    sam_name = sub_agency
                    taxonomy = "FEDERAL_FUNDING_AGENCY"
                    derivation_method = "LIVE_API_RESOLVED"
                    cursor.execute("INSERT OR REPLACE INTO ref_sam_federal_hierarchy VALUES (?, ?, ?, ?);", (org_upper, top_agency, sub_agency, cgac_code))
                    conn.commit()
                else:
                    # IDENTITY GATEKEEPER 4: API returned None -> Evaluate for definitive federal markers
                    is_explicit_gov = any(word in org_upper for word in ["DEPARTMENT", "AGENCY", "COMMISSION", "COMMAND", "ADMINISTRATION", "BUREAU", "OFFICE", "LABORATORY", "NIH", "NSF", "NASA", "FORCE", "ARMY", "NAVY", "COMMERCE", "ENERGY", "USDA", "AGRICULTURE", "ENDOWMENT", "SERVICE", "FED", "US "])
                    
                    if is_explicit_gov:
                        if any(w in org_upper for w in ["AGRICULTURE", "USDA", "NIFA", "FOREST SERVICE"]):
                            top_agency = "DEPARTMENT OF AGRICULTURE (USDA)"
                            sub_agency = "NATIONAL INSTITUTE OF FOOD AND AGRICULTURE (NIFA)" if "FOOD" in org_upper or "NIFA" in org_upper else ("FOREST SERVICE (FS)" if "FOREST" in org_upper else "DEPARTMENT OF AGRICULTURE (USDA)")
                        elif "ENDOWMENT" in org_upper:
                            top_agency = "NATIONAL FOUNDATION ON THE ARTS AND THE HUMANITIES"
                            sub_agency = "NATIONAL ENDOWMENT FOR THE ARTS" if "ARTS" in org_upper else "NATIONAL ENDOWMENT FOR THE HUMANITIES"
                        elif "HOMELAND SECURITY" in org_upper or "DHS" in org_upper:
                            top_agency, sub_agency = "DEPARTMENT OF HOMELAND SECURITY (DHS)", "DEPARTMENT OF HOMELAND SECURITY (DHS)"
                        elif "SOCIAL SECURITY" in org_upper or "SSA" in org_upper:
                            top_agency, sub_agency = "SOCIAL SECURITY ADMINISTRATION (SSA)", "SOCIAL SECURITY ADMINISTRATION (SSA)"
                        elif "FISH AND WILDLIFE" in org_upper:
                            top_agency, sub_agency = "DEPARTMENT OF INTERIOR (DOI)", "UNITED STATES FISH AND WILDLIFE SERVICE"
                        elif "DEFENSE THREAT REDUCTION" in org_upper or "DTRA" in org_upper or "DEFENSE" in org_upper or "ARMY" in org_upper or "NAVY" in org_upper or "AIR FORCE" in org_upper or "ONR" in org_upper or "DARPA" in org_upper:
                            top_agency = "DEPARTMENT OF DEFENSE (DOD)"
                            sub_agency = "US ARMY RESEARCH LABORATORY (ARL)" if "LABORATORY" in org_upper else ("US ARMY RESEARCH OFFICE (ARO)" if "OFFICE" in org_upper else ("DEFENSE THREAT REDUCTION AGENCY (DTRA)" if "THREAT" in org_upper else "DEPARTMENT OF THE ARMY"))
                        elif "HEALTH" in org_upper or "NIH" in org_upper or "CDC" in org_upper or "DHHS" in org_upper:
                            top_agency, sub_agency = "DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)", "NATIONAL INSTITUTES OF HEALTH (NIH)"
                        elif "SCIENCE FOUNDATION" in org_upper or "NSF" in org_upper:
                            top_agency, sub_agency = "NATIONAL SCIENCE FOUNDATION (NSF)", "NATIONAL SCIENCE FOUNDATION (NSF)"
                        elif "ENERGY" in org_upper or "DOE" in org_upper:
                            top_agency, sub_agency = "DEPARTMENT OF ENERGY (DOE)", "DEPARTMENT OF ENERGY (DOE)"
                        elif "COMMERCE" in org_upper or "NIST" in org_upper or "NOAA" in org_upper:
                            top_agency, sub_agency = "DEPARTMENT OF COMMERCE (DOC)", "DEPARTMENT OF COMMERCE (DOC)"
                        elif "NASA" in org_upper or "SPACE" in org_upper:
                            top_agency, sub_agency = "NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)", "NASA HEADQUARTERS"
                        else:
                            top_agency, sub_agency = "FEDERAL_DEPARTMENT", "UNMAPPED_FEDERAL_AGENCY"
                            
                        sam_name = sub_agency
                        taxonomy = "FEDERAL_FUNDING_AGENCY"
                        cache_code = "FALLBACK_GOV_LOCK"
                        derivation_method = "RESCUE_FALLBACK"
                    
                    else:
                        
                        top_agency, sub_agency = "N/A", "N/A"
                        sam_name = "N/A"
                        taxonomy = "NON_FEDERAL_RECIPIENT_ENTITY"
                        cache_code = "FALLBACK_NON_FED_LOCK"
                        derivation_method = "SYSTEM_NAME_GATE"
                    
                    cursor.execute("INSERT OR REPLACE INTO ref_sam_federal_hierarchy VALUES (?, ?, ?, ?);", (org_upper, top_agency, sub_agency, cache_code))
                    conn.commit()

    oracle_str = string_item if string_item in raw_oracle_names else "N/A"
    cayuse_str = string_item if string_item in raw_cayuse_names or string_item in raw_oracle_names else "N/A"

    cursor.execute("""
        INSERT OR IGNORE INTO staging_sponsor_inputs (oracle_sponsor_string, cayuse_sponsor_string, usaspending_top_tier, usaspending_sub_tier, sam_gov_certified_name, sponsor_taxonomy_class, match_derivation_method)
        VALUES (?, ?, ?, ?, ?, ?, ?);
    """, (oracle_str, cayuse_str, top_agency, sub_agency, sam_name, taxonomy, derivation_method))

conn.commit()

# ==========================================
# 6. GENERATE FINAL REPORT CROSSWALK GRID
# ==========================================
print("\n💾 Ingestion complete. Compressing views and extracting final load sheets...")
final_query = """
    SELECT 
        oracle_sponsor_string AS ORACLE_SPONSOR_STRING,
        cayuse_sponsor_string AS CAYUSE_SPONSOR_STRING,
        usaspending_top_tier AS USASPENDING_TOP_TIER_AGENCY,
        usaspending_sub_tier AS USASPENDING_SUB_TIER_AGENCY,
        sam_gov_certified_name AS SAM_GOV_CERTIFIED_NAME,
        sponsor_taxonomy_class AS SPONSOR_TAXONOMY_CLASS,
        match_derivation_method AS MATCH_DERIVATION_METHOD
    FROM staging_sponsor_inputs;
"""
master_crosswalk_df = pd.read_sql_query(final_query, conn)
conn.close()

try:
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as excel_writer:
        master_crosswalk_df.to_excel(excel_writer, sheet_name="Master Sponsor Crosswalk", index=False)
        worksheet = excel_writer.sheets["Master Sponsor Crosswalk"]
        cell_range = f"A1:G{len(master_crosswalk_df) + 1}"
        excel_table = Table(displayName="MasterSponsorCrosswalk", ref=cell_range)
        excel_table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        worksheet.add_table(excel_table)
    print(f"🎉 Success! 1-to-1 master crosswalk naming grid successfully saved to '{OUTPUT_FILE}'!")
except PermissionError:
    print(f"❌ Execution Blocked: Close '{OUTPUT_FILE}' in Excel before running this script.")