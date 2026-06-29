import os
import re
import sqlite3
import pandas as pd
from openpyxl.worksheet.table import Table, TableStyleInfo

# ==========================================
# ENVIRONMENT & AUTOMATED CONFIGURATION
# ==========================================
MASTER_INPUT = "Cayuse-Oracle award True Up - jth6-Copy.xlsx"
OUTPUT_FILE = "Master_Sponsor_Crosswalk_Index.xlsx"
DB_FILE = "rice_sponsor_mdm.db"

print("📊 Launching Master Relational Report Compiler (Database-Connected Version)...")

if not os.path.exists(MASTER_INPUT) or not os.path.exists(DB_FILE):
    print("❌ Error: Missing master input spreadsheet or verified database cache file.")
    exit()

conn = sqlite3.connect(DB_FILE)
cursor = conn.cursor()

# Self-healing database check to ensure federal tables exist
cursor.execute("""
    CREATE TABLE IF NOT EXISTS ref_sam_federal_hierarchy (
        agency_string_key TEXT PRIMARY KEY,
        top_tier_agency TEXT NOT NULL,
        sub_tier_agency TEXT NOT NULL,
        cgac_code TEXT DEFAULT 'N/A'
    );
""")

authority_reference_matrix = [
    ('USDA NATIONAL INSTITUTE OF FOOD AND AGRICULTURE', 'DEPARTMENT OF AGRICULTURE (USDA)', 'NATIONAL INSTITUTE OF FOOD AND AGRICULTURE (NIFA)', '012'),
    ('USDA FOREST SERVICE', 'DEPARTMENT OF AGRICULTURE (USDA)', 'FOREST SERVICE (FS)', '012'),
    ('UNITED STATES DEPARTMENT OF AGRICULTURE', 'DEPARTMENT OF AGRICULTURE (USDA)', 'DEPARTMENT OF AGRICULTURE (USDA)', '012'),
    ('ADVANCED RESEARCH PROJECTS AGENCY FOR HEALTH (ARPA-H)', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'ADVANCED RESEARCH PROJECTS AGENCY FOR HEALTH (ARPA-H)', '075'),
    ('NATIONAL INSTITUTES OF HEALTH', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'NATIONAL INSTITUTES OF HEALTH (NIH)', '075'),
    ('NATIONAL INSTITUTES OF HEALTH NIH DO NOT USE', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'NATIONAL INSTITUTES OF HEALTH (NIH)', '075'),
    ('DEPARTMENT OF HEALTH AND HUMAN SERVICES CENTERS FOR DISEASE CONTROL', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'CENTERS FOR DISEASE CONTROL AND PREVENTION (CDC)', '075'),
    ('DEPARTMENT OF HEALTH AND HUMAN RESOURCES HEALTH RESOURCES AND SERVICES ADMINISTRATION', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'HEALTH RESOURCES AND SERVICES ADMINISTRATION (HRSA)', '075'),
    ('DHHS HEALTH RESOURCES AND SERVICES ADMINISTRATION', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'HEALTH RESOURCES AND SERVICES ADMINISTRATION (HRSA)', '075'),
    ('DEPARTMENT OF HEALTH AND HUMAN SERVICES ADMINISTRATION FOR CHILDREN AND FAMILIES', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'ADMINISTRATION FOR CHILDREN AND FAMILIES (ACF)', '075'),
    ('DEPARTMENT OF HEALTH AND HUMAN SERVICES ADMINISTRATION FOR COMMUNITY LIVING', 'DEPARTMENT OF HEALTH AND HUMAN SERVICES (HHS)', 'ADMINISTRATION FOR COMMUNITY LIVING (ACL)', '075'),
    ('NATIONAL SCIENCE FOUNDATION', 'NATIONAL SCIENCE FOUNDATION (NSF)', 'NATIONAL SCIENCE FOUNDATION (NSF)', '049'),
    ('NASA HEADQUARTERS', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'NASA HEADQUARTERS', '080'),
    ('NASA JOHNSON SPACE CENTER (JSC)', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'NASA JOHNSON SPACE CENTER (JSC)', '080'),
    ('NASA LANGLEY RESEARCH CENTER', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'NASA LANGLEY RESEARCH CENTER', '080'),
    ('NATIONAL AERONAUTICS AND SPACE ADMINISTRATION GODDARD', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'NASA GODDARD SPACE FLIGHT CENTER', '080'),
    ('NASA SHARED SERVICES CENTER NSSC', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'NASA SHARED SERVICES CENTER (NSSC)', '080'),
    ('JET PROPULSION LABORATORY, CALIFORNIA INSTITUTE OF TECHNOLOGY', 'NATIONAL AERONAUTICS AND SPACE ADMINISTRATION (NASA)', 'JET PROPULSION LABORATORY (JPL)', '080'),
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
    ('THOMAS JEFFERSON NATIONAL ACCELERATOR FACILITY', 'DEPARTMENT OF ENERGY (DOE)', 'THOMAS JEFFERSON NATIONAL ACCELERATOR FACILITY', '089'),
    ('OFFICE OF ENERGY EFFICIENCY AND RENEWABLE ENERGY EERE', 'DEPARTMENT OF ENERGY (DOE)', 'OFFICE OF ENERGY EFFICIENCY AND RENEWABLE ENERGY (EERE)', '089'),
    ('UNIVERSITY OF CALIFORNIA ERNEST ORLANDO LAWRENCE BERKELEY NATIONAL LABORATORY', 'DEPARTMENT OF ENERGY (DOE)', 'LAWRENCE BERKELEY NATIONAL LABORATORY', '089'),
    ('DEPARTMENT OF COMMERCE MBDA', 'DEPARTMENT OF COMMERCE (DOC)', 'MINORITY BUSINESS DEVELOPMENT AGENCY (MBDA)', '013'),
    ('DEPARTMENT OF COMMERCE NIST 13060001 01', 'DEPARTMENT OF COMMERCE (DOC)', 'NATIONAL INSTITUTE OF STANDARDS AND TECHNOLOGY (NIST)', '013'),
    ('DEPARTMENT OF COMMERCE NIST NON LOC', 'DEPARTMENT OF COMMERCE (DOC)', 'NATIONAL INSTITUTE OF STANDARDS AND TECHNOLOGY (NIST)', '013'),
    ('DEPARTMENT OF COMMERCE US CENSUS BUREAU', 'DEPARTMENT OF COMMERCE (DOC)', 'BUREAU OF THE CENSUS', '013'),
    ('NATIONAL OCEANIC ATMOSPHERIC ADMINISTRATION', 'DEPARTMENT OF COMMERCE (DOC)', 'NATIONAL OCEANIC AND ATMOSPHERIC ADMINISTRATION (NOAA)', '013'),
    ('US ECONOMIC DEVELOPMENT ADMINISTRATION', 'DEPARTMENT OF COMMERCE (DOC)', 'ECONOMIC DEVELOPMENT ADMINISTRATION (EDA)', '013'),
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
    ('FACULTY FUND', 'N/A', 'N/A', 'FORCE_NON_FED'),
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

# Clear and rebuild reporting scratch staging table
cursor.execute("DROP TABLE IF EXISTS staging_sponsor_inputs;")
cursor.execute("""
    CREATE TABLE staging_sponsor_inputs (
        string_id INTEGER PRIMARY KEY AUTOINCREMENT,
        oracle_sponsor_string TEXT UNIQUE,
        cayuse_sponsor_string TEXT,
        usaspending_top_tier TEXT DEFAULT 'N/A',
        usaspending_sub_tier TEXT DEFAULT 'N/A',
        sam_gov_certified_name TEXT DEFAULT 'N/A',
        assigned_id TEXT DEFAULT 'N/A',
        sponsor_taxonomy_class TEXT,
        match_derivation_method TEXT
    );
""")
conn.commit()

# Loading manual portfolios
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

# Relational Join Loop Checking Persistent Database Cache Mappings
for string_item in filtered_strings:
    org_upper = string_item.upper().strip()
    
    oracle_row_matches = master_df[master_df['Funding Source Name'] == string_item]
    cayuse_row_matches = gsum_df[gsum_df['FLOW_THROUGH_SPONSOR'] == string_item]
    
    is_federal_context = (
        oracle_row_matches['AWARD_TERM_NAME'].astype(str).str.upper().str.contains('FEDERAL').any() or
        cayuse_row_matches['AWARD_TERM_NAME'].astype(str).str.upper().str.contains('FEDERAL').any()
    )

    assigned_id = "N/A"
    
    # 🧠 CORE CORRECTION: Enforce direct lookup into the ref_sam_entity_registry built by curation module
    if not is_federal_context:
        cursor.execute("SELECT verified_uei, verified_duns, sam_gov_legal_name, curation_status FROM ref_sam_entity_registry WHERE local_string_key = ?;", (org_upper,))
        non_fed_match = cursor.fetchone()
        
        if non_fed_match:
            uei, duns, legal_name, status = non_fed_match
            top_agency, sub_agency = "N/A", "N/A"
            sam_name = legal_name if legal_name != "N/A" else "N/A"
            assigned_id = uei if uei != "N/A" else duns
            taxonomy = "NON_FEDERAL_RECIPIENT_ENTITY"
            derivation_method = status
        else:
            top_agency, sub_agency, sam_name = "N/A", "N/A", "N/A"
            taxonomy = "NON_FEDERAL_RECIPIENT_ENTITY"
            derivation_method = "SYSTEM_TAXONOMY_GATE"
    else:
        cursor.execute("SELECT top_tier_agency, sub_tier_agency, cgac_code FROM ref_sam_federal_hierarchy WHERE agency_string_key = ?;", (org_upper,))
        fed_cache_match = cursor.fetchone()
        
        if fed_cache_match:
            cached_top, cached_sub, cached_code = fed_cache_match
            if cached_code == 'FORCE_NON_FED' or cached_top == 'N/A':
                cursor.execute("SELECT verified_uei, verified_duns, sam_gov_legal_name, curation_status FROM ref_sam_entity_registry WHERE local_string_key = ?;", (org_upper,))
                override_match = cursor.fetchone()
                
                top_agency, sub_agency = "N/A", "N/A"
                taxonomy = "NON_FEDERAL_RECIPIENT_ENTITY"
                if override_match:
                    uei, duns, legal_name, status = override_match
                    sam_name = legal_name
                    assigned_id = uei if uei != "N/A" else duns
                    derivation_method = status
                else:
                    sam_name = "N/A"
                    derivation_method = "EXACT_AUTHORITY_SEED"
            else:
                top_agency, sub_agency, sam_name = cached_top, cached_sub, cached_sub
                assigned_id = cached_code
                taxonomy = "FEDERAL_FUNDING_AGENCY"
                derivation_method = "EXACT_AUTHORITY_SEED"
        else:
            is_explicit_gov = any(word in org_upper for word in ["DEPARTMENT", "AGENCY", "COMMISSION", "COMMAND", "ADMINISTRATION", "BUREAU", "OFFICE", "LABORATORY", "NIH", "NSF", "NASA", "US "])
            if is_explicit_gov:
                top_agency, sub_agency = "FEDERAL_DEPARTMENT", "UNMAPPED_FEDERAL_AGENCY"
                sam_name = sub_agency
                taxonomy = "FEDERAL_FUNDING_AGENCY"
                derivation_method = "RESCUE_FALLBACK"
            else:
                top_agency, sub_agency, sam_name = "N/A", "N/A", "N/A"
                taxonomy = "NON_FEDERAL_RECIPIENT_ENTITY"
                derivation_method = "SYSTEM_TAXONOMY_GATE"

    oracle_str = string_item if string_item in raw_oracle_names else "N/A"
    cayuse_str = string_item if string_item in raw_cayuse_names or string_item in raw_oracle_names else "N/A"

    cursor.execute("""
        INSERT OR IGNORE INTO staging_sponsor_inputs (oracle_sponsor_string, cayuse_sponsor_string, usaspending_top_tier, usaspending_sub_tier, sam_gov_certified_name, assigned_id, sponsor_taxonomy_class, match_derivation_method)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?);
    """, (oracle_str, cayuse_str, top_agency, sub_agency, sam_name, assigned_id, taxonomy, derivation_method))

conn.commit()

# Multi-Tab Worksheet Data Grid Generation
master_crosswalk_df = pd.read_sql_query("""
    SELECT 
        oracle_sponsor_string AS ORACLE_SPONSOR_STRING,
        cayuse_sponsor_string AS CAYUSE_SPONSOR_STRING,
        usaspending_top_tier AS USASPENDING_TOP_TIER_AGENCY,
        usaspending_sub_tier AS USASPENDING_SUB_TIER_AGENCY,
        sam_gov_certified_name AS SAM_GOV_CERTIFIED_NAME,
        assigned_id AS ASSIGNED_ID,
        sponsor_taxonomy_class AS SPONSOR_TAXONOMY_CLASS,
        match_derivation_method AS MATCH_DERIVATION_METHOD
    FROM staging_sponsor_inputs;
""", conn)

fed_registry_df = pd.read_sql_query("""
    SELECT DISTINCT top_tier_agency AS TOP_TIER_AGENCY, sub_tier_agency AS SUB_TIER_AGENCY, cgac_code AS CGAC_CODE 
    FROM ref_sam_federal_hierarchy WHERE cgac_code != 'FORCE_NON_FED' AND top_tier_agency != 'N/A'
    ORDER BY top_tier_agency ASC, sub_tier_agency ASC;
""", conn)

non_fed_registry_df = pd.read_sql_query("""
    SELECT local_string_key AS LOCAL_SPONSOR_STRING, verified_uei AS VERIFIED_UEI, verified_duns AS VERIFIED_DUNS, sam_gov_legal_name AS SAM_GOV_LEGAL_NAME, curation_status AS DERIVATION_SOURCE
    FROM ref_sam_entity_registry ORDER BY local_string_key ASC;
""", conn)
conn.close()

rules_df = pd.DataFrame([
    ["EXACT_AUTHORITY_SEED", "Pre-Mapped Core Matches", "Locks hardcoded federal commands and verified authority seeds from the local master catalog."],
    ["VERIFIED_VIA_OFFLINE_NAME_MATCH", "sam.xlsx Global Name Index Join", "Alphanumerically verified and resolved the entity name against the 40MB shared drive dataset."],
    ["VERIFIED_VIA_SUPPLIER_DUNS", "OTBI Supplier File Join", "Cross-examined internal supplier metrics against the offline global tracking registry."],
    ["VERIFIED_VIA_AI_INPUT", "Validated AI Key Entry", "Successfully matched and verified an AI-predicted UEI code against the master file."],
    ["SYSTEM_TAXONOMY_GATE", "Generic Corporate Text Suffix Interception", "Flags non-federal context lines and clears direct federal operating column variables to N/A."]
], columns=["RULE_ID", "ROUTING_GATEWAYS", "LOGIC_DESCRIPTION"])

try:
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        master_crosswalk_df.to_excel(writer, sheet_name="Master Sponsor Crosswalk", index=False)
        t1 = Table(displayName="MasterSponsorCrosswalk", ref=f"A1:H{len(master_crosswalk_df) + 1}")
        t1.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        writer.sheets["Master Sponsor Crosswalk"].add_table(t1)
        
        fed_registry_df.to_excel(writer, sheet_name="Federal_Hierarchy_Registry", index=False)
        t2 = Table(displayName="FederalHierarchyRegistry", ref=f"A1:C{len(fed_registry_df) + 1}")
        t2.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        writer.sheets["Federal_Hierarchy_Registry"].add_table(t2)
        
        non_fed_registry_df.to_excel(writer, sheet_name="Non_Federal_Entity_Registry", index=False)
        t3 = Table(displayName="NonFederalEntityRegistry", ref=f"A1:E{len(non_fed_registry_df) + 1}")
        t3.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        writer.sheets["Non_Federal_Entity_Registry"].add_table(t3)
        
        rules_df.to_excel(writer, sheet_name="RULES", index=False)
        t4 = Table(displayName="RULES", ref=f"A1:C{len(rules_df) + 1}")
        t4.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        writer.sheets["RULES"].add_table(t4)
        
    print(f"🎉 Success! Multi-tab relational data grid successfully generated and saved to '{OUTPUT_FILE}'!")
except PermissionError:
    print(f"❌ Execution Blocked: Close '{OUTPUT_FILE}' before running this script.")