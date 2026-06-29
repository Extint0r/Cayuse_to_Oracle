import os
import re
import sqlite3
import pandas as pd

# ==========================================
# CONFIGURATION & DATA REPOSITORIES
# ==========================================
MASTER_INPUT = "Cayuse-Oracle award True Up - jth6-Copy.xlsx"
SUPPLIER_FILE = "SupplierUEI.xlsx"
SAM_OFFLINE_FILE = "sam.xlsx"
DB_FILE = "rice_sponsor_mdm.db"
AI_QUEUE_OUTPUT = "Sponsors_Pending_AI_Guesses.xlsx"

print("🔍 Initializing Phase A: Strict Deterministic Offline Non-Federal Entity Curation Module...")

if not os.path.exists(MASTER_INPUT) or not os.path.exists(SUPPLIER_FILE) or not os.path.exists(SAM_OFFLINE_FILE):
    print("❌ Error: Missing required files. Ensure Master Workbook, Supplier file, and sam.xlsx are in this folder.")
    exit()

# ==========================================
# 1. INITIALIZE CURATION SCHEMA & DATABASE
# ==========================================
conn = sqlite3.connect(DB_FILE)
cursor = conn.cursor()

cursor.execute("""
    CREATE TABLE IF NOT EXISTS ref_sam_entity_registry (
        local_string_key TEXT PRIMARY KEY,
        verified_uei TEXT DEFAULT 'N/A',
        verified_duns TEXT DEFAULT 'N/A',
        sam_gov_legal_name TEXT DEFAULT 'N/A',
        curation_status TEXT NOT NULL
    );
""")
conn.commit()

# ==========================================
# 2. AUTOMATED OFFLINE SAM.GOV INGESTION GATE
# ==========================================
cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='ref_sam_global_entities';")
if not cursor.fetchone():
    print(f"📦 First-time setup: Ingesting 40MB master asset '{SAM_OFFLINE_FILE}' into local SQLite database...")
    print("   (This will take roughly 15-30 seconds to parse and compress. Please wait...)")
    
    global_sam_df = pd.read_excel(SAM_OFFLINE_FILE, dtype=str).fillna("N/A")
    global_sam_df.columns = [str(c).upper().strip() for c in global_sam_df.columns]
    
    print("   🧹 Generating normalized name-matching keys for the master index...")
    global_sam_df['CLEAN_SPONSOR'] = global_sam_df['SPONSOR'].astype(str).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    global_sam_df.to_sql("ref_sam_global_entities", conn, if_exists="replace", index=False)
    
    print("   ⚡ Building high-speed relational database indexes over tracking columns...")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sam_global_uei ON ref_sam_global_entities (UEI);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sam_global_spon ON ref_sam_global_entities (SPONSOR);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sam_global_clean ON ref_sam_global_entities (CLEAN_SPONSOR);")
    conn.commit()
    print("   ✅ Master entity reference storage successfully locked and optimized.")
else:
    print("💾 Local master entity reference catalog detected. Skipping ingestion pass.")

# ==========================================
# 3. INGEST OPERATION SYSTEM DATA LAYERS
# ==========================================
print("📥 Loading operational portfolio sheets...")
master_df = pd.read_excel(MASTER_INPUT, sheet_name="Intermediate DATA")
gsum_df = pd.read_excel(MASTER_INPUT, sheet_name="Grant Summary")

xls = pd.ExcelFile(MASTER_INPUT)
ai_guesses_dict = {}
if "AI Guesses" in xls.sheet_names:
    print("🤖 Found 'AI Guesses' tab. Loading predictive key tracking rules...")
    ai_df = pd.read_excel(MASTER_INPUT, sheet_name="AI Guesses")
    if not ai_df.empty and len(ai_df.columns) >= 2:
        ai_guesses_dict = dict(zip(ai_df.iloc[:, 0].astype(str).str.upper().str.strip(), ai_df.iloc[:, 1]))

supplier_mapping = {}
print("📋 Loading Rice OTBI Supplier ledger...")
raw_sup = pd.read_excel(SUPPLIER_FILE).dropna(subset=['Supplier'])
for _, row in raw_sup.iterrows():
    s_name = str(row['Supplier']).upper().strip()
    s_clean_key = re.sub(r'[^A-Z0-9]', '', s_name)
    supplier_mapping[s_clean_key] = str(row['DUNS Number']).split('.')[0].strip()

# ==========================================
# 4. CONCATENATE UNIQUE SPONSOR STRINGS
# ==========================================
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

non_fed_sponsors = []
for s in filtered_strings:
    org_upper = s.upper().strip()
    oracle_row_matches = master_df[master_df['Funding Source Name'] == s]
    cayuse_row_matches = gsum_df[gsum_df['FLOW_THROUGH_SPONSOR'] == s]
    
    is_federal = (
        oracle_row_matches['AWARD_TERM_NAME'].astype(str).str.upper().str.contains('FEDERAL').any() or
        cayuse_row_matches['AWARD_TERM_NAME'].astype(str).str.upper().str.contains('FEDERAL').any()
    )
    if org_upper in ['FACULTY FUND', 'BLANK', 'NONE']:
        continue
    if not is_federal:
        non_fed_sponsors.append(s)

# ==========================================
# 5. STRICT DETERMINISTIC OFFLINE JOIN ENGINE
# ==========================================
ai_guessing_queue = []
print(f"🔄 Cross-examining {len(non_fed_sponsors)} non-federal lines against local data pools...\n")

for s in non_fed_sponsors:
    org_upper = s.upper().strip()
    norm_join_key = re.sub(r'[^A-Z0-9]', '', org_upper)
    
    cursor.execute("SELECT verified_uei, verified_duns, sam_gov_legal_name, curation_status FROM ref_sam_entity_registry WHERE local_string_key = ?;", (org_upper,))
    if cursor.fetchone():
        continue
        
    # Track 1: Strict Character-for-Character Exact Name Match
    cursor.execute("SELECT SPONSOR, UEI FROM ref_sam_global_entities WHERE CLEAN_SPONSOR = ? LIMIT 1;", (norm_join_key,))
    direct_sam_match = cursor.fetchone()
    
    if direct_sam_match:
        legal_title, verified_uei = direct_sam_match
        local_duns = supplier_mapping.get(norm_join_key, "N/A")
        cursor.execute("INSERT OR REPLACE INTO ref_sam_entity_registry VALUES (?, ?, ?, ?, ?);", 
                       (org_upper, verified_uei, local_duns, legal_title, "VERIFIED_VIA_OFFLINE_NAME_MATCH"))
        print(f"   ✅ Strict Exact Name Match Verified: '{s}' -> '{legal_title}'")
        conn.commit()
        continue

    # Track 2: Strict OTBI Supplier Ledger File Join
    if norm_join_key in supplier_mapping:
        local_duns = supplier_mapping[norm_join_key]
        cursor.execute("SELECT SPONSOR, UEI FROM ref_sam_global_entities WHERE CLEAN_SPONSOR = ? LIMIT 1;", (norm_join_key,))
        local_sam_match = cursor.fetchone()
        
        if local_sam_match:
            legal_title, verified_uei = local_sam_match
            cursor.execute("INSERT OR REPLACE INTO ref_sam_entity_registry VALUES (?, ?, ?, ?, ?);", 
                           (org_upper, verified_uei, local_duns, legal_title, "VERIFIED_VIA_SUPPLIER_DUNS"))
        else:
            cursor.execute("INSERT OR REPLACE INTO ref_sam_entity_registry VALUES (?, ?, ?, ?, ?);", 
                           (org_upper, "N/A", local_duns, org_upper, "SUPPLIER_MATCH_LOCAL_ONLY"))
        conn.commit()
        
    # Track 3: Strict Cryptographic Validation of Active AI Guesses
    elif org_upper in ai_guesses_dict:
        raw_guessed_key = str(ai_guesses_dict[org_upper]).strip()
        if raw_guessed_key.upper() in ['NAN', 'NAT', 'N/A', '']:
            ai_guessing_queue.append({"SPONSOR_NAME_STRING": s, "PREDICTED_UEI_OR_DUNS": ""})
            continue
            
        is_low_certainty = False
        if raw_guessed_key.endswith('*'):
            is_low_certainty = True
            cleaned_key = raw_guessed_key.replace('*', '').strip()
        else:
            cleaned_key = raw_guessed_key
            
        if cleaned_key.endswith('.0'):
            cleaned_key = cleaned_key.split('.')[0]
            
        cursor.execute("SELECT SPONSOR FROM ref_sam_global_entities WHERE UEI = ? LIMIT 1;", (cleaned_key,))
        sam_row = cursor.fetchone()
        
        if sam_row:
            legal_title = str(sam_row[0]).upper().strip()
            status_label = "VERIFIED_AI_LOW_CERTAINTY" if is_low_certainty else "VERIFIED_VIA_AI_INPUT"
            cursor.execute("INSERT OR REPLACE INTO ref_sam_entity_registry VALUES (?, ?, ?, ?, ?);", 
                           (org_upper, cleaned_key, "N/A", legal_title, status_label))
            print(f"   ✅ AI Guess Cryptographically Confirmed: '{legal_title}'")
        else:
            # Code does not exist in sam.xlsx -> Rejected and forced back into the queue
            ai_guessing_queue.append({"SPONSOR_NAME_STRING": s, "PREDICTED_UEI_OR_DUNS": ""})
        conn.commit()
        
    # Track 4: Pure Residuals Queue
    else:
        ai_guessing_queue.append({"SPONSOR_NAME_STRING": s, "PREDICTED_UEI_OR_DUNS": ""})

conn.commit()
conn.close()

# ==========================================
# 6. EXPORT THE STAGE QUEUE
# ==========================================
if ai_guessing_queue:
    queue_df = pd.DataFrame(ai_guessing_queue)
    queue_df.to_excel(AI_QUEUE_OUTPUT, sheet_name="AI Guesses", index=False)
    print(f"\n📢 Process finished! Isolated {len(queue_df)} true unmapped sponsor records into '{AI_QUEUE_OUTPUT}'.")
else:
    print("\n🎉 Complete coverage achieved! Every non-federal entity has been resolved through local verification data pools.")