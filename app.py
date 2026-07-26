import os
import gc
import json
import time
import shutil
import openpyxl
import pandas as pd
import streamlit as st
from io import BytesIO
from PIL import Image
from google import genai
from google.genai import types
from rapidfuzz import process, utils
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Universal OS | Multi-Store AI SaaS",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CLEAN MODERN SAAS STYLING ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Plus Jakarta Sans', -apple-system, sans-serif !important;
    }

    /* Suppress Input Instructions Overlap */
    div[data-testid="InputInstructions"] {
        display: none !important;
    }

    /* Container Spacing & Borders */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 3rem !important;
    }

    /* Metric Visual Upgrades */
    div[data-testid="stMetric"] {
        background-color: transparent !important;
    }
    div[data-testid="stMetricValue"] div {
        color: #38BDF8 !important;
        font-weight: 700 !important;
    }

    /* Clean Native Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 1px solid #1F2937;
        padding-bottom: 4px;
        margin-bottom: 16px;
    }

    .stTabs [data-baseweb="tab"] {
        height: 40px;
        border-radius: 8px;
        padding: 0 16px;
        font-weight: 600;
        font-size: 13px;
    }

    .stTabs [aria-selected="true"] {
        background-color: #4F46E5 !important;
        color: #FFFFFF !important;
    }
</style>
""", unsafe_allow_html=True)

# --- MULTI-STORE DIRECTORY & DATA ISOLATION ---
DATA_DIR = "stores_data"
os.makedirs(DATA_DIR, exist_ok=True)

def get_store_list():
    stores = [d for d in os.listdir(DATA_DIR) if os.path.isdir(os.path.join(DATA_DIR, d))]
    if not stores:
        default_path = os.path.join(DATA_DIR, "Universal_Hardware")
        os.makedirs(default_path, exist_ok=True)
        return ["Universal_Hardware"]
    return stores

def sanitize_store_name(name):
    return "".join([c if c.isalnum() else "_" for c in name.strip()])

# --- SIDEBAR: MULTI-TENANT WORKSPACE ---
st.sidebar.title("🏬 Store Directory")
existing_stores = get_store_list()

selected_store_slug = st.sidebar.selectbox(
    "Active Store Catalog",
    options=existing_stores,
    format_func=lambda x: x.replace("_", " ")
)

# Store Switch State Purge Guard
if "last_store_slug" not in st.session_state:
    st.session_state["last_store_slug"] = selected_store_slug

if st.session_state["last_store_slug"] != selected_store_slug:
    if "parsed_df" in st.session_state:
        del st.session_state["parsed_df"]
    st.session_state["last_store_slug"] = selected_store_slug
    st.rerun()

st.sidebar.divider()

# --- SIDEBAR STORE ACTIONS ---
side_tab1, side_tab2 = st.sidebar.tabs(["✏️ Rename Store", "➕ Add Store"])

with side_tab1:
    current_display = selected_store_slug.replace("_", " ")
    renamed_input = st.text_input("Current Store Name", value=current_display, key="rename_input_field")
    if st.button("Save New Name", use_container_width=True, type="secondary"):
        if renamed_input.strip() and renamed_input.strip() != current_display:
            new_slug = sanitize_store_name(renamed_input)
            old_path = os.path.join(DATA_DIR, selected_store_slug)
            new_path = os.path.join(DATA_DIR, new_slug)
            
            if not os.path.exists(new_path):
                shutil.move(old_path, new_path)
                st.session_state["last_store_slug"] = new_slug
                st.sidebar.success("Store name updated!")
                st.rerun()
            else:
                st.sidebar.error("A store with that name already exists.")

with side_tab2:
    new_store_name = st.text_input("New Store Title", placeholder="e.g. Metro Hardware", key="add_input_field")
    if st.button("Create Store", use_container_width=True, type="primary"):
        if new_store_name.strip():
            slug = sanitize_store_name(new_store_name)
            new_path = os.path.join(DATA_DIR, slug)
            os.makedirs(new_path, exist_ok=True)
            st.session_state["last_store_slug"] = slug
            st.sidebar.success(f"Store '{new_store_name}' created!")
            st.rerun()

st.sidebar.divider()

# --- API KEY AUTHENTICATION ---
api_key = None
try:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

if not api_key:
    api_key = os.environ.get("GEMINI_API_KEY")

if not api_key:
    st.error("⚠️ Gemini API Key missing. Please configure GEMINI_API_KEY in secrets or environment variables.")
    st.stop()

client = genai.Client(api_key=api_key)

# Active Store File Paths
CURRENT_STORE_DIR = os.path.join(DATA_DIR, selected_store_slug)
MEMORY_FILE = os.path.join(CURRENT_STORE_DIR, "vendor_mappings.json")
MASTER_FILE = os.path.join(CURRENT_STORE_DIR, "inventory_master.csv")
ACTIVE_STORE_DISPLAY = selected_store_slug.replace("_", " ").upper()

# --- HEADER SECTION (CLEAN & COHESIVE) ---
header_left, header_right = st.columns([3, 1])

with header_left:
    st.title("⚡ Universal OS")
    st.caption("Commercial Multi-Store AI Purchase Intake & Inventory Synchronizer")

with header_right:
    st.write("")
    st.info(f"📍 Active: **{ACTIVE_STORE_DISPLAY}**")

st.divider()

# --- STORE DATA LOADERS & PERSISTENCE ---
def load_json_memory():
    if os.path.exists(MEMORY_FILE):
        try:
            with open(MEMORY_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_json_memory(memory_dict):
    try:
        with open(MEMORY_FILE, "w") as f:
            json.dump(memory_dict, f, indent=4)
    except Exception as e:
        st.sidebar.error(f"Memory save alert: {e}")

@st.cache_data
def load_master(store_slug):
    master_path = os.path.join(DATA_DIR, store_slug, "inventory_master.csv")
    try:
        df = pd.read_csv(master_path)
        if "Selling_Price" not in df.columns:
            df["Selling_Price"] = 0.0
        return df
    except Exception:
        return pd.DataFrame({"Official_SKU_Name": [], "Category": [], "Default_Unit": [], "GST_Rate": [], "Selling_Price": []})

def save_master(df, store_slug):
    master_path = os.path.join(DATA_DIR, store_slug, "inventory_master.csv")
    df.to_csv(master_path, index=False)
    st.cache_data.clear()

master_df = load_master(selected_store_slug)
master_sku_list = master_df["Official_SKU_Name"].tolist() if not master_df.empty else []
mapping_memory = load_json_memory()

def match_sku(raw_name):
    cleaned_raw = raw_name.strip().upper()
    if cleaned_raw in mapping_memory:
        return mapping_memory[cleaned_raw], "🧠 Learned Memory"
    if master_sku_list:
        match, score, _ = process.extractOne(raw_name, master_sku_list, processor=utils.default_process)
        if score > 65:
            return match, f"🔍 Fuzzy ({int(score)}%)"
    return raw_name, "⚠️ New SKU"

def get_known_selling_price(sku_name):
    if not master_df.empty and "Selling_Price" in master_df.columns:
        matched = master_df[master_df["Official_SKU_Name"] == sku_name]
        if not matched.empty:
            price = matched.iloc[0]["Selling_Price"]
            if pd.notnull(price) and float(price) > 0:
                return float(price)
    return 0.0

# --- FAIL-SAFE AI ENGINE ---
def is_server_error(exception):
    err_str = str(exception).lower()
    return "503" in err_str or "unavailable" in err_str or "overloaded" in err_str or "429" in err_str or "resourceexhausted" in err_str

@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1.5, min=1, max=5),
    retry=retry_if_exception(is_server_error),
    reraise=True
)
def _call_gemini_with_retry(client, model_name, contents, config):
    return client.models.generate_content(
        model=model_name,
        contents=contents,
        config=config
    )

def extract_invoice_data(image):
    img_copy = image.copy()
    img_copy.thumbnail((1024, 1024))

    prompt = """
    Extract all purchase invoice items into strict JSON:
    {
        "Supplier Company Name": "Vendor Name",
        "Line Items": [
            {
                "Item Name": "description",
                "Quantity": 1.0,
                "Listed Base Rate": 0.0,
                "Listed Total Inclusive Rate": 0.0,
                "GST Rate": 18.0,
                "HSN Code": "",
                "Unit": "PCS"
            }
        ]
    }
    Rules: Rates and GST must be pure numbers. HSN Code as string. Missing values set to 0 or "". Default Unit to PCS or LTR.
    """
    
    config = types.GenerateContentConfig(response_mime_type="application/json")
    contents = [img_copy, prompt]
    candidate_models = ['gemini-3.5-flash-lite', 'gemini-3.5-flash', 'gemini-2.5-flash']
    
    last_error = None
    for model_name in candidate_models:
        try:
            response = _call_gemini_with_retry(client, model_name, contents, config)
            return json.loads(response.text)
        except Exception as e:
            last_error = e
            continue
            
    raise Exception(f"AI Service busy across models: {last_error}")

# --- WORKSPACE TABS ---
tab_parser, tab_master, tab_memory, tab_guide = st.tabs([
    "📥 Batch Invoice Parser", 
    "⚙️ Master Catalog",
    "📋 Vendor Memory", 
    "📖 Operating Guide"
])

# ==========================================
# TAB 1: BATCH INVOICE PARSER
# ==========================================
with tab_parser:
    sm1, sm2, sm3 = st.columns(3)
    with sm1:
        with st.container(border=True):
            st.metric("Master SKUs Registered", len(master_sku_list))
    with sm2:
        with st.container(border=True):
            st.metric("Learned Vendor Rules", len(mapping_memory))
    with sm3:
        with st.container(border=True):
            st.metric("Engine Status", "🟢 Active")

    st.write("")

    col_upload, col_info = st.columns([2.2, 1])
    
    with col_upload:
        with st.container(border=True):
            st.markdown("### 1. Ingestion Dropzone")
            st.caption("Upload purchase bills (PNG, JPG, JPEG) to extract line items.")
            uploaded_files = st.file_uploader(
                "Upload Bills",
                type=["jpg", "jpeg", "png"],
                accept_multiple_files=True,
                label_visibility="collapsed"
            )

    with col_info:
        with st.container(border=True):
            st.markdown("### ⚡ Ingestion Queue")
            if uploaded_files:
                st.success(f"📁 **{len(uploaded_files)} File(s)** Ready")
                st.caption("AI Ready to analyze invoice lines, HSN codes, and tax rates.")
            else:
                st.info("No files in queue.")
                st.caption("Drop purchase invoices to begin structured extraction.")

    if uploaded_files:
        st.write("")
        if st.button("🚀 Process Invoices with AI", type="primary", use_container_width=True):
            if "parsed_df" in st.session_state:
                del st.session_state["parsed_df"]
                
            all_parsed_items = []
            
            with st.status("Analyzing purchase bills with Multimodal AI...", expanded=True) as status_container:
                progress_bar = st.progress(0)
                
                for idx, file in enumerate(uploaded_files):
                    st.write(f"🔍 Processing **{file.name}** ({idx + 1}/{len(uploaded_files)})...")
                    try:
                        file.seek(0)
                        img = Image.open(file)
                        parsed_json = extract_invoice_data(img)
                        supplier = parsed_json.get("Supplier Company Name", "Unknown Supplier")
                        
                        for row in parsed_json.get("Line Items", []):
                            qty = float(row.get("Quantity") or 1.0)
                            gst_rate = float(row.get("GST Rate") or 18.0)
                            base_rate = float(row.get("Listed Base Rate") or 0.0)
                            total_inclusive = float(row.get("Listed Total Inclusive Rate") or 0.0)
                            hsn_sac = str(row.get("HSN Code") or "").strip()
                            
                            # Tax Reverse Math Handling
                            if base_rate > 0:
                                final_base = base_rate
                            elif total_inclusive > 0:
                                final_base = total_inclusive / (1 + (gst_rate / 100))
                            else:
                                final_base = 0.0
                                
                            raw_item_name = str(row.get("Item Name", "")).strip()
                            matched_sku, match_type = match_sku(raw_item_name)
                            
                            known_selling = get_known_selling_price(matched_sku)
                            
                            all_parsed_items.append({
                                "Supplier Name": supplier,
                                "Raw Vendor Item": raw_item_name,
                                "Official SKU": matched_sku,
                                "Match Status": match_type,
                                "Current Quantity": qty,
                                "Unit": str(row.get("Unit", "PCS")).upper(),
                                "HSN/SAC": hsn_sac,
                                "Category": "General",
                                "GST Rate": gst_rate,
                                "Purchase Price": round(final_base, 2),
                                "Selling Price": round(known_selling, 2)
                            })
                    except Exception as e:
                        st.error(f"Error reading {file.name}: {e}")
                        
                    progress_bar.progress((idx + 1) / len(uploaded_files))
                    
                gc.collect()
                status_container.update(label="✅ Batch Processing Complete!", state="complete", expanded=False)
                
            if all_parsed_items:
                st.session_state["parsed_df"] = pd.DataFrame(all_parsed_items)
                st.rerun()

    # --- REVIEW & EDIT WORKSPACE ---
    if "parsed_df" in st.session_state:
        st.divider()
        st.markdown("### 2. Live Inventory Audit Workspace")
        st.caption("Verify AI extraction, mapped SKUs, and rate details before exporting to accounting software.")
        
        df = st.session_state["parsed_df"]
        
        m1, m2, m3 = st.columns(3)
        with m1:
            with st.container(border=True):
                st.metric("Total Line Items", f"{len(df)} Items")
        with m2:
            with st.container(border=True):
                st.metric("Total Stock Quantity", f"{df['Current Quantity'].sum():,.0f} Units")
        with m3:
            with st.container(border=True):
                st.metric("Taxable Purchase Value", f"₹{(df['Purchase Price'] * df['Current Quantity']).sum():,.2f}")
        
        st.write("")
        
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Official SKU": st.column_config.SelectboxColumn("Official SKU Name", options=master_sku_list, required=True) if master_sku_list else "Official SKU",
                "Purchase Price": st.column_config.NumberColumn("Purchase Price (Excl. GST) ₹", format="₹%.2f"),
                "Selling Price": st.column_config.NumberColumn("Selling Price (Optional) ₹", format="₹%.2f"),
                "Current Quantity": st.column_config.NumberColumn("Current Quantity", min_value=0.1),
                "GST Rate": st.column_config.NumberColumn("GST Rate (%)", min_value=0, max_value=28),
                "HSN/SAC": st.column_config.TextColumn("HSN/SAC"),
            }
        )
        
        st.divider()
        if st.button("✅ Confirm Audit & Generate Excel Import File", type="primary", use_container_width=True):
            memory_updated = False
            master_updated = False
            current_master_skus = set(master_sku_list)
            
            for idx, row in edited_df.iterrows():
                raw = str(row["Raw Vendor Item"]).strip().upper()
                official = str(row["Official SKU"]).strip()
                
                # Update vendor memory mapping
                if raw and official and raw != official:
                    mapping_memory[raw] = official
                    memory_updated = True
                    
                # Auto-append new SKUs to Master Catalog if manually typed
                if official and official not in current_master_skus:
                    new_master_row = pd.DataFrame([{
                        "Official_SKU_Name": official,
                        "Category": str(row.get("Category", "General")),
                        "Default_Unit": str(row.get("Unit", "PCS")),
                        "GST_Rate": float(row.get("GST Rate", 18.0)),
                        "Selling_Price": float(row.get("Selling Price", 0.0))
                    }])
                    master_df = pd.concat([master_df, new_master_row], ignore_index=True)
                    current_master_skus.add(official)
                    master_updated = True

            if memory_updated:
                save_json_memory(mapping_memory)
                st.toast("🧠 Learned Vendor Mapping updated!")
                
            if master_updated:
                save_master(master_df, selected_store_slug)
                st.toast("⚙️ Master SKU catalog expanded!")
                
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Items"
            
            ws.append([ACTIVE_STORE_DISPLAY])
            ws.append(["Items"])
            ws.append([f"Generated On: {time.strftime('%d-%m-%Y %H:%M:%S')}"])
            ws.append([]) # Row 4
            
            exact_headers = [
                "S. No.", "Name", "Current Quantity", "Unit", "HSN/SAC",
                "Category", "GST Rate", "Selling Price", "Selling Price (Secondary)",
                "Purchase Price", "Purchase Price (Secondary)", "Secondary Unit", "Ratio"
            ]
            ws.append(exact_headers)
            
            for i, row in edited_df.iterrows():
                selling_val = float(row["Selling Price"]) if row["Selling Price"] > 0 else ""
                ws.append([
                    i + 1,
                    str(row["Official SKU"]),
                    float(row["Current Quantity"]),
                    str(row["Unit"]),
                    str(row["HSN/SAC"]),
                    str(row["Category"]),
                    float(row["GST Rate"]),
                    selling_val,
                    "",
                    float(row["Purchase Price"]),
                    "", "", ""
                ])
                
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            
            st.download_button(
                label=f"📥 Download Bulk Import Spreadsheet for {ACTIVE_STORE_DISPLAY}",
                data=buffer.getvalue(),
                file_name=f"{selected_store_slug}_Inventory_Import.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ==========================================
# TAB 2: STORE MASTER CATALOG MANAGER
# ==========================================
with tab_master:
    st.markdown(f"### ⚙️ Master Inventory Catalog ({ACTIVE_STORE_DISPLAY})")
    st.caption("Define official product SKUs, tax rates, and default units for this store.")
    
    col_add, col_list = st.columns([1, 2])
    
    with col_add:
        with st.container(border=True):
            st.markdown("#### ➕ Add New Master SKU")
            add_sku = st.text_input("SKU Name (e.g. Copper Wire 1.5mm)")
            add_cat = st.text_input("Category", value="General")
            add_unit = st.selectbox("Default Unit", options=["PCS", "BOX", "LTR", "KG", "NOS", "SET"])
            add_gst = st.selectbox("GST Rate (%)", options=[0, 5, 12, 18, 28], index=3)
            add_price = st.number_input("Selling Price ₹ (Optional)", min_value=0.0, step=10.0)
            
            if st.button("Save SKU to Catalog", use_container_width=True, type="primary"):
                if add_sku.strip():
                    clean_sku = add_sku.strip()
                    if clean_sku not in master_sku_list:
                        new_row = pd.DataFrame([{
                            "Official_SKU_Name": clean_sku,
                            "Category": add_cat,
                            "Default_Unit": add_unit,
                            "GST_Rate": add_gst,
                            "Selling_Price": add_price
                        }])
                        updated = pd.concat([master_df, new_row], ignore_index=True)
                        save_master(updated, selected_store_slug)
                        st.success(f"Added '{clean_sku}'!")
                        st.rerun()
                    else:
                        st.warning("SKU already exists.")
                    
    with col_list:
        with st.container(border=True):
            st.markdown("#### 📋 Catalog Register")
            if not master_df.empty:
                st.dataframe(master_df, use_container_width=True)
                st.write("")
                with st.expander("🗑️ Delete Catalog SKU"):
                    sku_to_delete = st.selectbox("Select SKU to Remove:", options=["-- None --"] + master_sku_list)
                    if st.button("Delete Selected SKU", use_container_width=True):
                        if sku_to_delete != "-- None --":
                            updated = master_df[master_df["Official_SKU_Name"] != sku_to_delete]
                            save_master(updated, selected_store_slug)
                            st.success(f"Removed '{sku_to_delete}'!")
                            st.rerun()
            else:
                st.info("Master catalog for this store is currently empty.")

# ==========================================
# TAB 3: VENDOR SKU MEMORY WORKSPACE
# ==========================================
with tab_memory:
    st.markdown(f"### 🧠 Learned AI Vendor Memory ({ACTIVE_STORE_DISPLAY})")
    st.caption("AI remembers how vendor-specific invoice descriptions map to your store SKUs.")
    
    if mapping_memory:
        mem_df = pd.DataFrame([
            {"Raw Vendor Item Description": k, "Mapped Store SKU": v}
            for k, v in mapping_memory.items()
        ])
        st.dataframe(mem_df, use_container_width=True)
        
        st.write("")
        if st.button("🗑️ Reset Store Memory Cache"):
            save_json_memory({})
            st.success("Memory cache reset!")
            st.rerun()
    else:
        st.info("No learned vendor mappings recorded yet for this store location.")

# ==========================================
# TAB 4: IMPORT GUIDE
# ==========================================
with tab_guide:
    st.markdown("### 📖 Standard Operating Procedure")
    with st.container(border=True):
        st.markdown("""
        ### How to Process & Sync Invoices:
        1. **Select Store:** Choose your active store location in the left sidebar directory.
        2. **Upload Bills:** Drop one or multiple purchase invoice photos in **Tab 1**.
        3. **Run AI Engine:** Click **Run AI Invoice Parsing Engine** to extract structured line items.
        4. **Audit Workspace:** Check quantities, HSN codes, purchase rates, and mapped SKUs.
        5. **Download Import File:** Generate the `.xlsx` spreadsheet.
        6. **Import to ERP:** Open your accounting or ERP software → **Items / Inventory** → **Bulk Import**, upload the `.xlsx` file.
        """)
