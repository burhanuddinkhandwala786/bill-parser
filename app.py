import os
import gc
import json
import time
import openpyxl
import pandas as pd
import streamlit as st
from io import BytesIO
from PIL import Image
from google import genai
from google.genai import types
from rapidfuzz import process, utils
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception

# --- ENTERPRISE SAAS PAGE CONFIGURATION & LUXURY LIGHT STYLING ---
st.set_page_config(
    page_title="Universal OS | Multi-Store AI Ingestion SaaS",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Executive Light Theme CSS System (Apple / macOS Studio Aesthetic)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    
    /* Global Reset & Base Canvas */
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
        color: #0F172A !important;
        -webkit-font-smoothing: antialiased;
    }
    
    .stApp {
        background-color: #F8FAFC !important;
    }

    /* Suppress Native Streamlit Header Overlaps */
    header[data-testid="stHeader"] {
        display: none !important;
    }
    
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2.5rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        max-width: 100% !important;
    }
    
    /* Hero Banner Styling - Sleek Metallic Glass */
    .saas-header {
        background: linear-gradient(135deg, #FFFFFF 0%, #F1F5F9 100%);
        padding: 24px 32px;
        border-radius: 20px;
        color: #0F172A;
        margin-bottom: 24px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.05);
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .saas-title-box {
        display: flex;
        flex-direction: column;
    }
    .saas-title {
        font-size: 26px;
        font-weight: 700;
        letter-spacing: -0.5px;
        margin: 0;
        color: #0F172A;
    }
    .saas-subtitle {
        font-size: 14px;
        color: #64748B;
        margin-top: 4px;
        margin-bottom: 0;
    }
    .badge-container {
        display: flex;
        flex-direction: column;
        align-items: flex-end;
        gap: 8px;
    }
    .status-badge {
        display: inline-block;
        background-color: #DCFCE7;
        color: #15803D;
        font-size: 11px;
        font-weight: 700;
        padding: 6px 14px;
        border-radius: 20px;
        border: 1px solid #BBF7D0;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .store-badge {
        display: inline-block;
        background-color: #EEF2FF;
        color: #4338CA;
        font-size: 12px;
        font-weight: 600;
        padding: 6px 14px;
        border-radius: 10px;
        border: 1px solid #C7D2FE;
    }

    /* Soft Glass File Dropzone */
    div[data-testid="stFileUploader"] {
        background: #FFFFFF !important;
        border: 1.5px dashed #CBD5E1 !important;
        border-radius: 16px !important;
        padding: 16px !important;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.03) !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #4338CA !important;
        background: #F8FAFC !important;
        box-shadow: 0 6px 16px rgba(67, 56, 202, 0.08) !important;
    }
    div[data-testid="stFileUploader"] section {
        background-color: transparent !important;
    }
    
    /* Metrics Visual Cards */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 16px !important;
        padding: 18px 20px !important;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.03) !important;
    }
    div[data-testid="stMetricLabel"] p {
        font-size: 11px !important;
        font-weight: 700 !important;
        color: #64748B !important;
        text-transform: uppercase !important;
        letter-spacing: 0.6px !important;
    }
    div[data-testid="stMetricValue"] div {
        font-size: 22px !important;
        font-weight: 700 !important;
        color: #0284C7 !important;
    }

    /* Executive Tabs Override */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: transparent;
        border-bottom: 1px solid #E2E8F0;
        padding-bottom: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 42px;
        border-radius: 12px;
        padding: 0 20px;
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        color: #64748B;
        font-weight: 600;
        font-size: 13px;
        transition: all 0.2s ease;
    }
    .stTabs [aria-selected="true"] {
        background-color: #4338CA !important;
        color: #FFFFFF !important;
        border: 1px solid #4338CA !important;
        box-shadow: 0 4px 12px rgba(67, 56, 202, 0.25) !important;
    }

    /* High-Contrast Glass Data Editor Grid */
    div[data-testid="stDataEditor"] {
        border: 1px solid #E2E8F0 !important;
        border-radius: 16px !important;
        overflow: hidden !important;
        background-color: #FFFFFF !important;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04) !important;
    }

    /* Clean Touch-Friendly Buttons */
    .stButton>button, .stDownloadButton>button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        padding: 10px 20px !important;
        border: 1px solid #CBD5E1 !important;
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04) !important;
        transition: all 0.2s ease-in-out !important;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        border-color: #4338CA !important;
        color: #4338CA !important;
        background-color: #F8FAFC !important;
    }
    button[kind="primary"] {
        background: linear-gradient(135deg, #4338CA 0%, #3730A3 100%) !important;
        border: none !important;
        color: #FFFFFF !important;
        box-shadow: 0 4px 14px rgba(67, 56, 202, 0.3) !important;
    }
    button[kind="primary"]:hover {
        background: linear-gradient(135deg, #3730A3 0%, #312E81 100%) !important;
        box-shadow: 0 6px 18px rgba(67, 56, 202, 0.4) !important;
    }

    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0 !important;
    }

    /* Card Panels */
    .info-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 20px;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.02);
        margin-bottom: 16px;
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

# --- SIDEBAR: MULTI-TENANT SWITCHER ---
st.sidebar.markdown("## 🏬 Store Management")
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

# Register New Store
with st.sidebar.expander("➕ Register New Store", expanded=False):
    new_store_name = st.text_input("New Store Name:")
    if st.button("Create Store Environment", use_container_width=True):
        if new_store_name.strip():
            slug = sanitize_store_name(new_store_name)
            new_path = os.path.join(DATA_DIR, slug)
            os.makedirs(new_path, exist_ok=True)
            st.success(f"Store '{new_store_name}' initialized!")
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 🔑 API Configuration")

api_key = None
try:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

if not api_key:
    api_key = os.environ.get("GEMINI_API_KEY")

if not api_key:
    api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

if not api_key:
    st.info("👋 Welcome! Enter your Gemini API Key in the sidebar or Streamlit Secrets to begin.")
    st.stop()

client = genai.Client(api_key=api_key)

# Active Store File Paths
CURRENT_STORE_DIR = os.path.join(DATA_DIR, selected_store_slug)
MEMORY_FILE = os.path.join(CURRENT_STORE_DIR, "vendor_mappings.json")
MASTER_FILE = os.path.join(CURRENT_STORE_DIR, "inventory_master.csv")
ACTIVE_STORE_DISPLAY = selected_store_slug.replace("_", " ").upper()

# Render Modernized SaaS Header
st.markdown(f"""
<div class="saas-header">
    <div class="saas-title-box">
        <p class="saas-title">⚡ Universal OS — AI Intake SaaS</p>
        <p class="saas-subtitle">Multi-Store Purchase Ingestion & Bulk Inventory Synchronizer</p>
    </div>
    <div class="badge-container">
        <span class="status-badge">🟢 Commercial SaaS Active</span>
        <span class="store-badge">📍 Active Store: <b>{ACTIVE_STORE_DISPLAY}</b></span>
    </div>
</div>
""", unsafe_allow_html=True)

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
    "⚙️ Store Master Catalog",
    "📋 Vendor SKU Memory", 
    "📖 Guide"
])

# ==========================================
# TAB 1: BATCH INVOICE PARSER
# ==========================================
with tab_parser:
    col_upload, col_info = st.columns([2, 1])
    
    with col_upload:
        st.markdown("### 1. Multi-Bill Image Intake")
        uploaded_files = st.file_uploader(
            f"Upload Purchase Bills for {ACTIVE_STORE_DISPLAY} (PNG, JPG, JPEG)",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True
        )

    with col_info:
        st.markdown("### ⚡ Batch Status")
        if uploaded_files:
            st.success(f"📁 {len(uploaded_files)} File(s) Ready for Ingestion")
        else:
            st.info("No files uploaded.")

    if uploaded_files:
        if st.button("🚀 Process & Parse Invoices with AI", type="primary", use_container_width=True):
            if "parsed_df" in st.session_state:
                del st.session_state["parsed_df"]
                
            all_parsed_items = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, file in enumerate(uploaded_files):
                status_text.caption(f"Parsing bill {idx + 1} of {len(uploaded_files)}: **{file.name}**...")
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
                
            status_text.empty()
            gc.collect()
                
            if all_parsed_items:
                st.session_state["parsed_df"] = pd.DataFrame(all_parsed_items)
                st.rerun()

    # --- REVIEW & EDIT WORKSPACE ---
    if "parsed_df" in st.session_state:
        st.write("---")
        st.markdown("### 2. Live Audit Workspace")
        st.caption("💡 Map raw vendor items to official SKUs. Selling Price defaults to 0.0 unless manually specified.")
        
        df = st.session_state["parsed_df"]
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Line Items", len(df))
        m2.metric("Total Units", f"{df['Current Quantity'].sum():,.0f}")
        m3.metric("Taxable Purchase Cost", f"₹{(df['Purchase Price'] * df['Current Quantity']).sum():,.2f}")
        
        st.write(" ")
        
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
        
        st.write("---")
        if st.button("✅ Generate & Download Excel File", type="primary", use_container_width=True):
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
                st.toast("🧠 Saved vendor mapping to learned memory!")
                
            if master_updated:
                save_master(master_df, selected_store_slug)
                st.toast("⚙️ New SKUs added to Master Catalog!")
                
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
                label=f"📥 Download File for {ACTIVE_STORE_DISPLAY}",
                data=buffer.getvalue(),
                file_name=f"{selected_store_slug}_Items_Import.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ==========================================
# TAB 2: STORE MASTER CATALOG MANAGER
# ==========================================
with tab_master:
    st.markdown(f"### ⚙️ Master Inventory Catalog ({ACTIVE_STORE_DISPLAY})")
    st.caption("Manage official SKUs for this store catalog.")
    
    col_add, col_list = st.columns([1, 2])
    
    with col_add:
        st.markdown('<div class="info-card">', unsafe_allow_html=True)
        st.markdown("#### ➕ Add New SKU")
        add_sku = st.text_input("SKU Name")
        add_cat = st.text_input("Category", value="General")
        add_unit = st.selectbox("Default Unit", options=["PCS", "BOX", "LTR", "KG", "NOS", "SET"])
        add_gst = st.selectbox("GST Rate (%)", options=[0, 5, 12, 18, 28], index=3)
        add_price = st.number_input("Selling Price ₹ (Optional)", min_value=0.0, step=10.0)
        
        if st.button("Save SKU to Master List", use_container_width=True, type="primary"):
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
        st.markdown('</div>', unsafe_allow_html=True)
                    
    with col_list:
        st.markdown("#### 📋 Store SKUs")
        if not master_df.empty:
            st.dataframe(master_df, use_container_width=True)
            st.write(" ")
            with st.expander("🗑️ Delete Existing SKU"):
                sku_to_delete = st.selectbox("Select SKU to Remove:", options=["-- None --"] + master_sku_list)
                if st.button("Delete Selected SKU", use_container_width=True):
                    if sku_to_delete != "-- None --":
                        updated = master_df[master_df["Official_SKU_Name"] != sku_to_delete]
                        save_master(updated, selected_store_slug)
                        st.success(f"Removed '{sku_to_delete}'!")
                        st.rerun()
        else:
            st.info("Master catalog for this store is empty.")

# ==========================================
# TAB 3: VENDOR SKU MEMORY WORKSPACE
# ==========================================
with tab_memory:
    st.markdown(f"### 🧠 Learned AI Vendor Memory ({ACTIVE_STORE_DISPLAY})")
    st.caption("Maps vendor bill descriptions to your store SKUs.")
    
    if mapping_memory:
        mem_df = pd.DataFrame([
            {"Raw Vendor Item Name": k, "Mapped Store SKU": v}
            for k, v in mapping_memory.items()
        ])
        st.dataframe(mem_df, use_container_width=True)
        
        st.write(" ")
        if st.button("🗑️ Clear Store Memory Cache"):
            save_json_memory({})
            st.success("Memory cleared!")
            st.rerun()
    else:
        st.info("No learned vendor mappings yet for this store.")

# ==========================================
# TAB 4: IMPORT GUIDE
# ==========================================
with tab_guide:
    st.markdown("### 📖 Import Guide")
    st.markdown("""
    <div class="info-card">
        <ol style="line-height: 1.8; margin-bottom: 0;">
            <li><b>Select Store:</b> Choose your active store in the sidebar or register a new one.</li>
            <li><b>Upload Bills:</b> Upload purchase invoice photos in <b>Tab 1</b> and click <b>Process Invoices</b>.</li>
            <li><b>Review & Audit:</b> Confirm quantities, HSN codes, and mapped SKUs.</li>
            <li><b>Download Excel:</b> Click <b>Download File</b> to generate your spreadsheet.</li>
            <li><b>Upload to ERP/Accounting:</b> Open your software → <b>Items / Inventory</b> → <b>Bulk Import</b>, select the generated <code>.xlsx</code> file, and upload.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)
