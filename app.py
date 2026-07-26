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

# --- ENTERPRISE SAAS PAGE CONFIGURATION & DESIGN SYSTEM ---
st.set_page_config(
    page_title="Universal OS | Enterprise Intake Workspace",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# High-End Dark Executive Design System
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    .stApp {
        background-color: #080C14;
        color: #E2E8F0;
    }
    
    /* Top Minimal Command Bar */
    .command-bar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: rgba(15, 23, 42, 0.75);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        padding: 12px 24px;
        border-radius: 12px;
        margin-bottom: 20px;
    }
    .brand-logo {
        font-size: 18px;
        font-weight: 700;
        letter-spacing: -0.4px;
        color: #F8FAFC;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .brand-badge {
        background: rgba(99, 102, 241, 0.15);
        color: #818CF8;
        border: 1px solid rgba(99, 102, 241, 0.3);
        font-size: 11px;
        font-weight: 600;
        padding: 2px 8px;
        border-radius: 6px;
    }
    .store-pill {
        background: #1E293B;
        border: 1px solid #334155;
        color: #38BDF8;
        font-size: 12px;
        font-weight: 600;
        padding: 4px 12px;
        border-radius: 20px;
    }
    
    /* Dashboard Surface Cards */
    .saas-card {
        background: #0F172A;
        border: 1px solid rgba(255, 255, 255, 0.07);
        border-radius: 14px;
        padding: 20px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.4);
        margin-bottom: 20px;
    }
    
    /* Sleek Metric Cards */
    div[data-testid="stMetric"] {
        background: #131D31 !important;
        border: 1px solid rgba(255, 255, 255, 0.06) !important;
        border-radius: 10px !important;
        padding: 14px !important;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 12px !important;
        font-weight: 600 !important;
        color: #94A3B8 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    div[data-testid="stMetricValue"] {
        font-size: 22px !important;
        font-weight: 700 !important;
        color: #38BDF8 !important;
    }

    /* Tab Custom Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: transparent;
        border-bottom: 1px solid rgba(255, 255, 255, 0.08);
        padding-bottom: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        border-radius: 8px;
        padding: 0 16px;
        background-color: #0F172A;
        border: 1px solid rgba(255, 255, 255, 0.05);
        color: #94A3B8;
        font-weight: 500;
        font-size: 13px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #6366F1 !important;
        color: #FFFFFF !important;
        border: none !important;
    }

    /* Clean Buttons */
    .stButton>button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        transition: all 0.2s ease;
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
st.sidebar.markdown("## 🏬 Store Manager")
existing_stores = get_store_list()

selected_store_slug = st.sidebar.selectbox(
    "Active Store Catalog",
    options=existing_stores,
    format_func=lambda x: x.replace("_", " ")
)

if "last_store_slug" not in st.session_state:
    st.session_state["last_store_slug"] = selected_store_slug

if st.session_state["last_store_slug"] != selected_store_slug:
    if "parsed_df" in st.session_state:
        del st.session_state["parsed_df"]
    st.session_state["last_store_slug"] = selected_store_slug
    st.rerun()

with st.sidebar.expander("➕ Register New Store"):
    new_store_name = st.text_input("Store Name:")
    if st.button("Create Store Environment"):
        if new_store_name.strip():
            slug = sanitize_store_name(new_store_name)
            os.makedirs(os.path.join(DATA_DIR, slug), exist_ok=True)
            st.success(f"Initialized '{new_store_name}'!")
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 🔑 API Access")

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
    st.info("👋 Enter your Gemini API Key in the sidebar or Streamlit Secrets to activate workspace.")
    st.stop()

client = genai.Client(api_key=api_key)

CURRENT_STORE_DIR = os.path.join(DATA_DIR, selected_store_slug)
MEMORY_FILE = os.path.join(CURRENT_STORE_DIR, "vendor_mappings.json")
MASTER_FILE = os.path.join(CURRENT_STORE_DIR, "inventory_master.csv")
ACTIVE_STORE_DISPLAY = selected_store_slug.replace("_", " ").upper()

# --- TOP MINIMAL COMMAND BAR ---
st.markdown(f"""
<div class="command-bar">
    <div class="brand-logo">
        ⚡ Universal OS <span class="brand-badge">PRO INTAKE</span>
    </div>
    <div class="store-pill">
        📍 Catalog: <b>{ACTIVE_STORE_DISPLAY}</b>
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
    "⚡ Batch Intake Workspace", 
    "⚙️ Master Catalog", 
    "🧠 AI Memory", 
    "📖 Integration Guide"
])

# ==========================================
# TAB 1: BATCH INGESTION WORKSPACE
# ==========================================
with tab_parser:
    col_intake, col_stats = st.columns([1.8, 1.2])
    
    with col_intake:
        st.markdown("#### 1. Ingestion Control")
        uploaded_files = st.file_uploader(
            "Drop purchase invoice photos (PNG, JPG, JPEG)",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True,
            label_visibility="collapsed"
        )
        
        if uploaded_files:
            if st.button("🚀 Process & Extract Batch Data", type="primary", use_container_width=True):
                if "parsed_df" in st.session_state:
                    del st.session_state["parsed_df"]
                    
                all_parsed_items = []
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for idx, file in enumerate(uploaded_files):
                    status_text.caption(f"Extracting {idx + 1}/{len(uploaded_files)}: **{file.name}**...")
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

    with col_stats:
        st.markdown("#### 2. Batch Health Overview")
        if "parsed_df" in st.session_state:
            df_temp = st.session_state["parsed_df"]
            m1, m2 = st.columns(2)
            m1.metric("Total Line Items", len(df_temp))
            m2.metric("Total Units", f"{df_temp['Current Quantity'].sum():,.0f}")
            st.metric("Taxable Purchase Cost", f"₹{(df_temp['Purchase Price'] * df_temp['Current Quantity']).sum():,.2f}")
        elif uploaded_files:
            st.info(f"📁 {len(uploaded_files)} File(s) staged. Click 'Process & Extract' to begin.")
        else:
            st.caption("Upload invoice photos on the left to initialize live audit workspace.")

    # --- AUDIT WORKSPACE GRID ---
    if "parsed_df" in st.session_state:
        st.markdown("---")
        st.markdown("#### 3. Data Review & Audit Workspace")
        st.caption("Double-click cells to adjust items, quantities, or tax values. New SKUs manually entered auto-save to catalog.")
        
        df = st.session_state["parsed_df"]
        
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
        
        st.write(" ")
        col_space, col_export = st.columns([2, 1])
        
        with col_export:
            if st.button("✅ Generate Excel Import", type="primary", use_container_width=True):
                memory_updated = False
                master_updated = False
                current_master_skus = set(master_sku_list)
                
                for idx, row in edited_df.iterrows():
                    raw = str(row["Raw Vendor Item"]).strip().upper()
                    official = str(row["Official SKU"]).strip()
                    
                    if raw and official and raw != official:
                        mapping_memory[raw] = official
                        memory_updated = True
                        
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
                    st.toast("🧠 Saved mapping to AI Memory!")
                    
                if master_updated:
                    save_master(master_df, selected_store_slug)
                    st.toast("⚙️ Updated Master Catalog!")
                    
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Items"
                
                ws.append([ACTIVE_STORE_DISPLAY])
                ws.append(["Items"])
                ws.append([f"Generated On: {time.strftime('%d-%m-%Y %H:%M:%S')}"])
                ws.append([]) # Empty Row 4
                
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
                    label=f"📥 Download Import File ({ACTIVE_STORE_DISPLAY})",
                    data=buffer.getvalue(),
                    file_name=f"{selected_store_slug}_Items_Import.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# ==========================================
# TAB 2: STORE MASTER CATALOG MANAGER
# ==========================================
with tab_master:
    st.markdown(f"#### Master Inventory Catalog ({ACTIVE_STORE_DISPLAY})")
    
    col_add, col_list = st.columns([1, 2])
    
    with col_add:
        st.markdown("##### ➕ Add Official SKU")
        add_sku = st.text_input("SKU Name")
        add_cat = st.text_input("Category", value="General")
        add_unit = st.selectbox("Default Unit", options=["PCS", "BOX", "LTR", "KG", "NOS", "SET"])
        add_gst = st.selectbox("GST Rate (%)", options=[0, 5, 12, 18, 28], index=3)
        add_price = st.number_input("Selling Price ₹ (Optional)", min_value=0.0, step=10.0)
        
        if st.button("Save SKU to Master"):
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
                    st.warning("SKU exists in master list.")
                    
    with col_list:
        st.markdown("##### 📋 Store Items")
        if not master_df.empty:
            st.dataframe(master_df, use_container_width=True)
            sku_to_delete = st.selectbox("Select SKU to Remove:", options=["-- None --"] + master_sku_list)
            if st.button("🗑️ Delete Selected SKU"):
                if sku_to_delete != "-- None --":
                    updated = master_df[master_df["Official_SKU_Name"] != sku_to_delete]
                    save_master(updated, selected_store_slug)
                    st.success(f"Removed '{sku_to_delete}'!")
                    st.rerun()
        else:
            st.info("Master catalog is empty.")

# ==========================================
# TAB 3: VENDOR SKU MEMORY WORKSPACE
# ==========================================
with tab_memory:
    st.markdown(f"#### Learned AI Mappings ({ACTIVE_STORE_DISPLAY})")
    
    if mapping_memory:
        mem_df = pd.DataFrame([
            {"Raw Vendor Item Name": k, "Mapped Store SKU": v}
            for k, v in mapping_memory.items()
        ])
        st.dataframe(mem_df, use_container_width=True)
        
        if st.button("🗑️ Clear Store Memory Cache"):
            save_json_memory({})
            st.success("Memory cleared!")
            st.rerun()
    else:
        st.info("No learned vendor mappings yet.")

# ==========================================
# TAB 4: IMPORT GUIDE
# ==========================================
with tab_guide:
    st.markdown("#### Accountune Integration Steps")
    st.markdown("""
    1. **Select Catalog:** Choose active store in sidebar.
    2. **Drop Bills:** Upload purchase photos in **Tab 1** and click **Process & Extract Batch Data**.
    3. **Review & Audit:** Confirm quantities, HSN codes, and mapped SKUs in the grid.
    4. **Download:** Click **Generate Accountune Excel Import**.
    5. **Upload:** Open Accountune $\rightarrow$ **Items / Inventory** $\rightarrow$ **Bulk Import**, select the generated `.xlsx` file, and upload!
    """)
