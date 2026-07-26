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

# --- ENTERPRISE PAGE CONFIGURATION & STYLING ---
st.set_page_config(page_title="Universal OS | Enterprise Intake", page_icon="⚡", layout="wide")

st.markdown("""
<style>
    .main-title {
        font-size: 26px;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 2px;
    }
    .sub-title {
        font-size: 14px;
        color: #64748B;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">⚡ AI Invoice-to-Excel Bulk Stock Engine</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Upload purchase bills → Audit & Review Rates → Export 100% Compliant Excel sheet for instant bulk import.</div>', unsafe_allow_html=True)

# --- SECURE API KEY INITIALIZATION ---
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
    st.warning("⚠️ API Key Missing! Enter your Gemini API Key in the sidebar or Streamlit Secrets.")
    st.stop()

client = genai.Client(api_key=api_key)

# --- MEMORY & MASTER INVENTORY FILES ---
MEMORY_FILE = "vendor_mappings.json"
MASTER_FILE = "inventory_master.csv"

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
def load_master():
    try:
        return pd.read_csv(MASTER_FILE)
    except Exception:
        return pd.DataFrame({"Official_SKU_Name": [], "Category": [], "Default_Unit": [], "GST_Rate": []})

master_df = load_master()
master_sku_list = master_df["Official_SKU_Name"].tolist() if not master_df.empty else []
mapping_memory = load_json_memory()

def match_sku(raw_name):
    """Checks learned vendor memory first, then applies fuzzy matching against official master list."""
    cleaned_raw = raw_name.strip().upper()
    
    # 1. Exact Learned Memory Match
    if cleaned_raw in mapping_memory:
        return mapping_memory[cleaned_raw], "🧠 Learned Memory"
    
    # 2. Fuzzy Match against Master List
    if master_sku_list:
        match, score, _ = process.extractOne(raw_name, master_sku_list, processor=utils.default_process)
        if score > 65:
            return match, f"🔍 Fuzzy ({int(score)}%)"
            
    return raw_name, "⚠️ New SKU"

# --- FAIL-SAFE HIGH-SPEED AI ENGINE WITH RETRIES & MODEL CASCADE ---
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
    # Downscale image in-memory for 5x faster network transmission and lower RAM usage
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

# --- WORKSPACE UI: STEP 1 INTAKE ---
st.subheader("1. Multi-Bill Intake")
uploaded_files = st.file_uploader(
    "Upload Supplier Purchase Invoices (Select One or Multiple PNG/JPG Files)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🚀 Process & Parse Bills", type="primary"):
        all_parsed_items = []
        progress_bar = st.progress(0)
        
        for idx, file in enumerate(uploaded_files):
            try:
                img = Image.open(file)
                parsed_json = extract_invoice_data(img)
                supplier = parsed_json.get("Supplier Company Name", "Unknown Supplier")
                
                for row in parsed_json.get("Line Items", []):
                    qty = float(row.get("Quantity") or 1.0)
                    gst_rate = float(row.get("GST Rate") or 18.0)
                    base_rate = float(row.get("Listed Base Rate") or 0.0)
                    total_inclusive = float(row.get("Listed Total Inclusive Rate") or 0.0)
                    hsn_sac = str(row.get("HSN Code") or "").strip()
                    
                    # Tax computation
                    if base_rate > 0:
                        final_base = base_rate
                        final_inclusive = base_rate * (1 + (gst_rate / 100))
                    elif total_inclusive > 0:
                        final_inclusive = total_inclusive
                        final_base = total_inclusive / (1 + (gst_rate / 100))
                    else:
                        final_base = 0.0
                        final_inclusive = 0.0
                        
                    raw_item_name = str(row.get("Item Name", "")).strip()
                    matched_sku, match_type = match_sku(raw_item_name)
                    suggested_sale = final_inclusive * 1.25 # Default 25% margin
                    
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
                        "Selling Price": round(suggested_sale, 2),
                    })
            except Exception as e:
                st.error(f"Error reading {file.name}: {e}")
                
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        # Free memory RAM after processing batch
        gc.collect()
            
        if all_parsed_items:
            st.session_state["parsed_df"] = pd.DataFrame(all_parsed_items)
            st.success("✅ Extraction complete! Review, edit, or adjust items below.")

# --- WORKSPACE UI: STEP 2 REVIEW & ACCOUNTUNE EXPORT ---
if "parsed_df" in st.session_state:
    st.write("---")
    st.subheader("2. Review & Adjust Before Export")
    st.info("💡 Double-click any cell below to edit item names, prices, quantities, HSN codes, or GST %. Learned SKU memory updates automatically!")
    
    df = st.session_state["parsed_df"]
    
    # Live Enterprise Summary Metric Cards
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Line Items", len(df))
    m2.metric("Total Bill Quantity", f"{df['Current Quantity'].sum():,.0f}")
    m3.metric("Total Taxable Cost", f"₹{df['Purchase Price'].sum():,.2f}")
    m4.metric("Est. Total Selling Value", f"₹{df['Selling Price'].sum():,.2f}")
    
    st.write(" ")
    
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Official SKU": st.column_config.SelectboxColumn("Official SKU Name", options=master_sku_list, required=True) if master_sku_list else "Official SKU",
            "Purchase Price": st.column_config.NumberColumn("Purchase Price (₹)", format="₹%.2f"),
            "Selling Price": st.column_config.NumberColumn("Selling Price (₹)", format="₹%.2f"),
            "Current Quantity": st.column_config.NumberColumn("Current Quantity", min_value=0.1),
            "GST Rate": st.column_config.NumberColumn("GST Rate (%)", min_value=0, max_value=28),
            "HSN/SAC": st.column_config.TextColumn("HSN/SAC"),
        }
    )
    
    st.write("---")
    st.markdown("### 📥 Generate Accountune Bulk Import File")
    
    if st.button("✅ Download Exact App-Compliant Import File", type="primary"):
        # Update learned memory mappings
        memory_updated = False
        for idx, row in edited_df.iterrows():
            raw = str(row["Raw Vendor Item"]).strip().upper()
            official = str(row["Official SKU"]).strip()
            if raw and official and raw != official:
                mapping_memory[raw] = official
                memory_updated = True
                
        if memory_updated:
            save_json_memory(mapping_memory)
            st.toast("🧠 Saved vendor SKU mapping to memory!")
            
        # Build openpyxl workbook matching the exact Accountune template layout
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Items"
        
        # Header Metadata Rows (Matching Accountune Template Rows 1 to 3)
        ws.append(["UNIVERSAL HARDWARE AND PLYWOOD STORE"])
        ws.append(["Items"])
        ws.append([f"Generated On: {time.strftime('%d-%m-%Y %H:%M:%S')}"])
        ws.append([]) # Empty Row 4
        
        # Row 5: Exact 13 Headers required by Accountune
        exact_headers = [
            "S. No.", "Name", "Current Quantity", "Unit", "HSN/SAC",
            "Category", "GST Rate", "Selling Price", "Selling Price (Secondary)",
            "Purchase Price", "Purchase Price (Secondary)", "Secondary Unit", "Ratio"
        ]
        ws.append(exact_headers)
        
        # Row 6+: Data Rows
        for i, row in edited_df.iterrows():
            ws.append([
                i + 1,                              # S. No.
                str(row["Official SKU"]),           # Name
                float(row["Current Quantity"]),     # Current Quantity
                str(row["Unit"]),                   # Unit
                str(row["HSN/SAC"]),                # HSN/SAC
                str(row["Category"]),               # Category
                float(row["GST Rate"]),             # GST Rate
                float(row["Selling Price"]),        # Selling Price
                "",                                 # Selling Price (Secondary)
                float(row["Purchase Price"]),       # Purchase Price
                "",                                 # Purchase Price (Secondary)
                "",                                 # Secondary Unit
                ""                                  # Ratio
            ])
            
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            label="📥 Download Compliant Excel File (Universal_Items_Import.xlsx)",
            data=buffer.getvalue(),
            file_name="Universal_Items_Import.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
