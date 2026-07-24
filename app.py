import os
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

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Enterprise Profit Engine", page_icon="📈", layout="wide")
st.title("📈 Universal Hardware & Paints - Profit & Inventory Engine")
st.caption("Military-Grade Invoice Audit, Dynamic Margin Engine & myBillBook Sync")

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

# --- CONFIGURATION & MEMORY FILES ---
MEMORY_FILE = "vendor_mappings.json"
MASTER_FILE = "inventory_master.csv"
MARGINS_FILE = "category_margins.json"

# Default category profit margin targets (%)
DEFAULT_MARGINS = {
    "Paints": 18.0,
    "Waterproofing": 20.0,
    "Hardware": 30.0,
    "Locks": 35.0,
    "Plywood": 15.0,
    "General": 25.0
}

def load_json_file(file_path, default_data):
    if os.path.exists(file_path):
        try:
            with open(file_path, "r") as f:
                return json.load(f)
        except Exception:
            return default_data
    return default_data

def save_json_file(file_path, data_dict):
    with open(file_path, "w") as f:
        json.dump(data_dict, f, indent=4)

mapping_memory = load_json_file(MEMORY_FILE, {})
category_margins = load_json_file(MARGINS_FILE, DEFAULT_MARGINS)

@st.cache_data
def load_master():
    try:
        df = pd.read_csv(MASTER_FILE)
        return df
    except Exception:
        return pd.DataFrame({"Official_SKU_Name": [], "Category": [], "Default_Unit": [], "GST_Rate": []})

master_df = load_master()
master_sku_list = master_df["Official_SKU_Name"].tolist() if not master_df.empty else []

def get_category_for_sku(sku_name):
    """Retrieves category from master inventory list."""
    if not master_df.empty and "Official_SKU_Name" in master_df.columns:
        match = master_df[master_df["Official_SKU_Name"] == sku_name]
        if not match.empty and "Category" in match.columns:
            return str(match["Category"].iloc[0])
    return "General"

def match_sku(raw_name):
    """Smart Matching Pipeline: Check Memory -> Fuzzy Search -> Raw Default."""
    cleaned_raw = raw_name.strip().upper()
    if cleaned_raw in mapping_memory:
        return mapping_memory[cleaned_raw], "🧠 Learned Memory"
    
    if master_sku_list:
        match, score, _ = process.extractOne(raw_name, master_sku_list, processor=utils.default_process)
        if score > 65:
            return match, f"🔍 Fuzzy ({int(score)}%)"
            
    return raw_name, "⚠️ New SKU"

# --- FAIL-SAFE AI PARSING ENGINE ---
def is_server_error(exception):
    err_str = str(exception).lower()
    return "503" in err_str or "unavailable" in err_str or "overloaded" in err_str or "429" in err_str

@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=2, min=2, max=10),
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
    prompt = """
    Analyze this commercial supplier invoice image with 100% precision.
    Extract all line items into a strict JSON array.
    
    For each item, identify if prices listed are GST Inclusive or Exclusive, and return JSON with these exact keys:
    [
        {
            "Item Name": "string description from bill",
            "Quantity": 1.0,
            "Listed Base Rate": 100.0,
            "Listed Total Inclusive Rate": 118.0,
            "GST Rate": 18.0,
            "Unit": "PCS"
        }
    ]
    
    CRITICAL RULES:
    1. Quantity, Listed Base Rate, Listed Total Inclusive Rate, and GST Rate MUST be numbers.
    2. If Listed Base Rate is missing, set it to 0.
    3. If Listed Total Inclusive Rate is missing, set it to 0.
    4. Default Unit to 'PCS' or 'LTR' if unspecified.
    """
    
    config = types.GenerateContentConfig(response_mime_type="application/json")
    contents = [image, prompt]
    candidate_models = ['gemini-3.5-flash', 'gemini-3.5-flash-lite', 'gemini-2.5-flash']
    
    last_error = None
    for model_name in candidate_models:
        try:
            response = _call_gemini_with_retry(client, model_name, contents, config)
            return json.loads(response.text)
        except Exception as e:
            last_error = e
            continue
            
    raise Exception(f"Server demand spike across all endpoints. Last Error: {last_error}")

# --- SIDEBAR: PROFIT MARGIN RULE SETTINGS ---
st.sidebar.header("⚙️ Profit Margin Target Rules")
st.sidebar.caption("Set default margin % targets per product category:")

for cat in list(category_margins.keys()):
    category_margins[cat] = st.sidebar.number_input(
        f"Margin % for {cat}",
        min_value=0.0,
        max_value=100.0,
        value=float(category_margins[cat]),
        step=1.0
    )

if st.sidebar.button("💾 Save Margin Rules"):
    save_json_file(MARGINS_FILE, category_margins)
    st.sidebar.success("Saved margin rules!")

# --- WORKSPACE INTERFACE ---
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Invoice Intake & Audit")
    uploaded_file = st.file_uploader("Upload Supplier Purchase Invoice (JPG/PNG)", type=["jpg", "jpeg", "png"])
    
    if uploaded_file:
        image = Image.open(uploaded_file)
        st.image(image, caption="Uploaded Invoice Preview", use_column_width=True)
        
        if st.button("🚀 Process & Audit Invoice", type="primary"):
            with st.spinner("Analyzing rates, applying category margins, and calculating profits..."):
                try:
                    raw_data = extract_invoice_data(image)
                    processed_items = []
                    
                    for row in raw_data:
                        qty = float(row.get("Quantity") or 1.0)
                        gst_rate = float(row.get("GST Rate") or 18.0)
                        base_rate = float(row.get("Listed Base Rate") or 0.0)
                        total_inclusive = float(row.get("Listed Total Inclusive Rate") or 0.0)
                        
                        # Calculate Base & Total Cost
                        if base_rate > 0:
                            final_base_cost = base_rate
                            final_inclusive_cost = base_rate * (1 + (gst_rate / 100))
                        elif total_inclusive > 0:
                            final_inclusive_cost = total_inclusive
                            final_base_cost = total_inclusive / (1 + (gst_rate / 100))
                        else:
                            final_base_cost = 0.0
                            final_inclusive_cost = 0.0
                            
                        raw_item_name = str(row.get("Item Name", "")).strip()
                        matched_sku, match_type = match_sku(raw_item_name)
                        category = get_category_for_sku(matched_sku)
                        
                        # Apply Category-Specific Target Margin %
                        target_margin_pct = category_margins.get(category, category_margins.get("General", 25.0))
                        
                        # Calculate Selling Price from Target Margin: Selling Price = Total Cost / (1 - Margin/100)
                        if target_margin_pct < 100:
                            suggested_sale_price = final_inclusive_cost / (1 - (target_margin_pct / 100))
                        else:
                            suggested_sale_price = final_inclusive_cost * 1.25
                            
                        processed_items.append({
                            "Raw Vendor Item Name": raw_item_name,
                            "Official SKU": matched_sku,
                            "Category": category,
                            "Qty": qty,
                            "Base Cost (Excl GST)": round(final_base_cost, 2),
                            "Total Cost (Incl GST)": round(final_inclusive_cost, 2),
                            "Selling Price": round(suggested_sale_price, 2),
                            "GST %": gst_rate,
                            "Unit": str(row.get("Unit", "PCS")).upper()
                        })
                        
                    st.session_state["parsed_df"] = pd.DataFrame(processed_items)
                    st.success("✅ Audit complete! Category margin rules applied.")
                except Exception as e:
                    st.error(f"❌ Error processing bill: {e}")

with col2:
    st.subheader("2. Advanced Profit & Control Panel")
    if "parsed_df" in st.session_state:
        df = st.session_state["parsed_df"]
        
        st.info("💡 Edit 'Official SKU' or 'Selling Price' below. Profit calculations update in real time.")
        
        # Interactive Editor
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Official SKU": st.column_config.SelectboxColumn("Official SKU Name", options=master_sku_list, required=True),
                "Category": st.column_config.SelectboxColumn("Category", options=list(category_margins.keys())),
                "Base Cost (Excl GST)": st.column_config.NumberColumn("Base Cost (₹)", format="₹%.2f"),
                "Total Cost (Incl GST)": st.column_config.NumberColumn("Total Cost (₹)", format="₹%.2f"),
                "Selling Price": st.column_config.NumberColumn("Selling Price (₹)", format="₹%.2f"),
                "Qty": st.column_config.NumberColumn("Qty", min_value=1),
                "GST %": st.column_config.NumberColumn("GST %", min_value=0, max_value=28),
            }
        )
        
        # --- DYNAMIC FINANCIAL CALCULATIONS ---
        # Recalculate metrics based on edited table values
        total_cost_sum = (edited_df["Total Cost (Incl GST)"] * edited_df["Qty"]).sum()
        total_revenue_sum = (edited_df["Selling Price"] * edited_df["Qty"]).sum()
        total_profit_sum = total_revenue_sum - total_cost_sum
        overall_margin_pct = (total_profit_sum / total_revenue_sum * 100) if total_revenue_sum > 0 else 0.0
        
        st.write("---")
        st.markdown("### 📊 Live Invoice Profit Analytics")
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Invoice Cost", f"₹{total_cost_sum:,.2f}")
        m2.metric("Total Selling Value", f"₹{total_revenue_sum:,.2f}")
        m3.metric("Net Invoice Profit", f"₹{total_profit_sum:,.2f}", delta=f"{overall_margin_pct:.1f}% Margin")
        m4.metric("Overall Margin %", f"{overall_margin_pct:.1f}%")
        
        # Warn if overall margin drops below 12%
        if overall_margin_pct < 12.0 and total_revenue_sum > 0:
            st.warning("⚠️ Low Margin Warning: Overall invoice profit margin is below 12%. Check individual item selling prices!")
            
        # --- EXPORT & MEMORY UPDATE ---
        if st.button("📥 Approve & Download for myBillBook", type="primary"):
            # Update memory JSON
            memory_updated = False
            for idx, row in edited_df.iterrows():
                raw = str(row["Raw Vendor Item Name"]).strip().upper()
                official = str(row["Official SKU"]).strip()
                
                if raw and official and raw != official:
                    mapping_memory[raw] = official
                    memory_updated = True
                    
            if memory_updated:
                save_json_file(MEMORY_FILE, mapping_memory)
                st.toast("🧠 Saved vendor mapping to memory!")
                
            # Prepare exact structure for myBillBook Excel Import
            output_df = pd.DataFrame()
            output_df["Item Name"] = edited_df["Official SKU"]
            output_df["Item Code / Barcode"] = ""
            output_df["Category"] = edited_df["Category"]
            output_df["Sale Price"] = edited_df["Selling Price"]
            output_df["Purchase Price"] = edited_df["Base Cost (Excl GST)"]
            output_df["Opening Stock"] = edited_df["Qty"]
            output_df["Measuring Unit"] = edited_df["Unit"]
            output_df["GST Tax Rate (%)"] = edited_df["GST %"]
            output_df["Tax Included in Sale Price?"] = "Yes"
            
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name="myBillBook Import")
                
            st.download_button(
                label="✅ Download Ready Excel File for myBillBook",
                data=buffer.getvalue(),
                file_name="myBillBook_Stock_Import.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
