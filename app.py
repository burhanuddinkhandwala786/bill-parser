import os
import json
import openpyxl
import pandas as pd
import streamlit as st
from io import BytesIO
from PIL import Image
from google import genai
from google.genai import types
from rapidfuzz import process, utils

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Enterprise Invoice Engine", page_icon="🛡️", layout="wide")
st.title("🛡️ Universal Hardware & Paints - Military-Grade Bill Parser")
st.caption("Fail-Safe Purchase Intake System for myBillBook Integration")

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

# --- MEMORY & MASTER DATA SYSTEM ---
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
    with open(MEMORY_FILE, "w") as f:
        json.dump(memory_dict, f, indent=4)

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
    """Smart Matching Pipeline: Check Memory -> Fuzzy Search -> Raw Default."""
    cleaned_raw = raw_name.strip().upper()
    
    # 1. Exact Memory Match
    if cleaned_raw in mapping_memory:
        return mapping_memory[cleaned_raw], "🧠 Learned Memory"
    
    # 2. Fuzzy Match against Master List
    if master_sku_list:
        match, score, _ = process.extractOne(raw_name, master_sku_list, processor=utils.default_process)
        if score > 65:
            return match, f"🔍 Fuzzy ({int(score)}%)"
            
    return raw_name, "⚠️ New SKU"

# --- MILITARY-GRADE AI PARSING PROMPT ---
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
            "Unit": "PCS",
            "Is Tax Explicitly Shown": true
        }
    ]
    
    CRITICAL RULES:
    1. Quantity, Listed Base Rate, Listed Total Inclusive Rate, and GST Rate MUST be numbers.
    2. If Listed Base Rate is missing on bill, set it to null or 0.
    3. If Listed Total Inclusive Rate is missing on bill, set it to null or 0.
    4. "Is Tax Explicitly Shown" must be true if base rate and GST amounts are written separately; false if only one total price is given.
    5. Default Unit to 'PCS' or 'LTR' if unspecified.
    """
    
    response = client.models.generate_content(
        model='gemini-3.5-flash',
        contents=[image, prompt],
        config=types.GenerateContentConfig(
            response_mime_type="application/json"
        )
    )
    return json.loads(response.text)

# --- WORKSPACE INTERFACE ---
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Invoice Intake & Verification")
    uploaded_file = st.file_uploader("Upload Supplier Purchase Invoice (JPG/PNG)", type=["jpg", "jpeg", "png"])
    
    if uploaded_file:
        image = Image.open(uploaded_file)
        st.image(image, caption="Uploaded Invoice Image", use_column_width=True)
        
        if st.button("🚀 Process & Audit Invoice", type="primary"):
            with st.spinner("AI Engine auditing invoice rates and tax structures..."):
                try:
                    raw_data = extract_invoice_data(image)
                    processed_items = []
                    
                    for row in raw_data:
                        qty = float(row.get("Quantity") or 1.0)
                        gst_rate = float(row.get("GST Rate") or 18.0)
                        base_rate = float(row.get("Listed Base Rate") or 0.0)
                        total_inclusive = float(row.get("Listed Total Inclusive Rate") or 0.0)
                        explicit_tax = row.get("Is Tax Explicitly Shown", False)
                        
                        # FAIL-SAFE CALCULATION LOGIC
                        if base_rate > 0:
                            # Mode A: Base rate is given on bill
                            final_base_cost = base_rate
                            final_inclusive_cost = base_rate * (1 + (gst_rate / 100))
                        elif total_inclusive > 0:
                            # Mode B: Only inclusive rate is given
                            final_inclusive_cost = total_inclusive
                            final_base_cost = total_inclusive / (1 + (gst_rate / 100))
                        else:
                            final_base_cost = 0.0
                            final_inclusive_cost = 0.0
                            
                        raw_item_name = str(row.get("Item Name", "")).strip()
                        matched_sku, match_type = match_sku(raw_item_name)
                        
                        # Default 25% profit margin logic if selling price is unknown
                        suggested_sale_price = final_inclusive_cost * 1.25
                        margin_pct = ((suggested_sale_price - final_inclusive_cost) / suggested_sale_price * 100) if suggested_sale_price > 0 else 0.0
                        
                        processed_items.append({
                            "Raw Vendor Item Name": raw_item_name,
                            "Official SKU": matched_sku,
                            "Match System": match_type,
                            "Qty": qty,
                            "Base Cost (Excl GST)": round(final_base_cost, 2),
                            "Total Cost (Incl GST)": round(final_inclusive_cost, 2),
                            "Selling Price": round(suggested_sale_price, 2),
                            "Margin %": round(margin_pct, 1),
                            "GST %": gst_rate,
                            "Unit": str(row.get("Unit", "PCS")).upper()
                        })
                        
                    st.session_state["parsed_df"] = pd.DataFrame(processed_items)
                    st.success("✅ Audit complete! Review extracted items in Panel 2.")
                except Exception as e:
                    st.error(f"❌ Error processing bill: {e}")

with col2:
    st.subheader("2. Review & Control Panel")
    if "parsed_df" in st.session_state:
        df = st.session_state["parsed_df"]
        
        st.info("💡 Review items below. The system automatically calculates Base Cost and Input Tax Credit for myBillBook.")
        
        # Interactive Editor with Column Rules
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Official SKU": st.column_config.SelectboxColumn("Official SKU Name", options=master_sku_list, required=True),
                "Base Cost (Excl GST)": st.column_config.NumberColumn("Base Cost (₹)", format="₹%.2f"),
                "Total Cost (Incl GST)": st.column_config.NumberColumn("Total Cost (₹)", format="₹%.2f"),
                "Selling Price": st.column_config.NumberColumn("Selling Price (₹)", format="₹%.2f"),
                "Margin %": st.column_config.NumberColumn("Margin", format="%.1f%%"),
                "Qty": st.column_config.NumberColumn("Qty", min_value=1),
                "GST %": st.column_config.NumberColumn("GST %", min_value=0, max_value=28),
            }
        )
        
        # Summary Analytics
        total_lines = len(edited_df)
        total_inv_value = (edited_df["Total Cost (Incl GST)"] * edited_df["Qty"]).sum()
        avg_margin = edited_df["Margin %"].mean()
        
        st.write("---")
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Lines", f"{total_lines}")
        m2.metric("Invoice Total", f"₹{total_inv_value:,.2f}")
        m3.metric("Average Profit Margin", f"{avg_margin:.1f}%")
        
        # --- EXPORT & MEMORY UPDATE ---
        if st.button("📥 Approve, Learn Memory & Download for myBillBook", type="primary"):
            # Update memory JSON
            memory_updated = False
            for idx, row in edited_df.iterrows():
                raw = str(row["Raw Vendor Item Name"]).strip().upper()
                official = str(row["Official SKU"]).strip()
                
                if raw and official and raw != official:
                    mapping_memory[raw] = official
                    memory_updated = True
                    
            if memory_updated:
                save_json_memory(mapping_memory)
                st.toast("🧠 Saved vendor mapping to memory!")
                
            # Prepare exact structure for myBillBook Excel Import
            output_df = pd.DataFrame()
            output_df["Item Name"] = edited_df["Official SKU"]
            output_df["Item Code / Barcode"] = ""
            output_df["Category"] = "General"
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
                label="✅ Download Ready Excel File",
                data=buffer.getvalue(),
                file_name="myBillBook_Stock_Import.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
