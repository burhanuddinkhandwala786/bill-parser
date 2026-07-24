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
st.set_page_config(page_title="Universal Business Intelligence Bridge", page_icon="⚡", layout="wide")
st.title("⚡ Universal Hardware & Paints - Enterprise Bill Engine")
st.caption("AI-Powered Purchase Intake & Inventory Optimization System")

# --- API KEY SETUP ---
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
    st.warning("Please enter your Gemini API Key in the sidebar or Streamlit Secrets.")
    st.stop()

client = genai.Client(api_key=api_key)

# --- MEMORY & MASTER DATA MANAGEMENT ---
MEMORY_FILE = "vendor_mappings.json"
MASTER_FILE = "inventory_master.csv"

def load_json_memory():
    if os.path.exists(MEMORY_FILE):
        with open(MEMORY_FILE, "r") as f:
            return json.load(f)
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
    """Smart matching: Checks learned memory first, then uses fuzzy string matching."""
    cleaned_raw = raw_name.strip().upper()
    
    # 1. Check if we previously learned this exact supplier abbreviation
    if cleaned_raw in mapping_memory:
        return mapping_memory[cleaned_raw], "Brain Memory (100% Match)"
    
    # 2. Fuzzy match against Master Inventory List
    if master_sku_list:
        match, score, _ = process.extractOne(raw_name, master_sku_list, processor=utils.default_process)
        if score > 65:
            return match, f"Fuzzy Match ({int(score)}%)"
            
    return raw_name, "New / Unmatched SKU"

# --- AI PARSING ENGINE ---
def extract_invoice_data(image):
    prompt = """
    Extract all line items from this purchase/vendor invoice bill into strict JSON.
    Return ONLY a JSON array of objects with these exact keys:
    [
        {
            "Item Name": "raw description from bill",
            "Quantity": 10.0,
            "Total Inclusive Purchase Price": 1180.0,
            "Sale Price": 1500.0,
            "GST Rate": 18,
            "Unit": "PCS"
        }
    ]
    Rules:
    - Quantity, Total Inclusive Purchase Price, Sale Price, and GST Rate MUST be pure numbers.
    - Total Inclusive Purchase Price is the final rate per unit including GST paid.
    - If Sale Price is missing on purchase bill, set it equal to Total Inclusive Purchase Price * 1.25 (25% default markup).
    - If Unit is unclear, default to 'PCS' or 'LTR'.
    """
    
    response = client.models.generate_content(
        model='gemini-3.5-flash',
        contents=[image, prompt],
        config=types.GenerateContentConfig(
            response_mime_type="application/json"
        )
    )
    return json.loads(response.text)

# --- MAIN WORKSPACE ---
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Bill Scanner & Intake")
    uploaded_file = st.file_uploader("Upload Image or Photo of Purchase Bill", type=["jpg", "jpeg", "png"])
    
    if uploaded_file:
        image = Image.open(uploaded_file)
        st.image(image, caption="Uploaded Bill Preview", use_column_width=True)
        
        if st.button("🚀 Analyze & Auto-Calculate Stock", type="primary"):
            with st.spinner("Processing invoice line items & calculating margins..."):
                try:
                    raw_data = extract_invoice_data(image)
                    
                    processed_items = []
                    for row in raw_data:
                        gst_rate = float(row.get("GST Rate", 18))
                        total_inclusive = float(row.get("Total Inclusive Purchase Price", 0))
                        qty = float(row.get("Quantity", 1))
                        
                        # Base Price (Excl. GST) = Total / (1 + GST/100)
                        base_cost = total_inclusive / (1 + (gst_rate / 100))
                        
                        raw_item_name = str(row.get("Item Name", "")).strip()
                        matched_sku, match_type = match_sku(raw_item_name)
                        
                        sale_price = float(row.get("Sale Price", total_inclusive * 1.25))
                        
                        # Profit Margin % = ((Sale Price - Total Cost) / Sale Price) * 100
                        profit_margin = ((sale_price - total_inclusive) / sale_price * 100) if sale_price > 0 else 0
                        
                        processed_items.append({
                            "Raw Item Name": raw_item_name,
                            "Official SKU": matched_sku,
                            "Match System": match_type,
                            "Qty": qty,
                            "Base Cost (Excl GST)": round(base_cost, 2),
                            "Total Cost (Incl GST)": round(total_inclusive, 2),
                            "Selling Price": round(sale_price, 2),
                            "Profit Margin %": round(profit_margin, 1),
                            "GST %": gst_rate,
                            "Unit": row.get("Unit", "PCS")
                        })
                        
                    st.session_state["parsed_df"] = pd.DataFrame(processed_items)
                    st.success("Analysis complete! Verification table ready.")
                except Exception as e:
                    st.error(f"Error processing invoice: {e}")

with col2:
    st.subheader("2. Profit Verification & Inventory Control")
    if "parsed_df" in st.session_state:
        df = st.session_state["parsed_df"]
        
        st.info("💡 Edit 'Official SKU' or 'Selling Price' below. The system automatically learns new mappings!")
        
        # Interactive Editor
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Official SKU": st.column_config.SelectboxColumn("Official SKU Name", options=master_sku_list, required=True),
                "Base Cost (Excl GST)": st.column_config.NumberColumn("Base Cost (₹)", format="₹%.2f"),
                "Total Cost (Incl GST)": st.column_config.NumberColumn("Total Cost (₹)", format="₹%.2f"),
                "Selling Price": st.column_config.NumberColumn("Selling Price (₹)", format="₹%.2f"),
                "Profit Margin %": st.column_config.NumberColumn("Profit Margin", format="%.1f%%"),
                "Qty": st.column_config.NumberColumn("Qty", min_value=1),
                "GST %": st.column_config.NumberColumn("GST %", min_value=0, max_value=28),
            }
        )
        
        # Calculate summary metrics
        total_items = len(edited_df)
        total_bill_val = (edited_df["Total Cost (Incl GST)"] * edited_df["Qty"]).sum()
        avg_margin = edited_df["Profit Margin %"].mean()
        
        st.write("---")
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Line Items", f"{total_items}")
        m2.metric("Total Invoice Value", f"₹{total_bill_val:,.2f}")
        m3.metric("Avg Profit Margin", f"{avg_margin:.1f}%")
        
        # --- EXPORT & LEARN BUTTON ---
        if st.button("📥 Approve, Save Memory & Download myBillBook Import"):
            # Update system memory with newly mapped abbreviations
            updated_memory = False
            for idx, row in edited_df.iterrows():
                raw = str(row["Raw Item Name"]).strip().upper()
                official = str(row["Official SKU"]).strip()
                
                if raw and official and raw != official:
                    mapping_memory[raw] = official
                    updated_memory = True
            
            if updated_memory:
                save_json_memory(mapping_memory)
                st.toast("🧠 Saved new supplier name mappings to AI Memory!")
            
            # Format output specifically for myBillBook Bulk Import
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
            
            # Export to Excel
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name="myBillBook Import")
            
            st.download_button(
                label="✅ Download Ready Excel File for myBillBook",
                data=buffer.getvalue(),
                file_name="myBillBook_Stock_Import.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )