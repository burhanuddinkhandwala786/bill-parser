import os
import gc
import json
import time
import shutil
import bcrypt
import openpyxl
import pandas as pd
import streamlit as st
from io import BytesIO
from PIL import Image
from google import genai
from google.genai import types
from rapidfuzz import process, utils
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception
from concurrent.futures import ThreadPoolExecutor
from sqlalchemy import create_engine, text

# REPORTLAB PDF ENGINE IMPORTS
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Universal OS | Multi-Store AI SaaS",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CLEAN THEME-SAFE CSS ---
st.markdown("""
<style>
    div[data-testid="InputInstructions"] {
        display: none !important;
    }
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2.5rem !important;
    }
    div[data-testid="stMetricValue"] div {
        color: #2563EB !important;
        font-weight: 700 !important;
    }
    div[data-testid="stColumn"] {
        display: flex;
        flex-direction: column;
    }
    div[data-testid="stColumn"] > div {
        flex: 1;
    }
    div[data-testid="stColumn"] div[data-testid="stVerticalBlockBorderWrapper"] {
        height: 100% !important;
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

if "last_store_slug" not in st.session_state:
    st.session_state["last_store_slug"] = selected_store_slug

if st.session_state["last_store_slug"] != selected_store_slug:
    if "parsed_df" in st.session_state:
        del st.session_state["parsed_df"]
    st.session_state["last_store_slug"] = selected_store_slug
    st.rerun()

st.sidebar.divider()

# --- STORE PROFILE / CONTACT PERSISTENCE ENGINE ---
CURRENT_STORE_DIR = os.path.join(DATA_DIR, selected_store_slug)
PROFILE_FILE = os.path.join(CURRENT_STORE_DIR, "store_profile.json")

def load_store_profile():
    if os.path.exists(PROFILE_FILE):
        try:
            with open(PROFILE_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {"phone_numbers": "", "business_title": selected_store_slug.replace("_", " ").upper()}

def save_store_profile(profile_dict):
    try:
        with open(PROFILE_FILE, "w") as f:
            json.dump(profile_dict, f, indent=4)
    except Exception as e:
        st.sidebar.error(f"Profile save alert: {e}")

store_profile = load_store_profile()

with st.sidebar.expander("📞 Business Contact Details (Saved)", expanded=False):
    saved_phone = store_profile.get("phone_numbers", "")
    phone_input = st.text_input("Owner Phone Number(s):", value=saved_phone, placeholder="e.g. +91 9876543210, +91 9123456789")
    if st.button("Save Contact Profile", use_container_width=True):
        store_profile["phone_numbers"] = phone_input.strip()
        save_store_profile(store_profile)
        st.sidebar.success("Business profile updated!")
        st.rerun()

# --- SIDEBAR STORE ACTIONS ---
with st.sidebar.expander("✏️ Rename Active Store", expanded=False):
    current_display = selected_store_slug.replace("_", " ")
    renamed_input = st.text_input("New Name:", value=current_display, key="rename_input_field")
    if st.button("Save Store Name", use_container_width=True, type="secondary"):
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

with st.sidebar.expander("➕ Register New Store", expanded=False):
    new_store_name = st.text_input("Store Title:", placeholder="e.g. Metro Hardware", key="add_input_field")
    if st.button("Create Environment", use_container_width=True, type="primary"):
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

MEMORY_FILE = os.path.join(CURRENT_STORE_DIR, "vendor_mappings.json")
MASTER_FILE = os.path.join(CURRENT_STORE_DIR, "inventory_master.csv")
ACTIVE_STORE_DISPLAY = selected_store_slug.replace("_", " ").upper()

# --- HEADER SECTION ---
header_left, header_right = st.columns([3, 1.5])

with header_left:
    st.title("⚡ Universal OS")
    st.caption("Commercial Multi-Store AI Purchase Intake & Inventory Synchronizer")

with header_right:
    st.caption("ACTIVE STORE CATALOG")
    st.markdown(f"📍 **{ACTIVE_STORE_DISPLAY}**")

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
    img_copy.thumbnail((1400, 1400), Image.Resampling.BILINEAR)
    
    buffer = BytesIO()
    if img_copy.mode in ("RGBA", "P"):
        img_copy = img_copy.convert("RGB")
    
    img_copy.save(buffer, format="JPEG", quality=85, optimize=True)
    buffer.seek(0)
    optimized_img = Image.open(buffer)

    prompt = """
    Extract all purchase invoice items into strict JSON format. CRITICAL INSTRUCTIONS:
    1. "Line Total Taxable Amount": The total base amount before GST for this item line (Printed Amount column).
    2. "Line Total Inclusive Amount": The total line amount including GST.
    3. "Unit Price (Excl. Tax)": Per-item rate before GST if printed.

    {
        "Supplier Company Name": "Vendor Name",
        "Line Items": [
            {
                "Item Name": "description",
                "Quantity": 1.0,
                "Unit Price (Excl. Tax)": 0.0,
                "Line Total Taxable Amount": 0.0,
                "Line Total Inclusive Amount": 0.0,
                "GST Rate": 18.0,
                "HSN Code": "",
                "Unit": "PCS"
            }
        ]
    }
    Rules: Rates and GST must be pure numbers. HSN Code as string. Missing values set to 0 or "". Default Unit to PCS, SQM, or LTR.
    """
    
    config = types.GenerateContentConfig(response_mime_type="application/json")
    contents = [optimized_img, prompt]
    
    candidate_models = ['gemini-2.5-flash', 'gemini-3.5-flash-lite', 'gemini-3.5-flash']
    
    last_error = None
    for model_name in candidate_models:
        try:
            response = _call_gemini_with_retry(client, model_name, contents, config)
            text = response.text.strip()
            if text.startswith("```json"):
                text = text[7:-3].strip()
            elif text.startswith("```"):
                text = text[3:-3].strip()
            return json.loads(text)
        except Exception as e:
            last_error = e
            continue
            
    raise Exception(f"AI Service error across models: {last_error}")

def process_single_file_raw(file_bytes):
    try:
        img = Image.open(BytesIO(file_bytes))
        return extract_invoice_data(img)
    except Exception as e:
        return {"ERROR": str(e)}

# --- PROFESSIONAL PDF QUOTATION GENERATOR ---
def generate_quotation_pdf(store_name, owner_phones, customer_name, quote_df, grand_total):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    elements = []
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'StoreTitle',
        parent=styles['Heading1'],
        fontSize=18,
        leading=22,
        textColor=colors.HexColor('#1E293B'),
        alignment=0
    )
    
    phone_style = ParagraphStyle(
        'PhoneStyle',
        parent=styles['Normal'],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor('#475569')
    )
    
    quote_title_style = ParagraphStyle(
        'QuoteHeader',
        parent=styles['Heading1'],
        fontSize=20,
        leading=24,
        textColor=colors.HexColor('#2563EB'),
        alignment=2
    )

    date_style = ParagraphStyle(
        'DateStyle',
        parent=styles['Normal'],
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#64748B'),
        alignment=2
    )
    
    # 1. Header Layout
    phone_text = f"Contact: {owner_phones}" if owner_phones else "Contact: N/A"
    header_data = [
        [
            Paragraph(f"<b>{store_name}</b>", title_style),
            Paragraph("<b>QUOTATION</b>", quote_title_style)
        ],
        [
            Paragraph(phone_text, phone_style),
            Paragraph(f"Date: {time.strftime('%d-%b-%Y %H:%M')}", date_style)
        ]
    ]
    
    header_table = Table(header_data, colWidths=[320, 220])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 15))
    
    # 2. Customer Section Box
    cust_style = ParagraphStyle('CustStyle', parent=styles['Normal'], fontSize=11, leading=15, textColor=colors.HexColor('#0F172A'))
    cust_p = Paragraph(f"<b>Quotation Prepared For:</b> {customer_name.upper()}", cust_style)
    elements.append(cust_p)
    elements.append(Spacer(1, 12))
    
    # 3. Item Table
    table_data = [["S.No", "Item Description", "Qty", "Unit", "Unit Price (₹)", "Total Value (₹)"]]
    
    for idx, row in quote_df.iterrows():
        table_data.append([
            str(idx + 1),
            Paragraph(str(row["Item Name"]), styles['Normal']),
            f"{row['Quantity']:,.2f}",
            str(row["Unit"]),
            f"₹{row['Customer Unit Price (₹)']:,.2f}",
            f"₹{row['Total Value (₹)']:,.2f}"
        ])
    
    # Total Row
    table_data.append(["", "", "", "", "Grand Total:", f"₹{grand_total:,.2f}"])
    
    item_table = Table(table_data, colWidths=[35, 235, 55, 45, 85, 85])
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2563EB')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('ALIGN', (2,0), (-1,-1), 'RIGHT'),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('TOPPADDING', (0,0), (-1,0), 6),
        ('GRID', (0,0), (-1,-2), 0.5, colors.HexColor('#E2E8F0')),
        ('FONTNAME', (-2,-1), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (-2,-1), (-1,-1), 10),
        ('BACKGROUND', (-2,-1), (-1,-1), colors.HexColor('#F8FAFC')),
        ('TOPPADDING', (0,1), (-1,-1), 6),
        ('BOTTOMPADDING', (0,1), (-1,-1), 6),
    ]))
    
    elements.append(item_table)
    elements.append(Spacer(1, 25))
    
    # 4. Footer Note
    footer_style = ParagraphStyle('FooterNote', parent=styles['Italic'], fontSize=8, leading=11, textColor=colors.HexColor('#64748B'), alignment=1)
    elements.append(Paragraph("This is a computer-generated estimate quotation. Prices are subject to availability at purchase.", footer_style))
    
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

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
    sm1.metric("Master SKUs Registered", len(master_sku_list))
    sm2.metric("Learned Vendor Rules", len(mapping_memory))
    sm3.metric("AI Engine Status", "🟢 Ready")

    st.divider()

    col_upload, col_info = st.columns([2, 1])
    
    with col_upload:
        with st.container(border=True):
            st.subheader("1. Ingestion Dropzone")
            st.caption("Upload purchase bills (PNG, JPG, JPEG) to extract line items.")
            uploaded_files = st.file_uploader(
                "Upload Bills",
                type=["jpg", "jpeg", "png"],
                accept_multiple_files=True,
                label_visibility="collapsed"
            )

    with col_info:
        with st.container(border=True):
            st.subheader("⚡ Ingestion Queue")
            if uploaded_files:
                st.success(f"📁 **{len(uploaded_files)} File(s)** Staged")
                st.caption("Parallel AI engine ready to parse files concurrently.")
            else:
                st.info("No files queued.")
                st.caption("Drop purchase invoices to begin structured extraction.")

    if uploaded_files:
        st.write("")
        if st.button("🚀 Process Invoices with Fast AI Engine", type="primary", use_container_width=True):
            if "parsed_df" in st.session_state:
                del st.session_state["parsed_df"]
                
            all_parsed_items = []
            file_bytes_list = [f.read() for f in uploaded_files]
            
            with st.status("Parsing purchase bills concurrently with Multimodal AI...", expanded=True) as status_container:
                with ThreadPoolExecutor(max_workers=min(len(file_bytes_list), 10)) as executor:
                    raw_results = list(executor.map(process_single_file_raw, file_bytes_list))
                    
                for parsed_json in raw_results:
                    if "ERROR" in parsed_json:
                        st.error(f"Failed to process a file: {parsed_json['ERROR']}")
                        continue
                        
                    supplier = parsed_json.get("Supplier Company Name", "Unknown Supplier")
                    
                    for row in parsed_json.get("Line Items", []):
                        qty = float(row.get("Quantity") or 1.0)
                        if qty <= 0: qty = 1.0
                        
                        gst_rate = float(row.get("GST Rate") or 18.0)
                        unit_price = float(row.get("Unit Price (Excl. Tax)") or 0.0)
                        line_taxable = float(row.get("Line Total Taxable Amount") or 0.0)
                        line_inclusive = float(row.get("Line Total Inclusive Amount") or 0.0)
                        hsn_sac = str(row.get("HSN Code") or "").strip()
                        
                        # GROUND TRUTH RECONCILIATION MATH
                        if line_taxable > 0:
                            total_taxable_item = line_taxable
                            final_base = line_taxable / qty
                        elif line_inclusive > 0:
                            total_taxable_item = line_inclusive / (1 + (gst_rate / 100))
                            final_base = total_taxable_item / qty
                        elif unit_price > 0:
                            final_base = unit_price
                            total_taxable_item = unit_price * qty
                        else:
                            final_base = 0.0
                            total_taxable_item = 0.0
                            
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
                            "Purchase Price": round(final_base, 2), # Excl. GST Rate for ERP Import
                            "Line Total Taxable": round(total_taxable_item, 2),
                            "Selling Price": round(known_selling, 2)
                        })
                    
                gc.collect()
                status_container.update(label="✅ Ingestion & Batch Extraction Complete!", state="complete", expanded=False)
                
            if all_parsed_items:
                st.session_state["parsed_df"] = pd.DataFrame(all_parsed_items)
                st.rerun()

    # --- REVIEW & EDIT WORKSPACE ---
    if "parsed_df" in st.session_state:
        st.divider()
        st.subheader("2. Live Inventory Audit Workspace")
        st.caption("Verify AI extraction, mapped SKUs, and rate details before exporting to accounting software.")
        
        df = st.session_state["parsed_df"]
        
        df["Line Total (Excl. GST)"] = (df["Purchase Price"] * df["Current Quantity"]).round(2)
        df["GST Tax Amount"] = (df["Line Total (Excl. GST)"] * (df["GST Rate"] / 100)).round(2)
        df["Line Total (Incl. GST)"] = (df["Line Total (Excl. GST)"] + df["GST Tax Amount"]).round(2)
        df["Unit Cost (GST Paid) ₹"] = (df["Line Total (Incl. GST)"] / df["Current Quantity"]).round(2)
        
        total_taxable = df["Line Total (Excl. GST)"].sum()
        total_gst = df["GST Tax Amount"].sum()
        grand_total_incl_tax = total_taxable + total_gst

        uom_groups = df.groupby("Unit")["Current Quantity"].sum()
        uom_summary_str = " | ".join([f"{val:,.2f} {unit}" for unit, val in uom_groups.items()])

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Line Items", f"{len(df)} Items")
        m2.metric("Stock Quantities by UOM", uom_summary_str)
        m3.metric("Taxable Base (Excl. GST)", f"₹{total_taxable:,.2f}")
        m4.metric("Grand Total (GST Paid)", f"₹{grand_total_incl_tax:,.2f}", delta=f"GST Tax: ₹{total_gst:,.2f}")
        
        st.write("")
        
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Official SKU": st.column_config.SelectboxColumn("Official SKU Name", options=master_sku_list, required=True) if master_sku_list else "Official SKU",
                "Unit Cost (GST Paid) ₹": st.column_config.NumberColumn("Unit Cost (GST Paid) ₹", format="₹%.2f", disabled=True),
                "Purchase Price": st.column_config.NumberColumn("Unit Rate (Excl. GST) ₹", format="₹%.2f"),
                "Line Total (Excl. GST)": st.column_config.NumberColumn("Line Total (Excl. GST) ₹", format="₹%.2f", disabled=True),
                "Line Total (Incl. GST)": st.column_config.NumberColumn("Line Total (Incl. GST) ₹", format="₹%.2f", disabled=True),
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
                st.toast("🧠 Learned Vendor Mapping updated!")
                
            if master_updated:
                save_master(master_df, selected_store_slug)
                st.toast("⚙️ Master SKU catalog expanded!")
                
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Items"
            
            ws.cell(row=1, column=1, value=ACTIVE_STORE_DISPLAY)
            ws.cell(row=2, column=1, value="Items")
            ws.cell(row=3, column=1, value=f"Generated On: {time.strftime('%d-%m-%Y %H:%M:%S')}")
            
            exact_headers = [
                "S. No.", "Name", "Current Quantity", "Unit", "HSN/SAC",
                "Category", "GST Rate", "Selling Price", "Selling Price (Secondary)",
                "Purchase Price", "Purchase Price (Secondary)", "Secondary Unit", "Ratio"
            ]
            
            for col_num, header_title in enumerate(exact_headers, 1):
                ws.cell(row=5, column=col_num, value=header_title)
            
            for i, row in edited_df.iterrows():
                row_idx = 6 + i
                selling_val = float(row["Selling Price"]) if float(row["Selling Price"]) > 0 else ""
                purchase_val = float(row["Purchase Price"])
                
                ws.cell(row=row_idx, column=1, value=i + 1)
                ws.cell(row=row_idx, column=2, value=str(row["Official SKU"]))
                ws.cell(row=row_idx, column=3, value=float(row["Current Quantity"]))
                ws.cell(row=row_idx, column=4, value=str(row["Unit"]))
                ws.cell(row=row_idx, column=5, value=str(row["HSN/SAC"]))
                ws.cell(row=row_idx, column=6, value=str(row["Category"]))
                ws.cell(row=row_idx, column=7, value=float(row["GST Rate"]))
                ws.cell(row=row_idx, column=8, value=selling_val)
                ws.cell(row=row_idx, column=9, value="")
                ws.cell(row=row_idx, column=10, value=purchase_val)
                ws.cell(row=row_idx, column=11, value="")
                ws.cell(row=row_idx, column=12, value="")
                ws.cell(row=row_idx, column=13, value="")

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

        # --- ON-THE-GO CUSTOMER QUOTATION (PDF) ---
        st.write("")
        with st.expander("📄 Generate Customer Quotation PDF (On-the-go)", expanded=False):
            st.caption("Instantly generate a customer PDF quotation by applying a custom markup % directly to your final GST-Paid cost.")
            
            col_q1, col_q2, col_q3 = st.columns([2, 2, 2])
            
            with col_q1:
                customer_name = st.text_input("Customer Name / Reference", value="Walk-in Customer")
            
            with col_q2:
                markup_pct = st.number_input("Markup Percentage (%)", min_value=0.0, value=15.0, step=1.0)
            
            with col_q3:
                current_phone = store_profile.get("phone_numbers", "")
                phone_override = st.text_input("Business Phone(s):", value=current_phone, help="Saved permanently per store directory.")
                if phone_override != current_phone:
                    store_profile["phone_numbers"] = phone_override.strip()
                    save_store_profile(store_profile)

            if st.button("Generate & Download PDF Quotation", type="primary", use_container_width=True):
                quote_items = []
                total_quote_value = 0.0
                
                for idx, row in edited_df.iterrows():
                    base_gst_paid_cost = float(row["Unit Cost (GST Paid) ₹"])
                    markup_price = round(base_gst_paid_cost * (1 + (markup_pct / 100)), 2)
                    qty = float(row["Current Quantity"])
                    line_total = round(markup_price * qty, 2)
                    
                    total_quote_value += line_total
                    
                    quote_items.append({
                        "Item Name": str(row["Official SKU"]),
                        "Quantity": qty,
                        "Unit": str(row["Unit"]),
                        "Customer Unit Price (₹)": markup_price,
                        "Total Value (₹)": line_total
                    })
                
                quote_df = pd.DataFrame(quote_items)
                
                # Render On-screen Preview
                st.markdown(f"**Previewing PDF Quotation for:** {customer_name.upper()}")
                st.dataframe(quote_df, use_container_width=True)
                st.metric(f"Total Customer Quotation Amount (Markup: {markup_pct}%)", f"₹{total_quote_value:,.2f}")
                
                # Build PDF File
                pdf_bytes = generate_quotation_pdf(
                    store_name=ACTIVE_STORE_DISPLAY,
                    owner_phones=phone_override,
                    customer_name=customer_name,
                    quote_df=quote_df,
                    grand_total=total_quote_value
                )
                
                st.download_button(
                    label=f"📄 Download Formal Quotation PDF for {customer_name}",
                    data=pdf_bytes,
                    file_name=f"Quotation_{customer_name.replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

# ==========================================
# TAB 2: STORE MASTER CATALOG MANAGER
# ==========================================
with tab_master:
    st.subheader(f"⚙️ Master Inventory Catalog ({ACTIVE_STORE_DISPLAY})")
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
    st.subheader(f"🧠 Learned AI Vendor Memory ({ACTIVE_STORE_DISPLAY})")
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
    st.subheader("📖 Standard Operating Procedure")
    with st.container(border=True):
        st.markdown("""
        ### How to Process & Sync Invoices:
        1. **Select Store:** Choose your active store location in the left sidebar directory.
        2. **Upload Bills:** Drop one or multiple purchase invoice photos in **Tab 1**.
        3. **Run AI Engine:** Click **Run AI Invoice Parsing Engine** to extract structured line items.
        4. **Audit Workspace:** Check quantities, HSN codes, purchase rates, and mapped SKUs.
        5. **Generate Quote PDF:** Open the Quotation expander to generate & download a PDF for a customer.
        6. **Download Import File:** Generate the `.xlsx` spreadsheet.
        7. **Import to ERP:** Open your accounting software → **Items / Inventory** → **Bulk Import**, upload the `.xlsx` file.
        """)
