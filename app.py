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
from PIL import Image, ImageEnhance
from google import genai
from google.genai import types
from rapidfuzz import process, utils
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception
from concurrent.futures import ThreadPoolExecutor
from sqlalchemy import create_engine, text

# --- SAFE COOKIE MANAGER IMPORT ---
try:
    import extra_streamlit_components as stx
    COOKIE_SUPPORT_AVAILABLE = True
except ImportError:
    COOKIE_SUPPORT_AVAILABLE = False

# --- SAFE REPORTLAB PDF IMPORTS ---
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

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

# --- DATABASE ENGINE SETUP ---
@st.cache_resource
def get_db_engine():
    db_url = st.secrets["SUPABASE_DB_URL"]
    return create_engine(db_url, pool_pre_ping=True)

def sanitize_store_slug(name):
    return "".join([c if c.isalnum() else "_" for c in name.strip()]).lower()

def get_or_create_store_id(store_slug: str, display_name: str = None) -> int:
    """Gets the database ID for a store slug, creating it if missing."""
    engine = get_db_engine()
    slug = sanitize_store_slug(store_slug)
    if not display_name:
        display_name = store_slug.replace("_", " ").title()
    
    with engine.begin() as conn:
        result = conn.execute(
            text("SELECT id FROM stores WHERE slug = :slug"),
            {"slug": slug}
        ).fetchone()
        
        if result:
            return result[0]
        
        insert_res = conn.execute(
            text("INSERT INTO stores (slug, display_name) VALUES (:slug, :display_name) RETURNING id"),
            {"slug": slug, "display_name": display_name}
        )
        return insert_res.fetchone()[0]

def get_store_phone(store_slug: str) -> str:
    """Fetches business phone numbers saved in database for a store safely."""
    engine = get_db_engine()
    try:
        with engine.connect() as conn:
            res = conn.execute(
                text("SELECT phone FROM stores WHERE slug = :slug"),
                {"slug": store_slug}
            ).fetchone()
            if res and res[0]:
                return res[0]
    except Exception:
        pass
    return ""

def save_store_phone(store_slug: str, phone: str):
    """Saves updated business phone numbers to database safely."""
    engine = get_db_engine()
    try:
        with engine.begin() as conn:
            conn.execute(
                text("UPDATE stores SET phone = :phone WHERE slug = :slug"),
                {"phone": phone.strip(), "slug": store_slug}
            )
    except Exception:
        pass

# --- AUTHENTICATION & MULTI-TENANT SESSION MANAGEMENT ---
def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def check_password(password: str, hashed: str) -> bool:
    return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))

def register_user(store_name: str, email: str, password: str):
    engine = get_db_engine()
    slug = sanitize_store_slug(store_name)
    hashed_pwd = hash_password(password)
    
    with engine.begin() as conn:
        existing = conn.execute(
            text("SELECT id FROM stores WHERE email = :email OR slug = :slug"),
            {"email": email.strip().lower(), "slug": slug}
        ).fetchone()
        
        if existing:
            return False, "Store name or Email already registered!"
            
        conn.execute(
            text("""
                INSERT INTO stores (slug, display_name, email, password)
                VALUES (:slug, :display_name, :email, :password)
            """),
            {
                "slug": slug,
                "display_name": store_name.strip(),
                "email": email.strip().lower(),
                "password": hashed_pwd
            }
        )
    return True, "Store registered successfully! Please log in."

def authenticate_user(email: str, password: str):
    engine = get_db_engine()
    with engine.connect() as conn:
        user = conn.execute(
            text("SELECT slug, display_name, password, phone FROM stores WHERE email = :email"),
            {"email": email.strip().lower()}
        ).fetchone()
        
        if user and user[2] and check_password(password, user[2]):
            return {
                "slug": user[0], 
                "display_name": user[1],
                "phone": user[3] if len(user) > 3 and user[3] else ""
            }
    return None

def get_user_by_slug(slug: str):
    engine = get_db_engine()
    try:
        with engine.connect() as conn:
            user = conn.execute(
                text("SELECT slug, display_name, phone FROM stores WHERE slug = :slug"),
                {"slug": slug.strip().lower()}
            ).fetchone()
            
            if user:
                return {
                    "slug": user[0],
                    "display_name": user[1],
                    "phone": user[2] if len(user) > 2 and user[2] else ""
                }
    except Exception:
        pass
    return None

def reset_user_password(email: str, new_password: str):
    engine = get_db_engine()
    hashed_pwd = hash_password(new_password)
    
    with engine.begin() as conn:
        user = conn.execute(
            text("SELECT id FROM stores WHERE email = :email"),
            {"email": email.strip().lower()}
        ).fetchone()
        
        if not user:
            return False, "No store registered with this business email address."
            
        conn.execute(
            text("UPDATE stores SET password = :password WHERE email = :email"),
            {"password": hashed_pwd, "email": email.strip().lower()}
        )
    return True, "Password updated successfully! Please log in with your new password."

# Initialize Cookie Manager
cookie_manager = stx.CookieManager(key="app_cookie_mgr") if COOKIE_SUPPORT_AVAILABLE else None

# Session State Initialization
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "user_store" not in st.session_state:
    st.session_state["user_store"] = None

# Auto-Login from Browser Cookie
if not st.session_state["authenticated"] and cookie_manager:
    try:
        saved_store_slug = cookie_manager.get(cookie="store_session")
        if saved_store_slug:
            store_data = get_user_by_slug(saved_store_slug)
            if store_data:
                st.session_state["authenticated"] = True
                st.session_state["user_store"] = store_data
    except Exception:
        pass

# --- LOGIN / SIGNUP / RESET SCREEN ---
if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("⚡ Universal OS")
        st.subheader("Commercial AI Invoice Intake & ERP Synchronizer")
        st.write("")
        
        auth_tab1, auth_tab2, auth_tab3 = st.tabs(["🔒 Partner Login", "✨ Register New Store", "🔑 Forgot Password"])
        
        with auth_tab1:
            with st.form("login_form"):
                login_email = st.text_input("Store Email")
                login_password = st.text_input("Password", type="password")
                submit_login = st.form_submit_button("Log In", use_container_width=True, type="primary")
                
                if submit_login:
                    if login_email and login_password:
                        store_data = authenticate_user(login_email, login_password)
                        if store_data:
                            st.session_state["authenticated"] = True
                            st.session_state["user_store"] = store_data
                            
                            if cookie_manager:
                                try:
                                    cookie_manager.set("store_session", store_data["slug"], max_age=30*24*3600)
                                except Exception:
                                    pass
                                
                            st.success(f"Welcome back, {store_data['display_name']}!")
                            st.rerun()
                        else:
                            st.error("Invalid email or password.")
                    else:
                        st.warning("Please fill in both email and password.")

        with auth_tab2:
            with st.form("register_form"):
                reg_store = st.text_input("Store Name (e.g. Universal Hardware)")
                reg_email = st.text_input("Business Email")
                reg_password = st.text_input("Create Password", type="password")
                submit_reg = st.form_submit_button("Create Account & Store Environment", use_container_width=True)
                
                if submit_reg:
                    if reg_store and reg_email and reg_password:
                        success, msg = register_user(reg_store, reg_email, reg_password)
                        if success:
                            st.success(msg)
                        else:
                            st.error(msg)
                    else:
                        st.warning("All fields are required.")

        with auth_tab3:
            with st.form("reset_password_form"):
                reset_email = st.text_input("Registered Business Email")
                new_pwd = st.text_input("New Password", type="password")
                confirm_pwd = st.text_input("Confirm New Password", type="password")
                submit_reset = st.form_submit_button("Reset Password", use_container_width=True)
                
                if submit_reset:
                    if not reset_email or not new_pwd:
                        st.warning("Please enter your email and a new password.")
                    elif new_pwd != confirm_pwd:
                        st.error("Passwords do not match!")
                    else:
                        success, msg = reset_user_password(reset_email, new_pwd)
                        if success:
                            st.success(msg)
                        else:
                            st.error(msg)
                            
    st.stop()

# Active user details post-login
selected_store_slug = st.session_state["user_store"]["slug"]
ACTIVE_STORE_DISPLAY = st.session_state["user_store"]["display_name"].upper()

# --- SIDEBAR (LOGGED IN USER) ---
st.sidebar.title(f"🏬 {st.session_state['user_store']['display_name']}")
st.sidebar.caption(f"Active Store ID: `{selected_store_slug}`")
st.sidebar.divider()

if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state["authenticated"] = False
    st.session_state["user_store"] = None
    if cookie_manager:
        try:
            cookie_manager.delete("store_session")
        except Exception:
            pass
    if "parsed_df" in st.session_state:
        del st.session_state["parsed_df"]
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

# --- HEADER SECTION ---
header_left, header_right = st.columns([3, 1.5])

with header_left:
    st.title("⚡ Universal OS")
    st.caption("Commercial Multi-Store AI Purchase Intake & Inventory Synchronizer")

with header_right:
    st.caption("ACTIVE STORE CATALOG")
    st.markdown(f"📍 **{ACTIVE_STORE_DISPLAY}**")

st.divider()

# --- FAST BULK-OPTIMIZED DATABASE STORE LOADERS ---
def load_json_memory(store_slug: str) -> dict:
    """Loads learned vendor item mappings from Supabase."""
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    query = "SELECT raw_name, mapped_sku FROM vendor_mappings WHERE store_id = :store_id"
    try:
        with engine.connect() as conn:
            df = pd.read_sql(text(query), conn, params={"store_id": store_id})
        if df.empty:
            return {}
        return dict(zip(df["raw_name"], df["mapped_sku"]))
    except Exception:
        return {}

def save_json_memory(store_slug: str, memory_dict: dict):
    """Saves learned vendor item mappings to Supabase using fast bulk upserts."""
    if not memory_dict:
        return
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    records = [{"store_id": store_id, "raw_name": k, "mapped_sku": v} for k, v in memory_dict.items()]
    
    with engine.begin() as conn:
        conn.execute(
            text("""
                INSERT INTO vendor_mappings (store_id, raw_name, mapped_sku)
                VALUES (:store_id, :raw_name, :mapped_sku)
                ON CONFLICT (store_id, raw_name) 
                DO UPDATE SET mapped_sku = EXCLUDED.mapped_sku
            """),
            records
        )

@st.cache_data
def load_master(store_slug: str) -> pd.DataFrame:
    """Loads master SKU catalog from Supabase for a given store."""
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    query = """
        SELECT official_sku_name as "Official_SKU_Name", 
               category as "Category", 
               default_unit as "Default_Unit", 
               gst_rate as "GST_Rate", 
               selling_price as "Selling_Price"
        FROM master_skus
        WHERE store_id = :store_id
    """
    try:
        with engine.connect() as conn:
            df = pd.read_sql(text(query), conn, params={"store_id": store_id})
        if df.empty:
            return pd.DataFrame(columns=["Official_SKU_Name", "Category", "Default_Unit", "GST_Rate", "Selling_Price"])
        return df
    except Exception:
        return pd.DataFrame(columns=["Official_SKU_Name", "Category", "Default_Unit", "GST_Rate", "Selling_Price"])

def save_master(df: pd.DataFrame, store_slug: str):
    """Saves updated master SKU catalog to Supabase in a single fast bulk transaction."""
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    
    with engine.begin() as conn:
        conn.execute(
            text("DELETE FROM master_skus WHERE store_id = :store_id"),
            {"store_id": store_id}
        )
        if not df.empty:
            records = [
                {
                    "store_id": store_id,
                    "official_sku_name": str(row["Official_SKU_Name"]),
                    "category": str(row.get("Category", "General")),
                    "default_unit": str(row.get("Default_Unit", "PCS")),
                    "gst_rate": float(row.get("GST_Rate", 18.0)),
                    "selling_price": float(row.get("Selling_Price", 0.0))
                }
                for _, row in df.iterrows()
            ]
            conn.execute(
                text("""
                    INSERT INTO master_skus (store_id, official_sku_name, category, default_unit, gst_rate, selling_price)
                    VALUES (:store_id, :official_sku_name, :category, :default_unit, :gst_rate, :selling_price)
                """),
                records
            )
    st.cache_data.clear()

master_df = load_master(selected_store_slug)
master_sku_list = master_df["Official_SKU_Name"].dropna().tolist() if not master_df.empty else []
mapping_memory = load_json_memory(selected_store_slug)

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
    if img_copy.mode in ("RGBA", "P"):
        img_copy = img_copy.convert("RGB")
        
    img_copy.thumbnail((1400, 1400), Image.Resampling.BILINEAR)
    
    try:
        enhancer = ImageEnhance.Contrast(img_copy)
        img_copy = enhancer.enhance(1.3)
        sharpener = ImageEnhance.Sharpness(img_copy)
        img_copy = sharpener.enhance(1.4)
    except Exception:
        pass

    buffer = BytesIO()
    img_copy.save(buffer, format="JPEG", quality=85, optimize=True)
    buffer.seek(0)
    optimized_img = Image.open(buffer)

    prompt = """
    You are an expert OCR and financial vision system built for retail & wholesale purchase bill ingestion.
    Your task is to extract line items from purchase invoices—including printed, handwritten, low-light, faded thermal receipts, crumpled paper, or skewed photos.

    CRITICAL EXTRACTION INSTRUCTIONS:
    1. "Supplier Company Name": Extract the vendor/distributor company name at the header. If unclear or handwritten, infer best title or use "Unknown Supplier".
    2. "Line Items": Extract every purchased item row from the table or receipt list.
       - "Item Name": Full product title or description. Read faint or handwritten pen marks carefully.
       - "Quantity": Pure numeric value (e.g. 1.0, 10, 0.5). If missing or unreadable, default to 1.0.
       - "Unit Price (Excl. Tax)": Per-item rate before GST if explicitly printed.
       - "Discount Amount": Any cash/trade discount deducted for this item line. If none, 0.0.
       - "Line Total Taxable Amount": The line base total before GST (Printed Amount column).
       - "Line Total Inclusive Amount": The total line amount including GST.
       - "GST Rate": GST tax percentage as a pure number (0, 5, 12, 18, 28). Default to 18.0 if unstated.
       - "HSN Code": HSN or SAC code as string. Empty string "" if missing.
       - "Unit": Unit of measure (PCS, BOX, LTR, KG, NOS, SET, SQM, MTR, PKT). Default to "PCS".

    ROBUSTNESS & HANDWRITING RULES:
    - Faded / Low Contrast Text: Infer numbers by cross-checking quantity * rate - discount = total where possible.
    - Handwritten Text: Treat pen strokes and annotations as primary text if printed text is crossed out or modified.
    - Pure Numbers Only: Rates, quantities, amounts, discounts, and GST must be numbers (no currency symbols like ₹ or Rs).

    OUTPUT SCHEMA (STRICT JSON ONLY):
    {
        "Supplier Company Name": "Vendor Name",
        "Line Items": [
            {
                "Item Name": "description",
                "Quantity": 1.0,
                "Unit Price (Excl. Tax)": 0.0,
                "Discount Amount": 0.0,
                "Line Total Taxable Amount": 0.0,
                "Line Total Inclusive Amount": 0.0,
                "GST Rate": 18.0,
                "HSN Code": "",
                "Unit": "PCS"
            }
        ]
    }
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

# --- REPORTLAB PDF QUOTATION GENERATOR ---
def generate_quotation_pdf(store_name: str, phone_str: str, customer_name: str, quote_df: pd.DataFrame, grand_total: float) -> bytes:
    if not REPORTLAB_AVAILABLE:
        raise ModuleNotFoundError("reportlab library is required for PDF generation.")
        
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=letter, 
        rightMargin=36, 
        leftMargin=36, 
        topMargin=36, 
        bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'DocTitle', 
        parent=styles['Heading1'], 
        fontSize=20, 
        leading=24, 
        textColor=colors.HexColor('#1E3A8A'), 
        fontName='Helvetica-Bold',
        spaceAfter=4
    )
    subtitle_style = ParagraphStyle(
        'DocSubtitle', 
        parent=styles['Normal'], 
        fontSize=9.5, 
        leading=13, 
        textColor=colors.HexColor('#4B5563')
    )
    meta_label = ParagraphStyle(
        'MetaLabel', 
        parent=styles['Normal'], 
        fontSize=9, 
        leading=12, 
        textColor=colors.HexColor('#374151')
    )

    story.append(Paragraph("OFFICIAL QUOTATION", title_style))
    story.append(Paragraph(f"<b>Issued By:</b> {store_name.upper()}", subtitle_style))
    if phone_str:
        story.append(Paragraph(f"<b>Contact / Mobile:</b> {phone_str}", subtitle_style))
    story.append(Spacer(1, 10))
    
    meta_data = [
        [
            Paragraph(f"<b>Customer Ref:</b> {customer_name}", meta_label), 
            Paragraph(f"<b>Date:</b> {time.strftime('%d-%m-%Y %H:%M')}", meta_label)
        ]
    ]
    meta_table = Table(meta_data, colWidths=[310, 230])
    meta_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F3F4F6')),
        ('PADDING', (0, 0), (-1, -1), 6),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LINEBELOW', (0, 0), (-1, -1), 1, colors.HexColor('#E5E7EB')),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 14))

    hdr_left = ParagraphStyle('HdrL', fontSize=9, leading=11, textColor=colors.white, fontName='Helvetica-Bold', alignment=0)
    hdr_center = ParagraphStyle('HdrC', fontSize=9, leading=11, textColor=colors.white, fontName='Helvetica-Bold', alignment=1)
    hdr_right = ParagraphStyle('HdrR', fontSize=9, leading=11, textColor=colors.white, fontName='Helvetica-Bold', alignment=2)

    cell_left = ParagraphStyle('CellL', fontSize=8.5, leading=11, textColor=colors.HexColor('#111827'), alignment=0)
    cell_center = ParagraphStyle('CellC', fontSize=8.5, leading=11, textColor=colors.HexColor('#111827'), alignment=1)
    cell_right = ParagraphStyle('CellR', fontSize=8.5, leading=11, textColor=colors.HexColor('#111827'), alignment=2)

    table_data = [[
        Paragraph("S.No", hdr_center), 
        Paragraph("Item Description", hdr_left), 
        Paragraph("Qty", hdr_center), 
        Paragraph("Unit", hdr_center), 
        Paragraph("Unit Price (Rs.)", hdr_right), 
        Paragraph("Total (Rs.)", hdr_right)
    ]]

    for i, row in quote_df.iterrows():
        unit_price = float(row['Customer Unit Price (₹)'])
        line_total = float(row['Total Value (₹)'])
        
        table_data.append([
            Paragraph(str(i + 1), cell_center),
            Paragraph(str(row["Item Name"]), cell_left),
            Paragraph(f"{float(row['Quantity']):g}", cell_center),
            Paragraph(str(row["Unit"]), cell_center),
            Paragraph(f"Rs. {unit_price:,.2f}", cell_right),
            Paragraph(f"Rs. {line_total:,.2f}", cell_right)
        ])

    item_table = Table(table_data, colWidths=[35, 225, 45, 45, 95, 95])
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2563EB')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, 0), 6),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('TOPPADDING', (0, 1), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E5E7EB')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F9FAFB')])
    ]))
    story.append(item_table)
    story.append(Spacer(1, 14))

    total_style = ParagraphStyle(
        'GrandTotal', 
        parent=styles['Heading2'], 
        fontSize=13, 
        leading=16, 
        textColor=colors.HexColor('#1E3A8A'), 
        fontName='Helvetica-Bold',
        alignment=2
    )
    story.append(Paragraph(f"<b>Grand Total: Rs. {grand_total:,.2f}</b>", total_style))

    doc.build(story)
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
                        discount = float(row.get("Discount Amount") or 0.0)
                        line_taxable = float(row.get("Line Total Taxable Amount") or 0.0)
                        line_inclusive = float(row.get("Line Total Inclusive Amount") or 0.0)
                        hsn_sac = str(row.get("HSN Code") or "").strip()
                        
                        if line_taxable > 0:
                            total_taxable_item = line_taxable
                            final_base = line_taxable / qty
                        elif line_inclusive > 0:
                            total_taxable_item = line_inclusive / (1 + (gst_rate / 100))
                            final_base = total_taxable_item / qty
                        elif unit_price > 0:
                            total_taxable_item = (unit_price * qty) - discount
                            final_base = total_taxable_item / qty if qty > 0 else 0.0
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
                            "Purchase Price": round(final_base, 2),
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
        
        # Ensure correct numeric types
        df["Current Quantity"] = pd.to_numeric(df["Current Quantity"], errors='coerce').fillna(1.0)
        df["Purchase Price"] = pd.to_numeric(df["Purchase Price"], errors='coerce').fillna(0.0)
        df["GST Rate"] = pd.to_numeric(df["GST Rate"], errors='coerce').fillna(18.0)
        df["Selling Price"] = pd.to_numeric(df.get("Selling Price", 0.0), errors='coerce').fillna(0.0)
        
        # Calculate full financial metrics
        df["Line Total (Excl. GST)"] = (df["Purchase Price"] * df["Current Quantity"]).round(2)
        df["GST Tax Amount"] = (df["Line Total (Excl. GST)"] * (df["GST Rate"] / 100)).round(2)
        df["Line Total (Incl. GST)"] = (df["Line Total (Excl. GST)"] + df["GST Tax Amount"]).round(2)
        df["Unit Cost (GST Paid) ₹"] = (df["Line Total (Incl. GST)"] / df["Current Quantity"]).round(2)
        
        total_taxable = df["Line Total (Excl. GST)"].sum()
        total_gst = df["GST Tax Amount"].sum()
        grand_total_incl_tax = total_taxable + total_gst

        uom_groups = df.groupby("Unit")["Current Quantity"].sum()
        uom_summary_str = " | ".join([f"{val:,.2f} {unit}" for unit, val in uom_groups.items()])

        # Executive Metrics Header
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Line Items", f"{len(df)} Items")
        m2.metric("Stock Quantities by UOM", uom_summary_str)
        m3.metric("Taxable Base (Excl. GST)", f"₹{total_taxable:,.2f}")
        m4.metric("Grand Total (GST Paid)", f"₹{grand_total_incl_tax:,.2f}", delta=f"GST Tax: ₹{total_gst:,.2f}")
        
        st.write("")
        
        # DUAL TOTALS AUDIT TABLE VIEW
        display_columns = [
            "Match Status",
            "Raw Vendor Item",
            "Official SKU",
            "Current Quantity",
            "Unit",
            "HSN/SAC",
            "Purchase Price",
            "Line Total (Excl. GST)",
            "GST Rate",
            "Unit Cost (GST Paid) ₹",
            "Line Total (Incl. GST)",
            "Selling Price"
        ]
        
        available_cols = [c for c in display_columns if c in df.columns]
        
        edited_display_df = st.data_editor(
            df[available_cols],
            num_rows="dynamic",
            use_container_width=True,
            key="audit_editor",
            column_config={
                "Match Status": st.column_config.TextColumn("Match Status", disabled=True),
                "Raw Vendor Item": st.column_config.TextColumn("Raw Vendor Item", disabled=True),
                "Official SKU": st.column_config.SelectboxColumn("Official SKU Name", options=master_sku_list, required=True) if master_sku_list else "Official SKU",
                "Current Quantity": st.column_config.NumberColumn("Qty", min_value=0.1, format="%.2f"),
                "Unit": st.column_config.TextColumn("Unit"),
                "HSN/SAC": st.column_config.TextColumn("HSN/SAC"),
                "Purchase Price": st.column_config.NumberColumn("Unit Rate (Excl. GST) ₹", format="₹%.2f"),
                "Line Total (Excl. GST)": st.column_config.NumberColumn("Total (Excl. GST) ₹", format="₹%.2f", disabled=True),
                "GST Rate": st.column_config.NumberColumn("GST %", min_value=0, max_value=28, format="%d%%"),
                "Unit Cost (GST Paid) ₹": st.column_config.NumberColumn("Unit Cost (GST Paid) ₹", format="₹%.2f", disabled=True),
                "Line Total (Incl. GST)": st.column_config.NumberColumn("Total (Incl. GST) ₹", format="₹%.2f", disabled=True),
                "Selling Price": st.column_config.NumberColumn("Selling Price ₹", format="₹%.2f"),
            }
        )
        
        # Safely handle row modifications/additions/deletions
        df_updated = edited_display_df.copy()
        df_updated["Current Quantity"] = pd.to_numeric(df_updated["Current Quantity"], errors='coerce').fillna(1.0)
        df_updated["Purchase Price"] = pd.to_numeric(df_updated["Purchase Price"], errors='coerce').fillna(0.0)
        df_updated["GST Rate"] = pd.to_numeric(df_updated["GST Rate"], errors='coerce').fillna(18.0)
        df_updated["Selling Price"] = pd.to_numeric(df_updated.get("Selling Price", 0.0), errors='coerce').fillna(0.0)

        # Re-compute dependent line totals across all rows
        df_updated["Line Total (Excl. GST)"] = (df_updated["Purchase Price"] * df_updated["Current Quantity"]).round(2)
        df_updated["GST Tax Amount"] = (df_updated["Line Total (Excl. GST)"] * (df_updated["GST Rate"] / 100)).round(2)
        df_updated["Line Total (Incl. GST)"] = (df_updated["Line Total (Excl. GST)"] + df_updated["GST Tax Amount"]).round(2)
        df_updated["Unit Cost (GST Paid) ₹"] = (df_updated["Line Total (Incl. GST)"] / df_updated["Current Quantity"]).round(2)
        
        st.session_state["parsed_df"] = df_updated

        st.divider()
        if st.button("✅ Confirm Audit & Generate Excel Import File", type="primary", use_container_width=True):
            memory_updated = False
            master_updated = False
            current_master_skus = set(master_sku_list)
            
            for idx, row in df_updated.iterrows():
                raw = str(row.get("Raw Vendor Item", "")).strip().upper()
                official = str(row.get("Official SKU", "")).strip()
                
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
                save_json_memory(selected_store_slug, mapping_memory)
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
            
            for i, row in df_updated.iterrows():
                row_idx = 6 + i
                selling_val = float(row["Selling Price"]) if float(row["Selling Price"]) > 0 else ""
                purchase_val = float(row["Purchase Price"])
                
                ws.cell(row=row_idx, column=1, value=i + 1)
                ws.cell(row=row_idx, column=2, value=str(row["Official SKU"]))
                ws.cell(row=row_idx, column=3, value=float(row["Current Quantity"]))
                ws.cell(row=row_idx, column=4, value=str(row["Unit"]))
                ws.cell(row=row_idx, column=5, value=str(row["HSN/SAC"]))
                ws.cell(row=row_idx, column=6, value=str(row.get("Category", "General")))
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

        # --- ON-THE-GO CUSTOMER QUOTATION ---
        st.write("")
        with st.expander("📄 Generate Customer Quotation (On-the-go)", expanded=False):
            st.caption("Instantly generate a customer quotation by applying a custom markup % directly to your final GST-Paid cost.")
            
            saved_phone = st.session_state["user_store"].get("phone") or get_store_phone(selected_store_slug)
            
            col_q1, col_q2, col_q3 = st.columns([2, 2, 1])
            with col_q1:
                customer_name = st.text_input("Customer Name / Reference", value="Walk-in Customer")
            with col_q2:
                biz_phone_input = st.text_input("Business Phone Number(s)", value=saved_phone, placeholder="e.g. +91 9876543210, +91 9123456789")
            with col_q3:
                markup_pct = st.number_input("Markup (%)", min_value=0.0, value=15.0, step=1.0)
            
            if biz_phone_input.strip() != saved_phone.strip():
                save_store_phone(selected_store_slug, biz_phone_input.strip())
                st.session_state["user_store"]["phone"] = biz_phone_input.strip()
            
            if st.button("Preview & Generate PDF Quotation", type="primary", use_container_width=True):
                quote_items = []
                total_quote_value = 0.0
                
                for idx, row in df_updated.iterrows():
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
                st.markdown(f"**Previewing Quote for:** {customer_name}")
                st.dataframe(quote_df, use_container_width=True)
                st.metric(f"Total Quotation Value (Markup: {markup_pct}% on GST-Paid Cost)", f"₹{total_quote_value:,.2f}")
                
                if REPORTLAB_AVAILABLE:
                    pdf_bytes = generate_quotation_pdf(
                        store_name=st.session_state['user_store']['display_name'],
                        phone_str=biz_phone_input.strip(),
                        customer_name=customer_name,
                        quote_df=quote_df,
                        grand_total=total_quote_value
                    )
                    
                    st.download_button(
                        label=f"📄 Download PDF Quotation for {customer_name}",
                        data=pdf_bytes,
                        file_name=f"Quotation_{customer_name.replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                else:
                    st.warning("⚠️ 'reportlab' library is missing in environment. Please update your requirements.txt file to enable PDF downloading.")

# ==========================================
# TAB 2: STORE MASTER CATALOG MANAGER
# ==========================================
with tab_master:
    st.subheader(f"⚙️ Master Inventory Catalog ({ACTIVE_STORE_DISPLAY})")
    
    # --- BULK CATALOG IMPORT EXPANDER ---
    with st.expander("📥 Bulk Import Catalog via Excel / CSV", expanded=False):
        st.caption("Upload a spreadsheet containing your full product catalog to register multiple SKUs at once.")
        
        uploaded_catalog = st.file_uploader(
            "Upload Catalog (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="bulk_catalog_uploader"
        )
        
        if uploaded_catalog is not None:
            try:
                if uploaded_catalog.name.endswith(".csv"):
                    imported_df = pd.read_csv(uploaded_catalog)
                else:
                    imported_df = pd.read_excel(uploaded_catalog)
                
                st.markdown("**Uploaded File Preview:**")
                st.dataframe(imported_df.head(5), use_container_width=True)
                
                col_sku = st.selectbox("Select SKU Name Column:", options=imported_df.columns)
                
                col_cat, col_unit, col_gst, col_price = st.columns(4)
                with col_cat:
                    opt_cat = st.selectbox("Category Column (Optional):", options=["-- None --"] + list(imported_df.columns))
                with col_unit:
                    opt_unit = st.selectbox("Default Unit Column (Optional):", options=["-- None --"] + list(imported_df.columns))
                with col_gst:
                    opt_gst = st.selectbox("GST Rate Column (Optional):", options=["-- None --"] + list(imported_df.columns))
                with col_price:
                    opt_price = st.selectbox("Selling Price Column (Optional):", options=["-- None --"] + list(imported_df.columns))
                
                if st.button("🚀 Import All SKUs into Master Catalog", type="primary", use_container_width=True):
                    new_records = []
                    for _, row in imported_df.iterrows():
                        sku_val = str(row[col_sku]).strip() if pd.notnull(row[col_sku]) else ""
                        if not sku_val or sku_val.lower() == "nan":
                            continue
                            
                        cat_val = str(row[opt_cat]).strip() if opt_cat != "-- None --" and pd.notnull(row[opt_cat]) else "General"
                        unit_val = str(row[opt_unit]).strip().upper() if opt_unit != "-- None --" and pd.notnull(row[opt_unit]) else "PCS"
                        
                        try:
                            gst_val = float(row[opt_gst]) if opt_gst != "-- None --" and pd.notnull(row[opt_gst]) else 18.0
                        except ValueError:
                            gst_val = 18.0
                            
                        try:
                            price_val = float(row[opt_price]) if opt_price != "-- None --" and pd.notnull(row[opt_price]) else 0.0
                        except ValueError:
                            price_val = 0.0
                            
                        new_records.append({
                            "Official_SKU_Name": sku_val,
                            "Category": cat_val,
                            "Default_Unit": unit_val,
                            "GST_Rate": gst_val,
                            "Selling_Price": price_val
                        })
                    
                    if new_records:
                        bulk_df = pd.DataFrame(new_records)
                        combined = pd.concat([master_df, bulk_df], ignore_index=True)
                        combined = combined.drop_duplicates(subset=["Official_SKU_Name"], keep="last")
                        
                        save_master(combined, selected_store_slug)
                        st.success(f"Successfully imported {len(new_records)} SKUs into {ACTIVE_STORE_DISPLAY}!")
                        st.rerun()
                    else:
                        st.error("No valid SKU rows found in the uploaded file.")
                        
            except Exception as err:
                st.error(f"Error parsing bulk catalog file: {err}")

    st.divider()

    col_add, col_list = st.columns([1, 2])
    
    with col_add:
        with st.container(border=True):
            st.markdown("#### ➕ Add Single Master SKU")
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
            engine = get_db_engine()
            store_id = get_or_create_store_id(selected_store_slug)
            with engine.begin() as conn:
                conn.execute(
                    text("DELETE FROM vendor_mappings WHERE store_id = :store_id"),
                    {"store_id": store_id}
                )
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
        5. **Generate Quote (Optional):** Open the Quotation expander to quickly quote a customer.
        6. **Download Import File:** Generate the `.xlsx` spreadsheet.
        7. **Import to ERP:** Open your accounting software → **Items / Inventory** → **Bulk Import**, upload the `.xlsx` file.
        """)
