import os
import gc
import json
import time
import base64
import random
import hashlib
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

# Optional PyMuPDF (fitz) Import for Catalog Extraction
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

# Optional Groq SDK Import for Secondary Failover
try:
    from groq import Groq
    GROQ_AVAILABLE = True
except ImportError:
    GROQ_AVAILABLE = False

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

# --- GLOBAL SESSION STATE INITIALIZATION ---
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "user_store" not in st.session_state:
    st.session_state["user_store"] = None
if "ocr_file_hash_cache" not in st.session_state:
    st.session_state["ocr_file_hash_cache"] = {}
if "processed_invoice_keys" not in st.session_state:
    st.session_state["processed_invoice_keys"] = set()
if "catalog_df" not in st.session_state:
    st.session_state["catalog_df"] = None

# --- DESIGN SYSTEM ---
st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
    :root {
        --uos-primary: #4F46E5;
        --uos-primary-600: #4338CA;
        --uos-primary-700: #3730A3;
        --uos-primary-50: #EEF2FF;
        --uos-primary-100: #E0E7FF;
        --uos-accent: #10B981;
        --uos-canvas: #F8FAFC;
        --uos-surface: #FFFFFF;
        --uos-surface-2: #F1F5F9;
        --uos-border: #E2E8F0;
        --uos-text: #0F172A;
        --uos-text-secondary: #475569;
        --uos-text-muted: #64748B;
        --uos-radius: 12px;
        --uos-shadow-xs: 0 1px 2px rgba(15, 23, 42, 0.04);
        --uos-focus-ring: 0 0 0 3px rgba(79, 70, 229, 0.20);
    }
    html, body, [class*="css"], .stApp, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
        background: var(--uos-canvas) !important;
        color: var(--uos-text) !important;
    }
    [data-testid="stHeader"] { background: transparent !important; }
    .block-container { padding-top: 1.5rem !important; padding-bottom: 3rem !important; max-width: 1400px; }
    div[data-testid="InputInstructions"] { display: none !important; }
    h1, h2, h3, h4, h5, h6 { font-family: 'Inter', sans-serif !important; color: var(--uos-text) !important; font-weight: 700 !important; }
    [data-testid="stSidebar"] { background: var(--uos-surface) !important; border-right: 1px solid var(--uos-border); }
    .uos-store-avatar { display: flex; align-items: center; gap: 12px; padding: 4px 2px 14px 2px; }
    .uos-store-avatar .uos-avatar-circle {
        width: 42px; height: 42px; border-radius: 12px;
        background: linear-gradient(135deg, var(--uos-primary) 0%, var(--uos-primary-700) 100%);
        color: #FFFFFF; display: flex; align-items: center; justify-content: center; font-weight: 700;
    }
    .uos-store-pill {
        display: inline-flex; align-items: center; gap: 8px; padding: 6px 14px 6px 10px;
        background: var(--uos-primary-50); color: var(--uos-primary); border: 1px solid #E0E7FF;
        border-radius: 999px; font-size: 0.82rem; font-weight: 600;
    }
    .uos-store-pill .uos-dot { width: 8px; height: 8px; border-radius: 50%; background: var(--uos-accent); }
    .stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid var(--uos-border); background: transparent !important; }
    .stTabs [data-baseweb="tab"] { padding: 12px 18px !important; font-weight: 500 !important; color: var(--uos-text-muted) !important; }
    .stTabs [data-baseweb="tab"][aria-selected="true"] { color: var(--uos-primary) !important; font-weight: 600 !important; border-bottom-color: var(--uos-primary) !important; }
    [data-testid="stMetric"] { background: var(--uos-surface) !important; border: 1px solid var(--uos-border) !important; border-radius: var(--uos-radius) !important; padding: 18px !important; }
    .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
        border-radius: var(--uos-radius) !important; font-weight: 600 !important; border: 1px solid var(--uos-border) !important;
    }
    .stButton > button[kind="primary"] { background: var(--uos-primary) !important; color: #FFFFFF !important; }
</style>
""", unsafe_allow_html=True)

# --- DATABASE ENGINE SETUP ---
@st.cache_resource
def get_db_engine():
    db_url = st.secrets["SUPABASE_DB_URL"]
    return create_engine(db_url, pool_pre_ping=True, pool_size=10, max_overflow=20)

def sanitize_store_slug(name):
    return "".join([c if c.isalnum() else "_" for c in name.strip()]).lower()

@st.cache_data(ttl=3600)
def get_or_create_store_id(store_slug: str, display_name: str = None) -> int:
    engine = get_db_engine()
    slug = sanitize_store_slug(store_slug)
    if not display_name:
        display_name = store_slug.replace("_", " ").title()
    with engine.begin() as conn:
        result = conn.execute(text("SELECT id FROM stores WHERE slug = :slug"), {"slug": slug}).fetchone()
        if result: return result[0]
        insert_res = conn.execute(text("INSERT INTO stores (slug, display_name) VALUES (:slug, :display_name) RETURNING id"), {"slug": slug, "display_name": display_name})
        return insert_res.fetchone()[0]

@st.cache_data(ttl=600)
def get_store_phone(store_slug: str) -> str:
    engine = get_db_engine()
    try:
        with engine.connect() as conn:
            res = conn.execute(text("SELECT phone FROM stores WHERE slug = :slug"), {"slug": store_slug}).fetchone()
            if res and res[0]: return res[0]
    except Exception:
        pass
    return ""

def save_store_phone(store_slug: str, phone: str):
    engine = get_db_engine()
    try:
        with engine.begin() as conn:
            conn.execute(text("UPDATE stores SET phone = :phone WHERE slug = :slug"), {"phone": phone.strip(), "slug": store_slug})
        get_store_phone.clear()
    except Exception:
        pass

# --- AUTHENTICATION ---
def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def check_password(password: str, hashed: str) -> bool:
    return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))

def register_user(store_name: str, email: str, password: str):
    engine = get_db_engine()
    slug = sanitize_store_slug(store_name)
    hashed_pwd = hash_password(password)
    with engine.begin() as conn:
        existing = conn.execute(text("SELECT id FROM stores WHERE email = :email OR slug = :slug"), {"email": email.strip().lower(), "slug": slug}).fetchone()
        if existing: return False, "Store name or Email already registered!"
        conn.execute(text("INSERT INTO stores (slug, display_name, email, password) VALUES (:slug, :display_name, :email, :password)"), {"slug": slug, "display_name": store_name.strip(), "email": email.strip().lower(), "password": hashed_pwd})
    return True, "Store registered successfully! Please log in."

def authenticate_user(email: str, password: str):
    engine = get_db_engine()
    with engine.connect() as conn:
        user = conn.execute(text("SELECT slug, display_name, password, phone FROM stores WHERE email = :email"), {"email": email.strip().lower()}).fetchone()
        if user and user[2] and check_password(password, user[2]):
            return {"slug": user[0], "display_name": user[1], "phone": user[3] if len(user) > 3 and user[3] else ""}
    return None

def get_user_by_slug(slug: str):
    engine = get_db_engine()
    try:
        with engine.connect() as conn:
            user = conn.execute(text("SELECT slug, display_name, phone FROM stores WHERE slug = :slug"), {"slug": slug.strip().lower()}).fetchone()
            if user: return {"slug": user[0], "display_name": user[1], "phone": user[2] if len(user) > 2 and user[2] else ""}
    except Exception:
        pass
    return None

def reset_user_password(email: str, new_password: str):
    engine = get_db_engine()
    hashed_pwd = hash_password(new_password)
    with engine.begin() as conn:
        user = conn.execute(text("SELECT id FROM stores WHERE email = :email"), {"email": email.strip().lower()}).fetchone()
        if not user: return False, "No store registered with this business email address."
        conn.execute(text("UPDATE stores SET password = :password WHERE email = :email"), {"password": hashed_pwd, "email": email.strip().lower()})
    return True, "Password updated successfully!"

if not st.session_state["authenticated"]:
    url_session = st.query_params.get("session", None)
    if url_session:
        if isinstance(url_session, list): url_session = url_session[0]
        store_data = get_user_by_slug(str(url_session))
        if store_data:
            st.session_state["authenticated"] = True
            st.session_state["user_store"] = store_data

if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
            <div class="uos-auth-hero" style="text-align:center; padding:20px 0;">
                <div class="uos-auth-mark" style="font-size:2.5rem;">⚡</div>
                <h1 class="uos-auth-title">Universal OS</h1>
                <p class="uos-auth-tagline">High-Performance Enterprise AI Pricing &amp; ERP Engine</p>
            </div>
        """, unsafe_allow_html=True)
        auth_tab1, auth_tab2, auth_tab3 = st.tabs(["🔒 Partner Login", "✨ Register Store", "🔑 Reset Password"])
        with auth_tab1:
            with st.form("login_form"):
                login_email = st.text_input("Store Email")
                login_password = st.text_input("Password", type="password")
                if st.form_submit_button("Log In", use_container_width=True, type="primary"):
                    if login_email and login_password:
                        store_data = authenticate_user(login_email, login_password)
                        if store_data:
                            st.session_state["authenticated"] = True
                            st.session_state["user_store"] = store_data
                            st.query_params["session"] = store_data["slug"]
                            st.rerun()
                        else: st.error("Invalid credentials.")
                    else: st.warning("Fill in all fields.")
        with auth_tab2:
            with st.form("register_form"):
                reg_store = st.text_input("Store Name")
                reg_email = st.text_input("Business Email")
                reg_password = st.text_input("Password", type="password")
                if st.form_submit_button("Create Account", use_container_width=True):
                    if reg_store and reg_email and reg_password:
                        success, msg = register_user(reg_store, reg_email, reg_password)
                        if success: st.success(msg)
                        else: st.error(msg)
                    else: st.warning("All fields required.")
        with auth_tab3:
            with st.form("reset_form"):
                remail = st.text_input("Registered Email")
                rpwd = st.text_input("New Password", type="password")
                if st.form_submit_button("Reset Password", use_container_width=True):
                    success, msg = reset_user_password(remail, rpwd)
                    if success: st.success(msg)
                    else: st.error(msg)
    st.stop()

selected_store_slug = st.session_state["user_store"]["slug"]
ACTIVE_STORE_DISPLAY = st.session_state["user_store"]["display_name"].upper()

# --- SIDEBAR ---
_display_name = st.session_state['user_store']['display_name']
_initials = "".join([w[0].upper() for w in _display_name.split()[:2] if w]) or "US"
st.sidebar.markdown(f"""
    <div class="uos-store-avatar">
        <div class="uos-avatar-circle">{_initials}</div>
        <div class="uos-avatar-meta">
            <div class="uos-avatar-name">{_display_name}</div>
            <div class="uos-avatar-slug">{selected_store_slug}</div>
        </div>
    </div>
""", unsafe_allow_html=True)
st.sidebar.divider()
if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state["authenticated"] = False
    st.session_state["user_store"] = None
    st.query_params.clear()
    st.rerun()

# --- API POOL SETUP ---
def get_api_key_pool():
    keys = []
    for i in range(1, 6):
        key = st.secrets.get(f"GEMINI_API_KEY_{i}") or os.environ.get(f"GEMINI_API_KEY_{i}")
        if key: keys.append(key)
    if not keys:
        single = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
        if single: keys.append(single)
    return keys

API_KEYS_POOL = get_api_key_pool()
if not API_KEYS_POOL:
    st.error("⚠️ Gemini API Key missing.")
    st.stop()

# --- LOADERS ---
@st.cache_data(ttl=3600)
def load_json_memory(store_slug: str) -> dict:
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    try:
        with engine.connect() as conn:
            df = pd.read_sql(text("SELECT raw_name, mapped_sku FROM vendor_mappings WHERE store_id = :store_id"), conn, params={"store_id": store_id})
        if df.empty: return {}
        return dict(zip(df["raw_name"], df["mapped_sku"]))
    except Exception:
        return {}

def save_json_memory(store_slug: str, memory_dict: dict):
    if not memory_dict: return
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    records = [{"store_id": store_id, "raw_name": k, "mapped_sku": v} for k, v in memory_dict.items()]
    with engine.begin() as conn:
        conn.execute(text("INSERT INTO vendor_mappings (store_id, raw_name, mapped_sku) VALUES (:store_id, :raw_name, :mapped_sku) ON CONFLICT (store_id, raw_name) DO UPDATE SET mapped_sku = EXCLUDED.mapped_sku"), records)
    load_json_memory.clear()

@st.cache_data(ttl=3600)
def load_master(store_slug: str) -> pd.DataFrame:
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    try:
        with engine.connect() as conn:
            df = pd.read_sql(text('SELECT official_sku_name as "Official_SKU_Name", category as "Category", default_unit as "Default_Unit", gst_rate as "GST_Rate", selling_price as "Selling_Price" FROM master_skus WHERE store_id = :store_id'), conn, params={"store_id": store_id})
        if df.empty: return pd.DataFrame(columns=["Official_SKU_Name", "Category", "Default_Unit", "GST_Rate", "Selling_Price"])
        return df
    except Exception:
        return pd.DataFrame(columns=["Official_SKU_Name", "Category", "Default_Unit", "GST_Rate", "Selling_Price"])

def save_master(df: pd.DataFrame, store_slug: str):
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM master_skus WHERE store_id = :store_id"), {"store_id": store_id})
        if not df.empty:
            records = [{"store_id": store_id, "official_sku_name": str(row["Official_SKU_Name"]), "category": str(row.get("Category", "General")), "default_unit": str(row.get("Default_Unit", "PCS")), "gst_rate": float(row.get("GST_Rate", 18.0)), "selling_price": float(row.get("Selling_Price", 0.0))} for _, row in df.iterrows()]
            conn.execute(text("INSERT INTO master_skus (store_id, official_sku_name, category, default_unit, gst_rate, selling_price) VALUES (:store_id, :official_sku_name, :category, :default_unit, :gst_rate, :selling_price)"), records)
    load_master.clear()

def add_single_sku_direct(store_slug: str, sku_name: str, category: str, unit: str, gst_rate: float, selling_price: float):
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    with engine.begin() as conn:
        conn.execute(text("INSERT INTO master_skus (store_id, official_sku_name, category, default_unit, gst_rate, selling_price) VALUES (:store_id, :official_sku_name, :category, :default_unit, :gst_rate, :selling_price) ON CONFLICT (store_id, official_sku_name) DO UPDATE SET category = EXCLUDED.category, default_unit = EXCLUDED.default_unit, gst_rate = EXCLUDED.gst_rate, selling_price = EXCLUDED.selling_price"), {"store_id": store_id, "official_sku_name": sku_name, "category": category, "default_unit": unit, "gst_rate": gst_rate, "selling_price": selling_price})
    load_master.clear()

def bulk_upsert_audited_skus(store_slug: str, records: list):
    if not records: return
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    with engine.begin() as conn:
        for rec in records:
            conn.execute(text("INSERT INTO master_skus (store_id, official_sku_name, category, default_unit, gst_rate, selling_price) VALUES (:store_id, :official_sku_name, :category, :default_unit, :gst_rate, :selling_price) ON CONFLICT (store_id, official_sku_name) DO UPDATE SET category = EXCLUDED.category, default_unit = EXCLUDED.default_unit, gst_rate = EXCLUDED.gst_rate, selling_price = EXCLUDED.selling_price"), {"store_id": store_id, "official_sku_name": rec["official"], "category": rec["category"], "default_unit": rec["unit"], "gst_rate": rec["gst"], "selling_price": rec["sp"]})
    load_master.clear()

def delete_multiple_skus(store_slug: str, sku_list: list):
    if not sku_list: return
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM master_skus WHERE store_id = :store_id AND official_sku_name IN :sku_names"), {"store_id": store_id, "sku_names": tuple(sku_list)})
    load_master.clear()

master_df = load_master(selected_store_slug)
master_sku_list = master_df["Official_SKU_Name"].dropna().tolist() if not master_df.empty else []
mapping_memory = load_json_memory(selected_store_slug)

def match_sku(raw_name):
    cleaned_raw = raw_name.strip().upper()
    if cleaned_raw in mapping_memory: return mapping_memory[cleaned_raw]
    if master_sku_list:
        match, score, _ = process.extractOne(raw_name, master_sku_list, processor=utils.default_process)
        if score > 85: return match
        elif score >= 60: return f"⚠️ Needs Review: {match}"
    return raw_name

def get_known_selling_price(sku_name):
    clean_sku = sku_name.replace("⚠️ Needs Review: ", "").strip()
    if not master_df.empty and "Selling_Price" in master_df.columns:
        matched = master_df[master_df["Official_SKU_Name"] == clean_sku]
        if not matched.empty:
            price = matched.iloc[0]["Selling_Price"]
            if pd.notnull(price) and float(price) > 0: return float(price)
    return 0.0

@retry(stop=stop_after_attempt(2), wait=wait_exponential(multiplier=1, min=1, max=3), reraise=True)
def _call_gemini_with_retry(client, model_name, contents, config):
    return client.models.generate_content(model=model_name, contents=contents, config=config)

# --- LIGHTNING-FAST CACHED OCR PARSER ---
def extract_invoice_data_multiformat(file_bytes, mime_type="image/jpeg"):
    file_hash = hashlib.md5(file_bytes).hexdigest()
    if file_hash in st.session_state["ocr_file_hash_cache"]:
        return st.session_state["ocr_file_hash_cache"][file_hash]

    try:
        if "pdf" in mime_type.lower() and PYMUPDF_AVAILABLE:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            page = doc[0]
            pix = page.get_pixmap(dpi=100)
            img = Image.open(BytesIO(pix.tobytes("jpeg")))
            doc.close()
        else:
            img = Image.open(BytesIO(file_bytes))

        if img.mode in ("RGBA", "P"): img = img.convert("RGB")
        img.thumbnail((700, 700), Image.Resampling.BILINEAR)
        buffer = BytesIO()
        img.save(buffer, format="JPEG", quality=70, optimize=True)
        optimized_bytes = buffer.getvalue()
        del img
        gc.collect()

        contents = [Image.open(BytesIO(optimized_bytes))]
        prompt = """
        Extract supplier name, invoice number, date, and line items (Item Name, Primary Quantity, Unit, Printed Taxable Amount, GST Rate, HSN Code).
        OUTPUT STRICT JSON ONLY:
        {
            "Supplier Company Name": "", "Invoice Number": "", "Invoice Date": "",
            "Line Items": [{"Item Name": "", "Primary Quantity": 1.0, "Unit": "PCS", "Printed Taxable Amount": 0.0, "GST Rate": 18.0, "HSN Code": ""}]
        }
        """
        contents.append(prompt)
        config = types.GenerateContentConfig(response_mime_type="application/json")
        models = ['gemini-2.5-flash-lite', 'gemini-2.5-flash']

        for api_key in API_KEYS_POOL:
            try:
                client = genai.Client(api_key=api_key)
                for m in models:
                    try:
                        res = _call_gemini_with_retry(client, m, contents, config)
                        text_res = res.text.strip()
                        if text_res.startswith("```json"): text_res = text_res[7:-3].strip()
                        elif text_res.startswith("```"): text_res = text_res[3:-3].strip()
                        parsed = json.loads(text_res)
                        st.session_state["ocr_file_hash_cache"][file_hash] = parsed
                        return parsed
                    except Exception: continue
            except Exception: continue
        raise Exception("AI quota exhausted.")
    except Exception as e:
        return {"ERROR": str(e)}

def process_single_item_tuple(tup):
    return extract_invoice_data_multiformat(tup[0], tup[1])

# --- TAB 2 CATALOG PARSER (LIGHTNING SPEED & STRICT GROUPING) ---
def parse_single_catalog_file(file_tuple):
    file_bytes, mime_type = file_tuple
    catalog_items = []
    try:
        images_to_process = []
        if "pdf" in mime_type.lower() and PYMUPDF_AVAILABLE:
            pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
            for page_idx in range(len(pdf_doc)):
                page = pdf_doc[page_idx]
                pix = page.get_pixmap(dpi=100)
                images_to_process.append(Image.open(BytesIO(pix.tobytes("jpeg"))))
            pdf_doc.close()
        else:
            images_to_process.append(Image.open(BytesIO(file_bytes)))

        cat_prompt = """
        Extract the Supplier/Brand Company Name (e.g. DAARVI, GREENPLY) from the header, and every product row in the table.
        OUTPUT STRICT JSON ONLY:
        {
            "Supplier Name": "Brand Name",
            "Products": [{"Item Name": "", "Model Code": "", "Unit": "PCS", "GST Rate": 18.0, "Dealer Rate": 0.0}]
        }
        """
        config = types.GenerateContentConfig(response_mime_type="application/json")
        models = ['gemini-2.5-flash-lite', 'gemini-2.5-flash']

        for img_obj in images_to_process:
            if img_obj.mode in ("RGBA", "P"): img_obj = img_obj.convert("RGB")
            img_obj.thumbnail((700, 700), Image.Resampling.BILINEAR)
            buf = BytesIO()
            img_obj.save(buf, format="JPEG", quality=70, optimize=True)
            raw_bytes = buf.getvalue()

            contents = [Image.open(BytesIO(raw_bytes)), cat_prompt]
            parsed_page = None

            for key in API_KEYS_POOL:
                try:
                    client = genai.Client(api_key=key)
                    for m in models:
                        try:
                            res = _call_gemini_with_retry(client, m, contents, config)
                            text_res = res.text.strip()
                            if text_res.startswith("```json"): text_res = text_res[7:-3].strip()
                            elif text_res.startswith("```"): text_res = text_res[3:-3].strip()
                            parsed_page = json.loads(text_res)
                            break
                        except Exception: continue
                    if parsed_page: break
                except Exception: continue

            if parsed_page:
                supplier_name = str(parsed_page.get("Supplier Name", "GENERAL VENDOR")).strip().upper()
                for prod in parsed_page.get("Products", []):
                    item_name = str(prod.get("Item Name", "")).strip()
                    model_code = str(prod.get("Model Code", "")).strip()
                    full_title = f"{item_name} {model_code}".strip() if model_code and model_code not in item_name else item_name
                    rate = float(prod.get("Dealer Rate") or 0.0)
                    if full_title and rate >= 0:
                        catalog_items.append({
                            "Supplier / Brand": supplier_name,
                            "Model Name / Description": full_title,
                            "Catalog Rate ₹": rate,
                            "GST Rate %": float(prod.get("GST Rate") or 18.0),
                            "Unit": str(prod.get("Unit", "PCS")).upper()
                        })
    except Exception:
        pass
    return catalog_items

# --- PDF QUOTATION GENERATOR ---
def generate_quotation_pdf(store_name: str, phone_str: str, customer_name: str, quote_df: pd.DataFrame, grand_total: float) -> bytes:
    if not REPORTLAB_AVAILABLE: raise ModuleNotFoundError("reportlab missing.")
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=20, leading=24, textColor=colors.HexColor('#1E3A8A'), fontName='Helvetica-Bold')
    subtitle_style = ParagraphStyle('DocSubtitle', parent=styles['Normal'], fontSize=9.5, leading=13, textColor=colors.HexColor('#4B5563'))
    
    story.append(Paragraph("OFFICIAL QUOTATION", title_style))
    story.append(Paragraph(f"<b>Issued By:</b> {store_name.upper()}", subtitle_style))
    if phone_str: story.append(Paragraph(f"<b>Contact:</b> {phone_str}", subtitle_style))
    story.append(Spacer(1, 10))

    table_data = [[Paragraph("S.No", styles['Normal']), Paragraph("Item Description", styles['Normal']), Paragraph("Qty", styles['Normal']), Paragraph("Unit", styles['Normal']), Paragraph("Final Rate (₹)", styles['Normal']), Paragraph("Total (₹)", styles['Normal'])]]
    for i, row in quote_df.iterrows():
        table_data.append([Paragraph(str(i+1), styles['Normal']), Paragraph(str(row["Item Name"]), styles['Normal']), Paragraph(f"{float(row['Quantity']):g}", styles['Normal']), Paragraph(str(row["Unit"]), styles['Normal']), Paragraph(f"₹{float(row['Customer Unit Price (₹)']):,.2f}", styles['Normal']), Paragraph(f"₹{float(row['Total Value (₹)']):,.2f}", styles['Normal'])])
    
    t = Table(table_data, colWidths=[30, 240, 50, 50, 80, 90])
    t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2563EB')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB'))]))
    story.append(t)
    story.append(Spacer(1, 14))
    story.append(Paragraph(f"<b>Grand Total: ₹{grand_total:,.2f}</b>", ParagraphStyle('GT', parent=styles['Heading2'], alignment=2, textColor=colors.HexColor('#1E3A8A'))))
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# --- TABS ---
tab_parser, tab_catalog, tab_master, tab_memory, tab_guide = st.tabs([
    "📥 Batch Invoice Parser", "📚 Price Catalog Extractor", "⚙️ Master Catalog", "📋 Vendor Memory", "📖 Operating Guide"
])

# ==========================================
# TAB 1: BATCH INVOICE PARSER
# ==========================================
with tab_parser:
    sm1, sm2, sm3 = st.columns(3)
    sm1.metric("Master SKUs Registered", len(master_sku_list))
    sm2.metric("Learned Vendor Rules", len(mapping_memory))
    sm3.metric("AI Engine Status", "🟢 Turbo Ready")
    st.divider()

    col_upload, col_info = st.columns([2, 1])
    with col_upload:
        with st.container(border=True):
            st.markdown('<h3><span class="uos-num-badge">1</span>High-Speed Ingestion Dropzone</h3>', unsafe_allow_html=True)
            upload_mode = st.radio("Input Mode:", ["📸 Direct Camera Snap", "📁 File Upload"], horizontal=True, key="d_mode")
            gst_bill_type = st.radio("Tax Mode:", ["Taxable GST Bill", "Non-GST / Net Rate Bill (0% Tax)"], horizontal=True, key="d_tax")
            
            staged_files = []
            if upload_mode == "📸 Direct Camera Snap":
                cam = st.camera_input("Snap Bill", key="cam_snap")
                if cam: staged_files.append((cam.getvalue(), "image/jpeg"))
            else:
                ups = st.file_uploader("Upload Bills", type=["jpg", "jpeg", "png", "pdf"], accept_multiple_files=True, key="up_bills")
                if ups:
                    for f in ups:
                        staged_files.append((f.read(), "application/pdf" if f.name.endswith(".pdf") else "image/jpeg"))

    with col_info:
        with st.container(border=True):
            st.subheader("⚡ Queue Status")
            if staged_files: st.success(f"📁 **{len(staged_files)} File(s)** Ready")
            else: st.info("Queue is empty.")

    if staged_files:
        if st.button("🚀 Process Invoices at Turbo Speed", type="primary", use_container_width=True):
            all_items = []
            with st.status("Executing concurrent parallel multi-threading OCR...", expanded=True) as status:
                with ThreadPoolExecutor(max_workers=6) as executor:
                    raw_results = list(executor.map(process_single_item_tuple, staged_files))
                
                for res in raw_results:
                    if "ERROR" in res: continue
                    sup = res.get("Supplier Company Name", "Unknown")
                    for row in res.get("Line Items", []):
                        qty = float(row.get("Primary Quantity") or 1.0)
                        if qty <= 0: qty = 1.0
                        taxable = float(row.get("Printed Taxable Amount") or 0.0)
                        gst = 0.0 if gst_bill_type == "Non-GST / Net Rate Bill (0% Tax)" else float(row.get("GST Rate") or 18.0)
                        raw_name = str(row.get("Item Name", "")).strip()
                        sku = match_sku(raw_name)
                        all_items.append({
                            "Supplier Name": sup, "Raw Vendor Item": raw_name, "Official SKU": sku,
                            "Current Quantity": qty, "Unit": str(row.get("Unit", "PCS")).upper(),
                            "Purchase Price": round(taxable/qty if qty > 0 else 0, 2), "GST Rate": gst,
                            "Selling Price": round(get_known_selling_price(sku), 2)
                        })
                status.update(label="✅ Turbo Extraction Complete!", state="complete")
            if all_items:
                st.session_state["parsed_df"] = pd.DataFrame(all_items)
                st.rerun()

    if "parsed_df" in st.session_state:
        st.divider()
        st.markdown('<h3><span class="uos-num-badge">2</span>Live Audit Workspace</h3>', unsafe_allow_html=True)
        df = st.session_state["parsed_df"]
        df["Line Total (Excl. GST)"] = (df["Purchase Price"] * df["Current Quantity"]).round(2)
        df["GST Tax Amount"] = (df["Line Total (Excl. GST)"] * (df["GST Rate"] / 100)).round(2)
        df["Unit Cost (GST Paid) ₹"] = ((df["Line Total (Excl. GST)"] + df["GST Tax Amount"]) / df["Current Quantity"]).round(2)

        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="audit_ed")
        st.session_state["parsed_df"] = edited_df

        if st.button("✅ Generate ERP Bulk Import Excel", type="primary", use_container_width=True):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Items"
            ws.cell(row=1, column=1, value=ACTIVE_STORE_DISPLAY)
            headers = ["S. No.", "Name", "Current Quantity", "Unit", "HSN/SAC", "Category", "GST Rate", "Selling Price", "Selling Price (Secondary)", "Purchase Price", "Purchase Price (Secondary)", "Secondary Unit", "Ratio"]
            for col_idx, h in enumerate(headers, 1): ws.cell(row=5, column=col_idx, value=h)
            
            for i, r in edited_df.iterrows():
                r_idx = 6 + i
                ws.cell(row=r_idx, column=1, value=i+1)
                ws.cell(row=r_idx, column=2, value=str(r["Official SKU"]).replace("⚠️ Needs Review: ", ""))
                ws.cell(row=r_idx, column=3, value=float(r["Current Quantity"]))
                ws.cell(row=r_idx, column=4, value=str(r["Unit"]))
                ws.cell(row=r_idx, column=6, value="General")
                ws.cell(row=r_idx, column=7, value=float(r["GST Rate"]))
                ws.cell(row=r_idx, column=8, value=float(r["Selling Price"]))
                ws.cell(row=r_idx, column=10, value=float(r["Purchase Price"]))

            buf = BytesIO()
            wb.save(buf)
            buf.seek(0)
            st.download_button("📥 Download ERP Spreadsheet", data=buf.getvalue(), file_name=f"{selected_store_slug}_ERP_Import.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# ==========================================
# TAB 2: PRICE CATALOG & MODEL EXTRACTOR (TURBO MULTI-FILE GROUPED)
# ==========================================
with tab_catalog:
    st.subheader(f"📚 High-Speed Multi-Vendor Price List Extractor ({ACTIVE_STORE_DISPLAY})")
    st.caption("Drop multiple price lists simultaneously. The system uses high-speed parallel threading and clusters items automatically by company/brand name.")

    col1, col2 = st.columns([2, 1])
    with col1:
        catalog_files = st.file_uploader("Upload Price Lists (PDF or Images)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_cat_up")
    with col2:
        catalog_tax_type = st.radio("Catalog Tax Mode:", ["Taxable GST Rate", "Non-GST / Net Rate"], horizontal=True, key="cat_tax_mode")

    if catalog_files:
        file_tuples = [(f.read(), "application/pdf" if f.name.endswith(".pdf") else "image/jpeg") for f in catalog_files]
        if st.button("🚀 Turbo-Extract All Price Lists in Parallel", type="primary", use_container_width=True, key="btn_turbo_cat"):
            extracted_catalogs = []
            with st.status(f"Executing parallel multi-threaded extraction across {len(file_tuples)} price lists...", expanded=True) as status_box:
                with ThreadPoolExecutor(max_workers=6) as executor:
                    batch_results = list(executor.map(parse_single_catalog_file, file_tuples))
                
                for res_list in batch_results:
                    for item in res_list:
                        if catalog_tax_type == "Non-GST / Net Rate": item["GST Rate %"] = 0.0
                        extracted_catalogs.append(item)
                status_box.update(label=f"✅ Extracted & Organized {len(extracted_catalogs)} items successfully!", state="complete")

            if extracted_catalogs:
                df_cat = pd.DataFrame(extracted_catalogs)
                df_cat_sorted = df_cat.sort_values(by=["Supplier / Brand", "Model Name / Description"]).reset_index(drop=True)
                st.session_state["catalog_df"] = df_cat_sorted
                st.rerun()

    if "catalog_df" in st.session_state and st.session_state["catalog_df"] is not None and not st.session_state["catalog_df"].empty:
        st.divider()
        st.markdown("#### 📋 Brand-Grouped Price List Preview")
        
        edited_cat = st.data_editor(st.session_state["catalog_df"], num_rows="dynamic", use_container_width=True, key="cat_edit_table")
        
        c1, c2 = st.columns(2)
        with c1:
            if st.button("📥 Import Grouped Items into Master Catalog", type="primary", use_container_width=True, key="btn_imp_master"):
                records = []
                for _, r in edited_cat.iterrows():
                    sku_name = f"[{r.get('Supplier / Brand', 'GENERIC')}] {r['Model Name / Description']}".strip()
                    if sku_name:
                        records.append({
                            "Official_SKU_Name": sku_name, "Category": str(r.get("Supplier / Brand", "General")),
                            "Default_Unit": str(r.get("Unit", "PCS")).upper(), "GST_Rate": float(r.get("GST Rate %", 18.0)),
                            "Selling_Price": float(r.get("Catalog Rate ₹", 0.0))
                        })
                if records:
                    bulk = pd.DataFrame(records)
                    comb = pd.concat([master_df, bulk], ignore_index=True).drop_duplicates(subset=["Official_SKU_Name"], keep="last")
                    save_master(comb, selected_store_slug)
                    st.toast("Successfully imported into Master Catalog!")
                    st.rerun()

        with c2:
            wb_c = openpyxl.Workbook()
            ws_c = wb_c.active
            ws_c.title = "Price Lists"
            ws_c.cell(row=1, column=1, value=ACTIVE_STORE_DISPLAY)
            headers_c = ["S. No.", "Supplier / Brand", "Model Name / Description", "Catalog Rate ₹", "GST Rate %", "Unit"]
            for idx, h in enumerate(headers_c, 1): ws_c.cell(row=5, column=idx, value=h)
            
            for i, r in edited_cat.iterrows():
                r_id = 6 + i
                ws_c.cell(row=r_id, column=1, value=i+1)
                ws_c.cell(row=r_id, column=2, value=str(r.get("Supplier / Brand", "")).strip())
                ws_c.cell(row=r_id, column=3, value=str(r["Model Name / Description"]).strip())
                ws_c.cell(row=r_id, column=4, value=float(r.get("Catalog Rate ₹", 0.0)))
                ws_c.cell(row=r_id, column=5, value=float(r.get("GST Rate %", 18.0)))
                ws_c.cell(row=r_id, column=6, value=str(r.get("Unit", "PCS")).upper())

            buf_c = BytesIO()
            wb_c.save(buf_c)
            buf_c.seek(0)
            st.download_button("📥 Download Brand-Grouped Excel File", data=buf_c.getvalue(), file_name=f"{selected_store_slug}_Grouped_Price_Lists.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_cat_excel")

# ==========================================
# TAB 3: MASTER CATALOG
# ==========================================
with tab_master:
    st.subheader(f"⚙️ Master Inventory Catalog ({ACTIVE_STORE_DISPLAY})")
    col_add, col_list = st.columns([1, 2])
    with col_add:
        with st.container(border=True):
            st.markdown("#### ➕ Add Single SKU")
            as_sku = st.text_input("SKU Name")
            as_cat = st.text_input("Category", value="General")
            as_unit = st.selectbox("Unit", ["PCS", "BOX", "LTR", "KG", "NOS", "SET"])
            as_gst = st.selectbox("GST %", [0, 5, 12, 18, 28], index=3)
            as_prc = st.number_input("Selling Price ₹", min_value=0.0)
            if st.button("Save SKU", type="primary", use_container_width=True):
                if as_sku.strip():
                    add_single_sku_direct(selected_store_slug, as_sku.strip(), as_cat, as_unit, float(as_gst), float(as_prc))
                    st.toast("SKU Saved!")
                    st.rerun()
    with col_list:
        with st.container(border=True):
            st.markdown("#### 📋 Catalog Register")
            if not master_df.empty:
                s_query = st.text_input("🔍 Search SKUs...", key="m_search")
                m_view = master_df.copy()
                if s_query:
                    m_view = m_view[m_view["Official_SKU_Name"].str.lower().str.contains(s_query.lower())]
                m_view.insert(0, "Del", False)
                ed_m = st.data_editor(m_view, num_rows="fixed", hide_index=True, use_container_width=True, key="master_ed")
                to_del = ed_m[ed_m["Del"] == True]["Official_SKU_Name"].tolist()
                if to_del and st.button(f"🗑️ Delete Selected ({len(to_del)})", type="primary"):
                    delete_multiple_skus(selected_store_slug, to_del)
                    st.rerun()
            else: st.info("Catalog is empty.")

# ==========================================
# TAB 4: VENDOR MEMORY & TAB 5: GUIDE
# ==========================================
with tab_memory:
    st.subheader("🧠 Learned AI Vendor Memory")
    if mapping_memory:
        st.dataframe(pd.DataFrame([{"Raw Vendor Item": k, "Mapped SKU": v} for k, v in mapping_memory.items()]), use_container_width=True)
        if st.button("Reset Memory"):
            engine = get_db_engine()
            store_id = get_or_create_store_id(selected_store_slug)
            with engine.begin() as conn: conn.execute(text("DELETE FROM vendor_mappings WHERE store_id = :store_id"), {"store_id": store_id})
            load_json_memory.clear()
            st.rerun()
    else: st.info("No memory logs recorded.")

with tab_guide:
    st.subheader("📖 Standard Operating Procedure")
    st.markdown("Use **Tab 2** to drop multiple price lists in parallel. They will automatically be sorted by company/brand name and outputted into organized Excel spreadsheets.")
