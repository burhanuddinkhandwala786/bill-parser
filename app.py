import os
import gc
import json
import time
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
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
    :root {
        --uos-primary: #0F172A;
        --uos-primary-hover: #1E293B;
        --uos-accent: #2563EB;
        --uos-accent-subtle: #EFF6FF;
        --uos-success: #10B981;
        --uos-canvas: #FAFAFA;
        --uos-surface: #FFFFFF;
        --uos-surface-hover: #F8FAFC;
        --uos-border: #E4E4E7;
        --uos-border-subtle: #F4F4F5;
        --uos-text: #09090B;
        --uos-text-secondary: #52525B;
        --uos-text-muted: #71717A;
        --uos-radius: 10px;
        --uos-radius-lg: 14px;
        --uos-shadow-subtle: 0 1px 3px 0 rgba(0, 0, 0, 0.02), 0 1px 2px -1px rgba(0, 0, 0, 0.02);
        --uos-shadow-card: 0 4px 6px -1px rgba(0, 0, 0, 0.02), 0 2px 4px -2px rgba(0, 0, 0, 0.02);
    }

    /* Global Typography & Canvas */
    html, body, [class*="css"], .stApp, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
        -webkit-font-smoothing: antialiased;
        background-color: var(--uos-canvas) !important;
        color: var(--uos-text) !important;
    }
    [data-testid="stHeader"] { background: transparent !important; }
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 4rem !important;
        max-width: 1320px !important;
    }
    div[data-testid="InputInstructions"] { display: none !important; }

    /* Premium Headings */
    h1, h2, h3, h4, h5, h6, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
        font-family: 'Inter', sans-serif !important;
        color: var(--uos-text) !important;
        letter-spacing: -0.03em !important;
        font-weight: 700 !important;
    }
    p, span, label, [data-testid="stWidgetLabel"] label {
        color: var(--uos-text-secondary) !important;
        font-size: 0.9rem !important;
    }
    [data-testid="stCaptionContainer"], small {
        color: var(--uos-text-muted) !important;
    }
    hr, [data-testid="stDivider"] {
        border-color: var(--uos-border) !important;
        margin: 1.5rem 0 !important;
    }

    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: var(--uos-surface) !important;
        border-right: 1px solid var(--uos-border) !important;
    }
    [data-testid="stSidebar"] hr {
        border-color: var(--uos-border-subtle) !important;
    }

    /* Clean Card Containers (No Patchy Overlaps) */
    [data-testid="stVerticalBlockBorderWrapper"], [data-testid="stForm"], [data-testid="stExpander"] {
        background-color: var(--uos-surface) !important;
        border: 1px solid var(--uos-border) !important;
        border-radius: var(--uos-radius-lg) !important;
        padding: 20px !important;
        box-shadow: var(--uos-shadow-subtle) !important;
    }

    /* High-End Metric Cards */
    [data-testid="stMetric"] {
        background-color: var(--uos-surface) !important;
        border: 1px solid var(--uos-border) !important;
        border-radius: var(--uos-radius-lg) !important;
        padding: 20px !important;
        box-shadow: var(--uos-shadow-subtle) !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.75rem !important;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: var(--uos-text-muted) !important;
        font-weight: 600 !important;
    }
    [data-testid="stMetricValue"] div {
        color: var(--uos-text) !important;
        font-weight: 800 !important;
        letter-spacing: -0.03em !important;
    }

    /* Clean Buttons */
    .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
        font-family: 'Inter', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.88rem !important;
        border-radius: var(--uos-radius) !important;
        padding: 10px 18px !important;
        border: 1px solid var(--uos-border) !important;
        background: var(--uos-surface) !important;
        color: var(--uos-text) !important;
        box-shadow: var(--uos-shadow-subtle) !important;
        transition: all 0.15s ease !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: var(--uos-surface-hover) !important;
        border-color: var(--uos-text-muted) !important;
    }
    .stButton > button[kind="primary"], .stFormSubmitButton > button[kind="primary"] {
        background: var(--uos-primary) !important;
        color: #FFFFFF !important;
        border-color: var(--uos-primary) !important;
        box-shadow: 0 2px 4px rgba(15, 23, 42, 0.1) !important;
    }
    .stButton > button[kind="primary"] p { color: #FFFFFF !important; }
    .stButton > button[kind="primary"]:hover {
        background: var(--uos-primary-hover) !important;
    }

    /* Input Fields & Forms */
    .stTextInput input, .stNumberInput input, .stTextArea textarea, 
    div[data-baseweb="input"] input, div[data-baseweb="select"] > div {
        background-color: var(--uos-surface) !important;
        color: var(--uos-text) !important;
        border: 1px solid var(--uos-border) !important;
        border-radius: var(--uos-radius) !important;
        padding: 10px 14px !important;
        box-shadow: none !important;
    }
    .stTextInput input:focus, .stNumberInput input:focus {
        border-color: var(--uos-accent) !important;
        box-shadow: 0 0 0 3px var(--uos-accent-subtle) !important;
    }

    /* Data Tables & Editors */
    [data-testid="stDataFrame"], [data-testid="stDataEditor"] {
        border: 1px solid var(--uos-border) !important;
        border-radius: var(--uos-radius) !important;
        background-color: var(--uos-surface) !important;
        overflow: hidden;
    }
    .stDataFrame [role="columnheader"] {
        background-color: var(--uos-surface-2) !important;
        font-weight: 600 !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 1px solid var(--uos-border);
        background: transparent !important;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 16px !important;
        background: transparent !important;
        font-weight: 500 !important;
        font-size: 0.9rem !important;
        color: var(--uos-text-muted) !important;
        border-bottom: 2px solid transparent !important;
    }
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        color: var(--uos-text) !important;
        font-weight: 700 !important;
        border-bottom-color: var(--uos-text) !important;
    }

    /* File Uploader Dropzone */
    [data-testid="stFileUploaderDropzone"] {
        background-color: var(--uos-surface-2) !important;
        border: 1px dashed var(--uos-border) !important;
        border-radius: var(--uos-radius-lg) !important;
    }
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

@st.cache_data(ttl=600)
def get_store_phone(store_slug: str) -> str:
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
    engine = get_db_engine()
    try:
        with engine.begin() as conn:
            conn.execute(
                text("UPDATE stores SET phone = :phone WHERE slug = :slug"),
                {"phone": phone.strip(), "slug": store_slug}
            )
        get_store_phone.clear()
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

# --- FAST SESSION & AUTO-LOGIN RESOLUTION ---
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "user_store" not in st.session_state:
    st.session_state["user_store"] = None

if not st.session_state["authenticated"]:
    url_session = st.query_params.get("session", None)
    if url_session:
        if isinstance(url_session, list):
            url_session = url_session[0]
        
        store_data = get_user_by_slug(str(url_session))
        if store_data:
            st.session_state["authenticated"] = True
            st.session_state["user_store"] = store_data

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
                login_email = st.text_input("Store Email", key="login_email_input")
                login_password = st.text_input("Password", type="password", key="login_pass_input")
                submit_login = st.form_submit_button("Log In", use_container_width=True, type="primary")
                
                if submit_login:
                    if login_email and login_password:
                        store_data = authenticate_user(login_email, login_password)
                        if store_data:
                            st.session_state["authenticated"] = True
                            st.session_state["user_store"] = store_data
                            st.query_params["session"] = store_data["slug"]
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

# --- SIDEBAR ---
st.sidebar.title(f"🏬 {st.session_state['user_store']['display_name']}")
st.sidebar.caption(f"Active Store ID: `{selected_store_slug}`")
st.sidebar.divider()

if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state["authenticated"] = False
    st.session_state["user_store"] = None
    st.query_params.clear()
    if "parsed_df" in st.session_state:
        del st.session_state["parsed_df"]
    st.rerun()

st.sidebar.divider()

# --- MULTI-KEY & MULTI-MODEL FAIL-SAFE API POOL SETUP ---
def get_api_key_pool():
    keys = []
    for i in range(1, 6):
        key = st.secrets.get(f"GEMINI_API_KEY_{i}") or os.environ.get(f"GEMINI_API_KEY_{i}")
        if key:
            keys.append(key)
            
    if not keys:
        single_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
        if single_key:
            keys.append(single_key)
            
    return keys

API_KEYS_POOL = get_api_key_pool()

if not API_KEYS_POOL:
    st.error("⚠️ Gemini API Key missing. Please configure GEMINI_API_KEY in secrets.")
    st.stop()

# --- HEADER SECTION ---
header_left, header_right = st.columns([3, 1.5])

with header_left:
    st.title("⚡ Universal OS")
    st.caption("Commercial Multi-Store AI Purchase Intake & Inventory Synchronizer")

with header_right:
    st.caption("ACTIVE STORE CATALOG")
    st.markdown(f"📍 **{ACTIVE_STORE_DISPLAY}**")

st.divider()

# --- FAST CACHED DATABASE STORE LOADERS ---
@st.cache_data(ttl=3600)
def load_json_memory(store_slug: str) -> dict:
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
    load_json_memory.clear()

@st.cache_data(ttl=3600)
def load_master(store_slug: str) -> pd.DataFrame:
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
    load_master.clear()

# --- ATOMIC OPTIMIZED SQL EXECUTORS ---
def add_single_sku_direct(store_slug: str, sku_name: str, category: str, unit: str, gst_rate: float, selling_price: float):
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    
    with engine.begin() as conn:
        conn.execute(
            text("""
                INSERT INTO master_skus (store_id, official_sku_name, category, default_unit, gst_rate, selling_price)
                VALUES (:store_id, :official_sku_name, :category, :default_unit, :gst_rate, :selling_price)
                ON CONFLICT (store_id, official_sku_name)
                DO UPDATE SET category = EXCLUDED.category, default_unit = EXCLUDED.default_unit, gst_rate = EXCLUDED.gst_rate, selling_price = EXCLUDED.selling_price
            """),
            {
                "store_id": store_id,
                "official_sku_name": sku_name,
                "category": category,
                "default_unit": unit,
                "gst_rate": gst_rate,
                "selling_price": selling_price
            }
        )
    load_master.clear()

def bulk_upsert_audited_skus(store_slug: str, records: list):
    if not records:
        return
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    
    # SINGLE ATOMIC BULK TRANSACTION
    with engine.begin() as conn:
        for rec in records:
            conn.execute(
                text("""
                    INSERT INTO master_skus (store_id, official_sku_name, category, default_unit, gst_rate, selling_price)
                    VALUES (:store_id, :official_sku_name, :category, :default_unit, :gst_rate, :selling_price)
                    ON CONFLICT (store_id, official_sku_name)
                    DO UPDATE SET category = EXCLUDED.category, default_unit = EXCLUDED.default_unit, gst_rate = EXCLUDED.gst_rate, selling_price = EXCLUDED.selling_price
                """),
                {
                    "store_id": store_id,
                    "official_sku_name": rec["official"],
                    "category": rec["category"],
                    "default_unit": rec["unit"],
                    "gst_rate": rec["gst"],
                    "selling_price": rec["sp"]
                }
            )
    load_master.clear()

def delete_multiple_skus(store_slug: str, sku_list: list):
    if not sku_list:
        return
    engine = get_db_engine()
    store_id = get_or_create_store_id(store_slug)
    
    # SAFE TUPLE PARAMETER BINDING (FIXES POSTGRES ARRAY SYNTAX BUG)
    with engine.begin() as conn:
        conn.execute(
            text("DELETE FROM master_skus WHERE store_id = :store_id AND official_sku_name = ANY(:sku_names)"),
            {"store_id": store_id, "sku_names": list(sku_list)}
        )
    load_master.clear()

master_df = load_master(selected_store_slug)
master_sku_list = master_df["Official_SKU_Name"].dropna().tolist() if not master_df.empty else []
mapping_memory = load_json_memory(selected_store_slug)

def match_sku(raw_name):
    cleaned_raw = raw_name.strip().upper()
    if cleaned_raw in mapping_memory:
        return mapping_memory[cleaned_raw]
    if master_sku_list:
        match, score, _ = process.extractOne(raw_name, master_sku_list, processor=utils.default_process)
        if score > 65:
            return match
    return raw_name

def get_known_selling_price(sku_name):
    if not master_df.empty and "Selling_Price" in master_df.columns:
        matched = master_df[master_df["Official_SKU_Name"] == sku_name]
        if not matched.empty:
            price = matched.iloc[0]["Selling_Price"]
            if pd.notnull(price) and float(price) > 0:
                return float(price)
    return 0.0

# --- FAIL-SAFE AI ENGINE WITH MULTI-KEY ROTATION ---
def is_server_error(exception):
    err_str = str(exception).lower()
    return "503" in err_str or "unavailable" in err_str or "overloaded" in err_str or "429" in err_str or "resourceexhausted" in err_str

@retry(
    stop=stop_after_attempt(2),
    wait=wait_exponential(multiplier=1, min=1, max=3),
    retry=retry_if_exception(is_server_error),
    reraise=True
)
def _call_gemini_with_retry(client, model_name, contents, config):
    return client.models.generate_content(
        model=model_name,
        contents=contents,
        config=config
    )

def extract_invoice_data_multiformat(file_bytes, mime_type="image/jpeg"):
    if "pdf" in mime_type.lower():
        file_part = types.Part.from_bytes(data=file_bytes, mime_type="application/pdf")
        contents = [file_part]
    else:
        img = Image.open(BytesIO(file_bytes))
        img_copy = img.copy()
        if img_copy.mode in ("RGBA", "P"):
            img_copy = img_copy.convert("RGB")
        img_copy.thumbnail((1024, 1024), Image.Resampling.BILINEAR)
        
        buffer = BytesIO()
        img_copy.save(buffer, format="JPEG", quality=85, optimize=True)
        buffer.seek(0)
        contents = [Image.open(buffer)]

    prompt = """
    You are an enterprise financial OCR system for wholesale, retail, plywood, hardware, and building material invoices.

    CRITICAL EXTRACTION & GROUND-TRUTH RULES:
    1. "Supplier Company Name": Main vendor/seller title from bill top header.
    2. "Invoice Number": Invoice/Bill number string if present, else "".
    3. "Invoice Date": Date string if present, else "".
    4. "Line Items": Extract every product row accurately.
       - "Item Name": Full product title or description. Read handwritten notes and pen edits carefully.
       - "Primary Quantity": Pure numeric count of physical pieces/sheets/boxes received (e.g., 6.0, 1.0, 270.12).
       - "Unit": Unit string (PCS, SQM, SQFT, BOX, KG, LTR, NOS, SET). Default to "PCS".
       - "Printed Taxable Amount": ABSOLUTE GROUND TRUTH. Extract the printed line subtotal/amount value before tax directly from the bill's "Amount" column (e.g. 7305.60, 63543.00, 2000.00). DO NOT multiply physical pieces * secondary unit rate!
       - "GST Rate": Total GST percentage as a pure number (0, 5, 12, 18, 28). If CGST (9%) & SGST (9%) are listed separately, SUM THEM to 18.0. Default to 18.0 if unstated.
       - "HSN Code": HSN/SAC code as string. If missing, "".

    OUTPUT SCHEMA (STRICT JSON ONLY):
    {
        "Supplier Company Name": "Vendor Title",
        "Invoice Number": "",
        "Invoice Date": "",
        "Line Items": [
            {
                "Item Name": "Description",
                "Primary Quantity": 1.0,
                "Unit": "PCS",
                "Printed Taxable Amount": 0.0,
                "GST Rate": 18.0,
                "HSN Code": ""
            }
        ]
    }
    """
    
    contents.append(prompt)
    config = types.GenerateContentConfig(response_mime_type="application/json")
    candidate_models = ['gemini-3.5-flash-lite', 'gemini-2.5-flash-lite', 'gemini-3.5-flash', 'gemini-2.5-flash']
    last_error = None
    
    for api_key in API_KEYS_POOL:
        try:
            client = genai.Client(api_key=api_key)
            for model_name in candidate_models:
                try:
                    response = _call_gemini_with_retry(client, model_name, contents, config)
                    text_res = response.text.strip()
                    if text_res.startswith("```json"):
                        text_res = text_res[7:-3].strip()
                    elif text_res.startswith("```"):
                        text_res = text_res[3:-3].strip()
                    return json.loads(text_res)
                except Exception as model_err:
                    last_error = model_err
                    continue
        except Exception as key_err:
            last_error = key_err
            continue
            
    raise Exception(f"AI Service error across all keys and models: {last_error}")

def process_single_item_tuple(item_tuple):
    file_bytes, mime_type = item_tuple
    try:
        return extract_invoice_data_multiformat(file_bytes, mime_type)
    except Exception as e:
        return {"ERROR": str(e)}

# --- REPORTLAB PDF GENERATOR ---
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

    hdr_left = ParagraphStyle('HdrL', fontSize=8.5, leading=11, textColor=colors.white, fontName='Helvetica-Bold', alignment=0)
    hdr_center = ParagraphStyle('HdrC', fontSize=8.5, leading=11, textColor=colors.white, fontName='Helvetica-Bold', alignment=1)
    hdr_right = ParagraphStyle('HdrR', fontSize=8.5, leading=11, textColor=colors.white, fontName='Helvetica-Bold', alignment=2)

    cell_left = ParagraphStyle('CellL', fontSize=8, leading=11, textColor=colors.HexColor('#111827'), alignment=0)
    cell_center = ParagraphStyle('CellC', fontSize=8, leading=11, textColor=colors.HexColor('#111827'), alignment=1)
    cell_right = ParagraphStyle('CellR', fontSize=8, leading=11, textColor=colors.HexColor('#111827'), alignment=2)

    table_data = [[
        Paragraph("S.No", hdr_center), 
        Paragraph("Item Description", hdr_left), 
        Paragraph("Qty", hdr_center), 
        Paragraph("Unit", hdr_center), 
        Paragraph("MRP (Rs.)", hdr_right),
        Paragraph("Disc (%)", hdr_center),
        Paragraph("Final Rate (Rs.)", hdr_right), 
        Paragraph("Total (Rs.)", hdr_right)
    ]]

    for i, row in quote_df.iterrows():
        mrp_val = float(row.get('MRP (₹)', row.get('Customer Unit Price (₹)', 0.0)))
        disc_val = float(row.get('Discount (%)', 0.0))
        final_rate = float(row['Customer Unit Price (₹)'])
        line_total = float(row['Total Value (₹)'])
        
        table_data.append([
            Paragraph(str(i + 1), cell_center),
            Paragraph(str(row["Item Name"]), cell_left),
            Paragraph(f"{float(row['Quantity']):g}", cell_center),
            Paragraph(str(row["Unit"]), cell_center),
            Paragraph(f"Rs. {mrp_val:,.2f}", cell_right),
            Paragraph(f"{disc_val:g}%" if disc_val > 0 else "-", cell_center),
            Paragraph(f"Rs. {final_rate:,.2f}", cell_right),
            Paragraph(f"Rs. {line_total:,.2f}", cell_right)
        ])

    item_table = Table(table_data, colWidths=[25, 175, 35, 35, 70, 45, 75, 80])
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
# TAB 1: BATCH INVOICE PARSER & QUOTATIONS
# ==========================================
with tab_parser:
    sm1, sm2, sm3 = st.columns(3)
    sm1.metric("Master SKUs Registered", len(master_sku_list))
    sm2.metric("Learned Vendor Rules", len(mapping_memory))
    sm3.metric("AI Engine Status", "🟢 Ready")

    st.divider()

    # --- EASY STREAMLINED QUOTATION GENERATOR ---
    with st.expander("📄 Generate Customer Quotation (On-the-go)", expanded=False):
        st.caption("Create a professional PDF quote instantly either by manual quick entry or auto-calculating from an audited purchase bill.")
        
        saved_phone = st.session_state["user_store"].get("phone") or get_store_phone(selected_store_slug)
        q_tab_manual, q_tab_bill = st.tabs(["✍️ Quick Manual Entry", "📥 Auto-Quote from Parsed Bill"])
        
        # --- MODE 1: MANUAL QUICK QUOTATION ---
        with q_tab_manual:
            col_m1, col_m2 = st.columns([2, 2])
            with col_m1:
                man_cust_name = st.text_input("Customer Name / Reference", value="Walk-in Customer", key="man_cust_input")
            with col_m2:
                man_phone_input = st.text_input("Contact Phone", value=saved_phone, key="man_phone_input")
                
            if man_phone_input.strip() != saved_phone.strip():
                save_store_phone(selected_store_slug, man_phone_input.strip())
                st.session_state["user_store"]["phone"] = man_phone_input.strip()

            col_mk1, col_mk2 = st.columns([1, 1])
            with col_mk1:
                manual_markup_pct = st.number_input("Profit Markup (% On Base MRP/Cost)", min_value=0.0, value=0.0, step=1.0, key="man_markup_input")
            with col_mk2:
                manual_disc_pct = st.number_input("Overall Discount (%)", min_value=0.0, max_value=100.0, value=0.0, step=1.0, key="man_disc_input")

            st.write("**Enter Quotation Items:**")
            
            if "manual_quote_data" not in st.session_state:
                st.session_state["manual_quote_data"] = pd.DataFrame([
                    {"Item Name": "Plywood 18mm Commercial", "Quantity": 2.0, "Unit": "PCS", "MRP / Base Rate (₹)": 1800.0},
                ])
                
            edited_manual_df = st.data_editor(
                st.session_state["manual_quote_data"],
                num_rows="dynamic",
                use_container_width=True,
                key="manual_quote_editor",
                column_config={
                    "Item Name": st.column_config.TextColumn("Item Description", required=True),
                    "Quantity": st.column_config.NumberColumn("Qty", min_value=0.01, format="%.2f", default=1.0),
                    "Unit": st.column_config.SelectboxColumn("Unit", options=["PCS", "BOX", "LTR", "KG", "NOS", "SET", "MTR", "SQM", "PKT", "BTL"], default="PCS"),
                    "MRP / Base Rate (₹)": st.column_config.NumberColumn("MRP / Base Rate (₹)", min_value=0.0, format="₹%.2f", default=0.0),
                }
            )
            
            st.session_state["manual_quote_data"] = edited_manual_df.copy()
            
            calc_manual_df = edited_manual_df.copy()
            calc_manual_df["Quantity"] = pd.to_numeric(calc_manual_df["Quantity"], errors='coerce').fillna(1.0)
            calc_manual_df["MRP (₹)"] = pd.to_numeric(calc_manual_df.get("MRP / Base Rate (₹)", 0.0), errors='coerce').fillna(0.0)
            calc_manual_df["Discount (%)"] = manual_disc_pct
            
            marked_up_mrp = (calc_manual_df["MRP (₹)"] * (1 + (manual_markup_pct / 100))).round(2)
            calc_manual_df["MRP (₹)"] = marked_up_mrp
            
            calc_manual_df["Customer Unit Price (₹)"] = (calc_manual_df["MRP (₹)"] * (1 - (calc_manual_df["Discount (%)"] / 100))).round(2)
            calc_manual_df["Total Value (₹)"] = (calc_manual_df["Quantity"] * calc_manual_df["Customer Unit Price (₹)"]).round(2)
            
            man_grand_total = calc_manual_df["Total Value (₹)"].sum()
            st.metric("Quotation Total", f"₹{man_grand_total:,.2f}")
            
            if st.button("📄 Generate & Download PDF Quotation", type="primary", use_container_width=True, key="btn_man_quote"):
                if calc_manual_df.empty or man_grand_total <= 0:
                    st.warning("Please enter at least one valid item with a price.")
                else:
                    if REPORTLAB_AVAILABLE:
                        pdf_bytes = generate_quotation_pdf(
                            store_name=st.session_state['user_store']['display_name'],
                            phone_str=man_phone_input.strip(),
                            customer_name=man_cust_name,
                            quote_df=calc_manual_df,
                            grand_total=man_grand_total
                        )
                        st.download_button(
                            label=f"⬇️ Click Here to Download PDF for {man_cust_name}",
                            data=pdf_bytes,
                            file_name=f"Quotation_{man_cust_name.replace(' ', '_')}.pdf",
                            mime="application/pdf",
                            type="primary",
                            use_container_width=True,
                            key="dl_man_pdf"
                        )
                    else:
                        st.warning("⚠️ 'reportlab' library is missing in environment.")

        # --- MODE 2: AUTO-QUOTATION FROM PARSED BILL ---
        with q_tab_bill:
            if "parsed_df" not in st.session_state or st.session_state["parsed_df"].empty:
                st.info("ℹ️ No parsed bill found. Upload a bill below or use 'Quick Manual Entry' above.")
            else:
                col_q1, col_q2, col_q3, col_q4 = st.columns([2, 2, 1, 1])
                with col_q1:
                    customer_name = st.text_input("Customer Name", value="Walk-in Customer", key="bill_cust_input")
                with col_q2:
                    biz_phone_input = st.text_input("Phone", value=saved_phone, key="bill_phone_input")
                with col_q3:
                    markup_pct = st.number_input("Profit Markup %", min_value=0.0, value=15.0, step=1.0, key="bill_markup_input")
                with col_q4:
                    disc_pct_bill = st.number_input("Discount %", min_value=0.0, max_value=100.0, value=0.0, step=1.0, key="bill_disc_input")
                
                if st.button("📄 Generate PDF Quote from Ingested Bill", type="primary", use_container_width=True, key="btn_bill_quote"):
                    df_bill_source = st.session_state["parsed_df"].copy()
                    quote_items = []
                    total_quote_value = 0.0
                    
                    for idx, row in df_bill_source.iterrows():
                        base_gst_paid_cost = float(row.get("Unit Cost (GST Paid) ₹", 0.0))
                        mrp_price = round(base_gst_paid_cost * (1 + (markup_pct / 100)), 2)
                        final_price = round(mrp_price * (1 - (disc_pct_bill / 100)), 2)
                        qty = float(row.get("Current Quantity", 1.0))
                        line_total = round(final_price * qty, 2)
                        
                        total_quote_value += line_total
                        
                        quote_items.append({
                            "Item Name": str(row.get("Official SKU", "")),
                            "Quantity": qty,
                            "Unit": str(row.get("Unit", "PCS")),
                            "MRP (₹)": mrp_price,
                            "Discount (%)": disc_pct_bill,
                            "Customer Unit Price (₹)": final_price,
                            "Total Value (₹)": line_total
                        })
                    
                    quote_df = pd.DataFrame(quote_items)
                    st.metric("Total Quotation Value", f"₹{total_quote_value:,.2f}")
                    
                    if REPORTLAB_AVAILABLE:
                        pdf_bytes = generate_quotation_pdf(
                            store_name=st.session_state['user_store']['display_name'],
                            phone_str=biz_phone_input.strip(),
                            customer_name=customer_name,
                            quote_df=quote_df,
                            grand_total=total_quote_value
                        )
                        st.download_button(
                            label=f"⬇️ Click Here to Download PDF for {customer_name}",
                            data=pdf_bytes,
                            file_name=f"Quotation_{customer_name.replace(' ', '_')}.pdf",
                            mime="application/pdf",
                            type="primary",
                            use_container_width=True,
                            key="dl_bill_pdf"
                        )

    st.divider()

    # --- INGESTION DROPZONE ---
    col_upload, col_info = st.columns([2, 1])
    
    with col_upload:
        with st.container(border=True):
            st.subheader("1. Ingestion Dropzone")
            st.caption("Upload PNG, JPG, or PDF files or click a photo directly using your phone/webcam.")
            
            c_mode, c_tax = st.columns([1, 1])
            with c_mode:
                upload_mode = st.radio(
                    "Input Mode:", 
                    ["📸 Direct Camera Snap", "📁 File Upload"], 
                    horizontal=True,
                    key="dropzone_mode_switch"
                )
            with c_tax:
                gst_bill_type = st.radio(
                    "Bill Tax Status:", 
                    ["Taxable GST Bill", "Non-GST / Net Rate Bill (0% Tax)"], 
                    horizontal=True,
                    key="gst_bill_type_toggle",
                    help="Non-GST overrides tax rate to 0% across all extracted items."
                )
            
            staged_file_tuples = []
            
            if upload_mode == "📸 Direct Camera Snap":
                cam_photo = st.camera_input("Take a photo of the purchase bill", key="direct_cam_input")
                if cam_photo is not None:
                    staged_file_tuples.append((cam_photo.getvalue(), "image/jpeg"))
            else:
                uploaded_files = st.file_uploader(
                    "Upload Bills (PDF or Images)",
                    type=["jpg", "jpeg", "png", "pdf"],
                    accept_multiple_files=True,
                    label_visibility="collapsed",
                    key="batch_file_uploader"
                )
                if uploaded_files:
                    for f in uploaded_files:
                        mtype = "application/pdf" if f.name.lower().endswith(".pdf") else "image/jpeg"
                        staged_file_tuples.append((f.read(), mtype))

    with col_info:
        with st.container(border=True):
            st.subheader("⚡ Ingestion Queue")
            if staged_file_tuples:
                st.success(f"📁 **{len(staged_file_tuples)} File(s)** Staged & Ready")
                tax_status_str = "0% Tax Override" if gst_bill_type == "Non-GST / Net Rate Bill (0% Tax)" else "Standard Tax Extraction"
                st.caption(f"Mode: **{tax_status_str}**")
            else:
                st.info("No files staged.")
                st.caption("Snap a photo or drop purchase bills to begin structured extraction.")

    if staged_file_tuples:
        st.write("")
        c_btn1, c_btn2 = st.columns([3, 1])
        with c_btn1:
            process_btn = st.button("🚀 Process Invoices with Fast AI Engine", type="primary", use_container_width=True)
        with c_btn2:
            if st.button("🧹 Clear Queue", use_container_width=True):
                if "parsed_df" in st.session_state:
                    del st.session_state["parsed_df"]
                st.rerun()

        if process_btn:
            if "parsed_df" in st.session_state:
                del st.session_state["parsed_df"]
                
            all_parsed_items = []
            
            with st.status("Parsing purchase bills concurrently with Multimodal AI...", expanded=True) as status_container:
                with ThreadPoolExecutor(max_workers=min(len(staged_file_tuples), 10)) as executor:
                    raw_results = list(executor.map(process_single_item_tuple, staged_file_tuples))
                    
                for parsed_json in raw_results:
                    if "ERROR" in parsed_json:
                        st.error(f"Failed to process a file: {parsed_json['ERROR']}")
                        continue
                        
                    supplier = parsed_json.get("Supplier Company Name", "Unknown Supplier")
                    inv_num = parsed_json.get("Invoice Number", "")
                    inv_date = parsed_json.get("Invoice Date", "")
                    
                    for row in parsed_json.get("Line Items", []):
                        qty = float(row.get("Primary Quantity") or 1.0)
                        if qty <= 0: qty = 1.0
                        
                        if gst_bill_type == "Non-GST / Net Rate Bill (0% Tax)":
                            gst_rate = 0.0
                        else:
                            gst_rate = float(row.get("GST Rate") or 18.0)
                            
                        printed_taxable = float(row.get("Printed Taxable Amount") or 0.0)
                        hsn_sac = str(row.get("HSN Code") or "").strip()
                        
                        base_rate_per_pc = printed_taxable / qty if qty > 0 else 0.0
                        
                        raw_item_name = str(row.get("Item Name", "")).strip()
                        matched_sku = match_sku(raw_item_name)
                        known_selling = get_known_selling_price(matched_sku)
                        
                        all_parsed_items.append({
                            "Supplier Name": supplier,
                            "Invoice No": inv_num,
                            "Invoice Date": inv_date,
                            "Raw Vendor Item": raw_item_name,
                            "Official SKU": matched_sku,
                            "Current Quantity": qty,
                            "Unit": str(row.get("Unit", "PCS")).upper(),
                            "HSN/SAC": hsn_sac,
                            "Category": "General",
                            "GST Rate": gst_rate,
                            "Purchase Price": round(base_rate_per_pc, 2),
                            "Line Total Taxable": round(printed_taxable, 2),
                            "Selling Price": round(known_selling, 2)
                        })
                    
                status_container.update(label="✅ Ingestion & Batch Extraction Complete!", state="complete", expanded=False)
                
            if all_parsed_items:
                st.session_state["parsed_df"] = pd.DataFrame(all_parsed_items)
                gc.collect()
                st.rerun()

    # --- REVIEW & EDIT WORKSPACE ---
    if "parsed_df" in st.session_state:
        st.divider()
        st.subheader("2. Live Inventory Audit Workspace")
        st.caption("Verify AI extraction, mapped SKUs, and landed costs before exporting to accounting software.")
        
        df = st.session_state["parsed_df"]
        
        # SELF-HEALING SCHEMA GUARANTEE
        required_cols = [
            "Raw Vendor Item", "Official SKU", "Current Quantity", "Unit", "HSN/SAC",
            "Purchase Price", "GST Rate", "Selling Price", "Category"
        ]
        for c in required_cols:
            if c not in df.columns:
                df[c] = "" if c in ["Raw Vendor Item", "Official SKU", "Unit", "HSN/SAC", "Category"] else 0.0

        df["Current Quantity"] = pd.to_numeric(df["Current Quantity"], errors='coerce').fillna(1.0)
        df["Purchase Price"] = pd.to_numeric(df["Purchase Price"], errors='coerce').fillna(0.0)
        df["GST Rate"] = pd.to_numeric(df["GST Rate"], errors='coerce').fillna(18.0)
        df["Selling Price"] = pd.to_numeric(df["Selling Price"], errors='coerce').fillna(0.0)
        
        df["Line Total (Excl. GST)"] = (df["Purchase Price"] * df["Current Quantity"]).round(2)
        df["GST Tax Amount"] = (df["Line Total (Excl. GST)"] * (df["GST Rate"] / 100)).round(2)
        df["Line Total (Incl. GST)"] = (df["Line Total (Excl. GST)"] + df["GST Tax Amount"]).round(2)
        df["Unit Cost (GST Paid) ₹"] = (df["Line Total (Incl. GST)"] / df["Current Quantity"]).round(2)
        
        total_taxable = df["Line Total (Excl. GST)"].sum()
        total_gst = df["GST Tax Amount"].sum()
        subtotal_incl_tax = total_taxable + total_gst

        uom_groups = df.groupby("Unit")["Current Quantity"].sum()
        uom_summary_str = " | ".join([f"{val:,.2f} {unit}" for unit, val in uom_groups.items()])

        c_m1, c_m2, c_m3, c_m4 = st.columns([1, 1, 1, 1])
        c_m1.metric("Total Line Items", f"{len(df)} Items")
        c_m2.metric("Stock Quantities by UOM", uom_summary_str)
        c_m3.metric("Taxable Base (Excl. GST)", f"₹{total_taxable:,.2f}")
        
        round_off_val = st.sidebar.number_input("Bill Round-Off Adjustment (₹)", value=0.0, step=0.05, format="%.2f", help="Adjust to match paper invoice total exactly.")
        final_bill_total = subtotal_incl_tax + round_off_val
        c_m4.metric("Grand Total (GST Paid)", f"₹{final_bill_total:,.2f}", delta=f"GST Tax: ₹{total_gst:,.2f}")
        
        st.write("")
        
        display_columns = [
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
        
        edited_display_df = st.data_editor(
            df[display_columns],
            num_rows="dynamic",
            use_container_width=True,
            key="audit_editor",
            column_config={
                "Raw Vendor Item": st.column_config.TextColumn("Raw Vendor Item", disabled=True),
                "Official SKU": st.column_config.SelectboxColumn("Official SKU Name", options=master_sku_list, required=True) if master_sku_list else "Official SKU",
                "Current Quantity": st.column_config.NumberColumn("Qty (Pcs/Sheets)", min_value=0.01, format="%.2f"),
                "Unit": st.column_config.TextColumn("Unit"),
                "HSN/SAC": st.column_config.TextColumn("HSN/SAC"),
                "Purchase Price": st.column_config.NumberColumn("Unit Rate (Excl. GST) ₹", format="₹%.2f"),
                "Line Total (Excl. GST)": st.column_config.NumberColumn("Total (Excl. GST) ₹", format="₹%.2f", disabled=True),
                "GST Rate": st.column_config.NumberColumn("GST %", min_value=0, max_value=28, format="%d%%"),
                "Unit Cost (GST Paid) ₹": st.column_config.NumberColumn("Landed Cost (GST Paid) ₹/Pc", format="₹%.2f", disabled=True),
                "Line Total (Incl. GST)": st.column_config.NumberColumn("Total (Incl. GST) ₹", format="₹%.2f", disabled=True),
                "Selling Price": st.column_config.NumberColumn("Selling Price (SP) ₹", format="₹%.2f"),
            }
        )
        
        df_updated = edited_display_df.copy()
        df_updated["Current Quantity"] = pd.to_numeric(df_updated["Current Quantity"], errors='coerce').fillna(1.0)
        df_updated["Purchase Price"] = pd.to_numeric(df_updated["Purchase Price"], errors='coerce').fillna(0.0)
        df_updated["GST Rate"] = pd.to_numeric(df_updated["GST Rate"], errors='coerce').fillna(18.0)
        df_updated["Selling Price"] = pd.to_numeric(df_updated.get("Selling Price", 0.0), errors='coerce').fillna(0.0)

        df_updated["Line Total (Excl. GST)"] = (df_updated["Purchase Price"] * df_updated["Current Quantity"]).round(2)
        df_updated["GST Tax Amount"] = (df_updated["Line Total (Excl. GST)"] * (df_updated["GST Rate"] / 100)).round(2)
        df_updated["Line Total (Incl. GST)"] = (df_updated["Line Total (Excl. GST)"] + df_updated["GST Tax Amount"]).round(2)
        df_updated["Unit Cost (GST Paid) ₹"] = (df_updated["Line Total (Incl. GST)"] / df_updated["Current Quantity"]).round(2)
        
        st.session_state["parsed_df"] = df_updated

        st.divider()
        if st.button("✅ Confirm Audit & Generate Excel Import File", type="primary", use_container_width=True):
            memory_updated = False
            upsert_records = []
            
            for idx, row in df_updated.iterrows():
                raw = str(row.get("Raw Vendor Item", "")).strip().upper()
                official = str(row.get("Official SKU", "")).strip()
                cat_val = str(row.get("Category", "General"))
                unit_val = str(row.get("Unit", "PCS"))
                gst_val = float(row.get("GST Rate", 18.0))
                sp_val = float(row.get("Selling Price", 0.0))
                
                if raw and official and raw != official:
                    mapping_memory[raw] = official
                    memory_updated = True
                
                if official:
                    upsert_records.append({
                        "official": official,
                        "category": cat_val,
                        "unit": unit_val,
                        "gst": gst_val,
                        "sp": sp_val
                    })

            # FAST BATCH TRANSACTION (ELIMINATES DB LOOP FREEZES)
            if upsert_records:
                bulk_upsert_audited_skus(selected_store_slug, upsert_records)

            if memory_updated:
                save_json_memory(selected_store_slug, mapping_memory)
                st.toast("🧠 Learned Vendor Mapping updated!")
                
            st.toast("⚙️ Master SKU catalog updated & synchronized!")
                
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
                raw_selling = row.get("Selling Price", 0.0)
                selling_val = float(raw_selling) if pd.notnull(raw_selling) and float(raw_selling) > 0 else ""
                purchase_val = float(row.get("Purchase Price", 0.0)) if pd.notnull(row.get("Purchase Price")) else 0.0
                
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

# ==========================================
# TAB 2: STORE MASTER CATALOG MANAGER
# ==========================================
with tab_master:
    st.subheader(f"⚙️ Master Inventory Catalog ({ACTIVE_STORE_DISPLAY})")
    
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
    
    # --- INSTANT SINGLE SKU SAVER ---
    with col_add:
        with st.container(border=True):
            st.markdown("#### ➕ Add Single Master SKU")
            add_sku = st.text_input("SKU Name (e.g. Copper Wire 1.5mm)")
            add_cat = st.text_input("Category", value="General")
            add_unit = st.selectbox("Default Unit", options=["PCS", "BOX", "LTR", "KG", "NOS", "SET"])
            add_gst = st.selectbox("GST Rate (%)", options=[0, 5, 12, 18, 28], index=3)
            add_price = st.number_input("Selling Price ₹ (Optional)", min_value=0.0, step=10.0)
            
            if st.button("⚡ Save SKU to Catalog", use_container_width=True, type="primary"):
                if add_sku.strip():
                    clean_sku = add_sku.strip()
                    add_single_sku_direct(selected_store_slug, clean_sku, add_cat, add_unit, float(add_gst), float(add_price))
                    st.toast(f"⚡ Saved '{clean_sku}' instantly!")
                    st.rerun()
                    
    # --- COMPACT CHECKBOX MULTI-DELETE TABLE WITH LIVE SEARCH ---
    with col_list:
        with st.container(border=True):
            st.markdown("#### 📋 Catalog Register")
            if not master_df.empty:
                search_query = st.text_input("🔍 Search Catalog SKUs...", placeholder="Type name or category...", key="catalog_search_bar")
                
                df_catalog_view = master_df.copy()
                if search_query.strip():
                    q = search_query.strip().lower()
                    mask = (
                        df_catalog_view["Official_SKU_Name"].astype(str).str.lower().str.contains(q) |
                        df_catalog_view["Category"].astype(str).str.lower().str.contains(q)
                    )
                    df_catalog_view = df_catalog_view[mask]

                df_catalog_view.insert(0, "Del", False)
                
                edited_catalog_df = st.data_editor(
                    df_catalog_view,
                    num_rows="fixed",
                    hide_index=True,
                    use_container_width=True,
                    key="master_catalog_editor",
                    column_config={
                        "Del": st.column_config.CheckboxColumn("Del", default=False, width="small"),
                        "Official_SKU_Name": st.column_config.TextColumn("Official SKU Name", disabled=True),
                        "Category": st.column_config.TextColumn("Category", disabled=True),
                        "Default_Unit": st.column_config.TextColumn("Unit", disabled=True, width="small"),
                        "GST_Rate": st.column_config.NumberColumn("GST %", format="%d%%", disabled=True, width="small"),
                        "Selling_Price": st.column_config.NumberColumn("Selling Price ₹", format="₹%.2f", disabled=True)
                    }
                )
                
                selected_for_delete = edited_catalog_df[edited_catalog_df["Del"] == True]["Official_SKU_Name"].tolist()
                
                if selected_for_delete:
                    st.write("")
                    if st.button(f"🗑️ Delete Selected ({len(selected_for_delete)} SKUs)", type="primary", use_container_width=True):
                        delete_multiple_skus(selected_store_slug, selected_for_delete)
                        st.toast(f"Deleted {len(selected_for_delete)} SKU(s)!")
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
            load_json_memory.clear()
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
        2. **Upload Bills:** Drop images, PDFs, or click a photo directly with your camera in **Tab 1**. Select whether the bill is Taxable GST or Non-GST.
        3. **Run AI Engine:** Click **Run AI Invoice Parsing Engine** to extract structured line items.
        4. **Audit Workspace:** Check quantities, HSN codes, landed purchase rates, and mapped SKUs.
        5. **Generate Quote (Optional):** Open the Quotation expander to quickly quote a customer with your Profit Markup (%).
        6. **Download Import File:** Generate the `.xlsx` spreadsheet.
        7. **Import to ERP:** Open your accounting software → **Items / Inventory** → **Bulk Import**, upload the `.xlsx` file.
        """)
