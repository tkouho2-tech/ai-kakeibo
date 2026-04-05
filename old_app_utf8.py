import streamlit as st
import jpholiday
import pandas as pd
import plotly.express as px
import gspread
import bcrypt
import os
import json
import io
import calendar
from datetime import datetime
from dateutil.relativedelta import relativedelta
from PIL import Image
from google import genai
from google.genai import types
import streamlit.components.v1 as components
import time

# ---------- 讒区・險ｭ螳・----------
SPREADSHEET_NAME = "Kakeibo_Data" # 螳滄圀縺ｮGoogle繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝亥錐縺ｫ蜷医ｏ縺帙※螟画峩縺励※縺上□縺輔＞
WORKSHEET_NAME = "users"
TRANSACTIONS_WORKSHEET_NAME = "transactions"

# ---------- 繧ｫ繝・ざ繝ｪ螳夂ｾｩ ----------
# AI蛻､蛻･繧・そ繝ｬ繧ｯ繝医・繝・け繧ｹ縺ｧ蛻ｩ逕ｨ縺吶ｋ縺溘ａ縺ｮ螟ｧ蛻・｡槭・蟆丞・鬘槭・隕ｪ蟄宣未菫ょｮ夂ｾｩ
EXPENSE_CATEGORIES = {
    "鬟滓攝雋ｻ": ["獄閧蛾｡・, "澄鬲夐｡・, "･ｬ驥手除譫懃黄", "忽荳ｻ鬟滄｡・, "些諠｣闖・, "･壼嵯荵ｳ陬ｽ蜩・, "･ｫ蜉蟾･鬟溷刀", "ｧりｪｿ蜻ｳ譁・, "梱蝸懷･ｽ蜩・, "笘暮｣ｲ譁・, "笶薙◎縺ｮ莉・],
    "螟夜｣溯ｲｻ": ["骨繝ｩ繝ｼ繝｡繝ｳ", "坤蜥碁｣・, "･｡荳ｭ闖ｯ", "黒繧､繧ｿ繝ｪ繧｢繝ｳ", "笘輔き繝輔ぉ", "瑳鬟ｲ驟・, "笶薙◎縺ｮ莉・],
    "譌･逕ｨ蜩・: ["ｧｻ豸郁怜刀", "ｧｺ謗・勁豢玲ｿｯ", "寫・剰｢句桁陬・, "笶薙◎縺ｮ莉・],
    "鄒主ｮｹ": ["ｧｴ繧ｱ繧｢逕ｨ蜩・, "嫡蛹也ｲｧ蜩・, "笨ゑｸ乗淵鬮ｪ", "笶薙◎縺ｮ莉・],
    "陦｣鬘・: ["装陦｣鬘・, "臓髱ｴ", "ｧ｣蟆冗黄", "笶薙◎縺ｮ莉・],
    "螳ｶ髮ｻ": ["銅螳ｶ髮ｻ", "捗蜻ｨ霎ｺ讖溷勣", "笶薙◎縺ｮ莉・],
    "譖ｸ邀・: ["答譖ｸ邀・, "槙・乗枚蜈ｷ", "笶薙◎縺ｮ莉・],
    "莠､騾夊ｲｻ": ["噬蜈ｬ蜈ｱ莠､騾・, "囓霆翫ち繧ｯ繧ｷ繝ｼ", "笵ｽ繧ｬ繧ｽ繝ｪ繝ｳ", "笶薙◎縺ｮ莉・],
    "菴丞ｱ・: ["寞・丞ｮｶ蜈ｷ", "匠菴丞ｱ・畑蜩・, "笶薙◎縺ｮ莉・],
    "螽ｯ讌ｽ": ["治螽ｯ讌ｽ", "耳繧ｰ繝・ぜ", "笶薙◎縺ｮ莉・],
    "謇区焚譁・: ["逃騾∵侭", "諜謇区焚譁・, "笶薙◎縺ｮ莉・],
    "繝壹ャ繝育畑蜩・: ["粋繝輔・繝・, "埒繝医う繝ｬ逕ｨ蜩・, "唱繝壹ャ繝亥現逋・, "笶薙◎縺ｮ莉・],
    "蛹ｻ逋・: ["唱逞・劼險ｺ逋・, "抽阮ｬ蜃ｦ譁ｹ", "忠讀懈渊蛛･險ｺ", "笶薙◎縺ｮ莉・],
    "蝨定敢繝ｻ讀咲黄": ["現闍励・遞ｮ", "ｪｴ隕ｳ闡画､咲黄", "ｧｱ蝨溘・閧･譁吶・驩｢", "屏・丞恍闃ｸ逕ｨ蜩・, "笶薙◎縺ｮ莉・],
    "蜑ｲ蠑輔・繝昴う繝ｳ繝亥茜逕ｨ": ["蜈ｱ騾壹・繧､繝ｳ繝亥茜逕ｨ", "蠎苓・迢ｬ閾ｪ繝昴う繝ｳ繝亥茜逕ｨ", "繧ｯ繝ｼ繝昴Φ蜑ｲ蠑・, "繧ｭ繝｣繝・す繝･繝舌ャ繧ｯ繝ｻ驍・・"],
    "豸郁ｲｻ遞・: ["8%", "10%", "笶薙◎縺ｮ莉・],
    "縺昴・莉・: ["刀譛ｪ蛻・｡・]
}

def get_categories_prompt_text():
    """AI・・emini遲会ｼ峨・繝励Ο繝ｳ繝励ヨ縺ｫ蝓九ａ霎ｼ繧縺溘ａ縺ｮ繧ｫ繝・ざ繝ｪ螳夂ｾｩ譁・ｭ怜・繧堤函謌・""
    text = "縲舌き繝・ざ繝ｪ繧ｷ繧ｹ繝・Β: 螟ｧ蛻・｡槭→蟆丞・鬘槭・繝ｪ繧ｹ繝医曾n"
    for major, minors in EXPENSE_CATEGORIES.items():
        text += f"- {major}: {', '.join(minors)}\n"
    text += "\n窶ｻ 蠢・★荳願ｨ倥・螟ｧ蛻・｡槭→蟆丞・鬘槭・邨・∩蜷医ｏ縺帙↓蠕薙▲縺ｦ縺上□縺輔＞縲・
    return text

# ---------- 繧ｻ繝・す繝ｧ繝ｳ迥ｶ諷九・蛻晄悄蛹・----------
if 'genai_client' not in st.session_state:
    st.session_state['genai_client'] = None

# API繧ｭ繝ｼ險ｭ螳夲ｼ・emini逕ｨ・・# 縺ｩ縺｡繧峨・譖ｸ縺肴婿縺ｧ繧ょ虚縺上ｈ縺・↓縺励∪縺・api_key = st.secrets.get("GEMINI_API_KEY") or st.secrets.get("general", {}).get("gemini_api_key")

if not api_key and "general" in st.secrets:
    api_key = st.secrets["general"].get("gemini_api_key")

if api_key:
    st.session_state['genai_client'] = genai.Client(api_key=api_key)

st.set_page_config(page_title="AI螳ｶ險育ｰｿ繧｢繝励Μ - 繝繝・す繝･繝懊・繝・, page_icon="投", layout="wide")

# 繝悶Λ繧ｦ繧ｶ縺ｮ閾ｪ蜍慕ｿｻ險ｳ縺ｮ隱､菴懷虚繧帝亟縺舌◆繧√？TML縺ｮ險隱櫁ｨｭ螳壹→鄙ｻ險ｳ諡貞凄險ｭ螳壹ｒ蠑ｷ蛹・components.html(
    """
    <script>
        // HTML隕∫ｴ縺ｮ險隱槭ｒ譌･譛ｬ隱槭↓蝗ｺ螳・        const html = window.parent.document.getElementsByTagName('html')[0];
        html.setAttribute('lang', 'ja');
        html.setAttribute('translate', 'no');
        html.classList.add('notranslate');

        // head縺ｫmeta繧ｿ繧ｰ繧呈諺蜈･縺励※Google鄙ｻ險ｳ繧呈・遉ｺ逧・↓諡貞凄
        const head = window.parent.document.getElementsByTagName('head')[0];
        if (!window.parent.document.querySelector('meta[name="google"][content="notranslate"]')) {
            const meta = window.parent.document.createElement('meta');
            meta.name = 'google';
            meta.content = 'notranslate';
            head.appendChild(meta);
        }
    </script>
    """,
    width=0,
    height=0,
)

# 蜈ｨ菴薙・繧ｹ繧ｿ繧､繝ｫ縺ｨ縺励※繧らｿｻ險ｳ諡貞凄繧帝←逕ｨ
st.markdown("""
    <style>
        /* 繧｢繝励Μ蜈ｨ菴薙〒鄙ｻ險ｳ繧堤┌蜉ｹ蛹・*/
        .stApp {
            unicode-bidi: isolate;
        }
        /* className繧剃ｽｿ縺｣縺滓拠蜷ｦ・井ｸ驛ｨ繝悶Λ繧ｦ繧ｶ蜷代￠・・*/
        .notranslate {
            translate: no !important;
        }
    </style>
""", unsafe_allow_html=True)

# ---------- 繧ｻ繝・す繝ｧ繝ｳ迥ｶ諷九・蛻晄悄蛹・----------
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'username' not in st.session_state:
    st.session_state['username'] = None
if 'current_month' not in st.session_state:
    st.session_state['current_month'] = datetime.today().replace(day=1)

# ---------- Google Sheets 謗･邯・----------
@st.cache_resource
def get_gspread_client():
    try:
        # 1. secrets.toml 縺九ｉ諠・ｱ繧定ｪｭ縺ｿ霎ｼ繧
        if "gcp_service_account" in st.secrets:
            # 霎樊嶌蠖｢蠑上↓螟画鋤・育判蜒・縺ｮ菫ｮ豁｣繧帝←逕ｨ・・            info = dict(st.secrets["gcp_service_account"])

            from google.oauth2.service_account import Credentials
            
            # 繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝域桃菴懊↓蠢・ｦ√↑讓ｩ髯撰ｼ医せ繧ｳ繝ｼ繝暦ｼ・            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            # 隱崎ｨｼ諠・ｱ縺ｮ菴懈・
            creds = Credentials.from_service_account_info(info, scopes=scopes)
            return gspread.authorize(creds)
            
        elif os.path.exists("credentials.json"):
            return gspread.service_account(filename="credentials.json")
            
        else:
            st.error("隱崎ｨｼ險ｭ螳壹′隕九▽縺九ｊ縺ｾ縺帙ｓ縲・)
            return None
            
    except Exception as e:
        st.error(f"隱崎ｨｼ繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: {e}")
        return None

# --- 繝ｪ繝医Λ繧､蜿ｯ閭ｽ縺ｪAPI蜻ｼ縺ｳ蜃ｺ縺励・繝ｫ繝代・ ---

def safe_gspread_call(func, *args, max_retries=3, delay=2, **kwargs):
    """API蜻ｼ縺ｳ蜃ｺ縺励ｒ繝ｪ繝医Λ繧､縺吶ｋ髢｢謨ｰ"""
    last_error = None
    for i in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            last_error = e
            # 荳譎ら噪縺ｪ謗･邯壹お繝ｩ繝ｼ縺ｮ蝣ｴ蜷医↓繝ｪ繝医Λ繧､
            if "RemoteDisconnected" in str(e) or "Connection aborted" in str(e) or "TimeoutError" in str(e):
                time.sleep(delay * (i + 1)) # 謖・焚繝舌ャ繧ｯ繧ｪ繝慕噪縺ｫ蠕・ｩ・                continue
            else:
                # 閾ｴ蜻ｽ逧・↑繧ｨ繝ｩ繝ｼ・郁ｪ崎ｨｼ遲会ｼ峨・縺吶＄縺ｫ荳翫￡繧・                raise e
    raise last_error

def safe_gemini_call(func, *args, max_retries=5, initial_delay=2, **kwargs):
    """Gemini API蜻ｼ縺ｳ蜃ｺ縺励ｒ繝ｪ繝医Λ繧､縺吶ｋ髢｢謨ｰ・・29/500/503繧ｨ繝ｩ繝ｼ蟇ｾ蠢懶ｼ・""
    last_error = None
    for i in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            last_error = e
            err_msg = str(e)
            # 429 RESOURCE_EXHAUSTED 縺ｾ縺溘・ 500/503 邉ｻ繧ｨ繝ｩ繝ｼ縺ｮ蝣ｴ蜷医↓繝ｪ繝医Λ繧､
            if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg or "500" in err_msg or "503" in err_msg:
                wait_time = initial_delay * (2 ** i) # 謖・焚繝舌ャ繧ｯ繧ｪ繝・                st.warning(f"迴ｾ蝨ｨ豺ｷ縺ｿ蜷医▲縺ｦ縺・∪縺呻ｼ・i+1}/{max_retries}蝗樒岼・峨・wait_time}遘貞ｾ後↓蜀崎ｩｦ陦後＠縺ｾ縺・..")
                time.sleep(wait_time)
                continue
            else:
                raise e
    raise last_error

def get_sheet(worksheet_name):
    client = get_gspread_client()
    if client is None:
        st.stop()
        
    try:
        # 繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医→謖・ｮ壹Ρ繝ｼ繧ｯ繧ｷ繝ｼ繝医↓謗･邯夲ｼ医Μ繝医Λ繧､莉倥″・・        def _open_sheet():
            return client.open(SPREADSHEET_NAME).worksheet(worksheet_name)
        
        return safe_gspread_call(_open_sheet)
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"繧ｨ繝ｩ繝ｼ: 繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝・'{SPREADSHEET_NAME}' 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲ゅけ繝ｬ繝・Φ繧ｷ繝｣繝ｫ縺ｮ繝｡繝ｼ繝ｫ繧｢繝峨Ξ繧ｹ ({client.auth.signer_email}) 縺ｨ繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医ｒ蜈ｱ譛峨＠縺ｦ縺上□縺輔＞縲・)
        st.stop()
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"繧ｨ繝ｩ繝ｼ: 繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝亥・縺ｫ '{worksheet_name}' 繧ｷ繝ｼ繝医′隕九▽縺九ｊ縺ｾ縺帙ｓ縲ゅす繝ｼ繝医ｒ譁ｰ隕丈ｽ懈・縺励※縺上□縺輔＞縲・)
        st.stop()
    except Exception as e:
        st.error(f"Google Sheets謗･邯壹お繝ｩ繝ｼ: {e}")
        st.stop()

def init_users_sheet(sheet):
    """蛻晄悄繧ｻ繝・ヨ繧｢繝・・・壹・繝・ム繝ｼ縺後↑縺・ｴ蜷医↓菴懈・縺吶ｋ"""
    try:
        headers = sheet.row_values(1)
        if not headers or headers[0] != "username":
            sheet.insert_row(["username", "password_hash"], 1)
    except Exception:
        # 繧ｷ繝ｼ繝医′遨ｺ縺ｮ蝣ｴ蜷医↓萓句､悶′逋ｺ逕溘☆繧句庄閭ｽ諤ｧ縺後≠繧九◆繧√√◎縺ｮ蝣ｴ蜷医・繝倥ャ繝繝ｼ繧定ｿｽ蜉
        sheet.insert_row(["username", "password_hash"], 1)

def init_transactions_sheet(sheet):
    """蛻晄悄繧ｻ繝・ヨ繧｢繝・・・壼叙蠑輔す繝ｼ繝医・繝倥ャ繝繝ｼ縺後↑縺・ｴ蜷医↓菴懈・縺吶ｋ"""
    try:
        headers = sheet.row_values(1)
        if not headers or headers[0] != "username":
            sheet.insert_row(["username", "date", "store_name", "item_name", "category", "subcategory", "amount"], 1)
        # 譌｢蟄倥す繝ｼ繝医〒 subcategory 蛻励′縺ｪ縺・ｴ蜷医〒繧ゅ・・ｬ｡霑ｽ蜉縺ｧ蟇ｾ蠢懷庄閭ｽ縺ｨ縺吶ｋ
    except Exception:
        sheet.insert_row(["username", "date", "store_name", "item_name", "category", "subcategory", "amount"], 1)

# ---------- 隱崎ｨｼ讖溯・ ----------
def register_user(username, password):
    sheet = get_sheet(WORKSHEET_NAME)
    init_users_sheet(sheet)
    
    # 莉墓ｧ倩ｦ∽ｻｶ: 繝ｦ繝ｼ繧ｶ繝ｼ蜷阪・ lower() 縺ｧ蜃ｦ逅・☆繧・    username = username.strip().lower()
    
    if not username or not password:
        return False, "繝ｦ繝ｼ繧ｶ繝ｼ蜷阪→繝代せ繝ｯ繝ｼ繝峨ｒ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・
        
    # 譌｢蟄倥Θ繝ｼ繧ｶ繝ｼ縺ｮ驥崎､・メ繧ｧ繝・け
    records = sheet.get_all_records()
    for row in records:
        if str(row.get("username", "")).lower() == username:
            return False, "縺薙・繝ｦ繝ｼ繧ｶ繝ｼ蜷阪・譌｢縺ｫ逋ｻ骭ｲ縺輔ｌ縺ｦ縺・∪縺吶・
            
    # 莉墓ｧ倩ｦ∽ｻｶ: 繝代せ繝ｯ繝ｼ繝峨・ bcrypt 縺ｧ繝上ャ繧ｷ繝･蛹悶☆繧・    salt = bcrypt.gensalt()
    hashed_password = bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')
    
    # 逋ｻ骭ｲ繝・・繧ｿ縺ｮ霑ｽ蜉
    sheet.append_row([username, hashed_password])
    return True, "逋ｻ骭ｲ縺悟ｮ御ｺ・＠縺ｾ縺励◆縲ゅΟ繧ｰ繧､繝ｳ繧ｿ繝悶°繧峨Ο繧ｰ繧､繝ｳ縺励※縺上□縺輔＞縲・

def authenticate_user(username, password):
    sheet = get_sheet(WORKSHEET_NAME)
    init_users_sheet(sheet)
    
    # 繝ｦ繝ｼ繧ｶ繝ｼ蜷咲・蜷医・縺溘ａ蟆乗枚蟄怜喧
    username = username.strip().lower()
    
    records = sheet.get_all_records()
    for row in records:
        if str(row.get("username", "")).lower() == username:
            stored_hash = str(row.get("password_hash", ""))
            # bcrypt 縺ｧ繝上ャ繧ｷ繝･繧堤・蜷・            if bcrypt.checkpw(password.encode('utf-8'), stored_hash.encode('utf-8')):
                return True
    return False

# ---------- 繝・・繧ｿ蜿門ｾ玲ｩ溯・ (蜈ｱ騾壹・繝ｫ繝代・) ----------
def get_clean_df(records, username):
    """
    繝ｬ繧ｳ繝ｼ繝峨°繧吋ataFrame繧剃ｽ懈・縺励√き繝ｩ繝蜷阪・豁｣隕丞喧(譌･譛ｬ隱槫ｯｾ蠢・縺ｨ繝ｦ繝ｼ繧ｶ繝ｼ繝輔ぅ繝ｫ繧ｿ繝ｪ繝ｳ繧ｰ繧定｡後≧
    """
    if not records:
        return pd.DataFrame()
        
    df = pd.DataFrame(records)
    
    # 繧ｯ繝ｬ繝ｳ繧ｸ繝ｳ繧ｰ・井ｸ崎ｦ√↑繧ｹ繝壹・繧ｹ遲牙炎髯､縲∝ｰ乗枚蟄怜喧縺励※繝槭ャ繝√Φ繧ｰ縺励ｄ縺吶￥縺吶ｋ・・    df.columns = df.columns.str.strip()
    
    # 譌｢蟄倥・繧ｫ繝ｩ繝蜷搾ｼ亥ｰ乗枚蟄暦ｼ峨→繧ｿ繝ｼ繧ｲ繝・ヨ縺ｮ繝槭ャ繝励ｒ菴懈・
    col_map = {c.lower(): c for c in df.columns}
    
    # --- 繧ｫ繝ｩ繝蜷阪・豁｣隕丞喧・域律譛ｬ隱槭・繝・ム繝ｼ繝ｻ螟ｧ譁・ｭ怜ｰ乗枚蟄励∈縺ｮ蟇ｾ蠢懶ｼ・---
    rename_rules = {
        "譌･莉・: "date",
        "date": "date",
        "繝ｦ繝ｼ繧ｶ繝ｼ蜷・: "username",
        "username": "username",
        "user": "username",
        "蠎苓・蜷・: "store_name",
        "蠎苓・": "store_name",
        "蝠・刀蜷・: "item_name",
        "蜀・ｮｹ": "item_name",
        "驥鷹｡・: "amount",
        "螟ｧ蛻・｡・: "category",
        "蟆丞・鬘・: "subcategory"
    }
    
    actual_rename = {}
    for key, target in rename_rules.items():
        # key縺後き繝ｩ繝蜷搾ｼ医◎縺ｮ縺ｾ縺ｾ縲√∪縺溘・蟆乗枚蟄暦ｼ峨↓蜷ｫ縺ｾ繧後※縺・ｋ縺狗｢ｺ隱・        if key in df.columns:
            actual_rename[key] = target
        elif key.lower() in col_map:
            actual_rename[col_map[key.lower()]] = target
            
    if actual_rename:
        df = df.rename(columns=actual_rename)
    
    # "username"縺ｧ繝輔ぅ繝ｫ繧ｿ (蠢・・
    if "username" in df.columns:
        # 蛟､閾ｪ菴薙・菴咏區繧ょ炎髯､縺励※豈碑ｼ・        df["username"] = df["username"].astype(str).str.strip().str.lower()
        df = df[df["username"] == username.lower()]
    else:
        return pd.DataFrame()
        
    if df.empty or "date" not in df.columns:
          return pd.DataFrame()

    # "date"蛻励ｒdatetime蝙九↓螟画鋤
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"])
    
    return df

# ---------- 繝・・繧ｿ蜿門ｾ玲ｩ溯・ ----------
def load_transactions_data(target_date, mode="monthly"):
    """
    謖・ｮ壹＠縺滓怦縺ｾ縺溘・蟷ｴ縺ｮ縲√Ο繧ｰ繧､繝ｳ繝ｦ繝ｼ繧ｶ繝ｼ縺ｮ繝・・繧ｿ繧貞叙蠕励☆繧・    mode: "monthly" (譛域ｬ｡) 縺ｾ縺溘・ "yearly" (蟷ｴ谺｡)
    """
    sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
    init_transactions_sheet(sheet)
    # 繝ｬ繧ｳ繝ｼ繝牙叙蠕励↓繝ｪ繝医Λ繧､繧帝←逕ｨ
    records = safe_gspread_call(sheet.get_all_records)
    
    # 蜈ｱ騾壹・繝ｫ繝代・縺ｧ繧ｯ繝ｬ繝ｳ繧ｸ繝ｳ繧ｰ縺ｨ繝輔ぅ繝ｫ繧ｿ繝ｪ繝ｳ繧ｰ
    curr_user = st.session_state.get('username', "")
    df = get_clean_df(records, curr_user)
    
    if df.empty:
         return pd.DataFrame()
    
    # 陦後う繝ｳ繝・ャ繧ｯ繧ｹ縺ｮ莉倅ｸ趣ｼ・ecords縺ｮ鬆・分縺ｫ蝓ｺ縺･縺擾ｼ・    # records縺ｮ繧､繝ｳ繝・ャ繧ｯ繧ｹ縺ｨdf縺ｮ繧､繝ｳ繝・ャ繧ｯ繧ｹ繧貞粋繧上○繧句ｿ・ｦ√′縺ゅｋ縺溘ａ縲√け繝ｬ繝ｳ繧ｸ繝ｳ繧ｰ蜑阪・records髟ｷ繧剃ｽｿ逕ｨ
    # records縺ｯ蜈ｨ繝ｦ繝ｼ繧ｶ繝ｼ蛻・≠繧九′縲‥f縺ｯ繝輔ぅ繝ｫ繧ｿ貂医∩縲・    # records縺ｫ縺ゅｋ蜈・・陦檎分蜿ｷ繧剃ｿ晄戟縺吶ｋ縺溘ａ縺ｫDataFrame菴懈・譎ゅ↓莉倅ｸ弱＠縺ｦ縺翫￥
    df_all_temp = pd.DataFrame(records)
    df_all_temp['_row_index'] = range(2, len(records) + 2)
    
    # df縺ｫrow_index繧堤ｵ仙粋
    # pd.merge繧剃ｽｿ縺・◆繧√∝・縺ｮ繧､繝ｳ繝・ャ繧ｯ繧ｹ繧貞茜逕ｨ
    df = df.join(df_all_temp[['_row_index']])
    
    # 譛滄俣縺ｧ繝輔ぅ繝ｫ繧ｿ
    if mode == "monthly":
        df = df[(df["date"].dt.year == target_date.year) & (df["date"].dt.month == target_date.month)]
    else:  # yearly
        df = df[df["date"].dt.year == target_date.year]
    
    # 驥鷹｡阪ｒ謨ｰ蛟､縺ｫ螟画鋤
    if "amount" in df.columns:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    else:
        df["amount"] = 0
        
    # --- 繧ｫ繝・ざ繝ｪ縺ｮ豁｣隕丞喧・磯寔險域凾繧・そ繝ｬ繧ｯ繝医・繝・け繧ｹ遲峨〒謖・ｮ壼､悶′蜃ｺ縺ｪ縺・ｈ縺・↓縺吶ｋ・・---
    # 螟ｧ蛻・｡槭・豁｣隕丞喧
    if "category" in df.columns:
        valid_majors = list(EXPENSE_CATEGORIES.keys())
        # 螳夂ｾｩ縺ｫ縺ｪ縺・､ｧ蛻・｡槭・縲後◎縺ｮ莉悶阪↓縺ｾ縺ｨ繧√ｋ
        df["category"] = df["category"].apply(lambda x: x if x in valid_majors else "縺昴・莉・)
        
    # 蟆丞・鬘槭・豁｣隕丞喧
    sub_cols = [c for c in ["subcategory", "sub_category", "蟆丞・鬘・] if c in df.columns]
    if sub_cols:
        sub_col = sub_cols[0]
        def normalize_sub(row):
            major = row.get("category", "縺昴・莉・)
            sub = str(row.get(sub_col, "")).strip()
            valid_subs = EXPENSE_CATEGORIES.get(major, sorted(EXPENSE_CATEGORIES["縺昴・莉・]))
            fallback = valid_subs[-1] if valid_subs else "笶薙◎縺ｮ莉・
            
            # 螳悟・縺ｫ荳閾ｴ縺吶ｋ縺・            if sub in valid_subs:
                return sub
                
            # 繧｢繧､繧ｳ繝ｳ縺ｪ縺励↑縺ｩ縺ｮ驛ｨ蛻・ｸ閾ｴ繧呈爾縺・            for v_sub in valid_subs:
                # 邨ｵ譁・ｭ励ｒ髯､縺・◆繝・く繧ｹ繝医〒驛ｨ蛻・ｸ閾ｴ縺吶ｋ縺狗｢ｺ隱・                text_only = "".join([c for c in v_sub if c.isalnum() or c in "鬘樒黄鬟溷刀譛ｪ蛻・｡槭◎縺ｮ莉・"]) 
                if text_only and (text_only in sub or sub in text_only) and len(sub) > 0:
                    return v_sub
                    
            return fallback

        df[sub_col] = df.apply(normalize_sub, axis=1)
        
    return df

def get_transaction_range(username):
    """繝ｦ繝ｼ繧ｶ繝ｼ縺ｮ蜈ｨ繝・・繧ｿ縺ｮ譛蟆丞ｹｴ譛医→譛螟ｧ蟷ｴ譛医ｒ蜿門ｾ励＠縲√そ繝・す繝ｧ繝ｳ縺ｫ菫晄戟縺吶ｋ"""
    # 菫ｮ豁｣繧貞叉譎ょ渚譏縺輔○繧九◆繧√∽ｸ譎ら噪縺ｫ繧ｭ繝｣繝・す繝･繧堤┌蜉ｹ蛹悶☆繧九°縲∝ｼｷ蛻ｶ繝ｪ繝輔Ξ繝・す繝･繧呈検繧
    # 縺薙％縺ｧ縺ｯ縲√ｂ縺嶺ｸ肴紛蜷医′縺ゅｌ縺ｰ蜀榊叙蠕励☆繧九ｈ縺・↓繧ｬ繝ｼ繝峨ｒ蠑ｷ蛹・    if 'date_range' in st.session_state and st.session_state.get('last_range_fetch_user') == username:
        if st.session_state['date_range']: # 遨ｺ縺ｧ縺ｪ縺・％縺ｨ繧堤｢ｺ隱・            return st.session_state['date_range']
    
    sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
    records = safe_gspread_call(sheet.get_all_records)
    
    # 蜈ｱ騾壹・繝ｫ繝代・縺ｧ繧ｯ繝ｬ繝ｳ繧ｸ繝ｳ繧ｰ縺ｨ繝輔ぅ繝ｫ繧ｿ繝ｪ繝ｳ繧ｰ
    df_user = get_clean_df(records, username)
    
    if df_user.empty:
        return []
        
    # 繝ｦ繝九・繧ｯ縺ｪ蟷ｴ譛医ｒ謚ｽ蜃ｺ縺励※繧ｽ繝ｼ繝・    df_user["year_month"] = df_user["date"].dt.to_period("M").dt.to_timestamp()
    available_months = sorted(df_user["year_month"].unique().tolist())
    
    st.session_state['date_range'] = available_months
    st.session_state['last_range_fetch_user'] = username
    return available_months

def render_year_navigation():
    """蟷ｴ谺｡髮・ｨ育畑縺ｮ蟷ｴ驕ｸ謚槭リ繝薙ご繝ｼ繧ｷ繝ｧ繝ｳ繧定｡ｨ遉ｺ縺吶ｋ (繝・・繧ｿ縺後≠繧句ｹｴ縺ｮ縺ｿ遘ｻ蜍募庄閭ｽ)"""
    curr = st.session_state.get('current_month', datetime.today().replace(day=1))
    current_user = st.session_state.get("username", "")
    current_menu = st.session_state.get("menu_selection", "繧ｫ繝ｬ繝ｳ繝繝ｼ")
    
    # 繝・・繧ｿ縺悟ｭ伜惠縺吶ｋ繝ｦ繝九・繧ｯ縺ｪ譛医Μ繧ｹ繝医ｒ蜿門ｾ・    available_months = get_transaction_range(current_user)
    available_years = sorted(list(set([m.year for m in available_months])))
    
    # 蜑榊ｾ後・蟷ｴ繧呈､懃ｴ｢
    prev_y = next((y for y in reversed(available_years) if y < curr.year), None)
    next_y = next((y for y in available_years if y > curr.year), None)
    
    has_prev = prev_y is not None
    has_next = next_y is not None
    
    prev_date_str = f"{prev_y}-01-01" if has_prev else ""
    next_date_str = f"{next_y}-01-01" if has_next else ""
    
    # 譛域ｬ｡縺ｨ蜷梧ｧ倥・CSS繧帝←逕ｨ縺励※1陦後↓蜿弱ａ繧・    st.markdown("""
        <style>
            /* st.columns 縺ｮ隕ｪ隕∫ｴ縺ｫ蟇ｾ縺励※讓ｪ荳ｦ縺ｳ繧貞ｼｷ蛻ｶ */
            div[data-testid="stHorizontalBlock"] {
                display: flex !important;
                flex-direction: row !important;
                flex-wrap: nowrap !important;
                align-items: center !important;
                justify-content: center !important;
                gap: 2px !important;
            }
            /* 蜑榊ｹｴ繝ｻ鄙悟ｹｴ繝懊ち繝ｳ縺ｮ繧ｫ繝ｩ繝 */
            div[data-testid="stHorizontalBlock"] > div:nth-child(1),
            div[data-testid="stHorizontalBlock"] > div:nth-child(3) {
                flex: 1 1 0% !important;
                width: auto !important;
                min-width: 0 !important;
            }
            /* 蠖灘ｹｴ繝昴ャ繝励が繝ｼ繝舌・縺ｮ繧ｫ繝ｩ繝 */
            div[data-testid="stHorizontalBlock"] > div:nth-child(2) {
                flex: 0 0 auto !important;
                width: 130px !important;
                min-width: 130px !important;
            }
            /* 繝昴ャ繝励が繝ｼ繝舌・繝懊ち繝ｳ縺ｮ菴咏區繧貞炎繧・*/
            div[data-testid="stPopover"] > button {
                padding-left: 2px !important;
                padding-right: 2px !important;
            }
            /* 繝昴ャ繝励が繝ｼ繝舌・繝懊ち繝ｳ縺ｮ繝・く繧ｹ繝医し繧､繧ｺ縺ｨ謚倥ｊ霑斐＠遖∵ｭ｢ */
            div[data-testid="stPopover"] > button p {
                font-size: 0.95rem !important;
                white-space: nowrap !important;
                margin: 0 !important;
            }
        </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col1:
        if has_prev:
            st.markdown(f"""
                <div style='display: flex; align-items: center; justify-content: flex-end; font-size: 0.9rem; white-space: nowrap;'>
                    <a href="/?date={prev_date_str}&user={current_user}&menu={current_menu}" target="_self" style='text-decoration: none; color: #007bff; font-weight: bold;'>笳 蜑榊ｹｴ</a>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='text-align: right; color: #ccc; font-size: 0.9rem;'>笳 蜑榊ｹｴ</div>", unsafe_allow_html=True)
            
    with col2:
        # 譛域ｬ｡縺ｨ繧ｹ繧ｿ繧､繝ｫ繧貞粋繧上○縺ｦ繝昴ャ繝励が繝ｼ繝舌・蛹・        with st.popover(curr.strftime('%Y蟷ｴ 笆ｼ'), use_container_width=True):
            # 驕ｸ謚槫庄閭ｽ縺ｪ蟷ｴ縺ｮ繝ｪ繧ｹ繝医ｒ菴懈・ (繝・・繧ｿ遽・峇蜀・
            if available_years:
                nav_years = sorted(available_years, reverse=True)
            else:
                nav_years = [curr.year]

            list_html = "<div style='text-align: center;'>"
            for y in nav_years:
                y_str = f"{y}-01-01"
                y_label = f"{y}蟷ｴ"
                is_current = (y == curr.year)
                
                bg_color = "#e6f2ff" if is_current else "transparent"
                font_weight = "bold" if is_current else "normal"
                color = "#0056b3" if is_current else "#333"
                
                link = f"<a href='/?date={y_str}&user={current_user}&menu={current_menu}' target='_self' style='display: block; padding: 10px; margin: 2px 0; border-radius: 4px; background-color: {bg_color}; color: {color}; text-decoration: none; font-weight: {font_weight}; font-size: 1.1rem;'>{y_label}</a>"
                list_html += link
            list_html += "</div>"
            st.markdown(list_html, unsafe_allow_html=True)
            
    with col3:
        if has_next:
            st.markdown(f"""
                <div style='display: flex; align-items: center; justify-content: flex-start; font-size: 0.9rem; white-space: nowrap;'>
                    <a href="/?date={next_date_str}&user={current_user}&menu={current_menu}" target="_self" style='text-decoration: none; color: #007bff; font-weight: bold;'>鄙悟ｹｴ 笆ｶ</a>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='text-align: left; color: #ccc; font-size: 0.9rem;'>鄙悟ｹｴ 笆ｶ</div>", unsafe_allow_html=True)
    
    st.markdown("---")

def render_month_navigation():
    """蜈ｨ讖溯・蜈ｱ騾壹・譛磯∈謚槭リ繝薙ご繝ｼ繧ｷ繝ｧ繝ｳ縺ｨ譛磯俣蜷郁ｨ医ｒ陦ｨ遉ｺ縺吶ｋ"""
    # 迴ｾ蝨ｨ縺ｮ譛医ｒ蜿門ｾ・    if 'current_month' not in st.session_state:
        st.session_state['current_month'] = datetime.today().replace(day=1)
    
    curr = st.session_state['current_month']
    
    # 蟷ｴ譛医Μ繧ｹ繝医・菴懈・ (2023蟷ｴ1譛医°繧臥ｿ悟ｹｴ譛ｫ縺ｾ縺ｧ縲∵・鬆・
    start_date = datetime(2023, 1, 1)
    end_date = datetime.today().replace(day=1) + relativedelta(years=1, month=12)
    
    month_options = []
    temp_date = start_date
    while temp_date <= end_date:
        month_options.append(temp_date)
        temp_date += relativedelta(months=1)
    
    # 髯埼・ｼ域眠縺励＞鬆・ｼ峨↓荳ｦ縺ｳ譖ｿ縺医ｋ
    month_options.reverse()
    
    # 陦ｨ遉ｺ逕ｨ縺ｮ繝ｩ繝吶Ν菴懈・
    month_labels = [dt.strftime('%Y蟷ｴ%m譛・) for dt in month_options]
    
    # 迴ｾ蝨ｨ縺ｮ繧､繝ｳ繝・ャ繧ｯ繧ｹ繧貞叙蠕・    try:
        current_idx = month_options.index(curr.replace(day=1))
    except ValueError:
        # 荳・′荳隕九▽縺九ｉ縺ｪ縺・ｴ蜷医・繝ｪ繧ｹ繝医・蜈磯ｭ・域怙譁ｰ・峨↓霑ｽ蜉縺励※蜀阪た繝ｼ繝・        month_options.append(curr.replace(day=1))
        month_options.sort(reverse=True)
        month_labels = [dt.strftime('%Y蟷ｴ%m譛・) for dt in month_options]
        current_idx = month_options.index(curr.replace(day=1))

    # 繝翫ン繧ｲ繝ｼ繧ｷ繝ｧ繝ｳUI縺ｮ繧ｹ繧ｿ繧､繝ｫ隱ｿ謨ｴ (繧ｹ繝槭・縺ｧ繧よｨｪ荳ｦ縺ｳ繧貞ｼｷ蛻ｶ縲√し繧､繧ｺ譛蟆丞喧)
    st.markdown("""
        <style>
            /* st.columns 縺ｮ隕ｪ隕∫ｴ縺ｫ蟇ｾ縺励※讓ｪ荳ｦ縺ｳ繧貞ｼｷ蛻ｶ */
            div[data-testid="stHorizontalBlock"] {
                display: flex !important;
                flex-direction: row !important;
                flex-wrap: nowrap !important;
                align-items: center !important;
                justify-content: center !important;
                gap: 2px !important;
            }
            /* 蜑肴怦繝ｻ鄙梧怦繝懊ち繝ｳ縺ｮ繧ｫ繝ｩ繝 */
            div[data-testid="stHorizontalBlock"] > div:nth-child(1),
            div[data-testid="stHorizontalBlock"] > div:nth-child(3) {
                flex: 1 1 0% !important;
                width: auto !important;
                min-width: 0 !important;
            }
            /* 蠖捺怦繝昴ャ繝励が繝ｼ繝舌・縺ｮ繧ｫ繝ｩ繝 */
            div[data-testid="stHorizontalBlock"] > div:nth-child(2) {
                flex: 0 0 auto !important;
                width: 130px !important; /* 繧ｹ繝槭・縺ｧ1陦後↓蜿弱∪繧九ｈ縺・↓蟷・ｒ蟆代＠謌ｻ縺・*/
                min-width: 130px !important;
            }
            /* 繝昴ャ繝励が繝ｼ繝舌・繝懊ち繝ｳ縺ｮ菴咏區繧貞炎繧・*/
            div[data-testid="stPopover"] > button {
                padding-left: 2px !important;
                padding-right: 2px !important;
            }
            /* 繝昴ャ繝励が繝ｼ繝舌・繝懊ち繝ｳ縺ｮ繝・く繧ｹ繝医し繧､繧ｺ縺ｨ謚倥ｊ霑斐＠遖∵ｭ｢ */
            div[data-testid="stPopover"] > button p {
                font-size: 0.95rem !important;
                white-space: nowrap !important;
                margin: 0 !important;
            }
        </style>
    """, unsafe_allow_html=True)

    # 繝翫ン繧ｲ繝ｼ繧ｷ繝ｧ繝ｳUI (3繧ｫ繝ｩ繝)
    current_user = st.session_state.get("username", "")
    current_menu = st.session_state.get("menu_selection", "繧ｫ繝ｬ繝ｳ繝繝ｼ")

    # 繝・・繧ｿ縺悟ｭ伜惠縺吶ｋ繝ｦ繝九・繧ｯ縺ｪ譛医Μ繧ｹ繝医ｒ蜿門ｾ・    available_months = get_transaction_range(current_user)
    
    # 蜑榊ｾ後・譛医ｒ繝ｪ繧ｹ繝医°繧画､懃ｴ｢ (繝・・繧ｿ縺後≠繧区怦縺ｸ繧ｸ繝｣繝ｳ繝・
    curr_month_start = curr.replace(day=1)
    prev_m = next((m for m in reversed(available_months) if m < curr_month_start), None)
    next_m = next((m for m in available_months if m > curr_month_start), None)
    
    has_prev = prev_m is not None
    has_next = next_m is not None
    
    prev_date_str = prev_m.strftime('%Y-%m-01') if has_prev else ""
    next_date_str = next_m.strftime('%Y-%m-01') if has_next else ""

    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col1:
        if has_prev:
            st.markdown(f"""
                <div style='display: flex; align-items: center; justify-content: flex-end; font-size: 0.9rem; white-space: nowrap;'>
                    <a href="/?date={prev_date_str}&user={current_user}&menu={current_menu}" target="_self" style='text-decoration: none; color: #007bff; font-weight: bold;'>笳 蜑肴怦</a>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='text-align: right; color: #ccc; font-size: 0.9rem;'>笳 蜑肴怦</div>", unsafe_allow_html=True)
            
    with col2:
        # 繝昴ャ繝励が繝ｼ繝舌・縺ｧ譛磯∈謚槭・繧ｿ繝ｳ鬚ｨ縺ｮUI繧剃ｽ懈・
        with st.popover(curr.strftime('%Y蟷ｴ%m譛・笆ｼ'), use_container_width=True):
            # 譁ｰ縺励＞鬆・↓譛医ｒ繝ｪ繧ｹ繝医い繝・・
            nav_months = sorted(available_months, reverse=True)
            if not nav_months:
                nav_months = [curr]
                
            list_html = "<div id='month-scroll-container' style='max-height: 250px; overflow-y: auto; text-align: center; border-radius: 5px;'>"
            
            for m in nav_months:
                m_str = m.strftime('%Y-%m-01')
                m_label = m.strftime('%Y蟷ｴ%m譛・)
                is_current = (m.year == curr.year and m.month == curr.month)
                
                bg_color = "#e6f2ff" if is_current else "transparent"
                font_weight = "bold" if is_current else "normal"
                color = "#0056b3" if is_current else "#333"
                id_attr = "id='current-month-link'" if is_current else ""
                
                # a繧ｿ繧ｰ縺ｫ繧医ｋ逕ｻ髱｢驕ｷ遘ｻ・医け繧ｨ繝ｪ繝代Λ繝｡繝ｼ繧ｿ譖ｴ譁ｰ・・                link = f"<a {id_attr} href='/?date={m_str}&user={current_user}&menu={current_menu}' target='_self' style='display: block; padding: 10px; margin: 2px 0; border-radius: 4px; background-color: {bg_color}; color: {color}; text-decoration: none; font-weight: {font_weight}; font-size: 1.1rem; transition: background 0.2s;'>{m_label}</a>"
                list_html += link
                
            list_html += "</div>"
            
            # JavaScript縺ｧ縲・幕縺・◆迸ｬ髢薙↓ current-month-link 縺ｮ菴咲ｽｮ縺ｾ縺ｧ閾ｪ蜍輔せ繧ｯ繝ｭ繝ｼ繝ｫ縺吶ｋ
            js = """
            <script>
            setTimeout(function() {
                var container = window.parent.document.getElementById('month-scroll-container');
                var target = window.parent.document.getElementById('current-month-link');
                if (container && target) {
                    var scrollPos = target.offsetTop - (container.clientHeight / 2) + (target.clientHeight / 2);
                    container.scrollTo({ top: scrollPos, behavior: 'instant' });
                }
            }, 50);
            </script>
            """
            
            st.markdown(list_html, unsafe_allow_html=True)
            components.html(js, height=0)
            
    with col3:
        if has_next:
            st.markdown(f"""
                <div style='display: flex; align-items: center; justify-content: flex-start; font-size: 0.9rem; white-space: nowrap;'>
                    <a href="/?date={next_date_str}&user={current_user}&menu={current_menu}" target="_self" style='text-decoration: none; color: #007bff; font-weight: bold;'>鄙梧怦 笆ｶ</a>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='text-align: left; color: #ccc; font-size: 0.9rem;'>鄙梧怦 笆ｶ</div>", unsafe_allow_html=True)

    # 繝・・繧ｿ縺ｮ隱ｭ縺ｿ霎ｼ縺ｿ
    with st.spinner("繝・・繧ｿ繧定ｪｭ縺ｿ霎ｼ縺ｿ荳ｭ..."):
        df = load_transactions_data(curr)
    
    # 蜷郁ｨ磯≡鬘阪・邂怜・
    monthly_total = 0
    if not df.empty and "amount" in df.columns:
        monthly_total = df['amount'].sum()

    # 譛磯俣蜷郁ｨ医・陦ｨ遉ｺ
    st.markdown(f"<p style='text-align: center; font-size: 1.2rem; font-weight: bold; margin-bottom: 10px; color: black;'>譛磯俣蜷郁ｨ域髪蜃ｺ: <span style='color: red;'>・･{int(monthly_total):,}</span></p>", unsafe_allow_html=True)
    st.markdown("---")
    
    return df

# ---------- 繝ｬ繧ｷ繝ｼ繝郁ｧ｣譫先ｩ溯・ ----------
def parse_receipt_with_gemini(image_file):
    try:
        img = Image.open(image_file)
        # 繝ｪ繧ｵ繧､繧ｺ・育洒霎ｺ繝ｻ髟ｷ霎ｺ縺ｨ繧ゅ↓驕ｩ蛻・↓蝨ｧ邵ｮ縲よ怙螟ｧ800px遞句ｺｦ縺ｫ縺励※API繧帝ｫ倬溷喧・・        img.thumbnail((800, 800))
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=85)
        img_byte_arr = img_byte_arr.getvalue()
        
        client = st.session_state.get('genai_client')
        if not client:
            return {"error": "API繧ｭ繝ｼ縺瑚ｨｭ螳壹＆繧後※縺・∪縺帙ｓ縲・}
            
        prompt = f"""
莉･荳九・逕ｻ蜒擾ｼ医Ξ繧ｷ繝ｼ繝医∪縺溘・鬆伜庶譖ｸ・峨°繧牙ｿ・ｦ√↑諠・ｱ繧呈歓蜃ｺ縺励∵・邏ｰ陦後＃縺ｨ縺ｫJSON蠖｢蠑上〒蜃ｺ蜉帙＠縺ｦ縺上□縺輔＞縲・
謚ｽ蜃ｺ鬆・岼・亥推譏守ｴｰ縺ｫ蟇ｾ縺励※・・
1. "store_name" : 蠎苓・蜷搾ｼ域枚蟄怜・縲∽ｸ肴・縺ｪ蝣ｴ蜷医・ ""・・2. "date" : 譌･莉假ｼ・YYY-MM-DD蠖｢蠑上∽ｸ肴・縺ｪ蝣ｴ蜷医・ ""・・3. "item_name" : 蝠・刀蜷阪∪縺溘・蜀・ｮｹ・域枚蟄怜・・・4. "amount" : 驥鷹｡搾ｼ域焚蛟､縺ｮ縺ｿ縲√き繝ｳ繝槭↑縺暦ｼ・5. "major_category" : 螟ｧ蛻・｡・6. "minor_category" : 蟆丞・鬘・
縲先怙蜆ｪ蜈医Ν繝ｼ繝ｫ縲・
逕ｻ蜒丞・縺ｫ逞・劼蜷阪↑縺ｩ縺ｮ蛹ｻ逋よｩ滄未縺ｮ蜷榊燕縲√≠繧九＞縺ｯ縲瑚ｨｺ逋よ・邏ｰ縲阪碁伜庶險ｼ・亥現逋よｩ滄未・峨阪→縺・▲縺滓枚蟄励′蜷ｫ縺ｾ繧後※縺・ｋ蝣ｴ蜷医・縺吶∋縺ｦ縺ｮ螟ｧ蛻・｡槭・蠑ｷ蛻ｶ逧・↓ "(13) 蛹ｻ逋・ 縺ｨ縺励∝ｰ丞・鬘槭・蜀・ｮｹ縺九ｉ縲交沛･逞・劼險ｺ逋ゅ阪交汳願脈蜃ｦ譁ｹ縲阪交汳画､懈渊蛛･險ｺ縲阪・縺・★繧後°繧呈耳隲悶＠縺ｦ險ｭ螳壹＠縺ｦ縺上□縺輔＞縲ゅ％繧御ｻ･螟悶・蛹ｻ逋らｳｻ縺ｮ蟆丞・鬘槭・逕滓・縺励↑縺・〒縺上□縺輔＞縲・
縲先ｶ郁ｲｻ遞弱・謚ｽ蜃ｺ繝ｫ繝ｼ繝ｫ縲・
繝ｬ繧ｷ繝ｼ繝亥・縺ｫ縲梧ｶ郁ｲｻ遞趣ｼ・%繧・0%縺ｪ縺ｩ・峨阪′譏守ｴｰ繧・・岼縺ｨ縺励※險倩ｼ峨＆繧後※縺・ｋ蝣ｴ蜷医√◎縺ｮ陦後ｒ1縺､縺ｮ譏守ｴｰ縺ｨ縺励※謚ｽ蜃ｺ縺励∝､ｧ蛻・｡槭ｒ "豸郁ｲｻ遞・ 縲∝ｰ丞・鬘槭ｒ縺昴・遞守紫・・8%" 繧・"10%"縺ｪ縺ｩ・峨→縺励※險ｭ螳壹＠縺ｦ縺上□縺輔＞縲・
縲仙粋險磯≡鬘阪・謨ｴ蜷域ｧ繝ｫ繝ｼ繝ｫ縲托ｼ磯㍾隕・ｼ・
繝ｬ繧ｷ繝ｼ繝医・縲悟粋險磯≡鬘阪阪→縲∵歓蜃ｺ縺励◆縺吶∋縺ｦ縺ｮ譏守ｴｰ縺ｮ縲碁≡鬘搾ｼ・mount・峨阪・蜷郁ｨ磯｡阪′縲∬ｨ育ｮ嶺ｸ雁ｿ・★螳悟・縺ｫ荳閾ｴ縺吶ｋ繧医≧縺ｫ縺励※縺上□縺輔＞縲・驥鷹｡阪′蜷医ｏ縺ｪ縺・ｴ蜷医・縲∵・邏ｰ陦後・蜑ｲ蠑輔ｄ蛟､蠑包ｼ医・繧､繝翫せ驥鷹｡阪〒謚ｽ蜃ｺ・峨・豸郁ｲｻ遞弱・蟆剰ｨ医↑縺ｩ縺ｮ縺・★繧後°繧定ｪｭ縺ｿ鬟帙・縺励※縺・ｋ縺玖ｪ､隱ｭ縺励※縺・ｋ蜿ｯ閭ｽ諤ｧ縺後≠繧翫∪縺吶りｪｭ縺ｿ鬟帙・縺励′縺ｪ縺・ｈ縺・√☆縺ｹ縺ｦ縺ｮ驥鷹｡崎ｦ∫ｴ繧呈ｼ上ｌ縺ｪ縺乗歓蜃ｺ縺励※縺上□縺輔＞縲・
縺昴ｌ莉･螟悶・蝣ｴ蜷医・縲∽ｻ･荳九・繧ｫ繝・ざ繝ｪ菴鍋ｳｻ縺ｫ蜴ｳ蟇・↓蠕薙▲縺ｦ縲∵・邏ｰ縺斐→縺ｫ驕ｩ蛻・↓蛻・｡槭＠縺ｦ縺上□縺輔＞縲・{get_categories_prompt_text()}

JSON縺ｮ蜃ｺ蜉帛ｽ｢蠑上・莉･荳九ｒ蜴ｳ螳医＠縺ｦ縺上□縺輔＞縲ゅ・繝ｼ繧ｯ繝繧ｦ繝ｳ縺ｮ ```json 縺ｪ縺ｩ縺ｯ蜷ｫ繧√ｋ縺壹∫ｴ皮ｲ九↑JSON譁・ｭ怜・・医が繝悶ず繧ｧ繧ｯ繝医・驟榊・・峨・縺ｿ繧定ｿ斐＠縺ｦ縺上□縺輔＞縲・[
  {{
    "store_name": "蠎苓・蜷・,
    "date": "YYYY-MM-DD",
    "item_name": "蝠・刀蜷・,
    "amount": 1000,
    "major_category": "螟ｧ蛻・｡・,
    "minor_category": "蟆丞・鬘・
  }}
]
"""
        # 繧ｷ繝ｳ繝励Ν縺ｪ隗｣譫舌Ο繧ｸ繝・け: 繝｢繝・Ν繧・gemini-2.5-flash 縺ｫ蝗ｺ螳壹＠縺ｦ蜊倡匱螳溯｡・        try:
            response = safe_gemini_call(
                client.models.generate_content,
                model='gemini-2.5-flash',
                contents=[
                    prompt,
                    types.Part.from_bytes(data=img_byte_arr, mime_type='image/jpeg')
                ]
            )
            
            response_text = response.text.strip()
            if response_text.startswith("```json"):
                response_text = response_text[7:]
            if response_text.startswith("```"):
                response_text = response_text[3:]
            if response_text.endswith("```"):
                response_text = response_text[:-3]
                
            result = json.loads(response_text.strip())
            
            # 驟榊・縺ｧ縺ｪ縺・ｴ蜷医・驟榊・縺ｫ縺吶ｋ
            if isinstance(result, dict):
                result = [result]
                
            return result
            
        except Exception as e:
            return {"error": f"繝ｬ繧ｷ繝ｼ繝医・隱ｭ縺ｿ蜿悶ｊ縺ｫ螟ｱ謨励＠縺ｾ縺励◆縲りｩｳ邏ｰ: {str(e)}"}
        
    except Exception as e:
        return {"error": str(e)}

def categorize_items_with_ai(items, store_name):
    """蝠・刀蜷阪Μ繧ｹ繝医→蠎苓・蜷阪°繧峨；emini API繧剃ｽｿ逕ｨ縺励※繧ｫ繝・ざ繝ｪ繧定・蜍募愛蛻･縺吶ｋ"""
    client = st.session_state.get('genai_client')
    if not client:
        return [{"major_category": "縺昴・莉・, "minor_category": "刀譛ｪ蛻・｡・} for _ in items]
        
    prompt = f"""
莉･荳九・蠎苓・縺ｧ雉ｼ蜈･縺励◆蝠・刀縺ｮ繝ｪ繧ｹ繝医↓縺､縺・※縲√◎繧後◇繧後・螟ｧ蛻・｡槭→蟆丞・鬘槭ｒ蛻､螳壹＠縺ｦJSON縺ｧ霑斐＠縺ｦ縺上□縺輔＞縲・
蠎苓・蜷・ {store_name}

縲舌き繝・ざ繝ｪ繧ｷ繧ｹ繝・Β: 螟ｧ蛻・｡槭→蟆丞・鬘槭・繝ｪ繧ｹ繝医・{get_categories_prompt_text()}

蜈･蜉帛膚蜩√Μ繧ｹ繝・
{json.dumps(items, ensure_ascii=False)}

蜃ｺ蜉帛ｽ｢蠑・(JSON驟榊・縺ｮ縺ｿ):
[
  {{"item_name": "蝠・刀蜷・, "major_category": "螟ｧ蛻・｡・, "minor_category": "蟆丞・鬘・}},
  ...
]
"""
    try:
        response = safe_gemini_call(
            client.models.generate_content,
            model='gemini-2.5-flash',
            contents=[prompt]
        )
        response_text = response.text.strip()
        # JSON驛ｨ蛻・・謚ｽ蜃ｺ
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0].strip()
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0].strip()
            
        return json.loads(response_text)
    except Exception:
        # 繧ｨ繝ｩ繝ｼ譎ゅ・縲後◎縺ｮ莉悶阪〒霑斐☆
        return [{"major_category": "縺昴・莉・, "minor_category": "刀譛ｪ蛻・｡・} for _ in items]

def render_transaction_breakdown(df, key_prefix):
    """
    螟ｧ蛻・｡槫挨縲∝ｺ苓・蛻･縲∝ｰ丞・鬘槫挨縺ｮ2谿ｵ髫弱い繧ｳ繝ｼ繝・ぅ繧ｪ繝ｳ繧定｡ｨ遉ｺ縺吶ｋ蜈ｱ騾夐未謨ｰ
    """
    if df.empty:
        st.info("繝・・繧ｿ縺後≠繧翫∪縺帙ｓ縲・)
        return

    # 陦ｨ遉ｺ繝代ち繝ｼ繝ｳ縺ｮ驕ｸ謚橸ｼ亥ｰ丞・鬘槫挨繧貞炎髯､・・    view_pattern = st.radio("陦ｨ遉ｺ繝代ち繝ｼ繝ｳ", ["蠎苓・蛻･", "螟ｧ蛻・｡槫挨"], horizontal=True, key=f"{key_prefix}_view_pattern")
    
    if view_pattern == "蠎苓・蛻･":
        store_col = "store_name" if "store_name" in df.columns else "store" if "store" in df.columns else None
        if store_col:
            store_grouped = df.groupby(store_col, as_index=False)["amount"].sum()
            store_grouped = store_grouped.sort_values(by="amount", ascending=False)
            
            for _, row in store_grouped.iterrows():
                store = row[store_col]
                total_amt_str = f"・･{int(row['amount']):,}"
                
                with st.expander(f"{store}・嘴total_amt_str}"):
                    store_df = df[df[store_col] == store].copy()
                    cat_grouped = store_df.groupby("category", as_index=False)["amount"].sum()
                    cat_grouped = cat_grouped.sort_values(by="amount", ascending=False)
                    
                    for _, cat_row in cat_grouped.iterrows():
                        cat = cat_row["category"]
                        cat_amt_str = f"・･{int(cat_row['amount']):,}"
                        
                        # 2谿ｵ髫守岼・壼､ｧ蛻・｡槭い繧ｳ繝ｼ繝・ぅ繧ｪ繝ｳ
                        sub_df = store_df[store_df["category"] == cat].copy()
                        sub_col = None
                        for col_name in ["subcategory", "sub_category", "蟆丞・鬘・]:
                            if col_name in sub_df.columns:
                                sub_col = col_name
                                break
                        
                        if sub_col:
                            sub_grouped = sub_df.groupby(sub_col, as_index=False)["amount"].sum()
                            sub_grouped = sub_grouped.sort_values(by="amount", ascending=False)
                            
                            with st.expander(f"  笏・{cat}・嘴cat_amt_str}"):
                                if key_prefix == "calendar":
                                    # 繧ｫ繝ｬ繝ｳ繝繝ｼ隧ｳ邏ｰ・亥ｺ苓・蛻･・峨・縺ｿ 3髫主ｱ､逶ｮ莉･髯阪ｒ繧ｫ繧ｹ繧ｿ繝HTML縺ｧ讌ｵ阮・｡ｨ遉ｺ
                                    for _, sub_row in sub_grouped.iterrows():
                                        sub_name = sub_row[sub_col]
                                        sub_amt_str = f"・･{int(sub_row['amount']):,}"
                                        
                                        # 3髫主ｱ､逶ｮ・亥ｰ丞・鬘橸ｼ峨→4髫主ｱ､逶ｮ・亥膚蜩∝錐・峨ｒ荳縺､縺ｮdetails繧ｿ繧ｰ縺ｫ縺ｾ縺ｨ繧√ｋ
                                        # 繧､繝ｳ繝・Φ繝医′縺ゅｋ縺ｨMarkdown縺ｮ繧ｳ繝ｼ繝峨ヶ繝ｭ繝・け縺ｨ隱､隱阪＆繧後ｋ縺溘ａ縲∝ｷｦ隧ｰ繧√↓縺吶ｋ
                                        html_str = f'<details style="margin: 1px 0;">'
                                        html_str += f'<summary style="background-color: #f0f2f6; padding: 2px 8px; margin: 0; border-left: 5px solid #007bff; font-size: 0.9rem; line-height: 1.2; list-style: none; cursor: pointer;">'
                                        html_str += f'L {sub_name}・嘴sub_amt_str}</summary>'
                                        html_str += f'<div style="padding-left: 10px;">'
                                        
                                        # 4髫主ｱ､逶ｮ・壼膚蜩∝錐・郁ｩｳ邏ｰ・・                                        item_df = sub_df[sub_df[sub_col] == sub_name].copy()
                                        item_col = "item_name" if "item_name" in item_df.columns else "item" if "item" in item_df.columns else None
                                        
                                        if item_col:
                                            item_grouped = item_df.groupby(item_col, as_index=False)["amount"].sum()
                                            item_grouped = item_grouped.sort_values(by="amount", ascending=False)
                                            for _, i_row in item_grouped.iterrows():
                                                i_name = i_row[item_col]
                                                i_amt = f"・･{int(i_row['amount']):,}"
                                                html_str += f'<div style="padding-left: 10px; font-size: 0.85rem; line-height: 1.1; margin: 0; color: #555;">笏・{i_name}・嘴i_amt}</div>'
                                        else:
                                            for _, i_row in item_df.iterrows():
                                                i_amt = f"・･{int(i_row['amount']):,}"
                                                html_str += f'<div style="padding-left: 10px; font-size: 0.85rem; line-height: 1.1; margin: 0; color: #555;">笏・{i_amt}</div>'
                                        
                                        html_str += "</div></details>"
                                        st.markdown(html_str, unsafe_allow_html=True)
                                else:
                                    # 縺昴・莉厄ｼ医ム繝・す繝･繝懊・繝臥ｭ会ｼ峨・ 3髫主ｱ､縺ｮ縺ｾ縺ｾ・亥､ｧ蛻・｡・> 蟆丞・鬘槭Μ繧ｹ繝茨ｼ・                                    sub_grouped_disp = sub_grouped.copy()
                                    sub_grouped_disp["amount"] = sub_grouped_disp["amount"].apply(lambda x: f"・･{int(x):,}")
                                    sub_grouped_disp.columns = ["蟆丞・鬘・, "驥鷹｡・]
                                    st.dataframe(sub_grouped_disp, use_container_width=True, hide_index=True)
                        else:
                            # 蟆丞・鬘槭′縺ｪ縺・ｴ蜷医・譏守ｴｰ
                            item_cols = [c for c in ["item_name", "item", "amount"] if c in sub_df.columns]
                            display_items = sub_df[item_cols].copy()
                            display_items["amount"] = display_items["amount"].apply(lambda x: f"・･{int(x):,}")
                            
                            with st.expander(f"  笏・{cat}・嘴cat_amt_str}"):
                                st.dataframe(display_items, use_container_width=True, hide_index=True)
        else:
            st.info("蠎苓・諠・ｱ縺後≠繧翫∪縺帙ｓ縲・)

    elif view_pattern == "螟ｧ蛻・｡槫挨":
        if "category" in df.columns:
            grouped_df = df.groupby("category", as_index=False)["amount"].sum()
            grouped_df = grouped_df.sort_values(by="amount", ascending=False)
            
            for _, row in grouped_df.iterrows():
                cat = row['category']
                total_amt_str = f"・･{int(row['amount']):,}"
                
                with st.expander(f"{cat}・嘴total_amt_str}"):
                    cat_df = df[df["category"] == cat].copy()
                    sub_col = None
                    for col_name in ["subcategory", "sub_category", "蟆丞・鬘・]:
                        if col_name in cat_df.columns:
                            sub_col = col_name
                            break
                    
                    if sub_col:
                        # 2谿ｵ髫守岼・壼ｰ丞・鬘槭い繧ｳ繝ｼ繝・ぅ繧ｪ繝ｳ・亥､ｧ蛻・｡・> 蟆丞・鬘・> 蝠・刀蜷搾ｼ・                        sub_grouped = cat_df.groupby(sub_col, as_index=False)["amount"].sum()
                        sub_grouped = sub_grouped.sort_values(by="amount", ascending=False)
                        
                        for _, sub_row in sub_grouped.iterrows():
                            sub_name = sub_row[sub_col]
                            sub_amt_str = f"・･{int(sub_row['amount']):,}"
                            
                            with st.expander(f"  笏・{sub_name}・嘴sub_amt_str}"):
                                item_df = cat_df[cat_df[sub_col] == sub_name].copy()
                                item_col = "item_name" if "item_name" in item_df.columns else "item" if "item" in item_df.columns else None
                                
                                if item_col:
                                    item_grouped = item_df.groupby(item_col, as_index=False)["amount"].sum()
                                    item_grouped = item_grouped.sort_values(by="amount", ascending=False)
                                    item_grouped["amount"] = item_grouped["amount"].apply(lambda x: f"・･{int(x):,}")
                                    item_grouped.columns = ["蝠・刀蜷・, "驥鷹｡・]
                                    st.dataframe(item_grouped, use_container_width=True, hide_index=True)
                                else:
                                    detail_df = item_df[["date", "amount"]].copy() if "date" in item_df.columns else item_df[["amount"]].copy()
                                    detail_df = detail_df.sort_values(by="amount", ascending=False)
                                    if "date" in detail_df.columns:
                                        detail_df["date"] = detail_df["date"].dt.strftime('%m/%d')
                                    detail_df["amount"] = detail_df["amount"].apply(lambda x: f"・･{int(x):,}")
                                    st.dataframe(detail_df, use_container_width=True, hide_index=True)
                    else:
                        display_df = cat_df.copy()
                        cols_to_keep = [c for c in ["date", "store_name", "store", "item_name", "item", "amount"] if c in display_df.columns]
                        display_df = display_df[cols_to_keep]
                        if "amount" in display_df.columns:
                            display_df = display_df.sort_values(by="amount", ascending=False)
                        if "date" in display_df.columns:
                            display_df["date"] = display_df["date"].dt.strftime('%m/%d')
                        display_df["amount"] = display_df["amount"].apply(lambda x: f"・･{int(x):,}")
                        st.dataframe(display_df, use_container_width=True, hide_index=True)
        else:
            st.warning("繧ｫ繝・ざ繝ｪ諠・ｱ縺後≠繧翫∪縺帙ｓ縲・)

# ---------- 髻ｳ螢ｰ讖溯・髢｢騾｣縺ｮ繝ｦ繝ｼ繝・ぅ繝ｪ繝・ぅ ----------
def render_speech_synthesis_button(text, key):
    """繝・く繧ｹ繝医ｒ隱ｭ縺ｿ荳翫￡繧九せ繝斐・繧ｫ繝ｼ繝懊ち繝ｳ繧定｡ｨ遉ｺ縺吶ｋ"""
    if not text:
        return
    
    # JavaScript縺ｫ繧医ｋ隱ｭ縺ｿ荳翫￡繝ｭ繧ｸ繝・け
    # 繧ｯ繝ｪ繝ｼ繝ｳ繧｢繝・・・域隼陦後↑縺ｩ縺ｮ髯､蜴ｻ・・    clean_text = text.replace("'", "\\'").replace("\n", " ")
    
    html_code = f"""
    <button id="btn-{key}" style="
        background: none;
        border: 1px solid #ddd;
        border-radius: 5px;
        padding: 2px 8px;
        cursor: pointer;
        font-size: 16px;
        margin-top: 5px;
        color: #555;
    " onclick="speak_{key}()" title="隱ｭ縺ｿ荳翫￡繧・>矧</button>
    
    <script>
    function speak_{key}() {{
        const btn = document.getElementById('btn-{key}');
        
        // 蜀咲函荳ｭ縺ｮ蝣ｴ蜷医・蛛懈ｭ｢
        if (window.speechSynthesis.speaking) {{
            window.speechSynthesis.cancel();
            btn.innerText = '矧';
            return;
        }}
        
        // iOS Safari蟇ｾ遲・ 荳蠎ｦ遨ｺ縺ｮcancel繧貞他縺ｶ縺薙→縺ｧ髻ｳ螢ｰ繧ｨ繝ｳ繧ｸ繝ｳ繧貞ｼｷ蛻ｶ逧・↓繧｢繧ｯ繝・ぅ繝悶↓縺吶ｋ
        window.speechSynthesis.cancel();
        
        // 蟆代＠驕・ｻｶ繧貞・繧後※縺九ｉ逋ｺ隧ｱ縺輔○繧具ｼ・OS蟇ｾ遲厄ｼ・        setTimeout(() => {{
            const uttr = new SpeechSynthesisUtterance('{clean_text}');
            uttr.lang = 'ja-JP';
            uttr.rate = 1.1;
            
            uttr.onstart = () => {{ btn.innerText = '竢ｹ'; btn.style.color = '#dc3545'; }};
            uttr.onend = () => {{ btn.innerText = '矧'; btn.style.color = '#555'; }};
            uttr.onerror = (e) => {{
                console.error("SpeechSynthesisError:", e);
                btn.innerText = '矧'; 
                btn.style.color = '#555'; 
            }};
            
            window.speechSynthesis.speak(uttr);
        }}, 50);
    }}
    </script>
    """
    components.html(html_code, height=45)

def render_voice_input_button(key_prefix):
    """髻ｳ螢ｰ蜈･蜉帙・繧ｿ繝ｳ繧定｡ｨ遉ｺ縺励∫ｵ先棡繧偵そ繝・す繝ｧ繝ｳ迥ｶ諷九↓霑斐☆"""
    # Streamlit縺ｮ繧ｻ繝・す繝ｧ繝ｳ迥ｶ諷九→縺ｮ讖区ｸ｡縺礼畑hidden field
    input_key = f"{key_prefix}_voice_input_result"
    
    html_code = f"""
    <div style="display: flex; align-items: center; margin-bottom: 10px;">
        <button id="mic-btn-{key_prefix}" style="
            background-color: #f0f2f6;
            border: 1px solid #dcdfe6;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            cursor: pointer;
            font-size: 20px;
            display: flex;
            justify-content: center;
            align-items: center;
            transition: all 0.3s;
        " onclick="startRecognition()">痔</button>
        <span id="status-{key_prefix}" style="margin-left: 10px; font-size: 14px; color: #666;"></span>
    </div>

    <script>
    function startRecognition() {{
        const btn = document.getElementById('mic-btn-{key_prefix}');
        const status = document.getElementById('status-{key_prefix}');
        const recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
        
        recognition.lang = 'ja-JP';
        recognition.interimResults = false;
        recognition.maxAlternatives = 1;

        recognition.onstart = () => {{
            btn.style.backgroundColor = '#ff4b4b';
            btn.style.color = 'white';
            status.innerText = '髻ｳ螢ｰ繧定ｪ崎ｭ倅ｸｭ... 隧ｱ縺励※縺上□縺輔＞';
        }};

        recognition.onspeechend = () => {{
            recognition.stop();
        }};

        recognition.onresult = (event) => {{
            const result = event.results[0][0].transcript;
            status.innerText = '隱崎ｭ伜ｮ御ｺ・ ' + result;
            
            // 隕ｪ繧ｦ繧｣繝ｳ繝峨え・・treamlit・峨・髫縺怜・蜉帙ヵ繧｣繝ｼ繝ｫ繝峨↓蛟､繧偵そ繝・ヨ縺励※騾∽ｿ｡
            // 縺溘□縺祐treamlit縺ｮ莉墓ｧ倅ｸ翫∫峩謗･繧ｻ繝・ヨ縺励※繧ょ渚蠢懊＠縺ｪ縺・ｴ蜷医′縺ゅｋ縺溘ａ縲・            // 繧ｫ繧ｹ繧ｿ繝繧､繝吶Φ繝医ｄ迚ｹ螳壹・DOM謫堺ｽ懊′蠢・ｦ・            window.parent.postMessage({{
                type: 'streamlit:set_component_value',
                value: result,
                key: '{input_key}'
            }}, '*');
            
            // 邁｡譏鍋噪縺ｪ譁ｹ豕輔→縺励※縲√ヶ繝ｩ繧ｦ繧ｶ縺ｮ繝励Ο繝ｳ繝励ヨ遲峨〒蛟､繧呈ｸ｡縺吶％縺ｨ繧ょ庄閭ｽ縺縺後・            // 縺薙％縺ｧ縺ｯStreamlit縺ｮ繧ｻ繝・す繝ｧ繝ｳ譖ｴ譁ｰ繧貞ｾ・▽
            setTimeout(() => {{
                // 繝壹・繧ｸ蜈ｨ菴薙↓繝｡繝・そ繝ｼ繧ｸ繧帝√ｋ
                const event = new CustomEvent('voiceInput', {{ detail: result }});
                window.parent.document.dispatchEvent(event);
            }}, 500);
        }};

        recognition.onerror = (event) => {{
            if (event.error === 'not-allowed') {{
                status.innerText = '繝槭う繧ｯ讓ｩ髯舌お繝ｩ繝ｼ: 險ｭ螳壹〒險ｱ蜿ｯ縺吶ｋ縺九、ndroid縺ｯHTTPS騾壻ｿ｡縺悟ｿ・ｦ√〒縺・;
            }} else {{
                status.innerText = '繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: ' + event.error;
            }}
            btn.style.backgroundColor = '#f0f2f6';
            btn.style.color = 'black';
        }};

        recognition.onend = () => {{
            btn.style.backgroundColor = '#f0f2f6';
            btn.style.color = 'black';
        }};

        recognition.start();
    }}
    </script>
    """
    
    # 隱崎ｭ倡ｵ先棡繧貞女縺大叙繧九◆繧√・繧ｳ繝ｳ繝昴・繝阪Φ繝・    # 豕ｨ諢・ Streamlit蜈ｬ蠑上・iframe縺九ｉ隕ｪ縺ｸ縺ｮ騾壻ｿ｡縺ｯ蛻ｶ髯舌′縺ゅｋ縺溘ａ縲・    # 螳滄圀縺ｫ縺ｯURL繝代Λ繝｡繝ｼ繧ｿ繧・√き繧ｹ繧ｿ繝繧ｳ繝ｳ繝昴・繝阪Φ繝医Λ繧､繝悶Λ繝ｪ縺ｪ縺励〒縺ｯ蟆代＠蟾･螟ｫ縺悟ｿ・ｦ√・    # 縺薙％縺ｧ縺ｯ縲∬ｪ崎ｭ倥＆繧後◆繝・く繧ｹ繝医ｒ荳譎ら噪縺ｫ陦ｨ遉ｺ縺励√Θ繝ｼ繧ｶ繝ｼ縺檎｢ｺ隱阪＠縺ｦ騾∽ｿ｡縺ｧ縺阪ｋ繧ｹ繧ｿ繧､繝ｫ縺ｫ縺吶ｋ縺九・    # 縺ゅｋ縺・・逶ｴ謗･繧ｻ繝・す繝ｧ繝ｳ縺ｫ譖ｸ縺崎ｾｼ繧縺溘ａ縺ｮ縲碁國縺励・繧ｿ繝ｳ縲咲噪縺ｪ繧｢繝励Ο繝ｼ繝√ｒ縺ｨ繧九・    
    components.html(html_code, height=60)
    
    # 邨先棡繧貞女縺大叙繧九◆繧√・螳滄ｨ鍋噪縺ｪ莉慕ｵ・∩
    # (螳滄圀縺ｫ縺ｯ st.chat_input 縺ｫ閾ｪ蜍輔〒豬√＠霎ｼ繧縺ｮ縺ｯJS縺ｮ繧ｻ繧ｭ繝･繝ｪ繝・ぅ蛻ｶ邏・ｸ企屮縺励＞縺溘ａ縲・    # 髻ｳ螢ｰ隱崎ｭ倥＆繧後◆繝・く繧ｹ繝医ｒ騾夂衍縺ｨ縺励※陦ｨ遉ｺ縺励√◎繧後ｒ蜈･蜉帶ｬ・↓蜿肴丐縺輔○繧九ぎ繧､繝峨ｒ蜃ｺ縺吶・縺檎樟螳溽噪)
    return None

# ---------- 繝壹・繧ｸUI縺ｮ螳溯｣・----------
def show_dashboard():
    # 繝倥ャ繝繝ｼ繧定｡ｨ遉ｺ縺吶ｋ縺溘ａ縺ｮ繝励Ξ繝ｼ繧ｹ繝帙Ν繝繝ｼ・医さ繝ｳ繝・リ・峨ｒ蜈医↓貅門ｙ
    header_placeholder = st.empty()

    # 蜈ｱ騾壹リ繝薙ご繝ｼ繧ｷ繝ｧ繝ｳ縺ｮ驕ｩ逕ｨ
    # 繝｢繝ｼ繝峨ｒ譏守､ｺ逧・↓謖・ｮ夲ｼ域怦谺｡・・    df = render_month_navigation()

    # 譛医・蛻・ｊ譖ｿ縺域桃菴懊′陦後ｏ繧後◆縲悟ｾ後阪・譛譁ｰ縺ｮ迥ｶ諷九〒繝倥ャ繝繝ｼ繧呈峩譁ｰ縺吶ｋ
    header_placeholder.markdown("#### 投 繝繝・す繝･繝懊・繝・(譛亥挨髮・ｨ・")

    if df.empty:
        st.info("窶ｻ莉頑怦縺ｮ繝・・繧ｿ縺ｯ縺ｾ縺縺ゅｊ縺ｾ縺帙ｓ縲・)
        return

    # 蛻・梵霆ｸ縺ｨ繧ｰ繝ｩ繝慕ｨｮ鬘槭・驕ｸ謚朸I
    col_a, col_b = st.columns(2)
    with col_a:
        analysis_axis = st.selectbox(
            "蛻・梵霆ｸ繧帝∈謚・, 
            ["螟ｧ蛻・｡槫挨", "蟆丞・鬘槫挨", "蠎苓・蛻･"], 
            index=0, 
            key="monthly_analysis_axis"
        )
    with col_b:
        graph_type = st.selectbox(
            "繧ｰ繝ｩ繝輔ｒ驕ｸ謚・,
            ["蜀・げ繝ｩ繝・, "譽偵げ繝ｩ繝・],
            index=0,
            key="monthly_graph_type"
        )

    # 驕ｸ謚槭↓蠢懊§縺ｦ髮・ｨ亥ｯｾ雎｡縺ｮ蛻励ｒ豎ｺ螳・    group_col = None
    title_label = ""
    
    if analysis_axis == "螟ｧ蛻・｡槫挨":
        group_col = "category"
        title_label = "螟ｧ蛻・｡槫挨驥鷹｡阪す繧ｧ繧｢"
    elif analysis_axis == "蟆丞・鬘槫挨":
        for col in ["subcategory", "sub_category", "蟆丞・鬘・]:
            if col in df.columns:
                group_col = col
                break
        title_label = "蟆丞・鬘槫挨驥鷹｡阪す繧ｧ繧｢"
    elif analysis_axis == "蠎苓・蛻･":
        for col in ["store_name", "store", "蠎苓・"]:
            if col in df.columns:
                group_col = col
                break
        title_label = "蠎苓・蛻･驥鷹｡阪す繧ｧ繧｢"

    if group_col and group_col in df.columns and "amount" in df.columns:
        if graph_type == "蜀・げ繝ｩ繝・:
            grouped_df = df.groupby(group_col, as_index=False)["amount"].sum()
            grouped_df = grouped_df.sort_values(by="amount", ascending=False)
            
            fig = px.pie(
                grouped_df, 
                values='amount', 
                names=group_col, 
                hole=0.4, 
                title=title_label,
                category_orders={group_col: grouped_df[group_col].tolist()} 
            )
            fig.update_traces(textposition='inside', textinfo='percent+label', sort=False)
            st.plotly_chart(fig, use_container_width=True)
            
            total_amount = grouped_df["amount"].sum()
            st.metric("蠖捺怦邱乗髪蜃ｺ鬘・, f"・･{int(total_amount):,}")
            
        else: # 譽偵げ繝ｩ繝輔・蝣ｴ蜷・            selected_year = st.session_state['current_month'].year
            selected_month = st.session_state['current_month'].month
            
            # 譌･莉倥＃縺ｨ縺ｮ髮・ｨ医・縺溘ａ縺ｫ譌･繧呈歓蜃ｺ
            # df縺ｯ縺吶〒縺ｫ蠖捺怦縺ｮ繝・・繧ｿ繧貞性繧薙〒縺・ｋ
            df_bar = df.copy()
            df_bar['day'] = df_bar['date'].dt.day
            df_bar['day_label'] = df_bar['day'].apply(lambda x: f"{x}譌･")
            
            # 謖・ｮ壹＆繧後◆蛻・梵霆ｸ縺ｧ譌･縺斐→縺ｮ繝・・繧ｿ繧偵げ繝ｫ繝ｼ繝怜喧
            daily_grouped = df_bar.groupby(['day', 'day_label', group_col], as_index=False)["amount"].sum()
            
            # 謖・ｮ壹・鬆・分繧剃ｿ昴▽縺溘ａ縲∝・菴薙・蜷郁ｨ磯｡埼・〒繧ｫ繝・ざ繝ｪ繝ｼ繧偵た繝ｼ繝医☆繧・            cat_sum = daily_grouped.groupby(group_col)["amount"].sum().sort_values(ascending=False).index.tolist()
            
            # 蠖捺怦縺ｮ譌･謨ｰ繧定ｨ育ｮ・(譛域忰縺ｾ縺ｧ繧ｫ繝ｬ繝ｳ繝繝ｼ騾壹ｊ陦ｨ遉ｺ)
            _, last_day = calendar.monthrange(selected_year, selected_month)
            all_days = [f"{i}譌･" for i in range(1, last_day + 1)]
            
            fig = px.bar(
                daily_grouped,
                x='day_label',
                y='amount',
                color=group_col,
                title=f"{selected_year}蟷ｴ{selected_month}譛・{title_label}譌･谺｡謗ｨ遘ｻ (遨堺ｸ翫￡譽偵げ繝ｩ繝・",
                labels={"amount": "驥鷹｡・, "day_label": "譌･", group_col: analysis_axis[:-1]},
                category_orders={"day_label": all_days, group_col: cat_sum}
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 蠖捺怦縺ｮ蜷郁ｨ磯≡鬘阪・縺昴・縺ｾ縺ｾ陦ｨ遉ｺ
            current_month_total = df['amount'].sum() if not df.empty else 0
            st.metric("蠖捺怦邱乗髪蜃ｺ鬘・, f"・･{int(current_month_total):,}")

        st.markdown("---")
        
        st.markdown("##### 繧ｫ繝・ざ繝ｪ蛻･蜀・ｨｳ (蠖捺怦)")
        render_transaction_breakdown(df, "dashboard")
    else:
        st.warning(f"蛻・梵縺ｫ蠢・ｦ√↑蛻暦ｼ・analysis_axis[:-1]}・峨′縺ゅｊ縺ｾ縺帙ｓ縲・)

def show_yearly_dashboard():
    # 繝倥ャ繝繝ｼ繧定｡ｨ遉ｺ縺吶ｋ縺溘ａ縺ｮ繝励Ξ繝ｼ繧ｹ繝帙Ν繝繝ｼ
    header_placeholder = st.empty()
    
    # 蟷ｴ谺｡繝翫ン繧ｲ繝ｼ繧ｷ繝ｧ繝ｳ繧定｡ｨ遉ｺ
    render_year_navigation()
    
    # 繝｡繧､繝ｳ繧ｿ繧､繝医Ν陦ｨ遉ｺ
    header_placeholder.markdown("#### 投 繝繝・す繝･繝懊・繝・(蟷ｴ谺｡髮・ｨ・")
    
    selected_year = st.session_state['current_month'].year
    target_date = datetime(selected_year, 1, 1)
    
    with st.spinner(f"{selected_year}蟷ｴ縺ｮ繝・・繧ｿ繧帝寔險井ｸｭ..."):
        # 蟷ｴ谺｡繝｢繝ｼ繝峨〒繝・・繧ｿ繧貞叙蠕・        df = load_transactions_data(target_date, mode="yearly")
        # 蜑榊ｹｴ豈碑ｼ・畑縺ｫ蜑榊ｹｴ繝・・繧ｿ繧ょ叙蠕・        prev_year_date = target_date - relativedelta(years=1)
        df_prev = load_transactions_data(prev_year_date, mode="yearly")

    if df.empty:
        st.info(f"窶ｻ{selected_year}蟷ｴ縺ｮ繝・・繧ｿ縺ｯ縺ｾ縺縺ゅｊ縺ｾ縺帙ｓ縲・)
        return

    # --- 繧ｰ繝ｩ繝戊｡ｨ遉ｺ驕ｸ謚・(譛域ｬ｡縺ｨ蜷梧ｧ倥↓2繧ｫ繝ｩ繝縺ｮ繝峨Ο繝・・繝繧ｦ繝ｳ) ---
    col_a, col_b = st.columns(2)
    with col_a:
        analysis_axis = st.selectbox(
            "蛻・梵霆ｸ繧帝∈謚・, 
            ["螟ｧ蛻・｡槫挨", "蟆丞・鬘槫挨", "蠎苓・蛻･"], 
            index=0, 
            key="yearly_analysis_axis"
        )
    with col_b:
        graph_type = st.selectbox(
            "繧ｰ繝ｩ繝輔ｒ驕ｸ謚・,
            ["蜀・げ繝ｩ繝・, "譽偵げ繝ｩ繝・, "蜑榊ｹｴ蟇ｾ豈・],
            index=0,
            key="yearly_graph_type"
        )

    # 驕ｸ謚槭↓蠢懊§縺ｦ髮・ｨ亥ｯｾ雎｡縺ｮ蛻励ｒ豎ｺ螳・    group_col = None
    title_label = ""
    if analysis_axis == "螟ｧ蛻・｡槫挨":
        group_col = "category"
        title_label = "螟ｧ蛻・｡槫挨"
    elif analysis_axis == "蟆丞・鬘槫挨":
        for col in ["subcategory", "sub_category", "蟆丞・鬘・]:
            if col in df.columns:
                group_col = col
                break
        title_label = "蟆丞・鬘槫挨"
    elif analysis_axis == "蠎苓・蛻･":
        for col in ["store_name", "store", "蠎苓・"]:
            if col in df.columns:
                group_col = col
                break
        title_label = "蠎苓・蛻･"

    if graph_type == "蜑榊ｹｴ蟇ｾ豈・:
        # 蠖灘ｹｴ繝・・繧ｿ縺ｮ譛亥挨髮・ｨ・        df['month'] = df['date'].dt.month
        monthly_summary = df.groupby('month', as_index=False)['amount'].sum()
        full_months = pd.DataFrame({'month': range(1, 13)})
        monthly_summary = pd.merge(full_months, monthly_summary, on='month', how='left').fillna(0)
        monthly_summary['month_label'] = monthly_summary['month'].apply(lambda x: f"{x}譛・)

        # 蜑榊ｹｴ繝・・繧ｿ縺ｮ譛亥挨髮・ｨ・        df_prev['month'] = df_prev['date'].dt.month
        prev_summary = df_prev.groupby('month', as_index=False)['amount'].sum()
        prev_summary = pd.merge(full_months, prev_summary, on='month', how='left').fillna(0)
        
        comparison_data = pd.DataFrame({
            '譛・: list(monthly_summary['month_label']) * 2,
            '驥鷹｡・: list(prev_summary['amount']) + list(monthly_summary['amount']),
            '蟷ｴ蠎ｦ': [f'{selected_year-1}蟷ｴ'] * 12 + [f'{selected_year}蟷ｴ'] * 12
        })
        
        fig = px.bar(comparison_data, x='譛・, y='驥鷹｡・, color='蟷ｴ蠎ｦ',
                     barmode='group',
                     title=f"{selected_year}蟷ｴ vs {selected_year-1}蟷ｴ 謾ｯ蜃ｺ豈碑ｼ・(譛域ｬ｡螻暮幕)")
        st.plotly_chart(fig, use_container_width=True)

    elif graph_type == "譽偵げ繝ｩ繝・:
        # 蠖灘ｹｴ縺ｮ譛亥挨謗ｨ遘ｻ (遨堺ｸ翫￡譽偵げ繝ｩ繝・
        df['month'] = df['date'].dt.month
        df['month_label'] = df['month'].apply(lambda x: f"{x}譛・)
        
        if group_col and group_col in df.columns:
            yearly_grouped = df.groupby(['month', 'month_label', group_col], as_index=False)["amount"].sum()
            cat_sum = yearly_grouped.groupby(group_col)["amount"].sum().sort_values(ascending=False).index.tolist()
            
            fig = px.bar(
                yearly_grouped,
                x='month_label',
                y='amount',
                color=group_col,
                title=f"{selected_year}蟷ｴ {title_label}謗ｨ遘ｻ (遨堺ｸ翫￡譽偵げ繝ｩ繝・",
                labels={"amount": "驥鷹｡・, "month_label": "譛・, group_col: analysis_axis[:-1]},
                category_orders={"month_label": [f"{i}譛・ for i in range(1, 13)], group_col: cat_sum}
            )
            st.plotly_chart(fig, use_container_width=True)
            
            year_total = df["amount"].sum()
            st.metric(f"{selected_year}蟷ｴ 邱乗髪蜃ｺ鬘・, f"・･{int(year_total):,}")
        else:
            st.warning(f"蛻・梵縺ｫ蠢・ｦ√↑蛻暦ｼ・analysis_axis[:-1]}・峨′縺ゅｊ縺ｾ縺帙ｓ縲・)

    else: # 蜀・げ繝ｩ繝・        if group_col and group_col in df.columns:
            cat_grouped = df.groupby(group_col, as_index=False)["amount"].sum()
            cat_grouped = cat_grouped.sort_values(by="amount", ascending=False)
            
            fig_pie = px.pie(cat_grouped, values='amount', names=group_col, hole=0.4,
                             title=f'{selected_year}蟷ｴ {title_label}謾ｯ蜃ｺ繧ｷ繧ｧ繧｢',
                             category_orders={group_col: cat_grouped[group_col].tolist()})
            fig_pie.update_traces(textposition='inside', textinfo='percent+label', sort=False)
            st.plotly_chart(fig_pie, use_container_width=True)
            
            year_total = cat_grouped["amount"].sum()
            st.metric(f"{selected_year}蟷ｴ 邱乗髪蜃ｺ鬘・, f"・･{int(year_total):,}")
        else:
            st.warning(f"蛻・梵縺ｫ蠢・ｦ√↑蛻暦ｼ・analysis_axis[:-1]}・峨′縺ゅｊ縺ｾ縺帙ｓ縲・)
    
    st.markdown("---")
    st.markdown("##### 繧ｫ繝・ざ繝ｪ蛻･蜀・ｨｳ (蟷ｴ谺｡)")
    render_transaction_breakdown(df, "yearly_dashboard")

def handle_menu_change():
    """繧ｵ繧､繝峨ヰ繝ｼ縺ｧ縺ｮ繝｡繝九Η繝ｼ螟画峩譎ゅ↓URL繝代Λ繝｡繝ｼ繧ｿ繧偵け繝ｪ繧｢縺励∝ｿ・ｦ√↓蠢懊§縺ｦ繝繝・す繝･繝懊・繝芽｡ｨ遉ｺ繧呈怦谺｡縺ｫ繝ｪ繧ｻ繝・ヨ縺吶ｋ"""
    if "date" in st.query_params:
        del st.query_params["date"]
    if "menu" in st.query_params:
        del st.query_params["menu"]
    
    # 繧ｻ繝・す繝ｧ繝ｳ蜀・・繝｡繝九Η繝ｼ驕ｸ謚槭ｒ遒ｺ隱搾ｼ・n_change譎らせ縺ｧ st.session_state.menu_selection 縺ｯ譖ｴ譁ｰ縺輔ｌ縺ｦ縺・ｋ・・    target_menu = st.session_state.get("menu_selection")
    # 莉墓ｧ假ｼ壹き繝ｬ繝ｳ繝繝ｼ縲√Ξ繧ｷ繝ｼ繝亥叙霎ｼ縲√Ξ繧ｷ繝ｼ繝井ｿｮ豁｣繧帝∈謚槭＠縺滄圀縲√ム繝・す繝･繝懊・繝蛾∈謚樒憾諷九ｒ繝ｪ繧ｻ繝・ヨ
    if target_menu in ["繧ｫ繝ｬ繝ｳ繝繝ｼ", "繝ｬ繧ｷ繝ｼ繝亥叙霎ｼ", "繝ｬ繧ｷ繝ｼ繝井ｿｮ豁｣"]:
        st.session_state["menu_selection_reset_flag"] = True # 繝輔Λ繧ｰ繧堤ｫ九※縺ｦ縺翫″縲∝ｾ後〒繝ｪ繧ｻ繝・ヨ繧剃ｿ・☆縺九∫峩謗･譖ｸ縺肴鋤縺医ｋ
        # 逶ｴ謗･譖ｸ縺肴鋤縺医ｋ縺ｨ辟｡髯舌Ν繝ｼ繝励・諱舌ｌ縺後≠繧九′縲√Λ繧ｸ繧ｪ繝懊ち繝ｳ縺ｮ蛟､繧呈桃菴懊☆繧九↓縺ｯ st.session_state.key 繧偵＞縺倥ｋ
        # 縺溘□縺・on_change 荳ｭ縺ｫ閾ｪ霄ｫ繧偵＞縺倥ｋ縺ｮ縺ｯ蛻ｶ髯舌′縺ゅｋ縺溘ａ縲［ain蛛ｴ縺ｧ蜃ｦ逅・☆繧区婿縺悟ｮ牙・縺ｪ蝣ｴ蜷医ｂ縺ゅｋ

def main():
    # URL繝代Λ繝｡繝ｼ繧ｿ縺ｮ蜷梧悄・医そ繝・す繝ｧ繝ｳ邯ｭ謖√・縺溘ａ蜀帝ｭ縺ｧ陦後≧・・    params = st.query_params
    
    # 繝ｭ繧ｰ繧､繝ｳ迥ｶ諷九・蠕ｩ蜈・    if "user" in params:
        st.session_state["username"] = params["user"]
        st.session_state["logged_in"] = True
        
    # --- 蛻晄悄蛹悶♀繧医・繧ｨ繝ｩ繝ｼ髦ｲ豁｢ ---
    if 'menu_selection' in st.session_state:
        # 蜿､縺・Γ繝九Η繝ｼ蜷搾ｼ遺霧莉倥″・峨′谿九▲縺ｦ縺・ｋ蝣ｴ蜷医・閾ｪ蜍募､画鋤
        mapping = {
            "繝ｬ繧ｷ繝ｼ繝域焔蜈･蜉・: "繝ｬ繧ｷ繝ｼ繝域焔蜈･蜉・,
            "繝ｬ繧ｷ繝ｼ繝井ｿｮ豁｣": "繝ｬ繧ｷ繝ｼ繝井ｿｮ豁｣"
        }
        if st.session_state['menu_selection'] in mapping:
            st.session_state['menu_selection'] = mapping[st.session_state['menu_selection']]
            
    selected_date_str = None # UnboundLocalError髦ｲ豁｢
    if "date" in params:
        # 謖・ｮ壹＆繧後◆譌･莉倥ｒ蜿門ｾ・        selected_date_str = params["date"]
        st.session_state['selected_date'] = selected_date_str
        
        # 譏守､ｺ逧・↓繝｡繝九Η繝ｼ縺梧欠螳壹＆繧後※縺・ｋ蝣ｴ蜷医・縺昴ｌ縺ｫ蠕薙≧
        if "menu" in params:
            st.session_state['menu_selection'] = params["menu"]
        else:
            # 繝｡繝九Η繝ｼ謖・ｮ壹′縺ｪ縺・ｴ蜷茨ｼ医き繝ｬ繝ｳ繝繝ｼ縺ｮ譌･莉倥け繝ｪ繝・け遲会ｼ峨・繧ｫ繝ｬ繝ｳ繝繝ｼ逕ｻ髱｢縺ｸ
            st.session_state['menu_selection'] = "繧ｫ繝ｬ繝ｳ繝繝ｼ"

        # URL縺ｮdate縺九ｉ陦ｨ遉ｺ譛・current_month)繧定・蜍募酔譛・        try:
            dt = datetime.strptime(selected_date_str, '%Y-%m-%d')
            st.session_state['current_month'] = dt.replace(day=1)
        except:
            pass
            
        # URL縺ｮ繝代Λ繝｡繝ｼ繧ｿ繧呈紛逅・ｼ井ｸ蠎ｦ驕ｩ逕ｨ縺励◆繧我ｸ榊ｿ・ｦ√↑繝ｪ繝ｭ繝ｼ繝峨ｒ髦ｲ縺舌◆繧√・驟肴・縺悟ｿ・ｦ√↑蝣ｴ蜷医ｂ縺ゅｋ縺後・        # 迴ｾ迥ｶ縺ｯ繝ｪ繝ｳ繧ｯ譁ｹ蠑上・縺溘ａ縲√％縺ｮ縺ｾ縺ｾ繧ｻ繝・す繝ｧ繝ｳ縺ｫ菫晄戟縺吶ｋ・・
    # 繝ｭ繧ｰ繧､繝ｳ貂医∩縺ｮ迥ｶ諷・    if st.session_state.get('logged_in', False):
        
        # 閾ｪ蜍慕判髱｢驕ｷ遘ｻ縺ｮ縺溘ａ縺ｮ繝ｪ繝繧､繝ｬ繧ｯ繝亥・逅・        if st.session_state.get('redirect_to_dashboard'):
            st.session_state['menu_selection'] = "繝繝・す繝･繝懊・繝会ｼ域怦谺｡髮・ｨ茨ｼ・
            st.session_state['redirect_to_dashboard'] = False
            
        # 繧ｵ繧､繝峨ヰ繝ｼ騾｣蜍輔Ο繧ｸ繝・け・郁・蜍募・繧頑崛縺茨ｼ・        # handle_menu_change 縺ｧ繧ｻ繝・ヨ縺輔ｌ縺溘ヵ繝ｩ繧ｰ繧偵メ繧ｧ繝・け
        if st.session_state.get("menu_selection_reset_flag"):
            # 繝ｪ繧ｻ繝・ヨ蟇ｾ雎｡繝｡繝九Η繝ｼ・医き繝ｬ繝ｳ繝繝ｼ縲√Ξ繧ｷ繝ｼ繝育ｳｻ・峨′驕ｸ縺ｰ繧後◆迴ｾ蝨ｨ縺ｮ迥ｶ諷九°繧峨・            # 谺｡縺ｫ繝繝・す繝･繝懊・繝峨↓謌ｻ縺｣縺滓凾縺ｫ縲梧怦谺｡髮・ｨ医阪↓縺ｪ繧九ｈ縺・↓蜀・Κ迥ｶ諷九ｒ縺・§繧・            # 縺溘□縺励∫樟迥ｶ縺ｮ radio 繝懊ち繝ｳ縺ｮ謖吝虚縺ｨ縺励※縲√後ｂ縺玲ｬ｡繝繝・す繝･繝懊・繝臥ｳｻ繧帝∈縺ｶ縺ｪ繧峨阪→縺・≧蛻ｶ蠕｡縺悟ｿ・ｦ・            st.session_state["menu_selection_reset_flag"] = False

        # 繧ｵ繧､繝峨ヰ繝ｼ繝｡繝九Η繝ｼ縺ｮ螳溯｣・        with st.sidebar:
            st.subheader("繝槭う繝九・ [Ver 3.1.0]")
            st.write(f"泊 繝ｦ繝ｼ繧ｶ繝ｼ: **{st.session_state['username']}**")
            st.markdown("---")
            if 'menu_selection' not in st.session_state:
                st.session_state['menu_selection'] = "繧ｫ繝ｬ繝ｳ繝繝ｼ"
            
            # 譌｢蟄倥・ "繝繝・す繝･繝懊・繝・ 繧・"繝繝・す繝･繝懊・繝会ｼ域怦谺｡髮・ｨ茨ｼ・ 縺ｫ鄂ｮ謠帙＠縲∝ｹｴ谺｡繧定ｿｽ蜉
            menu_options = [
                "繝繝・す繝･繝懊・繝会ｼ域怦谺｡髮・ｨ茨ｼ・, 
                "繝繝・す繝･繝懊・繝会ｼ亥ｹｴ谺｡髮・ｨ茨ｼ・, 
                "繧ｫ繝ｬ繝ｳ繝繝ｼ", 
                "繝ｬ繧ｷ繝ｼ繝亥叙霎ｼ", 
                "繝ｬ繧ｷ繝ｼ繝域焔蜈･蜉・, 
                "繝ｬ繧ｷ繝ｼ繝井ｿｮ豁｣", 
                "早AI逶ｸ隲・, 
                "繝倥Ν繝・, 
                "痘繝槭ル繝･繧｢繝ｫ"
            ]
            
            # 繝｡繝九Η繝ｼ縺ｮ繝ｪ繧ｻ繝・ヨ蜃ｦ逅・ｼ亥挨縺ｮ逕ｻ髱｢縺九ｉ謌ｻ縺｣縺ｦ縺阪◆縺ｨ縺咲畑・・            # 繧ゅ＠繧ｫ繝ｬ繝ｳ繝繝ｼ遲峨°繧峨後ム繝・す繝･繝懊・繝臥ｳｻ莉･螟悶阪ｒ邨檎罰縺励※謌ｻ縺｣縺ｦ縺阪◆蝣ｴ蜷医・            # 谺｡縺ｫ繝繝・す繝･繝懊・繝峨ｒ繧ｯ繝ｪ繝・け縺励◆縺ｨ縺阪↓縲梧怦谺｡縲阪↓縺励◆縺・→縺・≧隕∽ｻｶ縲・            # 逶ｴ蜑阪・蛟､繧剃ｿ晄戟縺励※縺翫″縲・・遘ｻ繧呈､懃衍縺吶ｋ
            if "last_menu_selection" not in st.session_state:
                st.session_state.last_menu_selection = st.session_state['menu_selection']
            
            menu_selection = st.radio(
                "讖溯・繧帝∈謚・,
                menu_options,
                key="menu_selection",
                on_change=handle_menu_change
            )
            st.session_state.last_menu_selection = menu_selection
            
            st.markdown("---")
            if st.button("繝ｭ繧ｰ繧｢繧ｦ繝・, use_container_width=True):
                # 繝ｭ繧ｰ繧｢繧ｦ繝域凾縺ｫURL繝代Λ繝｡繝ｼ繧ｿ縺ｨ繧ｻ繝・す繝ｧ繝ｳ迥ｶ諷九ｒ螳悟・縺ｫ繧ｯ繝ｪ繧｢縺吶ｋ
                st.query_params.clear()
                st.session_state.clear()
                st.rerun()

        # 繝｡繧､繝ｳ繧ｳ繝ｳ繝・Φ繝・・蛻・ｊ譖ｿ縺・        if menu_selection == "繝繝・す繝･繝懊・繝会ｼ域怦谺｡髮・ｨ茨ｼ・:
            show_dashboard()
        elif menu_selection == "繝繝・す繝･繝懊・繝会ｼ亥ｹｴ谺｡髮・ｨ茨ｼ・:
            show_yearly_dashboard()
        elif menu_selection == "繧ｫ繝ｬ繝ｳ繝繝ｼ":
            st.markdown("#### 套 繧ｫ繝ｬ繝ｳ繝繝ｼ")
            
            # 蜈ｱ騾壹リ繝薙ご繝ｼ繧ｷ繝ｧ繝ｳ縺ｮ驕ｩ逕ｨ
            df = render_month_navigation()
            
            daily_totals = {}
            if not df.empty and "date" in df.columns and "amount" in df.columns:
                df['day'] = df['date'].dt.day
                daily_totals = df.groupby('day')["amount"].sum().to_dict()
                
            year = st.session_state['current_month'].year
            month = st.session_state['current_month'].month

            # 繧ｫ繝ｬ繝ｳ繝繝ｼ縺ｮ騾ｱ縺ｮ髢句ｧ区屆譌･繧呈律譖懈律縺ｫ險ｭ螳壹＠縲√◎縺ｮ譛医・繧ｫ繝ｬ繝ｳ繝繝ｼ繝槭ヨ繝ｪ繝・け繧ｹ繧貞叙蠕・            calendar.setfirstweekday(calendar.SUNDAY)
            month_days = calendar.monthcalendar(year, month)

            # 繧ｻ繝・す繝ｧ繝ｳ迥ｶ諷九・蛻晄悄蛹厄ｼ磯∈謚樊律縺ｮ菫晄戟・・            if 'selected_date' not in st.session_state:
                st.session_state['selected_date'] = None

            # CSS螳夂ｾｩ・医Μ繝ｳ繧ｯ譁ｹ蠑上〒縺ｮ繝槭せ逶ｮ繝ｬ繧､繧｢繧ｦ繝茨ｼ・            st.markdown("""
<style>
.calendar-grid {
    display: grid;
    grid-template-columns: repeat(7, 1fr);
    gap: 4px;
    width: 100%;
    max-width: 400px;
    margin: 0 auto;
}
/* 繝ｪ繝ｳ繧ｯ繧偵・繧ｹ逶ｮ・域棧邱壻ｻ倥″・峨→縺励※讖溯・縺輔○繧・*/
.cal-link {
    display: block;
    position: relative;
    height: 65px;
    width: 100%;
    border: 1px solid #ccc;
    border-radius: 4px;
    background-color: #ffffff;
    text-decoration: none !important;
    color: inherit !important;
    margin-bottom: 5px;
    transition: all 0.2s ease;
}
.cal-link:hover {
    border-color: #007bff;
    background-color: #f0f7ff;
}
.selected-link {
    background-color: #e6f3ff !important;
    border: 2px solid #007bff !important;
}

/* 譌･莉假ｼ壼ｷｦ荳翫↓驟咲ｽｮ */
.cal-date {
    position: absolute;
    top: 4px; left: 8px;
    font-weight: bold;
    color: #333;
}

/* 驥鷹｡搾ｼ壼承荳九↓襍､蟄励〒驟咲ｽｮ */
.cal-amount {
    position: absolute;
    bottom: 2px; right: 2px;
    color: red !important;
    font-size: 11px;
    font-weight: bold;
}

/* 譖懈律繝倥ャ繝繝ｼ */
.weekday-header { text-align: center; font-weight: bold; padding: 5px 0; font-size: 0.85em; }
.sat-text { color: #3182ce; }
.sun-bg { background-color: #fff5f5; }
.sat-bg { background-color: #ebf8ff; }

/* 逾晄律蜷搾ｼ壻ｸｭ螟ｮ莉倩ｿ代↓驟咲ｽｮ */
.cal-holiday {
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    font-size: 8px;
    color: #e53e3e;
    width: 90%;
    text-align: center;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    pointer-events: none;
}
</style>
""", unsafe_allow_html=True)

            # 繧ｫ繝ｬ繝ｳ繝繝ｼ縺ｮ陦ｨ遉ｺ
            # 繧ｫ繝ｬ繝ｳ繝繝ｼ縺ｮHTML讒狗ｯ・            cal_html = '<div class="calendar-grid">'
            
            # 繝倥ャ繝繝ｼ・域屆譌･・・            for i, wd in enumerate(["譌･", "譛・, "轣ｫ", "豌ｴ", "譛ｨ", "驥・, "蝨・]):
                cls = "sun-text" if i == 0 else "sat-text" if i == 6 else ""
                cal_html += f'<div class="weekday-header {cls}">{wd}</div>'

            # 譌･莉倥・謠冗判
            for week in month_days:
                for i, day in enumerate(week):
                    if day == 0:
                        # 遨ｺ逋ｽ縺ｮ繝槭せ逶ｮ
                        cal_html += '<div></div>'
                    else:
                        amount = daily_totals.get(day, 0)
                        amount_text = f"・･{int(amount):,}" if amount > 0 else ""
                        date_obj = datetime(year, month, day).date()
                        date_str = date_obj.strftime('%Y-%m-%d')
                        is_selected = st.session_state.get('selected_date') == date_str
                        select_cls = "selected-link" if is_selected else ""
                        current_user = st.session_state.get("username", "")
                        
                        # 譖懈律縺翫ｈ縺ｳ逾晄律縺ｫ繧医ｋ閭梧勹濶ｲ縺ｮ蛻､螳・(i: 0=譌･, 6=蝨・
                        holiday_name = jpholiday.is_holiday_name(date_obj)
                        bg_cls = "sun-bg" if (i == 0 or holiday_name) else "sat-bg" if i == 6 else ""
                        
                        cal_html += f'<a href="/?date={date_str}&user={current_user}" target="_self" class="cal-link {select_cls} {bg_cls} notranslate" translate="no">'
                        cal_html += f'<div class="cal-date notranslate" translate="no">{day}</div>'
                        if holiday_name:
                            cal_html += f'<div class="cal-holiday notranslate" translate="no">{holiday_name}</div>'
                        cal_html += f'<div class="cal-amount notranslate" translate="no">{amount_text}</div>'
                        cal_html += '</a>'
            
            cal_html += '</div>'
            st.markdown(cal_html, unsafe_allow_html=True)

            # --- 蟇ｾ雎｡譌･縺ｮ隧ｳ邏ｰ陦ｨ遉ｺ ---
            selected_date = st.session_state.get('selected_date')
            if selected_date:
                # '2026-02-28' 縺ｮ繧医≧縺ｪ蠖｢蠑上°繧画律(day)繧呈歓蜃ｺ縺励※陦ｨ遉ｺ
                try:
                    display_day = int(selected_date.split("-")[-1])
                except:
                    display_day = selected_date

                # 隧ｲ蠖捺律縺ｮ繝・・繧ｿ繧偵ヵ繧｣繝ｫ繧ｿ繝ｪ繝ｳ繧ｰ
                # selected_date (YYYY-MM-DD) 縺ｨ荳閾ｴ縺吶ｋ縺九…urrent_month蜀・〒day縺御ｸ閾ｴ縺吶ｋ縺・                day_val = int(selected_date.split("-")[-1])
                
                day_df = pd.DataFrame()
                if not df.empty and 'date' in df.columns:
                    day_df = df[df['date'].dt.day == day_val].copy()
                
                # 蜷郁ｨ磯｡阪・險育ｮ・                day_total = int(day_df['amount'].sum()) if (not day_df.empty and 'amount' in day_df.columns) else 0
                # 蜷郁ｨ磯｡阪・險育ｮ・                day_total = int(day_df['amount'].sum()) if not day_df.empty else 0
                # 繝・じ繧､繝ｳ繧医ｊ鄙ｻ險ｳ蝗樣∩繧貞━蜈医＠縲√ロ繧､繝・ぅ繝悶↑繝槭・繧ｯ繝繧ｦ繝ｳ縺ｧ陦ｨ遉ｺ
                st.markdown(f"##### 搭 {display_day}譌･縺ｮ謾ｯ蜃ｺ隧ｳ邏ｰ (蜷郁ｨ・ ・･{day_total:,})")
                
                if not day_df.empty:
                    render_transaction_breakdown(day_df, "calendar")
                else:
                    st.info("縺薙・譌･縺ｮ謾ｯ蜃ｺ繝・・繧ｿ縺ｯ縺ゅｊ縺ｾ縺帙ｓ縲・)
            
            st.markdown("---")
            
        elif menu_selection == "繝ｬ繧ｷ繝ｼ繝亥叙霎ｼ":
            st.markdown("#### 萄 繝ｬ繧ｷ繝ｼ繝亥叙霎ｼ")
            
            # 蜈ｱ騾壹リ繝薙ご繝ｼ繧ｷ繝ｧ繝ｳ縺ｮ驕ｩ逕ｨ
            _ = render_month_navigation()
            
            st.info("逕ｻ蜒上ヵ繧｡繧､繝ｫ繧偵い繝・・繝ｭ繝ｼ繝峨＠縺ｦ繝ｬ繧ｷ繝ｼ繝医ｒ隗｣譫舌＠縺ｾ縺吶・)
            
            if "uploader_key" not in st.session_state:
                st.session_state.uploader_key = 0
                
            if "parsed_results" not in st.session_state:
                st.session_state.parsed_results = None
            
            uploaded_file = None
            
            file_img = st.file_uploader("繝ｬ繧ｷ繝ｼ繝医・逕ｻ蜒上ｒ繧｢繝・・繝ｭ繝ｼ繝・, type=['png', 'jpg', 'jpeg'], accept_multiple_files=False, key=f"uploader_{st.session_state.uploader_key}")
            if file_img:
                uploaded_file = file_img
            
            if uploaded_file is not None:
                st.image(uploaded_file, caption="蜿門ｾ励＠縺溘Ξ繧ｷ繝ｼ繝育判蜒・, width=300)
                
                # Streamlit縺ｮ繝懊ち繝ｳ縺ｫ濶ｲ繧偵▽縺代ｋ・医％縺ｮ繝壹・繧ｸ縺ｮ縺ｿ縺ｫ驕ｩ逕ｨ縺輔ｌ繧具ｼ・                st.markdown("""
                <style>
                    /* 逋ｻ骭ｲ繝懊ち繝ｳ(Primary) 繧帝搨濶ｲ縺ｫ */
                    div.stButton > button[kind="primary"] {
                        background-color: #007bff !important;
                        color: white !important;
                        border-color: #007bff !important;
                    }
                    div.stButton > button[kind="primary"]:hover {
                        background-color: #0056b3 !important;
                        border-color: #0056b3 !important;
                    }

                    /* 繧ｭ繝｣繝ｳ繧ｻ繝ｫ繝懊ち繝ｳ(Secondary) 繧定ｵ､濶ｲ縺ｫ */
                    div.stButton > button[kind="secondary"] {
                        background-color: #dc3545 !important;
                        color: white !important;
                        border-color: #dc3545 !important;
                    }
                    div.stButton > button[kind="secondary"]:hover {
                        background-color: #c82333 !important;
                        border-color: #c82333 !important;
                        color: white !important;
                    }
                </style>
                """, unsafe_allow_html=True)
                
                if st.session_state.parsed_results is None:
                    # 縺ｾ縺隗｣譫舌＠縺ｦ縺・↑縺・ｴ蜷・                    col1, col2 = st.columns(2)
                    with col1:
                        parse_btn = st.button("繝ｬ繧ｷ繝ｼ繝医ｒ隗｣譫舌☆繧・, type="primary", use_container_width=True)
                    with col2:
                        cancel_parse_btn = st.button("繧ｭ繝｣繝ｳ繧ｻ繝ｫ", type="secondary", use_container_width=True, key="cancel_upload")
                        
                    if cancel_parse_btn:
                        st.session_state.uploader_key += 1
                        st.rerun()
                        
                    if parse_btn:
                        try:
                            with st.spinner("逕ｻ蜒上ｒ隗｣譫蝉ｸｭ... Gemini縺瑚ｪｭ縺ｿ蜿悶▲縺ｦ縺・∪縺・):
                                results = parse_receipt_with_gemini(uploaded_file)
                                
                            if isinstance(results, list) and len(results) > 0 and isinstance(results[0], dict) and "error" in results[0]:
                                st.error(f"隗｣譫舌↓螟ｱ謨励＠縺ｾ縺励◆: {results[0]['error']}")
                            elif isinstance(results, dict) and "error" in results:
                                st.error(f"隗｣譫舌↓螟ｱ謨励＠縺ｾ縺励◆: {results['error']}")
                            else:
                                st.session_state.parsed_results = results
                                st.rerun()
                        except Exception as e:
                            st.error(f"隗｣譫仙・逅・ｸｭ縺ｫ莠域悄縺帙〓繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: {e}")
                
                else:
                    # 隗｣譫仙ｮ御ｺ・ｾ後√・繝ｬ繝薙Η繝ｼ縺ｨ遒ｺ隱咲判髱｢繧定｡ｨ遉ｺ
                    results = st.session_state.parsed_results
                    
                    if len(results) > 0:
                        preview_date = results[0].get("date", "")
                        preview_store = results[0].get("store_name", "")
                    else:
                        preview_date = ""
                        preview_store = ""
                        
                    total_amount = sum(int(item.get("amount", 0)) for item in results)
                    
                    # 螟ｧ蛻・｡槫挨縺ｮ蜀・ｨｳ繧帝寔險・                    category_totals = {}
                    for item in results:
                        cat = item.get("major_category", "縺昴・莉・)
                        # 豁｣隕丞喧蜃ｦ逅・ｒ驕ｩ逕ｨ縺励※螟ｧ蛻・｡槭ｒ謠・∴繧・                        majors = list(EXPENSE_CATEGORIES.keys())
                        final_major = "縺昴・莉・
                        for m in majors:
                            if m in cat or cat in m:
                                final_major = m
                                break
                        
                        amt = int(item.get("amount", 0))
                        category_totals[final_major] = category_totals.get(final_major, 0) + amt
                    
                    st.markdown("#### 搭 隗｣譫千ｵ先棡縺ｮ遒ｺ隱・)
                    st.write(f"**譌･莉・*: {preview_date}")
                    st.write(f"**蠎苓・**: {preview_store}")
                    st.write(f"**蜷郁ｨ磯≡鬘・*: ・･{total_amount:,}")
                    
                    # DataFrame縺ｧ荳隕ｧ陦ｨ遉ｺ
                    cat_df = pd.DataFrame([
                        {"螟ｧ蛻・｡・: k, "驥鷹｡・: f"・･{v:,}"} for k, v in category_totals.items()
                    ])
                    st.dataframe(cat_df, hide_index=True, use_container_width=True)
                    
                    st.markdown("---")
                    st.write("縺薙・蜀・ｮｹ縺ｧ逋ｻ骭ｲ縺励∪縺吶°・・)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        submit_btn = st.button("逋ｻ骭ｲ", type="primary", use_container_width=True)
                    with col2:
                        cancel_btn = st.button("繧ｭ繝｣繝ｳ繧ｻ繝ｫ", type="secondary", use_container_width=True)
                        
                    if cancel_btn:
                        st.session_state.parsed_results = None
                        st.session_state.uploader_key += 1
                        st.rerun()
                        
                    if submit_btn:
                        try:
                            with st.spinner("菫晏ｭ倅ｸｭ..."):
                                sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                init_transactions_sheet(sheet)
                                
                                written_count = 0
                                for item in results:
                                    # 繧ｫ繝・ざ繝ｪ縺ｮ豁｣隕丞喧・・4繧ｫ繝・ざ繝ｪ菴鍋ｳｻ縺ｫ蠑ｷ蛻ｶ・・                                    majors = list(EXPENSE_CATEGORIES.keys())
                                    major = str(item.get("major_category", "縺昴・莉・))
                                    final_major = "縺昴・莉・
                                    for m in majors:
                                        if m in major or major in m:
                                            final_major = m
                                            break
                                            
                                    minors = EXPENSE_CATEGORIES.get(final_major, EXPENSE_CATEGORIES["縺昴・莉・])
                                    minor = str(item.get("minor_category", "笶薙◎縺ｮ莉・))
                                    final_minor = minors[-1] if minors else "笶薙◎縺ｮ莉・
                                    for m in minors:
                                        text_only = "".join([c for c in m if c.isalnum() or c in "鬘樒黄鬟溷刀譛ｪ蛻・｡槭◎縺ｮ莉・"])
                                        if text_only and (text_only in minor or minor in text_only) and len(minor) > 0:
                                            final_minor = m
                                            break
                                            
                                    store_name = str(item.get("store_name", ""))
                                    item_name = str(item.get("item_name", ""))
                                    
                                    row_data = [
                                        str(st.session_state['username']),
                                        str(item.get("date", "")),
                                        str(store_name),
                                        str(item_name),
                                        str(final_major),
                                        str(final_minor),
                                        int(item.get("amount", 0))
                                    ]
                                    sheet.append_row(row_data)
                                    written_count += 1
                                
                                st.session_state.flash_message = f"笨・隗｣譫舌′螳御ｺ・＠縲＋written_count}莉ｶ縺ｮ繝・・繧ｿ繧剃ｿ晏ｭ倥＠縺ｾ縺励◆・・
                                
                                time.sleep(1)
                                
                                st.session_state.parsed_results = None
                                st.session_state.uploader_key += 1
                                st.rerun()
                                
                        except Exception as e:
                            st.error(f"菫晏ｭ倥お繝ｩ繝ｼ: {e}")

        elif menu_selection == "繝ｬ繧ｷ繝ｼ繝域焔蜈･蜉・:
            st.markdown("#### 統 繝ｬ繧ｷ繝ｼ繝域焔蜈･蜉・)
            
            # 繧ｻ繝・す繝ｧ繝ｳ迥ｶ諷九〒蜈･蜉帙ｒ邂｡逅・            if 'manual_input_form_id' not in st.session_state:
                st.session_state.manual_input_form_id = 0
            if 'manual_input_items' not in st.session_state:
                st.session_state.manual_input_items = [{"id": int(time.time() * 1000), "name": "", "amount": 0}]
            if 'manual_input_date' not in st.session_state:
                st.session_state.manual_input_date = datetime.today()
            if 'manual_input_store' not in st.session_state:
                st.session_state.manual_input_store = ""

            # IME蛻ｶ蠕｡縺ｨnumber_input縺ｮ繝懊ち繝ｳ髫縺礼畑CSS
            st.markdown("""
                <style>
                    /* 驥鷹｡榊・蜉帶ｬ・ｼ・umber_input・峨・ +/- 繝懊ち繝ｳ繧帝撼陦ｨ遉ｺ縺ｫ縺吶ｋ */
                    div[data-testid="stNumberInput"] button {
                        display: none !important;
                    }
                    div[data-testid="stNumberInput"] input {
                        ime-mode: disabled !important;
                    }
                    /* 繝悶Λ繧ｦ繧ｶ讓呎ｺ悶・繧ｹ繝斐Φ繝懊ち繝ｳ繧る撼陦ｨ遉ｺ縺ｫ縺吶ｋ */
                    input[type=number]::-webkit-inner-spin-button, 
                    input[type=number]::-webkit-outer-spin-button { 
                        -webkit-appearance: none; 
                        margin: 0; 
                    }
                    input[type=number] {
                        -moz-appearance: textfield;
                    }
                    /* 蜈･蜉帶ｬ・・菴薙ｒ讌ｵ髯舌∪縺ｧ繧ｳ繝ｳ繝代け繝医↓縺吶ｋ (1/3遞句ｺｦ縺ｫ) */
                    div[data-testid="stTextInput"] input, 
                    div[data-testid="stNumberInput"] input, 
                    div[data-testid="stDateInput"] input {
                        padding: 2px 8px !important;
                        min-height: 28px !important; /* 騾壼ｸｸ縺ｮ邏・/3繧堤岼讓吶↓ */
                        font-size: 14px !important;
                        line-height: 1.2 !important;
                        border: 1px solid #ccc !important; /* 鄂ｫ邱壹ｒ霑ｽ蜉 */
                        border-radius: 4px !important;
                    }
                    /* 蜷・・蜉幃・岼縺ｮ繝ｩ繝吶Ν縺ｮ菴咏區繧ょ炎繧・*/
                    div[data-testid="stWidgetLabel"] p {
                        font-size: 13px !important;
                        margin-bottom: 2px !important;
                    }
                    /* 1. 繧｢繝励Μ蜈ｨ菴薙→繝輔か繝ｼ繝縺ｮ閭梧勹繧堤區縲∵枚蟄励ｒ鮟偵↓蠑ｷ蛻ｶ */
                    .stApp, [data-testid="stForm"] {
                        background-color: #ffffff !important;
                        color: #000000 !important;
                    }
                    /* 2. 繧ｹ繝槭・逕ｻ髱｢・・68px莉･荳具ｼ峨〒縺ｮ繧ｰ繝ｪ繝・ラ蠑ｷ蛻ｶ */
                    @media (max-width: 768px) {
                        /* 譌･莉倥・蠎苓・蜷阪・谿ｵ */
                        div[data-testid="stForm"] > div:nth-child(2) div[data-testid="stHorizontalBlock"] {
                            display: grid !important;
                            grid-template-columns: 4fr 6fr !important;
                            gap: 5px !important;
                        }

                        /* 蝠・刀蜷阪・驥鷹｡阪・蜑企勁繝懊ち繝ｳ縺ｮ谿ｵ */
                        div[data-testid="stForm"] .row-widget.stHorizontalBlock {
                            display: grid !important;
                            grid-template-columns: 5fr 3fr 1fr !important; /* 繝懊ち繝ｳ繧呈怙蟆城剞縺ｫ */
                            gap: 4px !important;
                            width: 100% !important;
                        }
                        /* 3. 蜈･蜉帶ｬ・・譛蟆丞ｹ・ｒ蠑ｷ蛻ｶ隗｣髯､縺励※逕ｻ髱｢蜀・↓蜿弱ａ繧・*/
                        div[data-baseweb="input"], div[data-baseweb="base-input"] {
                            min-width: 0 !important;
                            width: 100% !important;
                        }
                        input {
                            padding: 6px 4px !important;
                            font-size: 14px !important;
                        }

                        /* 4. 繝ｩ繝吶Ν・域律莉倥∝ｺ苓・蜷阪↑縺ｩ・峨・譁・ｭ励ｂ蟆上＆縺・*/
                        label p {
                            font-size: 12px !important;
                        }
                    }
                </style>
            """, unsafe_allow_html=True)

            fid = st.session_state.manual_input_form_id
            # st.form 繧偵さ繝ｳ繝・リ縺ｫ螟画峩縺励※繝ｪ繧｢繧ｯ繝・ぅ繝悶↑謖吝虚繧貞庄閭ｽ縺ｫ縺吶ｋ
            with st.container():
                col_d, col_s = st.columns([4, 6])
                with col_d:
                    # key繧定ｿｽ蜉縺励※繝ｪ繧ｻ繝・ヨ蜿ｯ閭ｽ縺ｫ縺吶ｋ
                    input_date = st.date_input("譌･莉・, value=st.session_state.manual_input_date, key=f"mi_d_{fid}")
                with col_s:
                    # key繧定ｿｽ蜉縺励※繝ｪ繧ｻ繝・ヨ蜿ｯ閭ｽ縺ｫ縺吶ｋ
                    input_store = st.text_input("蠎苓・蜷・, value=st.session_state.manual_input_store, key=f"mi_s_{fid}", placeholder="蠎苓・蜷・)
                
                st.write("---")
                st.write("**譏守ｴｰ蜈･蜉・*")
                
                updated_items = []
                for i, item in enumerate(st.session_state.manual_input_items):
                    c1, c2, c3 = st.columns([5, 3, 1.5])
                    row_id = item.get("id", i) # 莠呈鋤諤ｧ縺ｮ縺溘ａ
                    with c1:
                        iname = st.text_input(f"蝠・刀蜷・{i+1}", value=item["name"], key=f"mi_n_{row_id}_{fid}", label_visibility="collapsed", placeholder="蝠・刀蜷・)
                    with c2:
                        # 驥鷹｡榊・蜉帶凾縺ｫ閾ｪ蜍輔〒谺｡縺ｮ陦後ｒ霑ｽ蜉縺吶ｋ繧ｳ繝ｼ繝ｫ繝舌ャ繧ｯ逕ｨ
                        def add_empty_row_if_last(idx=i):
                            if idx == len(st.session_state.manual_input_items) - 1:
                                # 驥鷹｡阪′蜈･蜉帙＆繧後◆繧画眠縺励＞陦後ｒ霑ｽ蜉・・D繧剃ｻ倅ｸ趣ｼ・                                new_id = int(time.time() * 1000) + len(st.session_state.manual_input_items)
                                st.session_state.manual_input_items.append({"id": new_id, "name": "", "amount": 0})

                        iamount = st.number_input(f"驥鷹｡・{i+1}", value=int(item["amount"]), step=1, key=f"mi_a_{row_id}_{fid}", label_visibility="collapsed", on_change=add_empty_row_if_last)
                    with c3:
                        # 蜑企勁繝懊ち繝ｳ縺ｫ遒ｺ隱阪ヵ繧ｧ繝ｼ繧ｺ繧定ｿｽ蜉
                        with st.popover("卵・・ if len(st.session_state.manual_input_items) > 1 else "ﾃ・, disabled=len(st.session_state.manual_input_items) <= 1):
                            st.write("縺薙・陦後ｒ蜑企勁縺励∪縺吶°・・)
                            if st.button("蜑企勁螳溯｡・, key=f"mi_del_manual_{row_id}_{fid}"): 
                                # ID繧貞・縺ｫ蜑企勁蟇ｾ雎｡繧堤音螳壹＠縺ｦ蜑企勁
                                st.session_state.manual_input_items = [itm for itm in st.session_state.manual_input_items if itm.get("id") != row_id]
                                st.rerun()
                    updated_items.append({"id": row_id, "name": iname, "amount": iamount})
                
                st.session_state.manual_input_items = updated_items
                
                # 繝懊ち繝ｳ鬘・                st.markdown("<br>", unsafe_allow_html=True) 
                col_btn_l, col_btn_r = st.columns(2)
                with col_btn_l:
                    submit_manual = st.button("逋ｻ骭ｲ", use_container_width=True, type="primary", key="submit_manual_input")
                with col_btn_r:
                    cancel_manual = st.button("繧ｭ繝｣繝ｳ繧ｻ繝ｫ", use_container_width=True, key="cancel_manual_input")
                
                if cancel_manual:
                    # 繝輔か繝ｼ繝ID繧呈峩譁ｰ縺励※蛻晄悄迥ｶ諷九↓謌ｻ縺・                    st.session_state.manual_input_form_id += 1
                    st.session_state.manual_input_items = [{"id": int(time.time() * 1000), "name": "", "amount": 0}]
                    st.session_state.manual_input_store = ""
                    st.session_state.manual_input_date = datetime.today()
                    st.rerun()

                if submit_manual:
                    # 蜈･蜉帙＆繧後※縺・ｋ繝・・繧ｿ縺ｮ縺ｿ繧呈歓蜃ｺ・亥膚蜩∝錐縺後≠繧翫√°縺､驥鷹｡阪′ 0 縺ｧ縺ｯ縺ｪ縺・ｂ縺ｮ・・                    valid_items = [itm for itm in st.session_state.manual_input_items if itm["name"].strip() != "" and itm["amount"] != 0]

                    if not input_store:
                        st.error("蠎苓・蜷阪ｒ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・)
                    elif not valid_items:
                        st.error("蟆代↑縺上→繧・莉ｶ莉･荳翫・譛牙柑縺ｪ繝・・繧ｿ繧貞・蜉帙＠縺ｦ縺上□縺輔＞縲・)
                    else:
                        with st.spinner("AI縺後き繝・ざ繝ｪ繧貞愛螳壻ｸｭ..."):
                            # 逋ｻ骭ｲ繝懊ち繝ｳ謚ｼ荳区凾縺ｫ譛牙柑縺ｪ譏守ｴｰ縺ｮ縺ｿ隗｣譫舌ｒ螳溯｡・                            item_names = [itm["name"] for itm in valid_items]
                            categories = categorize_items_with_ai(item_names, input_store)
                            
                            try:
                                sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                init_transactions_sheet(sheet)
                                
                                for itm, cat in zip(valid_items, categories):
                                    major = cat.get("major_category", "縺昴・莉・)
                                    minor = cat.get("minor_category", "刀譛ｪ蛻・｡・)
                                    
                                    row_data = [
                                        str(st.session_state['username']),
                                        str(input_date.strftime('%Y-%m-%d')),
                                        str(input_store),
                                        str(itm["name"]),
                                        str(major),
                                        str(minor),
                                        int(itm["amount"])
                                    ]
                                    sheet.append_row(row_data)
                                
                                st.success(f"笨・{len(st.session_state.manual_input_items)}莉ｶ縺ｮ繝・・繧ｿ繧堤匳骭ｲ縺励∪縺励◆・・)
                                # 繝輔か繝ｼ繝ID繧呈峩譁ｰ縺励※蜈ｨ繧ｦ繧｣繧ｸ繧ｧ繝・ヨ繧貞ｼｷ蛻ｶ繝ｪ繧ｻ繝・ヨ
                                st.session_state.manual_input_form_id += 1
                                st.session_state.manual_input_items = [{"id": int(time.time() * 1000), "name": "", "amount": 0}]
                                st.session_state.manual_input_store = ""
                                st.session_state.manual_input_date = datetime.today()
                                
                                time.sleep(1.5)
                                st.rerun()
                            except Exception as e:
                                st.error(f"逋ｻ骭ｲ繧ｨ繝ｩ繝ｼ: {e}")
        elif menu_selection == "繝ｬ繧ｷ繝ｼ繝井ｿｮ豁｣":
            st.markdown("#### 笞呻ｸ・繝ｬ繧ｷ繝ｼ繝井ｿｮ豁｣")
            
            # 蜈ｱ騾壹リ繝薙ご繝ｼ繧ｷ繝ｧ繝ｳ縺ｮ驕ｩ逕ｨ
            df = render_month_navigation()
                
            if df.empty:
                st.info("窶ｻ縺薙・譛医・繝・・繧ｿ縺ｯ縺ゅｊ縺ｾ縺帙ｓ縲・)
            else:
                # 蠎苓・蜷阪→蝠・刀蜷阪・繧ｫ繝ｩ繝繧貞虚逧・↓蛻､螳・                store_col = "store_name" if "store_name" in df.columns else "store" if "store" in df.columns else None
                item_col = "item_name" if "item_name" in df.columns else "item" if "item" in df.columns else "items" if "items" in df.columns else None
                
                if not store_col:
                    st.warning("繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医↓蠎苓・蜷搾ｼ・store_name' 縺ｾ縺溘・ 'store'・峨・蛻励′隕九▽縺九ｊ縺ｾ縺帙ｓ縲・)
                else:
                    # 繝ｬ繧ｷ繝ｼ繝亥腰菴阪↓髮・ｴ・ｼ域律莉倥→蠎苓・蜷阪′蜷後§繧ゅ・繧貞酔荳繝ｬ繧ｷ繝ｼ繝医→縺ｿ縺ｪ縺呻ｼ・                    receipts_df = df.groupby(["date", store_col], as_index=False).agg(
                        amount=("amount", "sum"),
                        譏守ｴｰ謨ｰ=("amount", "count")
                    )
                    receipts_df.columns = ["譌･莉・, "蠎苓・蜷・, "驥鷹｡榊粋險・, "譏守ｴｰ謨ｰ"]
                    receipts_df["譌･莉・] = receipts_df["譌･莉・].dt.strftime('%Y-%m-%d')
                    receipts_df["驥鷹｡榊粋險・] = receipts_df["驥鷹｡榊粋險・].apply(lambda x: int(x))
                    receipts_df = receipts_df.sort_values(by="譌･莉・, ascending=False).reset_index(drop=True)
                    
                    if "receipt_list_version" not in st.session_state:
                        st.session_state.receipt_list_version = 0

                    st.markdown("<p style='font-size: 0.85rem; font-weight: bold; margin-bottom: 5px;'>繝ｬ繧ｷ繝ｼ繝井ｸ隕ｧ陦ｨ・亥ｯｾ雎｡繝ｬ繧ｷ繝ｼ繝医ｒ驕ｸ謚槭＠縺ｦ縺上□縺輔＞・・/p>", unsafe_allow_html=True)
                    
                    # dataframe 驕ｸ謚・                    event = st.dataframe(
                        receipts_df, 
                        use_container_width=True, 
                        hide_index=True, 
                        selection_mode="single-row",
                        on_select="rerun",
                        key=f"receipt_list_df_{st.session_state.receipt_list_version}"
                    )
                    
                    if len(event.selection.rows) > 0:
                        selected_idx = event.selection.rows[0]
                        selected_receipt = receipts_df.iloc[selected_idx]
                        
                        # 邱丞粋險郁｡後′驕ｸ謚槭＆繧後◆蝣ｴ蜷医・隧ｳ邏ｰ陦ｨ遉ｺ縺励↑縺・                        if selected_receipt["譌･莉・] != "邱丞粋險・:
                            target_date = pd.to_datetime(selected_receipt["譌･莉・])
                            target_store = selected_receipt["蠎苓・蜷・]
                            
                            st.markdown("---")
                            st.write(f"##### 蟇ｾ雎｡繝ｬ繧ｷ繝ｼ繝域・邏ｰ・・{selected_receipt['譌･莉・]} - {target_store}")
                            
                            # 隧ｲ蠖薙Ξ繧ｷ繝ｼ繝医・譏守ｴｰ繧貞叙蠕・                            details = df[(df["date"] == target_date) & (df[store_col] == target_store)].copy()
                            
                            # session_state 荳翫・螟画峩迥ｶ諷九ｒ蛻晄悄蛹厄ｼ亥ｯｾ雎｡縺悟､峨ｏ縺｣縺溷ｴ蜷育畑・・                            receipt_key = f"{selected_receipt['譌･莉・]}_{target_store}"
                            if st.session_state.get('current_receipt_key') != receipt_key:
                                st.session_state['current_receipt_key'] = receipt_key
                                st.session_state['edit_data'] = {}
                                # 繝倥ャ繝繝ｼ諠・ｱ・域律莉倥・蠎苓・蜷搾ｼ峨ｂ縺薙％縺ｧ遒ｺ螳溘↓蛻晄悄蛹・                                st.session_state['edit_header'] = {
                                    "date": target_date.date(),
                                    "store": target_store
                                }
                            
                            for idx, row in details.iterrows():
                                row_index_gs = row["_row_index"]
                                
                                if row_index_gs not in st.session_state['edit_data']:
                                    major = row.get("category", "縺昴・莉・)
                                    # subcategory繧ｫ繝ｩ繝縺ｮ迚ｹ螳・                                    sub_cols = [c for c in ["subcategory", "sub_category", "蟆丞・鬘・] if c in df.columns]
                                    sub = row.get(sub_cols[0], "笶薙◎縺ｮ莉・) if sub_cols else "笶薙◎縺ｮ莉・
                                    
                                    st.session_state['edit_data'][row_index_gs] = {
                                        "name": row.get(item_col, "荳肴・縺ｪ蝠・刀") if item_col else "荳肴・縺ｪ蝠・刀",
                                        "amount": int(row.get("amount", 0)),
                                        "major": major,
                                        "minor": sub
                                    }
                                    
                            st.write("##### 譏守ｴｰ荳隕ｧ")
                            
                            # 髢ｲ隕ｧ繝｢繝ｼ繝峨→菫ｮ豁｣繝｢繝ｼ繝峨ｒ邂｡逅・☆繧鬼tate
                            # 繝ｬ繧ｷ繝ｼ繝医′蛻・ｊ譖ｿ繧上▲縺滓凾縺ｫ迥ｶ諷九ｒ繝ｪ繧ｻ繝・ヨ縺吶ｋ縺溘ａ縺ｮ繧ｭ繝ｼ蛻ｶ蠕｡
                            mode_key = f"mode_{receipt_key}"
                            if mode_key not in st.session_state:
                                st.session_state[mode_key] = False # 蛻晄悄險ｭ螳壹・髢ｲ隕ｧ繝｢繝ｼ繝・False)
                                
                            edit_mode = st.session_state[mode_key]
                            
                            if edit_mode:
                                # 縲蝉ｿｮ豁｣繝｢繝ｼ繝峨代・繝ｬ繧､繧｢繧ｦ繝・                                
                                # 譏守ｴｰ陦後ｒ1陦後・繧､繝ｳ繝ｩ繧､繝ｳ繝・く繧ｹ繝医・繧医≧縺ｫ陦ｨ遉ｺ縺輔○繧九◆繧√・CSS
                                st.markdown("""
                                <style>
                                    /* 繧ｦ繧｣繧ｸ繧ｧ繝・ヨ荳九・繝ｼ繧ｸ繝ｳ繧呈ｶ亥悉縺励※菴咏區繧貞ｮ悟・蜑企勁 */
                                    div[data-testid="stVerticalBlock"]:has(span#receipt-table-target):not(:has(div[data-testid="stVerticalBlock"] span#receipt-table-target)) div.stMarkdown,
                                    div[data-testid="stVerticalBlock"]:has(span#receipt-table-target):not(:has(div[data-testid="stVerticalBlock"] span#receipt-table-target)) div.stPopover {
                                        margin-bottom: 0 !important;
                                    }
                                    
                                    /* 繝昴ャ繝励が繝ｼ繝舌・・亥､ｧ蛻・｡槭・蟆丞・鬘槭・繧ｿ繝ｳ・峨・陦ｨ遉ｺ繧呈･ｵ蜉帙さ繝ｳ繝代け繝医↓ */
                                    div[data-testid="stPopover"] > button {
                                        padding: 0px 4px !important;
                                        font-size: 0.8em !important;
                                        min-height: 24px !important;
                                        width: 100% !important;
                                    }
                                    /* 蝠・刀蜷阪↑縺ｩ縺ｮ髟ｷ縺・ユ繧ｭ繧ｹ繝医′繝懊ち繝ｳ蜀・〒逵∫払縺輔ｌ縺ｪ縺・ｈ縺・↓隱ｿ謨ｴ */
                                    div[data-testid="stPopover"] > button div[data-testid="stMarkdownContainer"] p {
                                        white-space: normal !important;
                                        word-break: break-all !important;
                                        line-height: 1.2 !important;
                                    }
                                </style>
                                """, unsafe_allow_html=True)
                                
                                with st.container():
                                    st.markdown('<span id="receipt-table-target"></span>', unsafe_allow_html=True)
                                    
                                    # 繝ｬ繧ｷ繝ｼ繝医・繝・ム繝ｼ・域律莉倥・蠎苓・蜷搾ｼ峨・菫ｮ豁｣逕ｨ繝輔ぅ繝ｼ繝ｫ繝・                                    col_h1, col_h2 = st.columns(2)
                                    with col_h1:
                                        new_header_date = st.date_input("繝ｬ繧ｷ繝ｼ繝域律莉・, value=st.session_state['edit_header']['date'], key="edit_header_date")
                                    with col_h2:
                                        new_header_store = st.text_input("蠎苓・蜷・, value=st.session_state['edit_header']['store'], key="edit_header_store")
                                    
                                    st.write("---")
                                    
                                    modified = False
                                    if new_header_date != st.session_state['edit_header']['date'] or new_header_store != st.session_state['edit_header']['store']:
                                        st.session_state['edit_header']['date'] = new_header_date
                                        st.session_state['edit_header']['store'] = new_header_store
                                        modified = True
                                    
                                    for i, (idx, row) in enumerate(details.iterrows(), 1):
                                        row_index_gs = row["_row_index"]
                                        item_name = row.get(item_col, "荳肴・縺ｪ蝠・刀") if item_col else "荳肴・縺ｪ蝠・刀"
                                        # 隧ｳ邏ｰ逕ｻ髱｢縺ｧ縺ｯ譁・ｭ怜宛髯舌ｒ縺九￠縺ｪ縺・                                        
                                        edit_vals = st.session_state['edit_data'].get(row_index_gs)
                                        if not edit_vals:
                                            continue
                                        
                                        disp_name = edit_vals['name']
                                        disp_amount = edit_vals['amount']
                                        disp_major = edit_vals['major']
                                        disp_minor = edit_vals['minor']
                                        
                                        row_col0, row_col1, row_col2, row_col3, row_col4 = st.columns([0.4, 2, 1, 1, 1])
                                        
                                        with row_col0:
                                            st.markdown(f"<div style='font-size: 0.85em; padding-top: 5px;'>{i}.</div>", unsafe_allow_html=True)
                                        with row_col1:
                                            with st.popover(disp_name):
                                                new_name = st.text_input("蝠・刀蜷・, value=disp_name, key=f"nm_{row_index_gs}", label_visibility="collapsed")
                                        with row_col2:
                                            with st.popover(f"ﾂ･{disp_amount:,}"):
                                                new_amount = st.number_input("驥鷹｡・, value=int(disp_amount), step=1, key=f"amt_{row_index_gs}", label_visibility="collapsed")
                                        with row_col3:
                                            majors = list(EXPENSE_CATEGORIES.keys())
                                            default_major_idx = majors.index(disp_major) if disp_major in majors else majors.index("縺昴・莉・)
                                            with st.popover(disp_major):
                                                new_major = st.radio("螟ｧ蛻・｡・, majors, index=default_major_idx, key=f"maj_{r_idx_gs}" if 'r_idx_gs' in locals() else f"maj_{row_index_gs}", label_visibility="collapsed")
                                        with row_col4:
                                            minors = EXPENSE_CATEGORIES.get(new_major, EXPENSE_CATEGORIES["縺昴・莉・])
                                            default_minor_idx = minors.index(disp_minor) if disp_minor in minors else len(minors)-1
                                            with st.popover(disp_minor):
                                                new_minor = st.radio("蟆丞・鬘・, minors, index=default_minor_idx, key=f"min_{r_idx_gs}" if 'r_idx_gs' in locals() else f"min_{row_index_gs}", label_visibility="collapsed")
                                        
                                        if new_name != disp_name or new_amount != disp_amount or new_major != disp_major or new_minor != disp_minor:
                                            st.session_state['edit_data'][row_index_gs]["name"] = new_name
                                            st.session_state['edit_data'][row_index_gs]["amount"] = new_amount
                                            st.session_state['edit_data'][row_index_gs]["major"] = new_major
                                            st.session_state['edit_data'][row_index_gs]["minor"] = new_minor
                                            modified = True
                                        
                                if modified:
                                    st.rerun()
                                    
                                st.markdown("---")
                                
                                # 菫ｮ豁｣逕ｨ繝懊ち繝ｳ・育匳骭ｲ / 繧ｭ繝｣繝ｳ繧ｻ繝ｫ・・                                col1, col2 = st.columns(2)
                                with col1:
                                    if st.button("逋ｻ骭ｲ", use_container_width=True, key="save_receipt"):
                                        try:
                                            with st.spinner("菫晏ｭ倅ｸｭ..."):
                                                sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                                headers = sheet.row_values(1)
                                                amount_col_idx = headers.index("amount") + 1 if "amount" in headers else None
                                                category_col_idx = headers.index("category") + 1 if "category" in headers else None
                                                
                                                sub_col_idx = None
                                                for c in ["subcategory", "sub_category", "蟆丞・鬘・]:
                                                    if c in headers:
                                                        sub_col_idx = headers.index(c) + 1
                                                        break
                                                
                                                updates = []
                                                item_col_idx = headers.index(item_col) + 1 if item_col in headers else None
                                                date_col_idx = headers.index("date") + 1 if "date" in headers else None
                                                store_col_idx = headers.index(store_col) + 1 if store_col in headers else None
                                                
                                                # 繝倥ャ繝繝ｼ諠・ｱ縺ｮ蜿門ｾ・                                                new_date_str = st.session_state['edit_header']['date'].strftime('%Y-%m-%d')
                                                new_store_str = str(st.session_state['edit_header']['store']).strip()
                                                
                                                for r_idx_gs, vals in st.session_state['edit_data'].items():
                                                    if item_col_idx: updates.append(gspread.Cell(row=r_idx_gs, col=item_col_idx, value=str(vals["name"])))
                                                    if amount_col_idx: updates.append(gspread.Cell(row=r_idx_gs, col=amount_col_idx, value=int(vals["amount"])))
                                                    if category_col_idx: updates.append(gspread.Cell(row=r_idx_gs, col=category_col_idx, value=str(vals["major"])))
                                                    if sub_col_idx: updates.append(gspread.Cell(row=r_idx_gs, col=sub_col_idx, value=str(vals["minor"])))
                                                    
                                                    # 譌･莉倥→蠎苓・蜷阪・蜈ｨ陦後↓蟇ｾ縺励※譖ｴ譁ｰ
                                                    if date_col_idx: updates.append(gspread.Cell(row=r_idx_gs, col=date_col_idx, value=new_date_str))
                                                    if store_col_idx: updates.append(gspread.Cell(row=r_idx_gs, col=store_col_idx, value=new_store_str))
                                                    
                                                if updates:
                                                    sheet.update_cells(updates)
                                                    
                                                st.success("笨・繝ｬ繧ｷ繝ｼ繝域・邏ｰ繧呈峩譁ｰ縺励∪縺励◆")
                                                st.session_state['edit_data'] = {} # 繝ｪ繧ｻ繝・ヨ
                                                st.session_state[mode_key] = False # 髢ｲ隕ｧ繝｢繝ｼ繝峨↓謌ｻ縺・                                                st.session_state.receipt_list_version += 1 # 荳隕ｧ縺ｮ驕ｸ謚槭ｒ繝ｪ繧ｻ繝・ヨ
                                                time.sleep(1)
                                                st.rerun()
                                        except Exception as e:
                                            st.error(f"繧ｨ繝ｩ繝ｼ: {e}")
                                            
                                with col2:
                                    if st.button("繧ｭ繝｣繝ｳ繧ｻ繝ｫ", use_container_width=True, key="cancel_receipt_edit"):
                                        # 迥ｶ諷九ｒ繝ｪ繧ｻ繝・ヨ縺鈴夢隕ｧ繝｢繝ｼ繝峨↓謌ｻ繧・                                        st.session_state['current_receipt_key'] = "" # 繧ｭ繝ｼ繧堤ｩｺ縺ｫ縺励※蛻晄悄蛹門・逅・ｒ辟｡逅・ｄ繧雁・螳溯｡後＆縺帙ｋ
                                        st.session_state[mode_key] = False
                                        st.session_state.receipt_list_version += 1 # 荳隕ｧ縺ｮ驕ｸ謚槭ｒ繝ｪ繧ｻ繝・ヨ
                                        st.rerun()
                                
                            else:
                                # 縲宣夢隕ｧ繝｢繝ｼ繝峨代・繝ｬ繧､繧｢繧ｦ繝・                                
                                st.markdown("")
                                total_amount = 0
                                
                                # Markdown縺ｮ繝・・繝悶Ν繝倥ャ繝繝ｼ讒狗ｯ・                                table_md = "| No | 蝠・刀蜷・| 驥鷹｡・| 螟ｧ蛻・｡・| 蟆丞・鬘・|\n"
                                table_md += "|---|---|---:|---|---|\n"
                                
                                for i, (idx, row) in enumerate(details.iterrows(), 1):
                                    item_name = row.get(item_col, "荳肴・縺ｪ蝠・刀") if item_col else "荳肴・縺ｪ蝠・刀"
                                    # 蝠・刀蜷阪ｒ蜈ｨ隗・0譁・ｭ励∪縺ｧ縺ｫ蛻・ｊ隧ｰ繧・                                    display_item_name = item_name[:10] + "窶ｦ" if len(item_name) > 10 else item_name
                                    
                                    major = row.get("category", "縺昴・莉・)
                                    sub_cols = [c for c in ["subcategory", "sub_category", "蟆丞・鬘・] if c in df.columns]
                                    sub = row.get(sub_cols[0], "笶薙◎縺ｮ莉・) if sub_cols else "笶薙◎縺ｮ莉・
                                    amount = int(row.get("amount", 0))
                                    total_amount += amount
                                    
                                    # 蜷・｡後・繝・・繧ｿ繧定ｿｽ蜉 (驥鷹｡阪・蜀・｡ｨ遉ｺ縺ｯ荳崎ｦ・
                                    table_md += f"| {i} | {display_item_name} | {amount:,} | {major} | {sub} |\n"
                                
                                # 蜷郁ｨ郁｡後・霑ｽ蜉
                                table_md += f"| | **蜷郁ｨ・* | **{total_amount:,}** | | |\n"
                                
                                # 繝・・繝悶Ν縺ｮ謠冗判
                                st.markdown(table_md)
                                st.markdown("---")
                                
                                # 髢ｲ隕ｧ逕ｨ繧｢繧ｯ繧ｷ繝ｧ繝ｳ繝懊ち繝ｳ・井ｿｮ豁｣ / 蜑企勁・・                                col1, col2 = st.columns(2)
                                with col1:
                                    if st.button("菫ｮ豁｣", use_container_width=True, key="edit_receipt"):
                                        st.session_state[mode_key] = True
                                        st.rerun()
                                        
                                with col2:
                                    with st.popover("蜑企勁", use_container_width=True):
                                        st.write("譛ｬ蠖薙↓縺薙・繝ｬ繧ｷ繝ｼ繝医ｒ蜑企勁縺励∪縺吶°・・)
                                        if st.button("縺ｯ縺・∝炎髯､縺励∪縺・, use_container_width=True, key="delete_receipt_confirm"):
                                            try:
                                                with st.spinner("蜑企勁荳ｭ..."):
                                                    sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                                    # 荳九°繧蛾・↓蜑企勁縺吶ｋ・医う繝ｳ繝・ャ繧ｯ繧ｹ縺後★繧後↑縺・ｈ縺・↓・・                                                    rows_to_delete = sorted(list(st.session_state['edit_data'].keys()), reverse=True)
                                                    for r_idx in rows_to_delete:
                                                        sheet.delete_rows(r_idx)
                                                        
                                                    st.success("笨・繝ｬ繧ｷ繝ｼ繝医ｒ蜑企勁縺励∪縺励◆")
                                                    st.session_state['edit_data'] = {}
                                                    st.session_state[mode_key] = False
                                                    st.session_state.receipt_list_version += 1 # 荳隕ｧ縺ｮ驕ｸ謚槭ｒ繝ｪ繧ｻ繝・ヨ
                                                    time.sleep(1)
                                                    st.rerun()
                                            except Exception as e:
                                                st.error(f"繧ｨ繝ｩ繝ｼ: {e}")
                            
                            # CSS縺ｮ莉｣繧上ｊ縺ｫJS繧剃ｽｿ縺｣縺ｦ繧医ｊ遒ｺ螳溘↓繝懊ち繝ｳ縺ｮ濶ｲ繧貞､画峩
                            components.html("""
                            <script>
                            setInterval(() => {
                                const elements = window.parent.document.querySelectorAll('button');
                                elements.forEach(b => {
                                    const text = b.innerText.trim();
                                    if (text === '菫ｮ豁｣' || text === '逋ｻ骭ｲ') {
                                        b.style.backgroundColor = '#007bff';
                                        b.style.color = 'white';
                                        b.style.borderColor = '#007bff';
                                    }
                                    if (text === '蜑企勁' || text === '繧ｭ繝｣繝ｳ繧ｻ繝ｫ' || text === '蜑企勁螳溯｡・ || text === '縺ｯ縺・∝炎髯､縺励∪縺・) {
                                        b.style.backgroundColor = '#dc3545';
                                        b.style.color = 'white';
                                        b.style.borderColor = '#dc3545';
                                    }
                                });
                            }, 500);
                            </script>
                            """, height=0, width=0)

                            
        elif menu_selection == "早AI逶ｸ隲・:
            st.markdown("#### 早AI逶ｸ隲・ｼ亥ｰょｱ槭ヵ繧｡繧､繝翫Φ繧ｷ繝｣繝ｫ繝励Λ繝ｳ繝翫・・・)
            st.info("縺ゅ↑縺溘・螳ｶ險育ｰｿ繝・・繧ｿ縺ｫ蝓ｺ縺･縺・※縲、I縺悟・譫舌ｄ繧｢繝峨ヰ繧､繧ｹ繧定｡後＞縺ｾ縺吶・)
            
            # --- 繝・・繧ｿ縺ｮ貅門ｙ・亥・譛滄俣縺九ｉ繝ｭ繧ｰ繧､繝ｳ繝ｦ繝ｼ繧ｶ繝ｼ蛻・・縺ｿ謚ｽ蜃ｺ・・---
            @st.cache_data(ttl=300)
            def get_user_data_csv_for_ai(username):
                # 蜈ｨ繝・・繧ｿ繧貞叙蠕暦ｼ・oad_transactions_data繧呈ｵ∫畑縺帙★縲∝・譛滄俣繧貞ｯｾ雎｡縺ｫ縺吶ｋ縺溘ａ逶ｴ謗･蜿門ｾ暦ｼ・                sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                init_transactions_sheet(sheet)
                records = safe_gspread_call(sheet.get_all_records)
                
                if not records:
                    return ""
                
                df_all = pd.DataFrame(records)
                # 繧ｻ繧ｭ繝･繝ｪ繝・ぅ縺ｮ譛驥崎ｦ∬ｦ∽ｻｶ・夂樟蝨ｨ繝ｭ繧ｰ繧､繝ｳ縺励※縺・ｋ繝ｦ繝ｼ繧ｶ繝ｼ縺ｮ繝・・繧ｿ縺ｮ縺ｿ縺ｫ繝輔ぅ繝ｫ繧ｿ繝ｪ繝ｳ繧ｰ
                if "username" in df_all.columns:
                    df_user = df_all[df_all["username"].astype(str).str.lower() == username.lower()].copy()
                else:
                    return ""

                if df_user.empty:
                    return ""

                # 蠢・ｦ√↑繧ｫ繝ｩ繝縺ｮ縺ｿ謚ｽ蜃ｺ繝ｻ謨ｴ蠖｢
                # 縲悟ｯｾ雎｡蟷ｴ譛医∵律莉倥∝ｺ苓・蜷阪∝､ｧ蛻・｡槭∝ｰ丞・鬘槭・≡鬘阪・                df_user["date"] = pd.to_datetime(df_user["date"], errors="coerce")
                df_user = df_user.dropna(subset=["date"])
                df_user["蟇ｾ雎｡蟷ｴ譛・] = df_user["date"].dt.strftime('%Y-%m')
                df_user["譌･莉・] = df_user["date"].dt.strftime('%Y-%m-%d')
                
                # 陦ｨ遉ｺ逕ｨ繧ｫ繝ｩ繝縺ｮ繝ｪ繝阪・繝
                rename_map = {
                    "store_name": "蠎苓・蜷・,
                    "item_name": "蝠・刀蜷・,
                    "category": "螟ｧ蛻・｡・,
                    "subcategory": "蟆丞・鬘・,
                    "amount": "驥鷹｡・
                }
                # 蟄伜惠縺吶ｋ繧ｫ繝ｩ繝縺ｮ縺ｿ繝槭ャ繝斐Φ繧ｰ
                actual_rename = {k: v for k, v in rename_map.items() if k in df_user.columns}
                df_user = df_user.rename(columns=actual_rename)
                
                target_cols = ["蟇ｾ雎｡蟷ｴ譛・, "譌･莉・, "蠎苓・蜷・, "蝠・刀蜷・, "螟ｧ蛻・｡・, "蟆丞・鬘・, "驥鷹｡・]
                available_cols = [c for c in target_cols if c in df_user.columns]
                
                return df_user[available_cols].to_csv(index=False)
                
            csv_data_string = get_user_data_csv_for_ai(st.session_state['username'])
            if not csv_data_string:
                csv_data_string = "迴ｾ蝨ｨ縲∝盾辣ｧ縺ｧ縺阪ｋ螳ｶ險育ｰｿ繝・・繧ｿ縺ｯ縺ゅｊ縺ｾ縺帙ｓ縲・

            # --- 繝√Ε繝・ヨ繧ｻ繝・す繝ｧ繝ｳ縺ｨ繝｡繝・そ繝ｼ繧ｸ螻･豁ｴ縺ｮ蛻晄悄蛹・---
            if "ai_consult_messages" not in st.session_state:
                st.session_state.ai_consult_messages = []
            
            # --- 繝√Ε繝・ヨ螻･豁ｴ縺ｮ蛻晄悄蛹・(SDK縺梧悄蠕・☆繧九Μ繧ｹ繝亥ｽ｢蠑・ ---
            if "ai_consult_chat_history" not in st.session_state:
                st.session_state.ai_consult_chat_history = []

            # 譛蛻昴・繝｡繝・そ繝ｼ繧ｸ繧定ｿｽ蜉・亥ｱ･豁ｴ縺檎ｩｺ縺ｮ蝣ｴ蜷茨ｼ・            if not st.session_state.ai_consult_messages:
                st.session_state.ai_consult_messages.append({
                    "role": "assistant", 
                    "content": f"縺薙ｓ縺ｫ縺｡縺ｯ縲＋st.session_state['username']}縺輔ｓ・√≠縺ｪ縺溘・蟆ょｱ曦P縺ｧ縺吶ょ・譛滄俣縺ｮ繝・・繧ｿ繧定ｪｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆縲ゆｽ輔〒繧ら嶌隲・＠縺ｦ縺上□縺輔＞縺ｭ縲・
                })

            # 螻･豁ｴ縺ｮ陦ｨ遉ｺ
            for i, msg in enumerate(st.session_state.ai_consult_messages):
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
                    if msg["role"] == "assistant":
                        render_speech_synthesis_button(msg["content"], f"ai_{i}")
            
            # 髻ｳ螢ｰ蜈･蜉帙・繧ｿ繝ｳ
            render_voice_input_button("ai_consult")
            
            # 髻ｳ螢ｰ蜈･蜉帷ｵ先棡繧偵メ繝｣繝・ヨ蜈･蜉帶ｬ・↓蜿肴丐縺輔○繧九◆繧√・JS
            components.html("""
                <script>
                window.parent.document.addEventListener('voiceInput', function(e) {
                    const input = window.parent.document.querySelector('textarea[data-testid="stChatInputTextArea"]');
                    if (input) {
                        // Streamlit(React)縺ｮ蜀・Κ迥ｶ諷九→蜷梧悄縺輔○繧九◆繧√・繝上ャ繧ｯ
                        const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
                        nativeInputValueSetter.call(input, e.detail);
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        
                        // 騾∽ｿ｡繝懊ち繝ｳ繧偵け繝ｪ繝・け縺励※閾ｪ蜍暮∽ｿ｡
                        setTimeout(() => {
                            const btn = window.parent.document.querySelector('button[data-testid="stChatInputSubmitButton"]');
                            if (btn && !btn.disabled) {
                                btn.click();
                            } else {
                                // 騾∽ｿ｡繝懊ち繝ｳ縺梧､懃衍縺ｧ縺阪↑縺・ｴ蜷医・繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ・・nter繧ｭ繝ｼ縺ｮ繧ｷ繝溘Η繝ｬ繝ｼ繝茨ｼ・                                input.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true}));
                            }
                        }, 500);
                    }
                });
                </script>
            """, height=0)

            # 繝ｦ繝ｼ繧ｶ繝ｼ蜈･蜉・            if user_input := st.chat_input("雉ｪ蝠上ｒ蜈･蜉帙＠縺ｦ縺上□縺輔＞..."):
                st.session_state.ai_consult_messages.append({"role": "user", "content": user_input})
                with st.chat_message("user"):
                    st.markdown(user_input)
                    
                client = st.session_state.get('genai_client')
                if not client:
                    with st.chat_message("assistant"):
                        st.error("API繧ｭ繝ｼ縺瑚ｨｭ螳壹＆繧後※縺・↑縺・◆繧√∫嶌隲・ｒ髢句ｧ九〒縺阪∪縺帙ｓ縲・)
                else:
                    with st.chat_message("assistant"):
                        message_placeholder = st.empty()
                        message_placeholder.markdown("蛻・梵荳ｭ...")
                        
                        try:
                            # 繧ｷ繧ｹ繝・Β繝励Ο繝ｳ繝励ヨ繧帝・蠎ｦ讒狗ｯ会ｼ域怙譁ｰ繝・・繧ｿ繧貞渚譏縺輔○繧九◆繧・ｼ・                            system_prompt = f"""縺ゅ↑縺溘・繝ｦ繝ｼ繧ｶ繝ｼ蟆ょｱ槭・蜆ｪ遘縺ｪ繝輔ぃ繧､繝翫Φ繧ｷ繝｣繝ｫ繝励Λ繝ｳ繝翫・縺ｧ縺吶・莉･荳九・CSV繝・・繧ｿ縺ｯ縲√％縺ｮ繝ｦ繝ｼ繧ｶ繝ｼ・・st.session_state['username']}・牙倶ｺｺ縺ｮ螳ｶ險育ｰｿ繝・・繧ｿ縺ｧ縺吶・縺薙・繝・・繧ｿ縺ｫ縺ｯ縲悟膚蜩∝錐縲阪ｂ蜷ｫ縺ｾ繧後※縺翫ｊ縲√＞縺､縲√←縺薙〒縲∽ｽ輔ｒ雋ｷ縺｣縺溘°繧定ｩｳ邏ｰ縺ｫ謚頑升縺ｧ縺阪∪縺吶・繝ｦ繝ｼ繧ｶ繝ｼ縺九ｉ縺ｮ縲檎音螳壹・蝠・刀縺ｮ雉ｼ蜈･譎よ悄・井ｾ具ｼ夐ｶ剰ｉ繝翫Φ繧ｳ繝・・縺・▽雋ｷ縺｣縺滂ｼ滂ｼ峨阪ｄ縲悟膚蜩√・萓｡譬ｼ謗ｨ遘ｻ縲阪↑縺ｩ縺ｮ雉ｪ蝠上↓蟇ｾ縺励∵ｭ｣遒ｺ縺九▽隕ｪ霄ｫ縺ｫ遲斐∴縺ｦ縺上□縺輔＞縲・繝・・繧ｿ縺ｫ蟄伜惠縺励↑縺・耳貂ｬ縺ｯ驕ｿ縺代∫┌鬧・▲縺・・謖・遭繧・ｯ邏・・繧｢繝峨ヰ繧､繧ｹ縺ｪ縺ｩ繧らｩ肴･ｵ逧・↓陦後▲縺ｦ縺上□縺輔＞縲・
縲舌Θ繝ｼ繧ｶ繝ｼ縺ｮ螳ｶ險育ｰｿ繝・・繧ｿ縲・{csv_data_string}"""
                            
                            # 騾∽ｿ｡逶ｴ蜑阪〒繝√Ε繝・ヨ繧ｪ繝悶ず繧ｧ繧ｯ繝医ｒ縲悟ｱ･豁ｴ莉倥″縲阪〒菴懈・
                            chat = client.chats.create(
                                model='gemini-2.5-flash',
                                config=types.GenerateContentConfig(system_instruction=system_prompt),
                                history=st.session_state.ai_consult_chat_history
                            )
                            
                            # 429遲峨・繧ｨ繝ｩ繝ｼ繝上Φ繝峨Μ繝ｳ繧ｰ繧呈律譛ｬ隱槫喧
                            def _send():
                                return chat.send_message(user_input)
                            
                            try:
                                response = safe_gemini_call(_send)
                                response_text = response.text
                                message_placeholder.markdown(response_text)
                                
                                # 螻･豁ｴ繧呈峩譁ｰ・育判髱｢陦ｨ遉ｺ逕ｨ・・                                st.session_state.ai_consult_messages.append({"role": "assistant", "content": response_text})
                                
                                # 譛譁ｰ縺ｮ蝗樒ｭ斐↓繧りｪｭ縺ｿ荳翫￡繝懊ち繝ｳ繧定｡ｨ遉ｺ
                                render_speech_synthesis_button(response_text, "ai_latest")
                                
                                # SDK縺ｮ螻･豁ｴ繧偵そ繝・す繝ｧ繝ｳ縺ｫ蜷梧悄
                                st.session_state.ai_consult_chat_history = chat.get_history()
                                
                            except Exception as e:
                                err_msg = str(e)
                                if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                                    friendly_err = "迴ｾ蝨ｨAI縺ｮ騾壻ｿ｡縺梧ｷｷ縺ｿ蜷医▲縺ｦ縺・∪縺吶よ焚蜊∫ｧ貞ｾ・▲縺ｦ縺九ｉ蜀榊ｺｦ騾∽ｿ｡縺励※縺上□縺輔＞縲・
                                    st.warning(friendly_err)
                                else:
                                    st.error(f"繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: {e}")

                        except Exception as e:
                            st.error(f"莠域悄縺帙〓繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: {e}")

        elif menu_selection == "繝倥Ν繝・:
            st.markdown("#### 庁 繝倥Ν繝励・繧ｵ繝昴・繝・)
            
            st.info("繧｢繝励Μ縺ｮ讖溯・繧・ｽｿ縺・婿縲√ョ繝ｼ繧ｿ縺ｮ菫晏ｭ伜・縺ｪ縺ｩ縺ｫ縺､縺・※菴輔〒繧り◇縺・※縺上□縺輔＞・・)
            
            # --- 繝√Ε繝・ヨ螻･豁ｴ縺ｮ蛻晄悄蛹・---
            if "help_chat_history" not in st.session_state:
                st.session_state.help_chat_history = []
            
            if "help_messages" not in st.session_state:
                st.session_state.help_messages = [
                    {"role": "assistant", "content": "縺薙ｓ縺ｫ縺｡縺ｯ・、I螳ｶ險育ｰｿ繧｢繝励Μ縺ｮ繧ｵ繝昴・繝・I縺ｧ縺吶・n讖溯・縺ｮ菴ｿ縺・婿繧・√ョ繝ｼ繧ｿ縺後←縺薙↓菫晏ｭ倥＆繧後※縺・ｋ縺九↑縺ｩ縲∬ｳｪ蝠上′縺ゅｌ縺ｰ縺ｩ縺・◇・・}
                ]
                
            # 繝｡繝・そ繝ｼ繧ｸ螻･豁ｴ縺ｮ陦ｨ遉ｺ
            for i, msg in enumerate(st.session_state.help_messages):
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
                    if msg["role"] == "assistant":
                        render_speech_synthesis_button(msg["content"], f"help_{i}")
            
            # 髻ｳ螢ｰ蜈･蜉帙・繧ｿ繝ｳ
            render_voice_input_button("help")

            # 髻ｳ螢ｰ蜈･蜉帷ｵ先棡繧貞渚譏縺吶ｋ縺溘ａ縺ｮJS・・I逶ｸ隲・→蜷梧ｧ假ｼ・            components.html("""
                <script>
                window.parent.document.addEventListener('voiceInput', function(e) {
                    const input = window.parent.document.querySelector('textarea[data-testid="stChatInputTextArea"]');
                    if (input) {
                        const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
                        nativeInputValueSetter.call(input, e.detail);
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        setTimeout(() => {
                            const btn = window.parent.document.querySelector('button[data-testid="stChatInputSubmitButton"]');
                            if (btn && !btn.disabled) btn.click();
                            else input.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true}));
                        }, 500);
                    }
                });
                </script>
            """, height=0)

            # 繝ｦ繝ｼ繧ｶ繝ｼ蜈･蜉・            if user_input := st.chat_input("雉ｪ蝠上ｒ蜈･蜉帙＠縺ｦ縺上□縺輔＞...・井ｾ・ 繝ｬ繧ｷ繝ｼ繝医・縺ｩ縺・ｄ縺｣縺ｦ逋ｻ骭ｲ縺吶ｋ縺ｮ・滂ｼ・):
                # 繝ｦ繝ｼ繧ｶ繝ｼ縺ｮ繝｡繝・そ繝ｼ繧ｸ繧定｡ｨ遉ｺ縺励※螻･豁ｴ縺ｫ霑ｽ蜉
                st.session_state.help_messages.append({"role": "user", "content": user_input})
                with st.chat_message("user"):
                    st.markdown(user_input)
                    
                # API繧ｯ繝ｩ繧､繧｢繝ｳ繝医・蜿門ｾ・                client = st.session_state.get('genai_client')
                if not client:
                    with st.chat_message("assistant"):
                        st.error("API繧ｭ繝ｼ縺瑚ｨｭ螳壹＆繧後※縺・↑縺・◆繧√∝屓遲斐〒縺阪∪縺帙ｓ縲Ｔecrets.toml 繧堤｢ｺ隱阪＠縺ｦ縺上□縺輔＞縲・)
                else:
                    with st.chat_message("assistant"):
                        message_placeholder = st.empty()
                        message_placeholder.markdown("蝗樒ｭ斐ｒ逕滓・荳ｭ...")
                        
                        try:
                            # 1. 繧ｵ繝昴・繝・I逕ｨ縺ｮ繧ｷ繧ｹ繝・Β繝励Ο繝ｳ繝励ヨ・亥叙謇ｱ隱ｬ譏取嶌・・                            app_manual = """
縺ゅ↑縺溘・縲√％縺ｮ鬮俶ｩ溯・螳ｶ險育ｰｿ繧｢繝励Μ縲後・繧､繝九・縲阪・蜈ｬ蠑上し繝昴・繝・I縺ｧ縺吶・繝ｦ繝ｼ繧ｶ繝ｼ縺九ｉ讖溯・縺ｮ雉ｪ蝠上ｄ謫堺ｽ懈婿豕輔ｒ閨槭°繧後◆繧峨∽ｻ･荳九・諠・ｱ繧貞・縺ｫ縲∬ｦｪ蛻・°縺､蛻・°繧翫ｄ縺吶￥譯亥・縺励※縺上□縺輔＞縲・
縲・. 繝繝・す繝･繝懊・繝会ｼ域怦谺｡髮・ｨ茨ｼ峨・繝ｻ讎りｦ・ｼ壽怦髢薙・邱乗髪蜃ｺ縲∽ｺ育ｮ励・谿九ｊ縲∵律蛻･縺ｮ謾ｯ蜃ｺ謗ｨ遘ｻ繧偵げ繝ｩ繝輔〒遒ｺ隱阪〒縺阪∪縺吶・繝ｻ繧ｫ繝・ざ繝ｪ蛻･蜀・ｨｳ・夂判髱｢荳ｭ螟ｮ縺ｮ繝ｩ繧ｸ繧ｪ繝懊ち繝ｳ縺ｧ縲悟ｺ苓・蛻･縲阪悟､ｧ蛻・｡槫挨縲阪悟ｰ丞・鬘槫挨縲阪・3繝代ち繝ｼ繝ｳ縺ｫ蛻・ｊ譖ｿ縺亥庄閭ｽ縺ｧ縺吶・繝ｻ2谿ｵ髫手｡ｨ遉ｺ・壻ｾ九∴縺ｰ縲悟ｺ苓・蛻･縲阪ｒ驕ｸ縺ｶ縺ｨ蠎苓・縺御ｸｦ縺ｳ縲√け繝ｪ繝・け縺吶ｋ縺ｨ縺昴・蠎苓・縺ｧ雋ｷ縺｣縺溷､ｧ蛻・｡槭′陦ｨ遉ｺ縺輔ｌ縺ｾ縺吶・繝ｻ荳ｦ縺ｳ譖ｿ縺茨ｼ壽髪蜃ｺ縺ｮ螟壹＞鬆・ｼ磯剄鬆・ｼ峨↓閾ｪ蜍輔〒荳ｦ縺ｶ縺溘ａ縲√←縺薙↓荳逡ｪ縺企≡繧剃ｽｿ縺｣縺ｦ縺・ｋ縺倶ｸ逶ｮ縺ｧ繧上°繧翫∪縺吶・
縲・-2. 繝繝・す繝･繝懊・繝会ｼ亥ｹｴ谺｡髮・ｨ茨ｼ峨・繝ｻ讎りｦ・ｼ夐∈謚槭＠縺溘悟ｹｴ縲榊・菴薙・謾ｯ蜃ｺ繧呈怦縺斐→縺ｫ髮・ｨ医＠縺ｦ陦ｨ遉ｺ縺励∪縺吶・繝ｻ蟷ｴ谺｡謗ｨ遘ｻ・亥燕蟷ｴ蟇ｾ豈費ｼ会ｼ壼燕蟷ｴ縺ｨ縺ｮ謾ｯ蜃ｺ豈碑ｼ・ｒ譛医＃縺ｨ縺ｮ譽偵げ繝ｩ繝輔〒遒ｺ隱阪〒縺阪∪縺吶・繝ｻ蟷ｴ谺｡螟ｧ蛻・｡槫挨繧ｷ繧ｧ繧｢・・蟷ｴ髢薙・謾ｯ蜃ｺ繧貞､ｧ蛻・｡槭＃縺ｨ縺ｮ蜀・げ繝ｩ繝輔〒陦ｨ遉ｺ縺励∪縺吶・繝ｻ蟷ｴ谺｡繧ｫ繝・ざ繝ｪ蛻･蜀・ｨｳ・・蟷ｴ髢薙・蜈ｨ繝・・繧ｿ縺ｫ蟇ｾ縺励※縲・谿ｵ髫弱い繧ｳ繝ｼ繝・ぅ繧ｪ繝ｳ譁ｹ蠑擾ｼ亥､ｧ蛻・｡樞・蟆丞・鬘槭↑縺ｩ・峨〒隧ｳ邏ｰ繧堤｢ｺ隱阪〒縺阪∪縺吶・
縲・. 繧ｫ繝ｬ繝ｳ繝繝ｼ讖溯・縲・繝ｻ讎りｦ・ｼ壹き繝ｬ繝ｳ繝繝ｼ荳翫〒譌･縲・・謾ｯ蜃ｺ鬘阪ｒ荳隕ｧ縺ｧ縺阪∪縺吶ら･晄律縺ｯ襍､縺剰牡莉倥￠縺輔ｌ縺ｾ縺吶・繝ｻ隧ｳ邏ｰ遒ｺ隱搾ｼ壽律莉倥ｒ繧ｯ繝ｪ繝・け縺吶ｋ縺ｨ縲√◎縺ｮ譌･縺ｮ縲梧髪蜃ｺ譏守ｴｰ縲阪′繧｢繧ｳ繝ｼ繝・ぅ繧ｪ繝ｳ蠖｢蠑上〒陦ｨ遉ｺ縺輔ｌ縺ｾ縺吶・繝ｻ3縺､縺ｮ陦ｨ遉ｺ・壹き繝ｬ繝ｳ繝繝ｼ蜀・〒繧ゅ悟ｺ苓・蛻･縲阪悟､ｧ蛻・｡槫挨縲阪悟ｰ丞・鬘槫挨縲阪ｒ繝懊ち繝ｳ荳縺､縺ｧ蛻・ｊ譖ｿ縺医※蛻・梵縺ｧ縺阪∪縺吶・
縲・. 繝ｬ繧ｷ繝ｼ繝亥叙霎ｼ・・CR・峨・繝ｻ謫堺ｽ懶ｼ壹き繝｡繝ｩ縺ｧ謦ｮ縺｣縺溘Ξ繧ｷ繝ｼ繝育判蜒上ｒ繧｢繝・・繝ｭ繝ｼ繝峨☆繧九→縲、I縺後悟ｺ苓・蜷阪阪悟膚蜩∝錐縲阪碁≡鬘阪阪後き繝・ざ繝ｪ縲阪ｒ迸ｬ譎ゅ↓隗｣譫舌＠縺ｾ縺吶・繝ｻ遒ｺ隱搾ｼ夊ｧ｣譫千ｵ先棡繧堤｢ｺ隱阪・菫ｮ豁｣縺励※縲√◎縺ｮ縺ｾ縺ｾ螳ｶ險育ｰｿ縺ｫ逋ｻ骭ｲ縺ｧ縺阪∪縺吶・
縲・. 繝ｬ繧ｷ繝ｼ繝域焔蜈･蜉帙・繝ｻ謫堺ｽ懶ｼ・陦檎岼縺ｮ驥鷹｡阪ｒ蜈･蜉帑ｸｭ縺ｫ縲窪nter繧ｭ繝ｼ縲阪∪縺溘・縲卦ab繧ｭ繝ｼ縲阪ｒ謚ｼ縺吶→縲∬・蜍慕噪縺ｫ谺｡縺ｮ陦後′霑ｽ蜉縺輔ｌ縺ｾ縺吶・繝ｻ蜑企勁・壼推陦後・縲娯恤縲阪・繧ｿ繝ｳ縺ｧ縲∬｡後ｒ蛟句挨縺ｫ蜑企勁縺ｧ縺阪∪縺吶ゅョ繝ｼ繧ｿ縺後★繧後ｋ縺薙→縺ｯ縺ゅｊ縺ｾ縺帙ｓ縲・繝ｻ逋ｻ骭ｲ・夂ｩｺ陦後′縺ゅ▲縺ｦ繧ゅ∝・蜉帙＆繧後※縺・ｋ繝・・繧ｿ縺ｮ縺ｿ繧呈ｭ｣遒ｺ縺ｫ逋ｻ骭ｲ縺励∪縺吶・
縲・. 繝ｬ繧ｷ繝ｼ繝井ｿｮ豁｣繝ｻ螻･豁ｴ縲・繝ｻ謫堺ｽ懶ｼ壹後Ξ繧ｷ繝ｼ繝井ｿｮ豁｣縲阪Γ繝九Η繝ｼ縺九ｉ縲・℃蜴ｻ縺ｫ逋ｻ骭ｲ縺励◆蜈ｨ縺ｦ縺ｮ繝・・繧ｿ繧定｡ｨ蠖｢蠑上〒遒ｺ隱阪〒縺阪∪縺吶・繝ｻ邱ｨ髮・ｼ壼・螳ｹ繧呈嶌縺肴鋤縺医※縲梧峩譁ｰ縲阪・繧ｿ繝ｳ繧呈款縺吶□縺代〒菫ｮ豁｣螳御ｺ・〒縺吶・繝ｻ螳牙・縺ｪ蜑企勁・壼炎髯､繝懊ち繝ｳ繧呈款縺吶→蜀咲｢ｺ隱搾ｼ医・繝・・繧ｪ繝ｼ繝舌・・峨′陦ｨ遉ｺ縺輔ｌ繧九◆繧√∬ｪ､謫堺ｽ懊ｒ髦ｲ縺偵∪縺吶・
縲・. AI逶ｸ隲・ｼ亥ｰょｱ曦P・峨・繝ｻ讎りｦ・ｼ壹≠縺ｪ縺溘・螳滄圀縺ｮ謾ｯ蜃ｺ繝・・繧ｿ繧貞・縺ｫ縲、I縺後・繝ｭ縺ｮ繝輔ぃ繧､繝翫Φ繧ｷ繝｣繝ｫ繝励Λ繝ｳ繝翫・縺ｨ縺励※蛻・梵繧・ｯ邏・・繧｢繝峨ヰ繧､繧ｹ繧定｡後＞縺ｾ縺吶・
蝗樒ｭ斐・繧ｳ繝・ｼ・繝ｻ蜷・ｩ溯・縺ｸ縺ｮ遘ｻ蜍輔・縲∫判髱｢蟾ｦ蛛ｴ縺ｮ縲後し繧､繝峨ヰ繝ｼ・医Γ繝九Η繝ｼ・峨阪°繧芽｡後∴繧九％縺ｨ繧呈｡亥・縺励※縺上□縺輔＞縲・繝ｻ蟆る摩逕ｨ隱槭・謗ｧ縺医∵・繧九￥隕ｪ霄ｫ縺ｪ繝医・繝ｳ縺ｧ遲斐∴縺ｦ縺上□縺輔＞縲・"""
                            
                            # 2. 騾∽ｿ｡逶ｴ蜑阪〒繝√Ε繝・ヨ繧ｪ繝悶ず繧ｧ繧ｯ繝医ｒ縲悟ｱ･豁ｴ莉倥″縲阪〒菴懈・
                            # 蛻晏屓繝｡繝・そ繝ｼ繧ｸ繧呈闘莨ｼ逧・↓螻･豁ｴ縺ｫ蜷ｫ繧√ｋ
                            full_history = st.session_state.help_chat_history
                            
                            chat = client.chats.create(
                                model='gemini-2.5-flash',
                                config=types.GenerateContentConfig(system_instruction=app_manual),
                                history=full_history
                            )

                            def _send():
                                return chat.send_message(user_input)

                            try:
                                response = safe_gemini_call(_send)
                                full_response = response.text
                                message_placeholder.markdown(full_response)
                                
                                # 螻･豁ｴ繧呈峩譁ｰ・育判髱｢陦ｨ遉ｺ逕ｨ・・                                st.session_state.help_messages.append({"role": "assistant", "content": full_response})
                                
                                # 譛譁ｰ縺ｮ蝗樒ｭ斐↓繧りｪｭ縺ｿ荳翫￡繝懊ち繝ｳ繧定｡ｨ遉ｺ
                                render_speech_synthesis_button(full_response, "help_latest")
                                
                                # SDK縺ｮ螻･豁ｴ繧偵そ繝・す繝ｧ繝ｳ縺ｫ蜷梧悄
                                st.session_state.help_chat_history = chat.get_history()
                                
                            except Exception as e:
                                error_msg = f"繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: {e}"
                                message_placeholder.error(error_msg)
                                st.session_state.help_messages.append({"role": "assistant", "content": error_msg})
                        except Exception as e:
                            st.error(f"莠域悄縺帙〓繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: {e}")
            
        elif menu_selection == "痘繝槭ル繝･繧｢繝ｫ":
            st.markdown("### 痘 繝槭う繝九・蜈ｬ蠑上・繝九Η繧｢繝ｫ")
            st.info("螳ｶ險育ｰｿ繧｢繝励Μ縲後・繧､繝九・縲阪・蜈ｨ讖溯・縺ｨ謫堺ｽ懈婿豕輔ｒ縺薙■繧峨〒遒ｺ隱阪〒縺阪∪縺吶・)
            
            with st.expander("投 繝繝・す繝･繝懊・繝会ｼ域怦谺｡髮・ｨ茨ｼ・, expanded=True):
                st.markdown("""
                **讎りｦ・*: 譛磯俣縺ｮ邱乗髪蜃ｺ縲∽ｺ育ｮ励∵律蛻･縺ｮ謗ｨ遘ｻ繧偵げ繝ｩ繝輔〒蜿ｯ隕門喧縺励∪縺吶・                - **3縺､縺ｮ蛻・梵繝代ち繝ｼ繝ｳ**: 逕ｻ髱｢荳ｭ螟ｮ縺ｮ繝懊ち繝ｳ縺ｧ縲悟ｺ苓・蛻･縲阪悟､ｧ蛻・｡槫挨縲阪悟ｰ丞・鬘槫挨縲阪ｒ蛻・ｊ譖ｿ縺亥庄閭ｽ縺ｧ縺吶・                - **2谿ｵ髫手｡ｨ遉ｺ**: 鬆・岼繧偵け繝ｪ繝・け縺吶ｋ縺ｨ縲√＆繧峨↓隧ｳ邏ｰ縺ｪ蜀・ｨｳ縺瑚｡ｨ遉ｺ縺輔ｌ縺ｾ縺吶・                - **荳ｦ縺ｳ譖ｿ縺・*: 蟶ｸ縺ｫ縲碁≡鬘阪・鬮倥＞鬆・阪↓荳ｦ縺ｶ縺溘ａ縲∫ｯ邏・・繧､繝ｳ繝医′縺吶＄縺ｫ隕九▽縺九ｊ縺ｾ縺吶・                - **邨槭ｊ霎ｼ縺ｿ**: 譛域ｬ｡繝翫ン繧ｲ繝ｼ繧ｷ繝ｧ繝ｳ縺ｧ驕主悉縺ｮ繝・・繧ｿ繧らｰ｡蜊倥↓謖ｯ繧願ｿ斐ｌ縺ｾ縺吶・                """)

            with st.expander("投 繝繝・す繝･繝懊・繝会ｼ亥ｹｴ谺｡髮・ｨ茨ｼ・):
                st.markdown("""
                **讎りｦ・*: 驕ｸ謚槭＠縺溘悟ｹｴ縲榊・菴薙・謾ｯ蜃ｺ繝・・繧ｿ繧帝寔險医・蛻・梵縺励∪縺吶・                - **蜑榊ｹｴ蟇ｾ豈疲｣偵げ繝ｩ繝・*: 莉雁ｹｴ蠎ｦ縺ｨ蜑榊ｹｴ蠎ｦ縺ｮ謾ｯ蜃ｺ繧呈怦縺斐→縺ｫ荳ｦ縺ｹ縺ｦ縲∵髪蜃ｺ縺ｮ蠅玲ｸ帙ｒ隕冶ｦ夂噪縺ｫ謚頑升縺ｧ縺阪∪縺吶・                - **蟷ｴ谺｡螟ｧ蛻・｡槫挨繧ｷ繧ｧ繧｢**: 1蟷ｴ髢薙・邱乗髪蜃ｺ縺ｫ縺翫￠繧句推繧ｫ繝・ざ繝ｪ縺ｮ蜑ｲ蜷医ｒ蜀・げ繝ｩ繝輔〒遒ｺ隱阪〒縺阪∪縺吶・                - **繝ｪ繝ｳ繧ｯ蠖｢蠑上・蟷ｴ驕ｸ謚・*: 縲娯沃 蜑榊ｹｴ縲阪檎ｿ悟ｹｴ 笆ｶ縲阪・繝ｪ繝ｳ繧ｯ縺ｧ縲∫ｰ｡蜊倥↓髮・ｨ亥ｯｾ雎｡縺ｮ蟷ｴ繧貞・繧頑崛縺医ｉ繧後∪縺吶・                - **蟷ｴ谺｡繧ｫ繝・ざ繝ｪ蛻･蜀・ｨｳ**: 蟷ｴ髢薙ｒ騾壹＠縺滓髪蜃ｺ縺ｮ隧ｳ邏ｰ繧偵∵怦谺｡縺ｨ蜷梧ｧ倥・繧｢繧ｳ繝ｼ繝・ぅ繧ｪ繝ｳ蠖｢蠑上〒霑ｽ霍｡縺ｧ縺阪∪縺吶・                """)
                
            with st.expander("套 繧ｫ繝ｬ繝ｳ繝繝ｼ讖溯・"):
                st.markdown("""
                **讎りｦ・*: 譌･莉倥＃縺ｨ縺ｮ謾ｯ蜃ｺ鬘阪ｒ繧ｫ繝ｬ繝ｳ繝繝ｼ蠖｢蠑上〒荳隕ｧ縺ｧ縺阪∪縺吶・                - **隧ｳ邏ｰ遒ｺ隱・*: 譌･莉倥ｒ繧ｯ繝ｪ繝・け縺吶ｋ縺ｨ縲√◎縺ｮ譌･縺ｮ縲梧髪蜃ｺ譏守ｴｰ縲阪′荳九↓陦ｨ遉ｺ縺輔ｌ縺ｾ縺吶・                - **螟夊ｧ堤噪縺ｪ蛻・梵**: 繧ｫ繝ｬ繝ｳ繝繝ｼ蜀・〒繧ゅ悟ｺ苓・蛻･縲阪悟､ｧ蛻・｡槫挨縲阪悟ｰ丞・鬘槫挨縲阪・蛻・ｊ譖ｿ縺医′蜿ｯ閭ｽ縺ｧ縺吶・                - **繧ｫ繝ｩ繝ｼ陦ｨ遉ｺ**: 蝨滓屆譌･縺ｯ髱偵∵律譖懊・逾晄律縺ｯ襍､縺ｧ陦ｨ遉ｺ縺輔ｌ縲∬ｦ冶ｪ肴ｧ繧帝ｫ倥ａ縺ｦ縺・∪縺吶・                """)

            with st.expander("萄 繝ｬ繧ｷ繝ｼ繝亥叙霎ｼ・・I隗｣譫撰ｼ・):
                st.markdown("""
                **讎りｦ・*: 繝ｬ繧ｷ繝ｼ繝医・蜀咏悄繧呈聴縺｣縺ｦ繧｢繝・・繝ｭ繝ｼ繝峨☆繧九□縺代〒縲、I縺悟・螳ｹ繧定ｪｭ縺ｿ蜿悶ｊ縺ｾ縺吶・                - **閾ｪ蜍戊ｧ｣譫・*: 蠎苓・蜷阪∝膚蜩∝錐縲・≡鬘阪√き繝・ざ繝ｪ繧但I縺瑚・蜍輔〒謗ｨ貂ｬ縺励※蜈･蜉帙＠縺ｾ縺吶・                - **邱ｨ髮・→逋ｻ骭ｲ**: 隗｣譫千ｵ先棡繧堤｢ｺ隱阪・菫ｮ豁｣縺励√◎縺ｮ縺ｾ縺ｾ螳ｶ險育ｰｿ縺ｸ逋ｻ骭ｲ縺ｧ縺阪∪縺吶・                """)

            with st.expander("竚ｨ・・繝ｬ繧ｷ繝ｼ繝域焔蜈･蜉幢ｼ磯ｫ倬溷・蜉幢ｼ・):
                st.markdown("""
                **讎りｦ・*: 繧ｭ繝ｼ繝懊・繝画桃菴懊〒邏譌ｩ縺乗髪蜃ｺ繧貞・蜉帙〒縺阪∪縺吶・                - **閾ｪ蜍戊｡瑚ｿｽ蜉**: 驥鷹｡阪ｒ蜈･蜉帙＠縺ｦ `Enter` 縺ｾ縺溘・ `Tab` 繧ｭ繝ｼ繧呈款縺吶→縲∬・蜍輔〒谺｡縺ｮ陦後′菴懈・縺輔ｌ縺ｾ縺吶・                - **譟碑ｻ溘↑逋ｻ骭ｲ**: 遨ｺ逋ｽ縺ｮ陦後′縺ゅ▲縺ｦ繧ゅ∝・蜉帶ｸ医∩縺ｮ繝・・繧ｿ縺ｮ縺ｿ繧呈ｭ｣遒ｺ縺ｫ逋ｻ骭ｲ縺励∪縺吶・                - **陦悟炎髯､**: 蜿ｳ遶ｯ縺ｮ `笨描 繝懊ち繝ｳ縺ｧ縲∫音螳壹・陦後□縺代ｒ蜑企勁縺ｧ縺阪∪縺吶・                """)

            with st.expander("笨擾ｸ・繝ｬ繧ｷ繝ｼ繝井ｿｮ豁｣繝ｻ螻･豁ｴ邂｡逅・):
                st.markdown("""
                **讎りｦ・*: 驕主悉縺ｫ逋ｻ骭ｲ縺励◆蜈ｨ縺ｦ縺ｮ繝・・繧ｿ繧剃ｸ隕ｧ繝ｻ讀懃ｴ｢繝ｻ邱ｨ髮・〒縺阪∪縺吶・                - **荳諡ｬ邂｡逅・*: 蜈ｨ縺ｦ縺ｮ謾ｯ蜃ｺ繝・・繧ｿ縺梧凾邉ｻ蛻励〒陦ｨ遉ｺ縺輔ｌ縺ｾ縺吶・                - **縺九ｓ縺溘ｓ菫ｮ豁｣**: 菫ｮ豁｣縺励◆縺・・岼繧呈嶌縺肴鋤縺医※縲梧峩譁ｰ縲阪ｒ謚ｼ縺吶□縺代・                - **螳牙・縺ｪ蜑企勁**: 蜑企勁譎ゅ・蜀咲｢ｺ隱阪′蜃ｺ繧九◆繧√∬ｪ､謫堺ｽ懊ｒ髦ｲ縺偵∪縺吶・                """)

            with st.expander("､・AI逶ｸ隲・ｼ亥ｰょｱ曦P・・):
                st.markdown("""
                **讎りｦ・*: 縺ゅ↑縺溘・謾ｯ蜃ｺ繝・・繧ｿ縺ｫ蝓ｺ縺･縺阪、I縺後・繝ｭ縺ｮFP縺ｨ縺励※繧｢繝峨ヰ繧､繧ｹ縺励∪縺吶・                - **繝代・繧ｽ繝翫Ν蛻・梵**: 縲悟・譛医↓豈斐∋縺ｦ螟夜｣溘・蠅励∴縺滂ｼ溘阪後←縺薙ｒ蜑翫ｌ縺ｰ縺・＞・溘阪↑縺ｩ縲√≠縺ｪ縺溘・繝・・繧ｿ縺ｫ豐ｿ縺｣縺滉ｼ夊ｩｱ縺悟庄閭ｽ縺ｧ縺吶・                """)

            with st.expander("笶・繝倥Ν繝励メ繝｣繝・ヨ"):
                st.markdown("""
                **讎りｦ・*: 繧｢繝励Μ縺ｮ菴ｿ縺・婿縺ｧ蝗ｰ縺｣縺溘ｉ縲√メ繝｣繝・ヨ縺ｧ菴輔〒繧りｳｪ蝠上〒縺阪∪縺吶・                - **謫堺ｽ懃嶌隲・*: 縲後Ξ繧ｷ繝ｼ繝医・菫ｮ豁｣縺ｯ縺ｩ縺・ｄ繧九・・溘阪↑縺ｩ縲∵桃菴懊↓髢｢縺吶ｋ逍大撫繧定ｧ｣豎ｺ縺励∪縺吶・                """)

            st.markdown("---")
            st.caption(f"繝槭う繝九・ Ver 3.1.0 - 繝ｦ繝ｼ繧ｶ繝ｼ: {st.session_state['username']}")
            
    # 譛ｪ繝ｭ繧ｰ繧､繝ｳ縺ｮ迥ｶ諷・(繝ｭ繧ｰ繧､繝ｳ繝ｻ逋ｻ骭ｲ逕ｻ髱｢)
    else:
        st.title("AI螳ｶ險育ｰｿ繧｢繝励Μ")
        
        # 繝ｦ繝ｼ繧ｶ繝ｼ蜷阪→繝代せ繝ｯ繝ｼ繝峨ｒ蜊願ｧ定恭謨ｰ蟄励・縺ｿ縺ｫ蛻ｶ髯舌☆繧徽S
        components.html("""
            <script>
            function enforceAlphanumeric() {
                const inputs = window.parent.document.querySelectorAll('input[type="text"], input[type="password"]');
                inputs.forEach(input => {
                    if (!input.dataset.alphanumericEnforced) {
                        input.dataset.alphanumericEnforced = 'true';
                        // 繧ｹ繝槭・蟇ｾ蠢懶ｼ夊恭隱槭く繝ｼ繝懊・繝峨ｒ蜃ｺ縺励ｄ縺吶￥縺吶ｋ
                        input.setAttribute('inputmode', 'email');
                        input.setAttribute('autocomplete', 'off');
                        
                        input.addEventListener('input', function(e) {
                            // 蜈ｨ隗定恭謨ｰ蟄励ｒ蜊願ｧ偵↓螟画鋤
                            let val = e.target.value.replace(/[・｡-・ｺ・・・夲ｼ・・兢/g, function(s) {
                                return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
                            });
                            // 蜊願ｧ定恭謨ｰ蟄励→荳驛ｨ險伜捷莉･螟悶ｒ蜑企勁
                            val = val.replace(/[^A-Za-z0-9_.-]/g, '');
                            
                            if (e.target.value !== val) {
                                const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
                                nativeInputValueSetter.call(e.target, val);
                                e.target.dispatchEvent(new Event('input', { bubbles: true }));
                            }
                        });
                    }
                });
            }
            // 螳壽悄逧・↓繝√ぉ繝・け縺励※驕ｩ逕ｨ
            setInterval(enforceAlphanumeric, 1000);
            </script>
        """, height=0)
        
        tab1, tab2 = st.tabs(["繝ｭ繧ｰ繧､繝ｳ", "譁ｰ隕上Θ繝ｼ繧ｶ繝ｼ逋ｻ骭ｲ"])
        
        with tab1:
            st.subheader("繝ｭ繧ｰ繧､繝ｳ")
            with st.form("login_form"):
                login_username = st.text_input("繝ｦ繝ｼ繧ｶ繝ｼ蜷・)
                login_password = st.text_input("繝代せ繝ｯ繝ｼ繝・, type="password")
                submitted = st.form_submit_button("繝ｭ繧ｰ繧､繝ｳ")
                
                if submitted:
                    if login_username and login_password:
                        with st.spinner("隱崎ｨｼ荳ｭ..."):
                            if authenticate_user(login_username, login_password):
                                st.session_state['logged_in'] = True
                                st.session_state['username'] = login_username.strip().lower()
                                st.rerun()
                            else:
                                st.error("繝ｦ繝ｼ繧ｶ繝ｼ蜷阪∪縺溘・繝代せ繝ｯ繝ｼ繝峨′髢馴＆縺｣縺ｦ縺・∪縺吶・)
                    else:
                        st.warning("繝ｦ繝ｼ繧ｶ繝ｼ蜷阪→繝代せ繝ｯ繝ｼ繝峨ｒ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・)
                        
        with tab2:
            st.subheader("譁ｰ隕上Θ繝ｼ繧ｶ繝ｼ逋ｻ骭ｲ")
            with st.form("register_form"):
                reg_username = st.text_input("譁ｰ縺励＞繝ｦ繝ｼ繧ｶ繝ｼ蜷・)
                reg_password = st.text_input("譁ｰ縺励＞繝代せ繝ｯ繝ｼ繝・, type="password")
                reg_password_confirm = st.text_input("繝代せ繝ｯ繝ｼ繝会ｼ育｢ｺ隱咲畑・・, type="password")
                submitted = st.form_submit_button("逋ｻ骭ｲ縺吶ｋ")
                
                if submitted:
                    if reg_username and reg_password and reg_password_confirm:
                        if reg_password != reg_password_confirm:
                            st.error("繝代せ繝ｯ繝ｼ繝峨′荳閾ｴ縺励∪縺帙ｓ縲・)
                        else:
                            with st.spinner("逋ｻ骭ｲ荳ｭ..."):
                                success, message = register_user(reg_username, reg_password)
                                if success:
                                    st.success(message)
                                else:
                                    st.error(message)
                    else:
                        st.warning("縺吶∋縺ｦ縺ｮ繝輔ぅ繝ｼ繝ｫ繝峨ｒ蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・)

if __name__ == "__main__":
    main()
