import streamlit as st
import jpholiday
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import gspread
import bcrypt
import os
import json
import io
import calendar
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from PIL import Image, ImageOps
from google import genai
from google.genai import types
import time
import re
import base64
import xlsxwriter
from fixed_cost_expansion import show_fixed_cost_data_expansion, show_open_management_sheet, show_variable_cost_update
# ---------- 構成設定 ----------
from urllib.parse import urlparse

SPREADSHEET_NAME = "Kakeibo_Data"
WORKSHEET_NAME = "users"
TRANSACTIONS_WORKSHEET_NAME = "transactions"
USER_MASTER_WORKSHEET_NAME = "User_Master"
PAYMENT_MASTER_WORKSHEET_NAME = "Payment_Master"
BANK_MASTER_WORKSHEET_NAME = "Bank_Master"
CATEGORY_MASTER_WORKSHEET_NAME = "Category_Master"
FIXED_COST_MASTER_WORKSHEET_NAME = "Fixed_Cost_Master"

# ---------- カテゴリ定義 ----------
# AI判別やセレクトボックスで利用するための大分類・小分類の親子関係定義
DEFAULT_EXPENSE_CATEGORIES = {
    "食材費": ["🍖肉類", "🐟魚類", "🥬野菜果物", "🍚主食類", "🍱惣菜", "🥚卵乳製品", "🥫加工食品", "🧂調味料", "🍫嗜好品", "☕飲料", "❓その他"],
    "外食費": ["🍜ラーメン", "🍣和食", "🥡中華", "🍕イタリアン", "☕カフェ", "🍺飲酒", "❓その他"],
    "日用品": ["🧻消耗品", "🧺掃除洗濯", "🛍️袋包装", "❓その他"],
    "美容": ["🧴ケア用品", "💄化粧品", "✂️散髪", "❓その他"],
    "衣類": ["👕衣類", "👟靴", "🧣小物", "❓その他"],
    "家電": ["📺家電", "💻周辺機器", "❓その他"],
    "書籍": ["📚書籍", "🖊️文具", "❓その他"],
    "交通費": ["🚃公共交通", "🚗車タクシー", "⛽ガソリン", "❓その他"],
    "住居": ["🛋️家具", "🏠住居用品", "❓その他"],
    "娯楽": ["🎡娯楽", "🎨グッズ", "❓その他"],
    "手数料": ["📦送料", "💳手数料", "❓その他"],
    "ペット用品": ["🐈フード", "🚽トイレ用品", "🏥ペット医療", "❓その他"],
    "医療": ["🏥病院診療", "💊薬処方", "💉検査健診", "❓その他"],
    "園芸・植物": ["🌻花・苗・種", "🪴観葉植物", "🧱土・肥料・鉢", "🛠️園芸用品", "❓その他"],
    "割引・ポイント利用": ["共通ポイント利用", "店舗独自ポイント利用", "クーポン割引", "キャッシュバック・還元"],
    "消費税（外税）": ["外税8%", "外税10%", "外税？％"],
    "消費税（内税）": ["内税8%", "内税10%", "内税？％"],
    "納税額": ["確定申告納税額", "その他"],
    "その他": ["📁未分類"],
    "支払い方法": ["クレジットカード", "デビットカード", "電子マネー", "ポイント", "現金", "銀行振込", "未設定", "その他"]
}

# グラフ用配色定義 (大分類)
# 棒グラフと円グラフの色を統一するためのマスターマップ
CATEGORY_COLOR_MAP = {
    "食材費": "#00008B",   # 濃い青
    "ペット用品": "#87CEEB", # 水色
    "交通費": "#DC3545",   # 赤
    "外食費": "#FF8C00",   # ダークオレンジ
    "日用品": "#9370DB",   # ミディアムパープル
    "美容": "#FF69B4",     # ホットピンク
    "衣類": "#20B2AA",     # ライトシーグリーン
    "家電": "#708090",     # スレートグレー
    "書籍": "#8B4513",     # サドルブラウン
    "住居": "#556B2F",     # ダークオリーブグリーン
    "娯楽": "#FFD700",     # ゴールド
    "手数料": "#A9A9A9",   # ダークグレー
    "医療": "#228B22",     # フォレストグリーン
    "園芸・植物": "#32CD32", # ライムグリーン
    "その他": "#778899",   # ライトスレートグレー
    "消費税（外税）": "#BC8F8F", # ロージーブラウン
    "消費税（内税）": "#D2B48C", # タン
    "割引・ポイント利用": "#FFFF00", # 黄色
    "納税額": "#4682B4",      # スチールブルー
    "支払い方法": "#8A2BE2"      # ブルーバイオレット
}

@st.cache_data(ttl=600)
def get_categories():
    try:
        sheet = get_sheet(CATEGORY_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        records = safe_gspread_call(sheet.get_all_records)
        
        if not records:
            headers = ["major_category", "minor_category"]
            safe_gspread_call(sheet.insert_row, headers, 1)
            rows = []
            for major, minors in DEFAULT_EXPENSE_CATEGORIES.items():
                if minors:
                    for minor in minors:
                        rows.append([major, minor])
                else:
                    rows.append([major, ""])
            if rows:
                safe_gspread_call(sheet.append_rows, rows)
            return DEFAULT_EXPENSE_CATEGORIES
            
        new_categories = {}
        for row in records:
            major = str(row.get("major_category", "")).strip()
            minor = str(row.get("minor_category", "")).strip()
            if major:
                if major not in new_categories:
                    new_categories[major] = []
                if minor and minor not in new_categories[major]:
                    new_categories[major].append(minor)
                    
        return new_categories if new_categories else DEFAULT_EXPENSE_CATEGORIES
    except Exception as e:
        print(f"Error loading categories: {e}")
        return DEFAULT_EXPENSE_CATEGORIES

def get_categories_prompt_text():
    """AI（Gemini等）のプロンプトに埋め込むためのカテゴリ定義文字列を生成"""
    text = "【カテゴリシステム: 大分類と小分類のリスト】\n"
    for major, minors in get_categories().items():
        text += f"- {major}: {', '.join(minors)}\n"
    text += "\n※ 必ず上記の大分類と小分類の組み合わせに従ってください。"
    return text

def safe_money_int_cast(val):
    """
    金額文字列（カンマ、￥、小数点あり）を安全に整数に変換する。
    AIの出力が "1,380" や "1380.0" などの場合に int() で落ちるのを防ぐ。
    """
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return int(val)
    try:
        # 文字列クリーンアップ
        s = str(val).replace(",", "").replace("￥", "").replace("¥", "").strip()
        if not s:
            return 0
        # 小数点が含まれる場合は一旦floatにしてからintにする
        if "." in s:
            return int(float(s))
        return int(s)
    except Exception:
        return 0

# ---------- セッション状態の初期化 ----------
if 'genai_client' not in st.session_state:
    st.session_state['genai_client'] = None

# APIキー設定（Gemini用）
# どちらの書き方でも動くようにします
api_key = st.secrets.get("GEMINI_API_KEY") or st.secrets.get("general", {}).get("gemini_api_key")

if not api_key and "general" in st.secrets:
    api_key = st.secrets["general"].get("gemini_api_key")

if api_key:
    st.session_state['genai_client'] = genai.Client(api_key=api_key)

st.set_page_config(page_title="AI家計簿アプリ - ダッシュボード", page_icon="📊", layout="wide")

# ---------- セッション状態の初期化 ----------
if '_init_done' not in st.session_state:
    params = st.query_params
    need_rerun = False
    
    if "date" in params:
        try:
            d_val = params["date"]
            if isinstance(d_val, list): d_val = d_val[0]
            dt_obj = datetime.strptime(d_val, "%Y-%m-%d")
            st.session_state['current_month'] = dt_obj.replace(day=1)
            st.session_state['selected_date'] = d_val
            if 'date_range' in st.session_state:
                del st.session_state['date_range']
            need_rerun = True
        except:
            pass
            
    # --- URLパラメータによる自動ログイン機能を削除 (セキュリティ向上のため) ---
    # 以前は ?user=... で自動ログイン可能だったが、これを廃止し必ずログインフォームを通すようにする。
    # --------------------------------------------------------------------------
        need_rerun = True
        
    if "menu" in params:
        m_val = params["menu"]
        if isinstance(m_val, list): m_val = m_val[0]
        st.session_state['menu_selection'] = m_val
        need_rerun = True

    st.session_state['_init_done'] = True
    
    if need_rerun:
        st.query_params.clear()
        # ログイン状態が保持されている場合はURLにユーザー名を残す（スマホブラウザの再読込対策）
        if st.session_state.get('logged_in') and st.session_state.get('username'):
            st.query_params['user'] = st.session_state['username']
        st.rerun()

if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'username' not in st.session_state:
    st.session_state['username'] = None
if 'remember_me' not in st.session_state:
    st.session_state['remember_me'] = False
if 'current_month' not in st.session_state:
    st.session_state['current_month'] = datetime.today().replace(day=1)

# ---------- Google Sheets 接続 ----------
@st.cache_resource
def get_gspread_client():
    try:
        # 1. secrets.toml から情報を読み込む
        if "gcp_service_account" in st.secrets:
            # 辞書形式に変換（画像 の修正を適用）
            info = dict(st.secrets["gcp_service_account"])

            from google.oauth2.service_account import Credentials
            
            # スプレッドシート操作に必要な権限（スコープ）
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            # 認証情報の作成
            creds = Credentials.from_service_account_info(info, scopes=scopes)
            return gspread.authorize(creds)
            
        elif os.path.exists("credentials.json"):
            return gspread.service_account(filename="credentials.json")
            
        else:
            st.error("認証設定が見つかりません。")
            return None
            
    except Exception as e:
        st.error(f"認証エラーが発生しました: {e}")
        return None

# --- リトライ可能なAPI呼び出しヘルパー ---

def safe_gspread_call(func, *args, max_retries=5, delay=2, **kwargs):
    """API呼び出しをリトライする関数（429等のAPIError対応）"""
    last_error = None
    for i in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            last_error = e
            err_msg = str(e)
            err_type = type(e).__name__
            # 一時的な接続エラーや、API制限(429), サーバーエラー(500系)の場合にリトライ
            # Streamlit Cloud環境ではエラーメッセージがマスクされるため、例外の型名でも判定する
            should_retry = any(keyword in err_msg for keyword in [
                "RemoteDisconnected", "Connection aborted", "TimeoutError",
                "429", "Quota exceeded", "APIError", "500", "502", "503",
                "This app has encountered an error" # Streamlit Cloudのマスクメッセージ
            ]) or err_type == "APIError"
            
            if should_retry:
                wait_time = delay * (2 ** i) # 指数バックオフ
                time.sleep(wait_time)
                continue
            else:
                # 致命的なエラー（認証等、その他）はすぐに上げる
                raise e
    raise last_error

def safe_gemini_call(func, *args, max_retries=5, initial_delay=2, **kwargs):
    """Gemini API呼び出しをリトライする関数（429/500/503エラー対応）"""
    last_error = None
    for i in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            last_error = e
            err_msg = str(e)
            err_type = type(e).__name__
            # 429 RESOURCE_EXHAUSTED または 500/503 系エラーの場合にリトライ
            # Streamlit Cloud環境でのメッセージマスクにも対応
            should_retry = (
                "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg or 
                "500" in err_msg or "503" in err_msg or
                "This app has encountered an error" in err_msg or
                err_type in ["ResourceExhausted", "InternalServerError", "ServiceUnavailable"]
            )
            
            if should_retry:
                wait_time = initial_delay * (2 ** i) # 指数バックオフ
                st.warning(f"現在混み合っています（{i+1}/{max_retries}回目）。{wait_time}秒後に再試行します...")
                time.sleep(wait_time)
                continue
            else:
                raise e
    raise last_error

def get_sheet(worksheet_name, create_if_not_found=False):
    client = get_gspread_client()
    if client is None:
        st.stop()
        
    try:
        # スプレッドシートと指定ワークシートに接続（リトライ付き）
        def _open_sheet():
            return client.open(SPREADSHEET_NAME).worksheet(worksheet_name)
        
        return safe_gspread_call(_open_sheet)
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"エラー: スプレッドシート '{SPREADSHEET_NAME}' が見つかりません。クレデンシャルのメールアドレス ({client.auth.signer_email}) とスプレッドシートを共有してください。")
        st.stop()
    except gspread.exceptions.WorksheetNotFound:
        if create_if_not_found:
            try:
                def _create_sheet():
                    ss = client.open(SPREADSHEET_NAME)
                    return ss.add_worksheet(title=worksheet_name, rows="1000", cols="20")
                return safe_gspread_call(_create_sheet)
            except Exception as e:
                st.error(f"エラー: シート '{worksheet_name}' の自動作成に失敗しました: {e}")
                st.stop()
        else:
            st.error(f"エラー: スプレッドシート内に '{worksheet_name}' シートが見つかりません。シートを新規作成してください。")
            st.stop()
    except Exception as e:
        st.error(f"Google Sheets接続エラー: {e}")
        st.stop()

def get_user_master_sheet(username, worksheet_name, create_if_not_found=False):
    """ユーザー個別の『{username}_支払管理』スプレッドシート内のワークシートを取得する"""
    client = get_gspread_client()
    if client is None:
        return None
        
    ss_name = f"{username}_支払管理"
    try:
        # スプレッドシートを開く（リトライ付き）
        def _open_ss():
            return client.open(ss_name)
        ss = safe_gspread_call(_open_ss)
        
        # ワークシートを取得
        try:
            def _get_ws():
                return ss.worksheet(worksheet_name)
            return safe_gspread_call(_get_ws)
        except gspread.exceptions.WorksheetNotFound:
            if create_if_not_found:
                def _add_ws():
                    return ss.add_worksheet(title=worksheet_name, rows="1000", cols="20")
                return safe_gspread_call(_add_ws)
            else:
                return None
    except gspread.exceptions.SpreadsheetNotFound:
        # シート未作成の場合は None を返す（呼び出し側で作成を促すメッセージ等を表示）
        return None
    except Exception as e:
        print(f"get_user_master_sheet error: {e}")
        return None

def init_users_sheet(sheet):
    """初期セットアップ：ヘッダーがない場合に作成する"""
    try:
        headers = safe_gspread_call(sheet.row_values, 1)
        if not headers or headers[0] != "username":
            safe_gspread_call(sheet.insert_row, ["username", "password_hash"], 1)
    except Exception as e:
        print(f"Init users sheet error: {e}")

def init_transactions_sheet(sheet):
    """初期セットアップ：取引シートのヘッダーがない場合に作成する"""
    expected_headers = ["username", "date", "store_name", "item_name", "category", "subcategory", "amount", "update", "payment_method", "payment_type", "closing_date", "payment_month", "payment_date", "receipt_id", "memo"]
    try:
        headers = safe_gspread_call(sheet.row_values, 1)
        if not headers or headers[0] != "username":
            safe_gspread_call(sheet.insert_row, expected_headers, 1)
        elif len(headers) < len(expected_headers):
            # 不足しているヘッダーを追記する
            for i in range(len(headers), len(expected_headers)):
                safe_gspread_call(sheet.update_cell, 1, i + 1, expected_headers[i])
    except Exception as e:
        print(f"Init transactions sheet error: {e}")

def init_user_master_sheet(sheet):
    """初期セットアップ：User_Masterシートのヘッダーがない場合に作成する"""
    expected_headers = ["username", "name", "gender", "birthdate", "mbti", "occupation", "hobbies", "life_stance", "ai_base_instruction"]
    try:
        headers = safe_gspread_call(sheet.row_values, 1)
        if not headers or headers[0] != "username":
            safe_gspread_call(sheet.insert_row, expected_headers, 1)
        elif len(headers) < len(expected_headers):
            # 不足しているヘッダーを追記する
            for i in range(len(headers), len(expected_headers)):
                safe_gspread_call(sheet.update_cell, 1, i + 1, expected_headers[i])
    except Exception as e:
        print(f"Init user_master sheet error: {e}")

def init_payment_master_sheet(sheet):
    """初期セットアップ：Payment_Masterシートのヘッダーがない場合に作成する"""
    expected_headers = ["username", "payment_id", "name", "type", "closing_date", "payment_month", "payment_date", "is_credit_card", "credit_limit"]
    try:
        headers = safe_gspread_call(sheet.row_values, 1)
        if not headers or headers[0] != "username":
            safe_gspread_call(sheet.insert_row, expected_headers, 1)
        elif len(headers) < len(expected_headers):
            for i in range(len(headers), len(expected_headers)):
                safe_gspread_call(sheet.update_cell, 1, i + 1, expected_headers[i])
    except Exception as e:
        print(f"Init payment_master sheet error: {e}")

def init_fixed_cost_master_sheet(sheet):
    """初期セットアップ：Fixed_Cost_Masterシートのヘッダーがない場合に作成する"""
    expected_headers = [
        "username", "fixed_cost_id", "major_category", "payment_1", "payment_2", 
        "is_finite", "item_name", "amount", "fixed_or_variable", 
        "payment_month", "final_amount", "transfer_fee", "start_month", "completion_month"
    ]
    try:
        headers = safe_gspread_call(sheet.row_values, 1)
        if not headers or headers[0] != "username":
            safe_gspread_call(sheet.insert_row, expected_headers, 1)
        elif len(headers) < len(expected_headers):
            for i in range(len(headers), len(expected_headers)):
                safe_gspread_call(sheet.update_cell, 1, i + 1, expected_headers[i])
    except Exception as e:
        print(f"Init fixed_cost_master sheet error: {e}")



def init_bank_master_sheet(sheet):
    expected_headers = ["username", "bank_id", "bank_name", "balance"]
    try:
        headers = safe_gspread_call(sheet.row_values, 1)
        if not headers or headers[0] != "username":
            safe_gspread_call(sheet.insert_row, expected_headers, 1)
        elif len(headers) < len(expected_headers):
            for i in range(len(headers), len(expected_headers)):
                safe_gspread_call(sheet.update_cell, 1, i + 1, expected_headers[i])
    except Exception as e:
        print(f"Init bank_master sheet error: {e}")

def get_banks(username):
    try:
        sheet = get_user_master_sheet(username, BANK_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        if not sheet:
            return []
        init_bank_master_sheet(sheet)
        records = safe_gspread_call(sheet.get_all_records)
        banks = []
        if records:
            for row in records:
                # ユーザー個別シートなのでフィルタリングは不要だが、
                # 念のため既存ロジックを尊重
                if not row.get("username") or str(row.get("username", "")).lower() == username.lower():
                    try:
                        b_val = row.get("balance", 0)
                        row["balance"] = int(float(str(b_val).replace(',', '').replace('¥', '').replace('￥', ''))) if b_val != "" else 0
                    except ValueError:
                        row["balance"] = 0
                    banks.append(row)
        return banks
    except Exception as e:
        import streamlit as st
        st.error(f"銀行マスター取得エラー: {e}")
        return []

def add_bank(username, bank_id, bank_name, balance=0):
    try:
        sheet = get_user_master_sheet(username, BANK_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        if not sheet:
            return False
        init_bank_master_sheet(sheet)
        safe_gspread_call(sheet.append_row, [username, bank_id, bank_name, balance])
        return True
    except Exception as e:
        import streamlit as st
        st.error(f"銀行登録エラー: {e}")
        return False

def delete_bank(username, bank_id):
    try:
        sheet = get_user_master_sheet(username, BANK_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        if not sheet:
            return False
        init_bank_master_sheet(sheet)
        records = safe_gspread_call(sheet.get_all_records)
        row_idx = -1
        for i, row in enumerate(records):
            if str(row.get("bank_id", "")) == str(bank_id):
                row_idx = i + 2
                break
        if row_idx > 1:
            safe_gspread_call(sheet.delete_rows, row_idx)
            return True
        return False
    except Exception as e:
        import streamlit as st
        st.error(f"銀行削除エラー: {e}")
        return False
def update_bank(username, bank_id, bank_name, balance):
    try:
        sheet = get_user_master_sheet(username, BANK_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        if not sheet:
            return False
        init_bank_master_sheet(sheet)
        records = safe_gspread_call(sheet.get_all_records)
        row_idx = -1
        for i, row in enumerate(records):
            if str(row.get("bank_id", "")) == str(bank_id):
                row_idx = i + 2
                break
        
        if row_idx > 1:
            safe_gspread_call(sheet.update_cell, row_idx, 3, bank_name)
            safe_gspread_call(sheet.update_cell, row_idx, 4, balance)
            return True
        return False
    except Exception as e:
        import streamlit as st
        st.error(f"銀行更新エラー: {e}")
        return False

def get_fixed_costs(username):
    """ユーザーの固定費リストを取得する"""
    try:
        sheet = get_sheet(FIXED_COST_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        init_fixed_cost_master_sheet(sheet)
        records = safe_gspread_call(sheet.get_all_records)
        costs = []
        if records:
            for row in records:
                if str(row.get("username", "")).lower() == username.lower():
                    # 数値パース
                    for col in ["amount", "final_amount", "transfer_fee"]:
                        try:
                            val = row.get(col, "")
                            row[col] = int(float(str(val).replace(',', '').replace('¥', '').replace('￥', ''))) if val != "" else 0
                        except ValueError:
                            row[col] = 0
                    costs.append(row)
        return costs
    except Exception as e:
        st.error(f"固定費マスター取得エラー: {e}")
        return []

def add_fixed_cost(username, fixed_cost_id, major_category, payment_1, payment_2, 
                   is_finite, item_name, amount, fixed_or_variable, payment_month, 
                   final_amount, transfer_fee, start_month, completion_month):
    """新しい固定費を登録する"""
    try:
        sheet = get_sheet(FIXED_COST_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        init_fixed_cost_master_sheet(sheet)
        new_row = [
            username, fixed_cost_id, major_category, payment_1, payment_2,
            is_finite, item_name, amount, fixed_or_variable, payment_month, 
            final_amount, transfer_fee, start_month, completion_month
        ]
        safe_gspread_call(sheet.append_row, new_row)
        return True
    except Exception as e:
        st.error(f"固定費登録エラー: {e}")
        return False

def update_fixed_cost(username, fixed_cost_id, major_category, payment_1, payment_2, is_finite, item_name, amount, fixed_or_variable, payment_month, final_amount=0, transfer_fee=0, start_month="", completion_month=""):
    try:
        sheet = get_sheet(FIXED_COST_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        init_fixed_cost_master_sheet(sheet)
        records = safe_gspread_call(sheet.get_all_records)
        row_idx = -1
        for i, row in enumerate(records):
            if str(row.get("username", "")).lower() == username.lower() and str(row.get("fixed_cost_id", "")) == str(fixed_cost_id):
                row_idx = i + 2
                break
        
        if row_idx > 1:
            row_data = [username, fixed_cost_id, major_category, payment_1, payment_2, is_finite, item_name, amount, fixed_or_variable, payment_month, final_amount, transfer_fee, start_month, completion_month]
            safe_gspread_call(sheet.update, f"A{row_idx}:N{row_idx}", [row_data])
            return True
        return False
    except Exception as e:
        import streamlit as st
        st.error(f"固定費更新エラー: {e}")
        return False

def delete_fixed_cost(username, fixed_cost_id):
    """固定費を削除する"""
    try:
        sheet = get_sheet(FIXED_COST_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        init_fixed_cost_master_sheet(sheet)
        records = safe_gspread_call(sheet.get_all_records)
        row_idx = -1
        for i, row in enumerate(records):
            if str(row.get("username", "")).lower() == username.lower() and str(row.get("fixed_cost_id", "")) == str(fixed_cost_id):
                row_idx = i + 2
                break
        
        if row_idx > 1:
            safe_gspread_call(sheet.delete_rows, row_idx)
            return True
        return False
    except Exception as e:
        st.error(f"固定費削除エラー: {e}")
        return False

def get_payment_methods(username):
    """ユーザーの支払い方法リストを取得する"""
    try:
        sheet = get_sheet(PAYMENT_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        init_payment_master_sheet(sheet)
        records = safe_gspread_call(sheet.get_all_records)
        
        methods = []
        if records:
            for row in records:
                if str(row.get("username", "")).lower() == username.lower():
                    # boolean パース
                    is_cc_raw = str(row.get("is_credit_card", "False")).lower()
                    row["is_credit_card"] = True if is_cc_raw == "true" else False
                    
                    # limit パース
                    limit_raw = row.get("credit_limit", "")
                    try:
                        row["credit_limit"] = int(float(limit_raw)) if limit_raw else 0
                    except ValueError:
                        row["credit_limit"] = 0
                    
                    methods.append(row)
        return methods
    except Exception as e:
        st.error(f"支払い方法取得エラー: {e}")
        return []

def get_payment_details_for_transaction(username, method_name):
    """Transactionsに付与する支払い方法詳細（種類、締日、支払月、支払日）を取得する"""
    if not method_name or method_name == "未設定":
        return "未設定", "", "", ""
        
    methods = get_payment_methods(username)
    for m in methods:
        if m.get("name") == method_name:
            p_type = m.get("type", "その他")
            is_cc = m.get("is_credit_card", False)
            if is_cc or p_type == "クレジットカード":
                return (
                    p_type,
                    m.get("closing_date", ""),
                    m.get("payment_month", ""),
                    m.get("payment_date", "")
                )
            else:
                return (p_type, "", "", "")
    return "その他", "", "", ""

# --- クレジットカード利用額計算用ヘルパー ---
def get_last_day_of_month(target_date):
    """指定した月の日付から、その月の末日（datetime.date）を返す"""
    _, last_day = calendar.monthrange(target_date.year, target_date.month)
    return target_date.replace(day=last_day).date()

def calculate_credit_card_periods(target_date, closing_str, pay_month_str, pay_date_str):
    """
    指定月(target_date: datetime)周辺の、クレジットカードの3つの支払期間を算出する。
    戻り値 (リスト形式で3期間を返す):
      [
        {"label": "①当月支払", "start": date, "end": date, "pay_date": date},
        {"label": "②次回支払額", "start": date, "end": date, "pay_date": date},
        {"label": "③次回以降支払額", "start": date, "end": date, "pay_date": date}
      ]
    ※ 設定が未入力の場合は空リストを返す。
    """
    closing_str = str(closing_str) if closing_str else ""
    pay_month_str = str(pay_month_str) if pay_month_str else ""
    pay_date_str = str(pay_date_str) if pay_date_str else ""
    
    if not closing_str or not pay_month_str or not pay_date_str:
        return []
        
    # target_date を基本の「当月1日」とする
    base_calc_date = target_date.replace(day=1)
    
    # 支払い月のオフセット計算
    pay_month_offset = 0
    if "当月" in pay_month_str: pay_month_offset = 0
    elif "翌々月" in pay_month_str: pay_month_offset = 2
    elif "翌月" in pay_month_str: pay_month_offset = 1

    try:
        periods = []
        labels = ["①当月支払", "②次回支払額", "③次回以降支払額"]
        
        # i=0: 当月支払, i=1: 次回支払, i=2: 次回以降支払
        for i in range(3):
            # nヶ月後の支払月を求める
            payment_month_date = base_calc_date + relativedelta(months=i)
            # その支払月の対象となる利用「基準月」をオフセットから逆算する
            billing_base_date = payment_month_date - relativedelta(months=pay_month_offset)
            
            # --- 支払日の決定 ---
            if "末" in pay_date_str:
                pay_date = get_last_day_of_month(payment_month_date)
            else:
                p_day_str = pay_date_str.replace("日払い", "").replace("日", "").strip()
                p_day = int(p_day_str) if p_day_str.isdigit() else 27 # fallback
                pay_date = payment_month_date.replace(day=p_day).date()
                
            # --- 利用期間の決定 ---
            if "末" in closing_str:
                start_date = billing_base_date.date()
                end_date = get_last_day_of_month(billing_base_date)
            else:
                c_day_str = closing_str.replace("日締め", "").replace("日", "").strip()
                c_day = int(c_day_str) if c_day_str.isdigit() else 15 # fallback
                end_date = billing_base_date.replace(day=c_day).date()
                
                prev_m = billing_base_date - relativedelta(months=1)
                start_date = (prev_m.replace(day=c_day) + relativedelta(days=1)).date()
                
            periods.append({
                "label": labels[i],
                "start": start_date,
                "end": end_date,
                "pay_date": pay_date
            })
            
        return periods

    except Exception as e:
        print(f"Date Calc Error: {e}")
        return []

def save_payment_method(username, payment_data):
    """支払い方法を保存（新規追加または更新）する"""
    try:
        sheet = get_sheet(PAYMENT_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        init_payment_master_sheet(sheet)
        
        # payment_idで既存レコードを検索
        records = safe_gspread_call(sheet.get_all_records)
        row_idx = None
        if records:
            for i, row in enumerate(records):
                if str(row.get("username", "")).lower() == username.lower() and str(row.get("payment_id", "")) == str(payment_data.get("payment_id", "")):
                    row_idx = i + 2  # ヘッダー行+1でスプレッドシートの行番号(1始まり)
                    break
        
        row_data = [
            username.lower(),
            payment_data.get("payment_id", ""),
            payment_data.get("name", ""),
            payment_data.get("type", "現金"),
            payment_data.get("closing_date", ""),
            payment_data.get("payment_month", ""),
            payment_data.get("payment_date", ""),
            str(payment_data.get("is_credit_card", False)),
            str(payment_data.get("credit_limit", ""))
        ]
        
        if row_idx:
            # 既存の行を更新
            update_range = f"A{row_idx}:I{row_idx}"
            safe_gspread_call(sheet.update, range_name=update_range, values=[row_data])
        else:
            # 新規追加
            safe_gspread_call(sheet.append_row, row_data)
            
        return True, "支払い方法を保存しました。"
    except Exception as e:
        return False, f"支払い方法保存エラー: {e}"

def delete_payment_method(username, payment_id):
    """支払い方法を削除する"""
    try:
        sheet = get_sheet(PAYMENT_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        init_payment_master_sheet(sheet)
        
        records = safe_gspread_call(sheet.get_all_records)
        row_idx = None
        if records:
            for i, row in enumerate(records):
                if str(row.get("username", "")).lower() == username.lower() and str(row.get("payment_id", "")) == str(payment_id):
                    row_idx = i + 2
                    break
        
        if row_idx:
            safe_gspread_call(sheet.delete_rows, row_idx)
            return True, "支払い方法を削除しました。"
        return False, "削除対象が見つかりませんでした。"
    except Exception as e:
        return False, f"支払い方法削除エラー: {e}"

# ---------- 認証機能 ----------
def register_user(username, password):
    sheet = get_sheet(WORKSHEET_NAME)
    init_users_sheet(sheet)
    
    # 仕様要件: ユーザー名は lower() で処理する
    username = username.strip().lower()
    
    if not username or not password:
        return False, "ユーザー名とパスワードを入力してください。"
        
    # 既存ユーザーの重複チェック
    records = sheet.get_all_records()
    for row in records:
        if str(row.get("username", "")).lower() == username:
            return False, "このユーザー名は既に登録されています。"
            
    # 仕様要件: パスワードは bcrypt でハッシュ化する
    salt = bcrypt.gensalt()
    hashed_password = bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')
    
    # 登録データの追加
    sheet.append_row([username, hashed_password])
    
    # --- 新規ユーザー向け：デフォルトの支払い方法を自動登録 ---
    import uuid
    default_methods = [
        {"name": "未設定", "type": "未設定"},
        {"name": "現金", "type": "現金"},
        {"name": "PayPay", "type": "電子マネー"}
    ]
    for method in default_methods:
        payment_data = {
            "payment_id": str(uuid.uuid4()),
            "name": method["name"],
            "type": method["type"],
            "closing_date": "",
            "payment_month": "",
            "payment_date": "",
            "is_credit_card": False,
            "credit_limit": ""
        }
        save_payment_method(username, payment_data)
        
    return True, "登録が完了しました。ログインタブからログインしてください。"

def authenticate_user(username, password):
    sheet = get_sheet(WORKSHEET_NAME)
    init_users_sheet(sheet)
    
    # ユーザー名照合のため小文字化
    username = username.strip().lower()
    
    records = sheet.get_all_records()
    for row in records:
        if str(row.get("username", "")).lower() == username:
            stored_hash = str(row.get("password_hash", ""))
            # bcrypt でハッシュを照合
            if bcrypt.checkpw(password.encode('utf-8'), stored_hash.encode('utf-8')):
                return True
    return False

def change_user_password(username, old_password, new_password):
    """ユーザーのパスワードを変更する"""
    try:
        sheet = get_sheet(WORKSHEET_NAME)
        init_users_sheet(sheet)
        username = username.strip().lower()
        
        records = safe_gspread_call(sheet.get_all_records)
        row_idx = None
        stored_hash = None
        
        for i, row in enumerate(records):
            if str(row.get("username", "")).lower() == username:
                row_idx = i + 2  # データ行は2行目から
                stored_hash = str(row.get("password_hash", ""))
                break
                
        if not row_idx or not stored_hash:
            return False, "ユーザーが見つかりません。"
            
        import bcrypt
        if not bcrypt.checkpw(old_password.encode('utf-8'), stored_hash.encode('utf-8')):
            return False, "現在のパスワードが間違っています。"
            
        # 新しいパスワードをハッシュ化
        salt = bcrypt.gensalt()
        new_hashed = bcrypt.hashpw(new_password.encode('utf-8'), salt).decode('utf-8')
        
        # B列 (password_hash) を更新
        update_range = f"B{row_idx}"
        safe_gspread_call(sheet.update, range_name=update_range, values=[[new_hashed]])
        
        return True, "パスワードを正常に変更しました。次回から新しいパスワードでログインしてください。"
    except Exception as e:
        return False, f"パスワード変更エラー: {e}"

@st.cache_data(ttl=300)
def get_user_master_data(username):
    """ログインユーザーのプロフィール情報を取得する"""
    try:
        sheet = get_sheet(USER_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        init_user_master_sheet(sheet)
        records = safe_gspread_call(sheet.get_all_records)
        
        if not records:
            return None
            
        for row in records:
            if str(row.get("username", "")).lower() == username.lower():
                return row
        return None
    except Exception as e:
        st.error(f"プロフィール取得エラー: {e}")
        return None

def save_user_master_data(username, profile_data):
    """ログインユーザーのプロフィール情報を保存（新規または上書き）する"""
    try:
        sheet = get_sheet(USER_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        init_user_master_sheet(sheet)
        
        cell = safe_gspread_call(sheet.find, username.lower(), in_column=1)
        row_data = [
            username.lower(),
            profile_data.get("name", ""),
            profile_data.get("gender", ""),
            profile_data.get("birthdate", ""),
            profile_data.get("mbti", ""),
            profile_data.get("occupation", ""),
            profile_data.get("hobbies", ""),
            profile_data.get("life_stance", ""),
            profile_data.get("ai_base_instruction", "")
        ]
        
        if cell:
            # 既存の行を更新
            row_idx = cell.row
            # A列(1) から I列(9) までの範囲を更新
            update_range = f"A{row_idx}:I{row_idx}"
            safe_gspread_call(sheet.update, range_name=update_range, values=[row_data])
        else:
            # 新規追加
            safe_gspread_call(sheet.append_row, row_data)
            
        return True, "プロフィール設定を保存しました。"
    except Exception as e:
        return False, f"プロフィール保存エラー: {e}"

# ---------- データ取得機能 (共通ヘルパー) ----------

def get_clean_df(records, username):
    """
    レコードからDataFrameを作成し、カラム名の正規化(日本語対応)とユーザーフィルタリングを行う
    """
    if not records:
        return pd.DataFrame()
        
    df = pd.DataFrame(records)
    
    # クレンジング（不要なスペース等削除、小文字化してマッチングしやすくする）
    df.columns = df.columns.str.strip()
    
    # 既存のカラム名（小文字）とターゲットのマップを作成
    col_map = {c.lower(): c for c in df.columns}
    
    # --- カラム名の正規化（日本語ヘッダー・大文字小文字への対応） ---
    rename_rules = {
        "日付": "date",
        "date": "date",
        "ユーザー名": "username",
        "username": "username",
        "user": "username",
        "店舗名": "store_name",
        "店舗": "store_name",
        "商品名": "item_name",
        "内容": "item_name",
        "金額": "amount",
        "大分類": "category",
        "小分類": "subcategory",
        "支払い方法": "payment_method",
        "payment_method": "payment_method",
        "receipt_id": "receipt_id",
        "レシートid": "receipt_id"
    }
    
    actual_rename = {}
    for key, target in rename_rules.items():
        # keyがカラム名（そのまま、または小文字）に含まれているか確認
        if key in df.columns:
            actual_rename[key] = target
        elif key.lower() in col_map:
            actual_rename[col_map[key.lower()]] = target
            
    if actual_rename:
        df = df.rename(columns=actual_rename)
    
    # "username"でフィルタ (必須)
    if "username" in df.columns:
        # 値自体の余白も削除して比較
        df["username"] = df["username"].astype(str).str.strip().str.lower()
        df = df[df["username"] == username.lower()]
    else:
        return pd.DataFrame()
        
    if df.empty or "date" not in df.columns:
          return pd.DataFrame()

    # "date"列をdatetime型に変換
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"])
    
    # "payment_method"が空の場合は「未設定」とする
    if "payment_method" in df.columns:
        # 空文字やスペースのみの文字、nullを「未設定」にする
        df["payment_method"] = df["payment_method"].replace(r"^\s*$", "未設定", regex=True)
        df["payment_method"] = df["payment_method"].fillna("未設定")
    else:
        df["payment_method"] = "未設定"
    
    return df

# ---------- データ取得機能 ----------
def load_transactions_data(target_date, mode="monthly"):
    """
    指定した月または年の、ログインユーザーのデータを取得する
    mode: "monthly" (月次) または "yearly" (年次)
    """
    sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
    init_transactions_sheet(sheet)
    values = safe_gspread_call(sheet.get_all_values)
    if not values or len(values) < 2:
        return pd.DataFrame()
    
    # 重複・空ヘッダー対策: DataFrame化してからカラム名を付与
    headers = [h.strip() if h.strip() else f"empty_{i}" for i, h in enumerate(values[0])]
    # 重複がある場合はpandasが自動で .1, .2 を付けるが、明示的にハンドル
    records_df = pd.DataFrame(values[1:])
    # カラム数が一致しない場合のガード
    if records_df.shape[1] > len(headers):
        headers += [f"extra_{i}" for i in range(len(headers), records_df.shape[1])]
    records_df.columns = headers[:records_df.shape[1]]
    records = records_df.to_dict('records')
    
    # 共通ヘルパーでクレンジングとフィルタリング
    curr_user = st.session_state.get('username', "")
    df = get_clean_df(records, curr_user)
    
    if df.empty:
         return pd.DataFrame()
    
    # 行インデックスの付与（recordsの順番に基づく）
    # recordsのインデックスとdfのインデックスを合わせる必要があるため、クレンジング前のrecords長を使用
    # recordsは全ユーザー分あるが、dfはフィルタ済み。
    # recordsにある元の行番号を保持するためにDataFrame作成時に付与しておく
    df_all_temp = records_df.copy()
    df_all_temp['_row_index'] = range(2, len(records) + 2)
    
    # dfにrow_indexを結合
    # pd.mergeを使うため、元のインデックスを利用
    df = df.join(df_all_temp[['_row_index']])
    
    # 期間でフィルタ
    if mode == "monthly":
        df = df[(df["date"].dt.year == target_date.year) & (df["date"].dt.month == target_date.month)]
    elif mode == "yearly":
        df = df[df["date"].dt.year == target_date.year]
    # mode == "all" の場合は日付フィルタを行わず全期間を返す
    
    # 金額を数値に変換
    if "amount" in df.columns:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    else:
        df["amount"] = 0
        
    # --- カテゴリの正規化（集計時やセレクトボックス等で指定外が出ないようにする） ---
    # 大分類の正規化
    if "category" in df.columns:
        valid_majors = list(get_categories().keys())
        # 定義にない大分類は「その他」にまとめる
        df["category"] = df["category"].apply(lambda x: x if x in valid_majors else "その他")
        
    # 小分類の正規化
    sub_cols = [c for c in ["subcategory", "sub_category", "小分類"] if c in df.columns]
    if sub_cols:
        sub_col = sub_cols[0]
        def normalize_sub(row):
            major = row.get("category", "その他")
            sub = str(row.get(sub_col, "")).strip()
            valid_subs = get_categories().get(major, sorted(get_categories()["その他"]))
            fallback = valid_subs[-1] if valid_subs else "❓その他"
            
            # 完全に一致するか
            if sub in valid_subs:
                return sub
                
            # アイコンなしなどの部分一致を探す
            for v_sub in valid_subs:
                # 絵文字を除いたテキストで部分一致するか確認
                text_only = "".join([c for c in v_sub if c.isalnum() or c in "類物食品未分類その他%"]) 
                if text_only and (text_only in sub or sub in text_only) and len(sub) > 0:
                    return v_sub
                    
            return fallback

        df[sub_col] = df.apply(normalize_sub, axis=1)
        
    return df

def get_transaction_range(username):
    """ユーザーの全データの最小年月と最大年月を取得し、セッションに保持する"""
    # 修正を即時反映させるため、一時的にキャッシュを無効化するか、強制リフレッシュを挟む
    # ここでは、もし不整合があれば再取得するようにガードを強化
    if 'date_range' in st.session_state and st.session_state.get('last_range_fetch_user') == username:
        if st.session_state['date_range']: # 空でないことを確認
            return st.session_state['date_range']
    
    sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
    values = safe_gspread_call(sheet.get_all_values)
    if not values or len(values) < 2:
        return []
        
    headers = [h.strip() if h.strip() else f"empty_{i}" for i, h in enumerate(values[0])]
    records_df = pd.DataFrame(values[1:])
    if records_df.shape[1] > len(headers):
        headers += [f"extra_{i}" for i in range(len(headers), records_df.shape[1])]
    records_df.columns = headers[:records_df.shape[1]]
    records = records_df.to_dict('records')
    
    # 共通ヘルパーでクレンジングとフィルタリング
    df_user = get_clean_df(records, username)
    
    if df_user.empty:
        return []
        
    # ユニークな年月を抽出してソート
    df_user["year_month"] = df_user["date"].dt.to_period("M").dt.to_timestamp()
    available_months = sorted(df_user["year_month"].unique().tolist())
    
    st.session_state['date_range'] = available_months
    st.session_state['last_range_fetch_user'] = username
    return available_months

def render_year_navigation():
    """年次集計用の年選択ナビゲーションを表示する (データがある年のみ移動可能)"""
    curr = st.session_state.get('current_month', datetime.today().replace(day=1))
    current_user = st.session_state.get("username", "")
    current_menu = st.session_state.get("menu_selection", "カレンダー")
    
    # データが存在するユニークな月リストを取得
    available_months = get_transaction_range(current_user)
    available_years = sorted(list(set([m.year for m in available_months])))
    
    # 前後の年を検索
    prev_y = next((y for y in reversed(available_years) if y < curr.year), None)
    next_y = next((y for y in available_years if y > curr.year), None)
    
    has_prev = prev_y is not None
    has_next = next_y is not None
    
    prev_date_str = f"{prev_y}-01-01" if has_prev else ""
    next_date_str = f"{next_y}-01-01" if has_next else ""
    
    # 月次と同様のCSSを適用して1行に収める
    st.markdown("""
        <style>
            /* st.columns の親要素に対して横並びを強制 */
            div[data-testid="stHorizontalBlock"] {
                display: flex !important;
                flex-direction: row !important;
                flex-wrap: nowrap !important;
                align-items: center !important;
                justify-content: center !important;
                gap: 2px !important;
            }
            /* 前年・翌年ボタンのカラム */
            div[data-testid="stHorizontalBlock"] > div:nth-child(1),
            div[data-testid="stHorizontalBlock"] > div:nth-child(3) {
                flex: 1 1 0% !important;
                width: auto !important;
                min-width: 0 !important;
            }
            /* 当年ポップオーバーのカラム */
            div[data-testid="stHorizontalBlock"] > div:nth-child(2) {
                flex: 0 0 auto !important;
                width: 130px !important;
                min-width: 130px !important;
            }
            /* ポップオーバーボタンの余白を削る */
            div[data-testid="stPopover"] > button {
                padding-left: 2px !important;
                padding-right: 2px !important;
            }
            /* ポップオーバーボタンのテキストサイズと折り返し禁止 */
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
                    <a href="/?date={prev_date_str}&user={current_user}&menu={current_menu}" target="_self" style='text-decoration: none; color: #007bff; font-weight: bold;'>◀ 前年</a>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='text-align: right; color: #ccc; font-size: 0.9rem;'>◀ 前年</div>", unsafe_allow_html=True)
            
    with col2:
        # 月次とスタイルを合わせてポップオーバー化
        with st.popover(curr.strftime('%Y年 ▼'), use_container_width=True):
            # 選択可能な年のリストを作成 (データ範囲内)
            if available_years:
                nav_years = sorted(available_years, reverse=True)
            else:
                nav_years = [curr.year]

            list_html = "<div style='text-align: center;'>"
            for y in nav_years:
                y_str = f"{y}-01-01"
                y_label = f"{y}年"
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
                    <a href="/?date={next_date_str}&user={current_user}&menu={current_menu}" target="_self" style='text-decoration: none; color: #007bff; font-weight: bold;'>翌年 ▶</a>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='text-align: left; color: #ccc; font-size: 0.9rem;'>翌年 ▶</div>", unsafe_allow_html=True)
    
    st.markdown("---")

def render_month_navigation():
    """全機能共通の月選択ナビゲーションと月間合計を表示する"""
    # 現在の月を取得
    if 'current_month' not in st.session_state:
        st.session_state['current_month'] = datetime.today().replace(day=1)
    
    curr = st.session_state['current_month']
    
    # 年月リストの作成 (2023年1月から翌年末まで、昇順)
    start_date = datetime(2023, 1, 1)
    end_date = datetime.today().replace(day=1) + relativedelta(years=1, month=12)
    
    month_options = []
    temp_date = start_date
    while temp_date <= end_date:
        month_options.append(temp_date)
        temp_date += relativedelta(months=1)
    
    # 降順（新しい順）に並び替える
    month_options.reverse()
    
    # 表示用のラベル作成
    month_labels = [dt.strftime('%Y年%m月') for dt in month_options]
    
    # 現在のインデックスを取得
    try:
        current_idx = month_options.index(curr.replace(day=1))
    except ValueError:
        # 万が一見つからない場合はリストの先頭（最新）に追加して再ソート
        month_options.append(curr.replace(day=1))
        month_options.sort(reverse=True)
        month_labels = [dt.strftime('%Y年%m月') for dt in month_options]
        current_idx = month_options.index(curr.replace(day=1))

    # ナビゲーションUIのスタイル調整 (スマホでも横並びを強制、サイズ最小化)
    st.markdown("""
        <style>
            /* st.columns の親要素に対して横並びを強制 */
            div[data-testid="stHorizontalBlock"] {
                display: flex !important;
                flex-direction: row !important;
                flex-wrap: nowrap !important;
                align-items: center !important;
                justify-content: center !important;
                gap: 2px !important;
            }
            /* 前月・翌月ボタンのカラム */
            div[data-testid="stHorizontalBlock"] > div:nth-child(1),
            div[data-testid="stHorizontalBlock"] > div:nth-child(3) {
                flex: 1 1 0% !important;
                width: auto !important;
                min-width: 0 !important;
            }
            /* 当月ポップオーバーのカラム */
            div[data-testid="stHorizontalBlock"] > div:nth-child(2) {
                flex: 0 0 auto !important;
                width: 130px !important; /* スマホで1行に収まるように幅を少し戻す */
                min-width: 130px !important;
            }
            /* ポップオーバーボタンの余白を削る */
            div[data-testid="stPopover"] > button {
                padding-left: 2px !important;
                padding-right: 2px !important;
            }
            /* ポップオーバーボタンのテキストサイズと折り返し禁止 */
            div[data-testid="stPopover"] > button p {
                font-size: 0.95rem !important;
                white-space: nowrap !important;
                margin: 0 !important;
            }
        </style>
    """, unsafe_allow_html=True)

    # ナビゲーションUI (3カラム)
    current_user = st.session_state.get("username", "")
    current_menu = st.session_state.get("menu_selection", "カレンダー")

    # データが存在するユニークな月リストを取得
    available_months = get_transaction_range(current_user)
    
    # 前後の月をリストから検索 (データがある月へジャンプ)
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
                    <a href="/?date={prev_date_str}&user={current_user}&menu={current_menu}" target="_self" style='text-decoration: none; color: #007bff; font-weight: bold;'>◀ 前月</a>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='text-align: right; color: #ccc; font-size: 0.9rem;'>◀ 前月</div>", unsafe_allow_html=True)
            
    with col2:
        # ポップオーバーで月選択ボタン風のUIを作成
        with st.popover(curr.strftime('%Y年%m月 ▼'), use_container_width=True):
            # 新しい順に月をリストアップ
            nav_months = sorted(available_months, reverse=True)
            if not nav_months:
                nav_months = [curr]
                
            list_html = "<div id='month-scroll-container' style='max-height: 250px; overflow-y: auto; text-align: center; border-radius: 5px;'>"
            
            for m in nav_months:
                m_str = m.strftime('%Y-%m-01')
                m_label = m.strftime('%Y年%m月')
                is_current = (m.year == curr.year and m.month == curr.month)
                
                bg_color = "#e6f2ff" if is_current else "transparent"
                font_weight = "bold" if is_current else "normal"
                color = "#0056b3" if is_current else "#333"
                id_attr = "id='current-month-link'" if is_current else ""
                
                # aタグによる画面遷移（クエリパラメータ更新）
                link = f"<a {id_attr} href='/?date={m_str}&user={current_user}&menu={current_menu}' target='_self' style='display: block; padding: 10px; margin: 2px 0; border-radius: 4px; background-color: {bg_color}; color: {color}; text-decoration: none; font-weight: {font_weight}; font-size: 1.1rem; transition: background 0.2s;'>{m_label}</a>"
                list_html += link
                
            list_html += "</div>"
            
            # JavaScriptで、開いた瞬間に current-month-link の位置まで自動スクロールする
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
            
    with col3:
        if has_next:
            st.markdown(f"""
                <div style='display: flex; align-items: center; justify-content: flex-start; font-size: 0.9rem; white-space: nowrap;'>
                    <a href="/?date={next_date_str}&user={current_user}&menu={current_menu}" target="_self" style='text-decoration: none; color: #007bff; font-weight: bold;'>翌月 ▶</a>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='text-align: left; color: #ccc; font-size: 0.9rem;'>翌月 ▶</div>", unsafe_allow_html=True)

    # データの読み込み
    with st.spinner("データを読み込み中..."):
        df = load_transactions_data(curr)
    
    # 合計金額の算出 (消費税（内税）は二重計上防止のため除外)
    monthly_total = 0
    if not df.empty and "amount" in df.columns:
        agg_df = df[df["category"] != "消費税（内税）"] if "category" in df.columns else df
        monthly_total = agg_df['amount'].sum()

    # 月間合計の表示
    st.markdown(f"<p style='text-align: center; font-size: 1.2rem; font-weight: bold; margin-bottom: 10px; color: black;'>月間合計支出: <span style='color: red;'>￥{int(monthly_total):,}</span></p>", unsafe_allow_html=True)
    st.markdown("---")
    
    return df

# ---------- レシート解析機能 ----------
def parse_receipt_with_gemini(image_file, additional_instruction=""):
    try:
        img = Image.open(image_file)
        img = ImageOps.exif_transpose(img) # 向きを自動補正
        # リサイズ（短辺・長辺ともに適切に圧縮。最大800px程度にしてAPIを高速化）
        img.thumbnail((800, 800))
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=85)
        img_byte_arr = img_byte_arr.getvalue()
        
        client = st.session_state.get('genai_client')
        if not client:
            return {"error": "APIキーが設定されていません。"}
            
        prompt = f"""
以下の画像（レシートまたは領収書）から必要な情報を抽出し、明細行ごとにJSON形式で出力してください。
{additional_instruction}

抽出項目（各明細に対して）:
1. "store_name" : 店舗名（文字列、不明な場合は ""）
2. "date" : 日付（YYYY-MM-DD形式、不明な場合は ""）
3. "item_name" : 商品名または内容（文字列）
4. "amount" : 金額（数値のみ、カンマなし）
5. "major_category" : 大分類
6. "minor_category" : 小分類

【最優先ルール】:
画像内に病院名などの医療機関の名前、あるいは「診療明細」「領収証（医療機関）」といった文字が含まれている場合、
すべての大分類は強制的に "(13) 医療" とし、小分類は内容から「🏥病院診療」「💊薬処方」「💉検査健診」のいずれかを推論して設定してください。これ以外の医療系の小分類は生成しないでください。

【金額抽出の厳格ルール】:
レシートに記載されている各商品の金額（amount）は、内税・外税に関わらず、加工（税抜き計算など）せずに「レシートに記載された数値のまま」を抽出してください。

【消費税の抽出ルール】:
レシート内に「消費税（8%や10%など）」が明細や項目として記載されている場合、その行を1つの明細として抽出してください。
その際、レシート内に「内税」という言葉が含まれている場合は、大分類を "消費税（内税）" 、小分類を "内税 + 読み取った税率" （例: "内税8%", "内税10%"）と設定してください。
「内税」という言葉が含まれていない場合は、大分類を "消費税（外税）" 、小分類を "外税 + 読み取った税率" （例: "外税8%", "外税10%"）と設定してください。
税率が不明な場合は、小分類を "内税？％" または "外税？％" としてください。

【合計金額の整合性ルール】（重要）:
レシート内の「合計金額」と、抽出した明細の関係は以下の通りである必要があります：
1. 「（消費税(内税)以外のすべての明細の金額） + （消費税(外税)の金額）」の合計が、レシートの「合計金額」と完全に一致すること。
2. 「消費税（内税）」はすでに商品単価に含まれているため、レシート合計金額の計算（検証）においては「無視」してください。
AIが辻褄を合わせるために勝手に商品金額を調整（減額・増額）することは絶対に禁止します。
行・割引や値引（マイナス金額で抽出）・消費税・小計などのいずれかを読み飛ばしているか誤読している可能性があります。読み飛ばしがないよう、すべての金額要素を漏れなく抽出してください。

【合計金額の明示】:
抽出した明細の最後に、必ず「レシートの最終合計金額」を示す特別な明細を1つ追加してください。
その際、"major_category" は "合計"、"minor_category" は "総合計"、"item_name" は "合計金額" とし、"amount" にはレシートに印字された最終的な支払い合計額を設定してください。

【割引・ギフト等の特別ルール】:
レシート内に「Gift」や「ギフト」などのキーワードと金額が記載されている場合、それらは割引・ポイント利用として扱います。大分類は "割引・ポイント利用" とし、小分類は "ギフト" や "割引" 等に設定し、金額（amount）は必ず「マイナスの数値」（例: -500）として抽出してください。

【支払い方法・お釣りの除外ルール】（最重要）:
レシートの下部に記載される「現金」「お預り」「お釣り」「クレジットカード」「カード払い」「PayPay」などの【支払い方法・お釣り等】に関する行や金額は、購入した商品ではないため、明細として*絶対に*抽出しないでください。抽出すると金額合算エラーの原因となります。

それ以外の場合は、以下のカテゴリ体系に厳密に従って、明細ごとに適切に分類してください。
{get_categories_prompt_text()}

JSONの出力形式は以下を厳守してください。マークダウンの ```json などは含めるず、純粋なJSON文字列（オブジェクトの配列）のみを返してください。
[
  {{
    "store_name": "店舗名",
    "date": "YYYY-MM-DD",
    "item_name": "商品名",
    "amount": 1000,
    "major_category": "大分類",
    "minor_category": "小分類"
  }}
]
"""
        # シンプルな解析ロジック: モデルを gemini-2.5-flash に固定して単発実行
        try:
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
            
            # 配列でない場合は配列にする
            if isinstance(result, dict):
                result = [result]
                
            return result
            
        except Exception as e:
            return {"error": f"レシートの読み取りに失敗しました。詳細: {str(e)}"}
        
    except Exception as e:
        return {"error": str(e)}

def verify_receipt_checksum(results):
    """
    レシートの読み取り結果（JSON）について、計算が合うか検証する。
    $$最終合計額 = （各明細の合計） + （消費税 / 外税） ー （利用ポイント）
    一致するかどうか(bool)と、合計行を除外したクリーンな結果(list)を返す。
    """
    if not isinstance(results, list) or len(results) == 0:
        return False, results
        
    reported_total = None
    calculated_total = 0
    clean_results = []
    
    for item in results:
        # 合計行の抽出
        if item.get("major_category") == "合計" and item.get("minor_category") == "総合計":
            reported_total = safe_money_int_cast(item.get("amount", 0))
            continue
            
        clean_results.append(item)
        amt = safe_money_int_cast(item.get("amount", 0))
        cat = item.get("major_category", "")
        # 内税は単価に含まれるため計算から除外
        if "内税" in cat or cat == "消費税（内税）":
            continue
            
        # ポイント利用などは絶対値を引き算
        if cat == "割引・ポイント利用":
            calculated_total -= abs(amt)
        else:
            calculated_total += amt

    # 合計行が抽出されていない場合はチェック不能なのでとりあえずFalse
    if reported_total is None:
        return False, clean_results

    is_valid = (reported_total == calculated_total)
    if not is_valid:
        print(f"Checksum mismatch: Reported={reported_total}, Calculated={calculated_total}")
        for item in clean_results:
            print(f"  - {item.get('item_name')}: {item.get('amount')} ({item.get('major_category')})")

    return is_valid, clean_results

def categorize_items_with_ai(items, store_name):
    """商品名リストと店舗名から、Gemini APIを使用してカテゴリを自動判別する"""
    client = st.session_state.get('genai_client')
    if not client:
        return [{"major_category": "その他", "minor_category": "📁未分類"} for _ in items]
        
    prompt = f"""
以下の店舗で購入した商品のリストについて、それぞれの大分類と小分類を判定してJSONで返してください。

店舗名: {store_name}

【カテゴリシステム: 大分類と小分類のリスト】
{get_categories_prompt_text()}

入力商品リスト:
{json.dumps(items, ensure_ascii=False)}

出力形式 (JSON配列のみ):
[
  {{"item_name": "商品名", "major_category": "大分類", "minor_category": "小分類"}},
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
        # JSON部分の抽出
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0].strip()
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0].strip()
            
        return json.loads(response_text)
    except Exception:
        # エラー時は「その他」で返す
        return [{"major_category": "その他", "minor_category": "📁未分類"} for _ in items]

def display_categories_as_html(target_df):
    import streamlit as st
    if target_df.empty:
        return
        
    cat_col = "category"
    if cat_col not in target_df.columns:
        st.warning("カテゴリ情報がありません。")
        return
        
    cat_grouped = target_df.groupby(cat_col, as_index=False)["amount"].sum()
    cat_grouped = cat_grouped.sort_values(by="amount", ascending=False)
    
    for _, cat_row in cat_grouped.iterrows():
        cat = cat_row[cat_col]
        cat_amt_str = f"￥{int(cat_row['amount']):,}"
        
        html_str = f'<details style="margin: 4px 0;">'
        html_str += f'<summary style="background-color: #ffffff; padding: 6px 10px; margin: 0; border: 1px solid #e0e0e0; border-radius: 4px; font-size: 0.95rem; font-weight: bold; list-style: none; cursor: pointer;">'
        html_str += f'📁 {cat}：{cat_amt_str}</summary>'
        html_str += f'<div style="padding-left: 15px; margin-top: 4px;">'
        
        sub_df = target_df[target_df[cat_col] == cat].copy()
        sub_col = None
        for col_name in ["subcategory", "sub_category", "小分類"]:
            if col_name in sub_df.columns:
                sub_col = col_name
                break
                
        if sub_col:
            sub_grouped = sub_df.groupby(sub_col, as_index=False)["amount"].sum()
            sub_grouped = sub_grouped.sort_values(by="amount", ascending=False)
            
            for _, sub_row in sub_grouped.iterrows():
                sub_name = sub_row[sub_col]
                sub_amt_str = f"￥{int(sub_row['amount']):,}"
                
                html_str += f'<details style="margin: 2px 0;">'
                html_str += f'<summary style="background-color: #f9fafb; padding: 3px 8px; margin: 0; border-left: 3px solid #007bff; font-size: 0.9rem; line-height: 1.2; list-style: none; cursor: pointer;">'
                html_str += f'L {sub_name}：{sub_amt_str}</summary>'
                html_str += f'<div style="padding-left: 10px; margin-top: 2px;">'
                
                item_df = sub_df[sub_df[sub_col] == sub_name].copy()
                item_col = "item_name" if "item_name" in item_df.columns else "item" if "item" in item_df.columns else None
                
                if item_col:
                    item_grouped = item_df.groupby(item_col, as_index=False)["amount"].sum()
                    item_grouped = item_grouped.sort_values(by="amount", ascending=False)
                    for _, i_row in item_grouped.iterrows():
                        i_name = i_row[item_col]
                        i_amt = f"￥{int(i_row['amount']):,}"
                        html_str += f'<div style="padding-left: 10px; font-size: 0.85rem; line-height: 1.2; margin: 2px 0; color: #555;">└ {i_name}：{i_amt}</div>'
                else:
                    for _, i_row in item_df.iterrows():
                        i_amt = f"￥{int(i_row['amount']):,}"
                        html_str += f'<div style="padding-left: 10px; font-size: 0.85rem; line-height: 1.2; margin: 2px 0; color: #555;">└ {i_amt}</div>'
                
                html_str += "</div></details>"
        else:
            item_cols = [c for c in ["item_name", "item", "amount"] if c in sub_df.columns]
            item_df = sub_df[item_cols].copy()
            item_col = "item_name" if "item_name" in item_df.columns else "item" if "item" in item_df.columns else None
            
            if item_col:
                item_grouped = item_df.groupby(item_col, as_index=False)["amount"].sum()
                item_grouped = item_grouped.sort_values(by="amount", ascending=False)
                for _, i_row in item_grouped.iterrows():
                    i_name = i_row[item_col]
                    i_amt = f"￥{int(i_row['amount']):,}"
                    html_str += f'<div style="padding-left: 10px; font-size: 0.85rem; line-height: 1.2; margin: 2px 0; color: #555;">└ {i_name}：{i_amt}</div>'
            else:
                for _, i_row in item_df.iterrows():
                    i_amt = f"￥{int(i_row['amount']):,}"
                    html_str += f'<div style="padding-left: 10px; font-size: 0.85rem; line-height: 1.2; margin: 2px 0; color: #555;">└ {i_amt}</div>'
                    
        html_str += "</div></details>"
        st.markdown(html_str, unsafe_allow_html=True)

def render_transaction_breakdown(df, key_prefix):
    """
    大分類別、店舗別、小分類別の2段階アコーディオンを表示する共通関数
    """
    if df.empty:
        st.info("データがありません。")
        return

    # 集計用のデータ（内税を除外して合計に反映させないようにする）
    df_agg = df.copy()
    if "category" in df_agg.columns:
        df_agg.loc[df_agg["category"] == "消費税（内税）", "amount"] = 0

    # 表示パターンの選択（小分類別を削除）
    view_pattern = st.radio("表示パターン", ["店舗別", "大分類別", "支払い方法別"], horizontal=True, key=f"{key_prefix}_view_pattern")
    
    if view_pattern == "店舗別":
        store_col = "store_name" if "store_name" in df.columns else "store" if "store" in df.columns else None
        if store_col:
            store_grouped = df_agg.groupby(store_col, as_index=False)["amount"].sum()
            store_grouped = store_grouped.sort_values(by="amount", ascending=False)
            
            for _, row in store_grouped.iterrows():
                store = row[store_col]
                total_amt_str = f"￥{int(row['amount']):,}"
                
                with st.expander(f"{store}：{total_amt_str}"):
                    store_df_agg = df_agg[df_agg[store_col] == store].copy()
                    store_df_disp = df[df[store_col] == store].copy()
                    
                    if key_prefix == "calendar":
                        # 2段階目：支払い方法アコーディオン
                        if "payment_method" not in store_df_agg.columns:
                            store_df_agg["payment_method"] = "未設定"
                            store_df_disp["payment_method"] = "未設定"
                            
                        pay_grouped = store_df_agg.groupby("payment_method", as_index=False)["amount"].sum()
                        pay_grouped = pay_grouped.sort_values(by="amount", ascending=False)
                        
                        for _, pay_row in pay_grouped.iterrows():
                            payment = pay_row["payment_method"]
                            pay_amt_str = f"￥{int(pay_row['amount']):,}"
                            
                            with st.expander(f"  └ 💳 {payment}：{pay_amt_str}"):
                                pay_df = store_df_disp[store_df_disp["payment_method"] == payment].copy()
                                
                                # 3段階目以降（カスタムHTML化）
                                display_categories_as_html(pay_df)
                    else:
                        # ダッシュボード等でも Store -> Major -> Minor -> Item を深い階層で表示する
                        display_categories_as_html(store_df_disp)
        else:
            st.info("店舗情報がありません。")

    elif view_pattern == "大分類別":
        if "category" in df.columns:
            # 大分類の集計は内税を含めて表示する
            grouped_df = df.groupby("category", as_index=False)["amount"].sum()
            grouped_df = grouped_df.sort_values(by="amount", ascending=False)
            
            for _, row in grouped_df.iterrows():
                cat = row['category']
                total_amt_str = f"￥{int(row['amount']):,}"
                
                with st.expander(f"{cat}：{total_amt_str}"):
                    cat_df = df[df["category"] == cat].copy()
                    sub_col = None
                    for col_name in ["subcategory", "sub_category", "小分類"]:
                        if col_name in cat_df.columns:
                            sub_col = col_name
                            break
                    
                    if sub_col:
                        # 2段階目：小分類アコーディオン（大分類 > 小分類 > 商品名）
                        # 小分類の集計は内税を含める
                        sub_grouped = cat_df.groupby(sub_col, as_index=False)["amount"].sum()
                        sub_grouped = sub_grouped.sort_values(by="amount", ascending=False)
                        
                        for _, sub_row in sub_grouped.iterrows():
                            sub_name = sub_row[sub_col]
                            sub_amt_str = f"￥{int(sub_row['amount']):,}"
                            
                            with st.expander(f"  └ {sub_name}：{sub_amt_str}"):
                                item_df = cat_df[cat_df[sub_col] == sub_name].copy()
                                item_col = "item_name" if "item_name" in item_df.columns else "item" if "item" in item_df.columns else None
                                
                                if item_col:
                                    item_grouped = item_df.groupby(item_col, as_index=False)["amount"].sum()
                                    item_grouped = item_grouped.sort_values(by="amount", ascending=False)
                                    item_grouped["amount"] = item_grouped["amount"].apply(lambda x: f"￥{int(x):,}")
                                    item_grouped.columns = ["商品名", "金額"]
                                    st.dataframe(item_grouped, use_container_width=True, hide_index=True)
                                else:
                                    detail_df = item_df[["date", "amount"]].copy() if "date" in item_df.columns else item_df[["amount"]].copy()
                                    detail_df = detail_df.sort_values(by="amount", ascending=False)
                                    if "date" in detail_df.columns:
                                        detail_df["date"] = detail_df["date"].dt.strftime('%m/%d')
                                    detail_df["amount"] = detail_df["amount"].apply(lambda x: f"￥{int(x):,}")
                                    st.dataframe(detail_df, use_container_width=True, hide_index=True)
                    else:
                        display_df = cat_df.copy()
                        cols_to_keep = [c for c in ["date", "store_name", "store", "item_name", "item", "amount"] if c in display_df.columns]
                        display_df = display_df[cols_to_keep]
                        if "amount" in display_df.columns:
                            display_df = display_df.sort_values(by="amount", ascending=False)
                        if "date" in display_df.columns:
                            display_df["date"] = display_df["date"].dt.strftime('%m/%d')
                        display_df["amount"] = display_df["amount"].apply(lambda x: f"￥{int(x):,}")
                        st.dataframe(display_df, use_container_width=True, hide_index=True)
        else:
            st.warning("カテゴリ情報がありません。")

    elif view_pattern == "支払い方法別":
        pm_col = "payment_method"
        if pm_col in df.columns:
            # fillna to avoid dropping NaNs in grouping
            df_pm_agg = df_agg.copy()
            df_pm_disp = df.copy()
            df_pm_agg[pm_col] = df_pm_agg[pm_col].fillna("未設定").replace(r'^\s*$', "未設定", regex=True)
            df_pm_disp[pm_col] = df_pm_disp[pm_col].fillna("未設定").replace(r'^\s*$', "未設定", regex=True)
            
            pm_grouped = df_pm_agg.groupby(pm_col, as_index=False)["amount"].sum()
            pm_grouped = pm_grouped[pm_grouped["amount"] > 0]
            pm_grouped = pm_grouped.sort_values(by="amount", ascending=False)
            
            store_col = "store_name" if "store_name" in df.columns else "store" if "store" in df.columns else None
            
            for _, row in pm_grouped.iterrows():
                pm = row[pm_col]
                total_amt_str = f"￥{int(row['amount']):,}"
                
                with st.expander(f"{pm}：{total_amt_str}"):
                    pm_filtered_agg = df_pm_agg[df_pm_agg[pm_col] == pm].copy()
                    pm_filtered_disp = df_pm_disp[df_pm_disp[pm_col] == pm].copy()
                    
                    if store_col:
                        # 2段階目：店舗
                        store_grouped = pm_filtered_agg.groupby(store_col, as_index=False)["amount"].sum()
                        store_grouped = store_grouped.sort_values(by="amount", ascending=False)
                        
                        for _, s_row in store_grouped.iterrows():
                            store_name = s_row[store_col]
                            s_amt_str = f"￥{int(s_row['amount']):,}"
                            
                            with st.expander(f"  └ {store_name}：{s_amt_str}"):
                                store_filtered_df = pm_filtered_disp[pm_filtered_disp[store_col] == store_name].copy()
                                
                                # 3段階目以降（大分類 -> 小分類 -> 商品）
                                display_categories_as_html(store_filtered_df)
                    else:
                        st.info("店舗情報がありません。")
        else:
            st.warning("支払い方法データがありません。")

# ---------- データダウンロード機能 ----------
def prepare_download_data(username):
    """
    全期間のトランザクション、プロフィール、カテゴリを取得し、
    ダウンロード用に「年」「月」列を追加したデータを準備する
    """
    try:
        # 1. 全トランザクション取得
        sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
        values = safe_gspread_call(sheet.get_all_values)
        if not values or len(values) < 2:
            df_transactions = pd.DataFrame()
        else:
            headers = [h.strip() if h.strip() else f"empty_{i}" for i, h in enumerate(values[0])]
            records_df = pd.DataFrame(values[1:])
            records_df.columns = headers[:records_df.shape[1]]
            records = records_df.to_dict('records')
            df_transactions = get_clean_df(records, username)
            
            if not df_transactions.empty:
                # 「年」「月」列を追加
                df_transactions["年"] = df_transactions["date"].dt.year
                df_transactions["月"] = df_transactions["date"].dt.month
                # 表示順序やフォーマットの調整
                df_transactions = df_transactions.sort_values("date", ascending=False)
        
        # 2. プロフィール取得
        profile = get_user_master_data(username)
        df_profile = pd.DataFrame([profile]) if profile else pd.DataFrame()
        
        # 3. カテゴリマスタ取得 (親子関係)
        cat_list = []
        for major, minors in get_categories().items():
            for minor in minors:
                cat_list.append({"大分類": major, "小分類": minor})
        df_categories = pd.DataFrame(cat_list)
        
        return df_transactions, df_profile, df_categories
    except Exception as e:
        st.error(f"データ準備エラー: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def generate_excel_download(username):
    """複数シートを持つExcelファイルをバイナリ形式で生成する"""
    df_tx, df_prof, df_cat = prepare_download_data(username)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not df_tx.empty:
            df_tx.to_excel(writer, index=False, sheet_name='実績データ')
        if not df_prof.empty:
            df_prof.to_excel(writer, index=False, sheet_name='ユーザープロフィール')
        if not df_cat.empty:
            df_cat.to_excel(writer, index=False, sheet_name='カテゴリマスタ')
            
        # xlsxwriterのワークブックオブジェクトを取得してフォーマット調整なども可能
        # 今回はシンプルに書き出しのみ
    
    return output.getvalue()

def generate_csv_download(username):
    """CSV (UTF-8-sig) を生成する"""
    df_tx, _, _ = prepare_download_data(username)
    if df_tx.empty:
        return None
    
    return df_tx.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

# ---------- 音声機能関連のユーティリティ ----------
import re

def render_speech_synthesis_button(text, key):
    """テキストを読み上げるスピーカーボタンを表示する"""
    if not text:
        return
    
    # 音声読み上げ用にMarkdown記号をクリーンアップ
    # '# ' や '### 'のような見出し記号を削除
    cleaned = re.sub(r'#+\s*', '', text)
    # '**' や '*' のような強調記号を削除
    cleaned = re.sub(r'\*{1,3}', '', cleaned)
    
    # JavaScriptによる読み上げロジック用のエスケープと改行処理
    clean_text = cleaned.replace("'", "\\'").replace("\n", " ")
    
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
    " onclick="speak_{key}()" title="読み上げる">🔊</button>
    
    <script>
    function speak_{key}() {{
        const btn = document.getElementById('btn-{key}');
        
        // 再生中の場合は停止
        if (window.speechSynthesis.speaking) {{
            window.speechSynthesis.cancel();
            btn.innerText = '🔊';
            return;
        }}
        
        // iOS Safari対策: 一度空のcancelを呼ぶことで音声エンジンを強制的にアクティブにする
        window.speechSynthesis.cancel();
        
        // 少し遅延を入れてから発話させる（iOS対策）
        setTimeout(() => {{
            const uttr = new SpeechSynthesisUtterance('{clean_text}');
            uttr.lang = 'ja-JP';
            uttr.rate = 1.1;
            
            uttr.onstart = () => {{ btn.innerText = '⏹'; btn.style.color = '#dc3545'; }};
            uttr.onend = () => {{ btn.innerText = '🔊'; btn.style.color = '#555'; }};
            uttr.onerror = (e) => {{
                console.error("SpeechSynthesisError:", e);
                btn.innerText = '🔊'; 
                btn.style.color = '#555'; 
            }};
            
            window.speechSynthesis.speak(uttr);
        }}, 50);
    }}
    </script>
    """
    import streamlit.components.v1 as components
    components.html(html_code, height=45)

def render_voice_input_button(key_prefix):
    """チャット入力欄の右側にマイクボタンを動的に配置し、音声入力を可能にする"""
    html_code = f"""
    <script>
    (function() {{
        const parentDoc = window.parent.document;
        
        function injectMicButton() {{
            const chatInputArea = parentDoc.querySelector('[data-testid="stChatInput"]');
            if (!chatInputArea) return false;
            
            if (parentDoc.getElementById('injected-mic-btn-{key_prefix}')) return true;
            
            const micBtn = parentDoc.createElement('button');
            micBtn.id = 'injected-mic-btn-{key_prefix}';
            micBtn.innerHTML = '🎤';
            micBtn.title = '音声で入力';
            micBtn.style.cssText = `
                background: transparent;
                border: none;
                font-size: 22px;
                cursor: pointer;
                padding: 0 5px;
                margin-right: 5px;
                display: flex;
                align-items: center;
                justify-content: center;
                transition: transform 0.2s;
            `;
            
            micBtn.onmouseover = () => micBtn.style.transform = 'scale(1.2)';
            micBtn.onmouseout = () => micBtn.style.transform = 'scale(1)';
            
            micBtn.onclick = function() {{
                const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
                if (!SpeechRecognition) {{
                    alert('このブラウザは音声認識をサポートしていません。');
                    return;
                }}
                const recognition = new SpeechRecognition();
                recognition.lang = 'ja-JP';
                recognition.interimResults = false;
                recognition.maxAlternatives = 1;
                
                micBtn.innerHTML = '🔴';
                
                recognition.onresult = function(event) {{
                    const result = event.results[0][0].transcript;
                    const chatTextArea = parentDoc.querySelector('textarea[data-testid="stChatInputTextArea"]');
                    if (chatTextArea) {{
                        // ReactのVirtual DOMに値を正しく認識させるためのネイティブSetter呼び出し
                        const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
                        const currentVal = chatTextArea.value || "";
                        const newVal = currentVal ? currentVal + " " + result : result;
                        nativeInputValueSetter.call(chatTextArea, newVal);
                        
                        // ReactのonChangeイベントを発火
                        chatTextArea.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    }}
                    micBtn.innerHTML = '🎤';
                }};
                
                recognition.onerror = function(event) {{
                    console.error('Speech recognition error', event.error);
                    micBtn.innerHTML = '🎤';
                    if (event.error === 'not-allowed') {{
                        alert('マイクの使用が許可されていません。設定を確認するか、Androidの場合はHTTPSでアクセスしてください。');
                    }}
                }};
                
                recognition.onend = function() {{
                    micBtn.innerHTML = '🎤';
                }};
                
                recognition.start();
            }};
            
            // [data-testid="stChatInput"] の中にある送信ボタンの前にマイクボタンを挿入する
            const innerFlex = chatInputArea.children[0];
            if (innerFlex) {{
                innerFlex.insertBefore(micBtn, innerFlex.lastElementChild);
                return true;
            }}
            return false;
        }}
        
        // st.chat_input は後からレンダリングされる可能性があるため、出現するまで定期的にチェック
        let attempts = 0;
        const intervalId = setInterval(() => {{
            if (injectMicButton() || attempts > 40) {{
                clearInterval(intervalId);
            }}
            attempts++;
        }}, 500);
    }})();
    </script>
    """
    
    import streamlit.components.v1 as components
    components.html(html_code, height=0)

# ---------- ページUIの実装 ----------
def show_dashboard():
    # ヘッダーを表示するためのプレースホルダー（コンテナ）を先に準備
    header_placeholder = st.empty()

    # 共通ナビゲーションの適用
    # モードを明示的に指定（月次）
    df = render_month_navigation()

    # 月の切り替え操作が行われた「後」の最新の状態でヘッダーを更新する
    header_placeholder.markdown("#### 📊 ダッシュボード (月別集計)")

    if df.empty:
        st.info("※今月のデータはまだありません。")
        return

    # 集計用のデータ（内税を除外）
    df_agg = df[df["category"] != "消費税（内税）"] if "category" in df.columns else df

    # 分析軸とグラフ種類の選択UI
    col_a, col_b = st.columns(2)
    with col_a:
        analysis_axis = st.selectbox(
            "分析軸を選択", 
            ["大分類別", "小分類別", "店舗別", "消費税", "支払い方法"], 
            index=0, 
            key="monthly_analysis_axis"
        )
    with col_b:
        graph_type = st.selectbox(
            "グラフを選択",
            ["円グラフ", "棒グラフ"],
            index=0,
            key="monthly_graph_type"
        )

    # 選択に応じて集計対象の列を決定
    group_col = None
    title_label = ""
    
    if analysis_axis == "大分類別":
        group_col = "category"
        title_label = "大分類別金額シェア"
    elif analysis_axis == "小分類別":
        for col in ["subcategory", "sub_category", "小分類"]:
            if col in df.columns:
                group_col = col
                break
        title_label = "小分類別金額シェア"
    elif analysis_axis == "店舗別":
        for col in ["store_name", "store", "店舗"]:
            if col in df.columns:
                group_col = col
                break
        title_label = "店舗別金額シェア"
    elif analysis_axis == "消費税":
        group_col = "tax_group"
        title_label = "消費税率別金額シェア"
        # 税率マッピング関数の定義
        def map_tax(subcategory):
            sub = str(subcategory)
            if "10%" in sub: return "外税10%＋内税10%"
            elif "8%" in sub: return "外税8%+内税8%"
            else: return "外税？%+内税？%"
        
        # サブカテゴリ列を探す
        sub_col = "subcategory"
        for col in ["subcategory", "sub_category", "小分類"]:
            if col in df.columns:
                sub_col = col
                break
        
        # 【修正】カテゴリを消費税（外税・内税）に限定してフィルタリング
        df = df[df["category"].isin(["消費税（外税）", "消費税（内税）"])]
        if sub_col in df.columns:
            df["tax_group"] = df[sub_col].apply(map_tax)
        df_agg = df # 消費税軸の場合はこのフィルタ済みデータを全表示
    elif analysis_axis == "支払い方法":
        group_col = "payment_method"
        title_label = "支払い方法別金額シェア"

    if group_col and group_col in df.columns and "amount" in df.columns:
        if graph_type == "円グラフ":
            # 選択された軸が店舗別なら内税を除外し、大分類・小分類なら含めて表示する（二重計上防止だが内訳は正しく出す）
            df_for_chart = df_agg if analysis_axis == "店舗別" else df
            grouped_df = df_for_chart.groupby(group_col, as_index=False)["amount"].sum()
            # 金額が0以下のデータ（マイナス値や0円）を除外
            grouped_df = grouped_df[grouped_df["amount"] > 0]
            grouped_df = grouped_df.sort_values(by="amount", ascending=False)
            
            # 表示順序の固定
            tax_order = ["外税8%+内税8%", "外税10%＋内税10%", "外税？%+内税？%"]
            category_orders = {group_col: tax_order} if analysis_axis == "消費税" else {group_col: grouped_df[group_col].tolist()}

            fig = px.pie(
                grouped_df, 
                values='amount', 
                names=group_col, 
                hole=0.4, 
                title=title_label,
                category_orders=category_orders,
                color=group_col,
                color_discrete_map=CATEGORY_COLOR_MAP
            )
            fig.update_traces(
                textposition='inside', 
                textinfo='percent+label'
            )
            st.plotly_chart(fig, use_container_width=True)
            
            total_amount = df_agg["amount"].sum() if not df_agg.empty else 0
            st.metric("当月総支出額", f"￥{int(total_amount):,}")
            
        else: # 棒グラフの場合
            selected_year = st.session_state['current_month'].year
            selected_month = st.session_state['current_month'].month
            
            # 指定された分析軸で日ごとのデータをグループ化 (店舗別なら内税除外、それ以外は含める)
            df_bar = df_agg.copy() if analysis_axis == "店舗別" else df.copy()
            df_bar['day'] = df_bar['date'].dt.day
            df_bar['day_label'] = df_bar['day'].apply(lambda x: f"{x}日")
            daily_grouped = df_bar.groupby(['day', 'day_label', group_col], as_index=False)["amount"].sum()
            
            # 指定の順番を保つため、全体の合計額順でカテゴリーをソートする
            if analysis_axis == "消費税":
                cat_sum = ["外税8%+内税8%", "外税10%＋内税10%", "外税？%+内税？%"]
            else:
                cat_sum = daily_grouped.groupby(group_col)["amount"].sum().sort_values(ascending=False).index.tolist()
            
            # 当月の日数を計算 (月末までカレンダー通り表示)
            _, last_day = calendar.monthrange(selected_year, selected_month)
            all_days = [f"{i}日" for i in range(1, last_day + 1)]
            fig = px.bar(
                daily_grouped,
                x='day_label',
                y='amount',
                color=group_col,
                title=f"{selected_year}年{selected_month}月 {title_label}日次推移 (積上げ棒グラフ)",
                labels={"amount": "金額", "day_label": "日", group_col: analysis_axis[:-1]},
                category_orders={"day_label": all_days, group_col: cat_sum},
                color_discrete_map=CATEGORY_COLOR_MAP
            )

            # --- トレンド線（カテゴリ別累積推移）の追加 ---
            # 各カテゴリの積上げ高さに合わせるため、下層のカテゴリから順に合算した累積シリーズを維持する
            cumulative_daily = pd.Series(0.0, index=all_days)
            for cat in reversed(cat_sum): # 下層から積上げるためreversed
                cat_data = daily_grouped[daily_grouped[group_col] == cat]
                if not cat_data.empty:
                    # all_daysに合わせて補完し、その日のカテゴリ合計を出す
                    cat_data_full = pd.DataFrame({'day_label': all_days}).merge(cat_data, on='day_label', how='left').fillna({'amount': 0.0})
                    cumulative_daily += cat_data_full['amount']
                    
                    line_color = CATEGORY_COLOR_MAP.get(cat, None)
                    fig.add_trace(go.Scatter(
                        x=all_days,
                        y=cumulative_daily,
                        mode='lines+markers',
                        name=f'{cat} (累積推移)',
                        line=dict(color=line_color, width=1),
                        marker=dict(size=4),
                        showlegend=False # 棒グラフの凡例があるため非表示
                    ))

            if analysis_axis == "消費税":
                fig.update_xaxes(categoryorder='array', categoryarray=cat_sum)
            # マイナスのデータを黄色にする (棒グラフのトレースのみ対象)
            for trace in fig.data:
                if trace.type == 'bar' and hasattr(trace, 'y') and trace.y is not None:
                    orig_color = trace.marker.color
                    trace.marker.color = ['#FFFF00' if v is not None and (isinstance(v, (int, float)) and v < 0) else orig_color for v in trace.y]

            fig.update_yaxes(zerolinewidth=2, zerolinecolor='black')
            st.plotly_chart(fig, use_container_width=True)
            
            # 当月の合計金額はそのまま表示（内税除外）
            current_month_total = df_agg['amount'].sum() if not df_agg.empty else 0
            st.metric("当月総支出額", f"￥{int(current_month_total):,}")

        st.markdown("---")
        
        st.markdown("##### カテゴリ別内訳 (当月)")
        render_transaction_breakdown(df, "dashboard")
    else:
        st.warning(f"分析に必要な列（{analysis_axis[:-1]}）がありません。")

def show_credit_card_dashboard():
    st.markdown("#### 💳 クレジットカード利用状況")
    
    render_month_navigation()
    target_date = st.session_state.current_month
    
    methods = get_payment_methods(st.session_state['username'])
    credit_cards = [m for m in methods if str(m.get("is_credit_card", "False")).lower() == "true"]
    
    if not credit_cards:
        st.info("クレジットカードが登録されていません。サイドバーの「マスター設定」＞「支払方法設定」から登録してください。")
        return
        
    st.markdown("---")
    
    # ---------------------------------------------------------
    # ① カード選択機能
    # ---------------------------------------------------------
    card_names = [cc.get("name", "名称未設定") for cc in credit_cards]
    selected_card_name = st.selectbox("確認するクレジットカードを選択してください", options=card_names)
    
    selected_cc = next((cc for cc in credit_cards if cc.get("name") == selected_card_name), credit_cards[0])
    
    closing_str = selected_cc.get("closing_date", "")
    pay_m_str = selected_cc.get("payment_month", "")
    pay_d_str = selected_cc.get("payment_date", "")
    limit_str = selected_cc.get("credit_limit", "")
    
    # クレジットカードの計算ヘルパー呼び出し
    periods = calculate_credit_card_periods(
        target_date, closing_str, pay_m_str, pay_d_str
    )
    
    if not periods:
        st.warning("締日や支払日の設定が完了していません。設定画面から設定してください。")
        st.markdown("---")
        return
        
    with st.spinner("データ集計中..."):
        df_all = load_transactions_data(target_date, mode="all")
        
        # 該当カードのデータを抽出
        if not df_all.empty:
            df_cc = df_all[df_all["payment_method"] == selected_card_name]
            # 集計時の二重計上を防ぐため、内税を除外
            if "category" in df_cc.columns:
                df_cc = df_cc[df_cc["category"] != "消費税（内税）"]
        else:
            df_cc = pd.DataFrame()
        
        # 各期間の集計とデータフレーム保持
        today = date.today()
        for p in periods:
            if not df_cc.empty:
                mask = (df_cc["date"].dt.date >= p["start"]) & (df_cc["date"].dt.date <= p["end"])
                p["df"] = df_cc[mask]
                p["total"] = p["df"]["amount"].sum()
            else:
                p["df"] = pd.DataFrame()
                p["total"] = 0
                
            # 状態（支払済 / 支払予定）の判定
            if p["pay_date"] <= today:
                p["status_text"] = "支払済"
                p["status_color"] = "gray"
            else:
                p["status_text"] = "支払予定"
                p["status_color"] = "red"  # 未払いは目立つように
            
        # ---------------------------------------------------------
        # ② 支払額サマリーの表示（3行の垂直レイアウト）
        # ---------------------------------------------------------
        st.markdown("#### 📅 利用状況サマリー")
        
        for p in periods:
            pr_start = p["start"].strftime('%Y/%m/%d')
            pr_end = p["end"].strftime('%Y/%m/%d')
            p_date = p["pay_date"].strftime('%Y/%m/%d')
            
            st.markdown(
                f"<div style='border-left: 4px solid {p['status_color']}; padding-left: 10px; margin-bottom: 15px;'>"
                f"<div style='font-size: 1.1em; font-weight: bold;'>{p['label']}: ¥{int(p['total']):,} "
                f"<span style='font-size: 0.8em; color: {p['status_color']}; margin-left:10px;'>({p_date} {p['status_text']})</span></div>"
                f"<div style='color: gray; font-size: 0.85em; margin-top: 4px;'>利用期間: {pr_start} 〜 {pr_end}</div>"
                f"</div>",
                unsafe_allow_html=True
            )

        # 限度額の表示（設定されている場合）
        if limit_str and str(limit_str).isdigit():
            limit = int(limit_str)
            total_unpaid = sum(p["total"] for p in periods if p.get("status_text") != "支払済")
            ratio = min(total_unpaid / limit, 1.0) if limit > 0 else 0.0
            st.progress(ratio)
            st.caption(f"限度額: ¥{limit:,} に対する現在の利用率: {int(ratio*100)}%")
            remaining = limit - total_unpaid
            st.caption(f"現在の利用額合計: ¥{int(total_unpaid):,}　｜　残額: ¥{int(remaining):,}")

        st.markdown("<hr style='margin: 1.5em 0;'>", unsafe_allow_html=True)
        
        # ---------------------------------------------------------
        # ③ 明細表示の切り替え (3つの期間から選択)
        # ---------------------------------------------------------
        view_options = [p["label"] for p in periods]
        view_mode = st.radio("表示する明細を選択", view_options, horizontal=True)
        
        # 選択された期間のデータを特定
        selected_period = next(p for p in periods if p["label"] == view_mode)
        target_df = selected_period["df"]
        period_total = selected_period["total"]
        st.markdown(f"##### {view_mode} 内訳　小計：¥{int(period_total):,}")

        # ---------------------------------------------------------
        # ④ 5階層のドリルダウン表示 (日付 > 店舗 > 大分類 > 小分類 > 商標名)
        # ---------------------------------------------------------
        if target_df.empty:
            st.info("この期間の利用明細はありません。")
        else:
            # 欠損値対策
            target_df = target_df.fillna({
                "store_name": "不明な店舗", 
                "category": "その他", 
                "subcategory": "未分類", 
                "item_name": "不明な商品", 
                "amount": 0
            })
            
            # 第1階層: 日付＋店舗名 (pandas dt.date と store_name でマルチグループ化)
            target_df["date_str"] = target_df["date"].dt.strftime('%Y/%m/%d')
            target_df["date_disp"] = target_df["date"].dt.strftime('%m/%d')
            for (date_str, store), store_df in target_df.groupby(["date_str", "store_name"]):
                date_disp = store_df.iloc[0]["date_disp"]
                store_total = store_df["amount"].sum()
                store_name_disp = store if store else "不明な店舗"
                with st.expander(f"**{date_disp} 🏪 {store_name_disp}** （¥{int(store_total):,}）"):
                    
                    # 第2階層: 大分類
                    for major, major_df in store_df.groupby("category"):
                        major_total = major_df["amount"].sum()
                        major_disp = major if major else "その他"
                        with st.expander(f"📂 {major_disp} （¥{int(major_total):,}）"):
                            
                            # 第3階層: 小分類
                            for minor, minor_df in major_df.groupby("subcategory"):
                                minor_total = minor_df["amount"].sum()
                                minor_disp = minor if minor else "未分類"
                                with st.expander(f"📁 {minor_disp} （¥{int(minor_total):,}）"):
                                    
                                    # 第4階層: 商標名（商品名）の一覧
                                    for _, row in minor_df.iterrows():
                                        item = row["item_name"]
                                        amt = row["amount"]
                                        st.markdown(
                                            f"<div style='margin-left: 10px; padding: 4px 0; border-bottom: 1px dashed #eee;'>"
                                            f"・ <b>{item}</b> "
                                            f"<span style='float:right;'>¥{int(amt):,}</span></div>", 
                                            unsafe_allow_html=True
                                        )
                            
    st.markdown("---")

@st.cache_data(ttl=60)
def get_used_categories():
    try:
        sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
        records = safe_gspread_call(sheet.get_all_records)
        used_majors = set()
        used_minors = {}
        if not records:
            return used_majors, used_minors
        for row in records:
            major = str(row.get("category", "")).strip()
            sub = ""
            for c in ["subcategory", "sub_category", "小分類"]:
                if c in row:
                    sub = str(row[c]).strip()
                    break
                    
            if major:
                used_majors.add(major)
                if major not in used_minors:
                    used_minors[major] = set()
                if sub:
                    used_minors[major].add(sub)
        return used_majors, used_minors
    except Exception as e:
        print(f"Error checking used categories: {e}")
        return set(), {}

def save_categories_to_sheet(cats_dict):
    import streamlit as st
    try:
        sheet = get_sheet(CATEGORY_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        safe_gspread_call(sheet.clear)
        
        new_data = [["major_category", "minor_category"]]
        for major, minors in cats_dict.items():
            if minors:
                for minor in minors:
                    new_data.append([major, minor])
            else:
                new_data.append([major, ""])
                
        safe_gspread_call(sheet.update, range_name="A1", values=new_data)
        get_categories.clear() # Cache clear
        return True
    except Exception as e:
        st.error(f"カテゴリ保存エラー: {e}")
        return False

def show_category_master():
    """カテゴリマスター設定画面 (オーナー専用)"""
    import streamlit as st
    st.markdown("#### 📂 カテゴリマスター （大分類・小分類設定）")
    st.info("家計簿全体で使用される「大分類」と「小分類」の設定を行います。この画面はオーナー専用です。\n\n※ 既にレシートデータで登録済みのカテゴリは**変更・削除できません**。どうしても変更が必要な場合は先にレシート修正画面から修正してください。\n※ 新しく追加したカテゴリグラフの色は自動で割り当てられます。")
    
    current_cats = get_categories()
    used_majors, used_minors = get_used_categories()
    
    col1, padding, col2 = st.columns([4, 1, 6])
    
    with col1:
        st.markdown("##### 📁 【大分類の一覧】")
        st.caption("※ 編集・確認する大分類を以下から選択してください。")
        major_cats = list(current_cats.keys())
        
        if 'selected_major_cat' not in st.session_state:
            st.session_state.selected_major_cat = major_cats[0] if major_cats else None
            
        selected_major = st.radio("大分類選択", options=major_cats, key="cb_major_cat", label_visibility="collapsed")
        st.session_state.selected_major_cat = selected_major
        
        st.markdown("---")
        with st.expander("➕ 新しい大分類を追加する", expanded=False):
            new_major_name = st.text_input("大分類名（例: 日用品費）", key="new_major_input")
            if st.button("大分類を追加", use_container_width=True):
                new_major = new_major_name.strip()
                if new_major:
                    if new_major in current_cats:
                        st.warning("その大分類は既に存在します。")
                    else:
                        current_cats[new_major] = []
                        if save_categories_to_sheet(current_cats):
                            st.success(f"「{new_major}」を追加しました！")
                            st.session_state.selected_major_cat = new_major
                            import time
                            time.sleep(1)
                            st.rerun()
                else:
                    st.warning("大分類名を入力してください。")
                    
        with st.expander("⚠️ 大分類名を変更・削除する", expanded=False):
            if selected_major:
                st.markdown(f"**対象:** 「{selected_major}」")
                changed_major_name = st.text_input("新しい名前", value=selected_major, key="change_major_input")
                
                if st.button("名前を変更する", type="primary", use_container_width=True):
                    new_m_name = changed_major_name.strip()
                    if new_m_name and new_m_name != selected_major:
                        if new_m_name in current_cats:
                            st.warning("その名前は既に使用されています。")
                        elif selected_major in used_majors:
                            st.error(f"「{selected_major}」は既にレシートデータで登録済みの為、名前の変更はできません。")
                        else:
                            new_cats = {}
                            for k, v in current_cats.items():
                                if k == selected_major:
                                    new_cats[new_m_name] = v
                                else:
                                    new_cats[k] = v
                            if save_categories_to_sheet(new_cats):
                                st.success("名前を変更しました！")
                                st.session_state.selected_major_cat = new_m_name
                                import time
                                time.sleep(1)
                                st.rerun()
                
                st.markdown("<br>", unsafe_allow_html=True)
                st.warning(f"「{selected_major}」を削除しますか？紐づく小分類もすべて消去されます。")
                if st.button(f"「{selected_major}」を完全に削除する", use_container_width=True):
                    if selected_major in used_majors:
                        st.error(f"「{selected_major}」は既にレシートデータで登録済みの為、削除はできません。")
                    else:
                        del current_cats[selected_major]
                        if save_categories_to_sheet(current_cats):
                            st.success(f"「{selected_major}」を削除しました！")
                            current_majors = list(current_cats.keys())
                            st.session_state.selected_major_cat = current_majors[0] if current_majors else None
                            import time
                            time.sleep(1)
                            st.rerun()
                        
    with col2:
        if selected_major:
            st.markdown(f"##### 📄 【小分類の編集】")
            st.write(f"**現在の対象大分類:** 📁 {selected_major}")
            st.caption("※ 下の表を直接クリックして小分類名を追加・修正・削除できます。最下部の `+` マークで行を追加できます。")
            
            minors = current_cats.get(selected_major, [])
            import pandas as pd
            if minors:
                df_minors = pd.DataFrame([{"小分類": m} for m in minors])
            else:
                df_minors = pd.DataFrame(columns=["小分類"])
            
            edited_minors_df = st.data_editor(
                df_minors,
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_{selected_major}",
                column_config={
                    "小分類": st.column_config.TextColumn("小分類", required=True)
                }
            )
            
            if st.button(f"「{selected_major}」の小分類の変更を保存", type="primary", use_container_width=True):
                old_minors = set(minors)
                new_minors = []
                for _, row in edited_minors_df.iterrows():
                    minor_val = str(row.get("小分類", "")).strip()
                    if minor_val and minor_val != "nan":
                        if minor_val not in new_minors:
                            new_minors.append(minor_val)
                
                # Identify deleted or renamed categories (which appear as deletions)
                removed_minors = old_minors - set(new_minors)
                used_sub_for_major = used_minors.get(selected_major, set())
                in_use_removed = removed_minors.intersection(used_sub_for_major)
                
                if in_use_removed:
                    err_msg = ", ".join(in_use_removed)
                    st.error(f"以下の小分類は既にレシートデータで登録済みのため、変更・削除できません: {err_msg}")
                else:
                    with st.spinner("保存中..."):
                        current_cats[selected_major] = new_minors
                        if save_categories_to_sheet(current_cats):
                            st.success(f"「{selected_major}」の小分類を保存しました！設定をアプリに反映するにはページをリロードしてください。")
                            import time
                            time.sleep(1)
                            st.rerun()
        else:
            st.info("大分類が登録されていません。左側から大分類を追加してください。")

def show_data_check():
    """データチェック機能（現在はプレースホルダー）"""
    st.markdown("#### ✅ データチェック")
    st.info("データチェック機能は現在準備中です。データの整合性や異常値の自動検知機能を将来的に追加する予定です。")
    if st.button("戻る", use_container_width=True):
        st.session_state.menu_selection = "ダッシュボード（月次集計）"
        st.rerun()

def show_fixed_cost_management():
    """固定費管理（支払管理）シートの作成とリンク表示"""
    import streamlit as st
    import gspread
    
    st.markdown("#### 📑 支払管理シート新規作成")
    st.info("ログインアカウント専用の「固定費管理シート」を作成します。")
    
    username = st.session_state.get("username", "")
    sheet_name = f"{username}_支払管理"
    
    client = get_gspread_client()
    if client is None:
        st.error("Google Drive / Sheets への接続に失敗しました。")
        return
        
    try:
        # スプレッドシートが存在するか確認 (サービスアカウントに共有されているか)
        try:
            def _open_fixed_sheet():
                return client.open(sheet_name)
            sheet = safe_gspread_call(_open_fixed_sheet)
            st.info(f"固定費管理シート（{sheet_name}）は既に作成済です。")
            
        except gspread.exceptions.SpreadsheetNotFound:
            st.warning(f"現在、あなた（{username}）専用の支払管理シートはアプリと連携されていません。")
            
            # Googleサービスアカウントの容量制限を回避するため、
            # ユーザーの権限で動作する Google Apps Script (GAS) Webアプリを呼び出す
            if st.button("固定費管理シート（支払管理）を作成する", type="primary"):
                with st.spinner("シートを作成中...（10秒程度かかる場合があります）"):
                    try:
                        import requests
                        # ユーザーがデプロイしたGASウェブアプリのURL
                        gas_url = "https://script.google.com/macros/s/AKfycbydYbEYLvidlXeWI8MTcyGx2dC_RUHLnAGR2aDgxWigcpAniHh0izNuaOU_wn7V5PQf/exec"
                        
                        # サービスアカウントのアドレスを安全に取得
                        service_email = ""
                        try:
                            if "gcp_service_account" in st.secrets:
                                service_email = st.secrets["gcp_service_account"].get("client_email", "")
                            else:
                                import json, os
                                if os.path.exists("credentials.json"):
                                    with open("credentials.json", "r", encoding="utf-8") as f:
                                        creds_data = json.load(f)
                                        service_email = creds_data.get("client_email", "")
                        except:
                            pass
                            
                        # GASを呼び出してコピーと共有を実行させる
                        res = requests.post(gas_url, data={
                            "action": "copy",
                            "newName": sheet_name,
                            "shareEmail": service_email
                        }, timeout=45)
                        
                        data = res.json()
                        if data.get("status") == "success":
                            # 作成が成功したら、少し待機してからリロードを促す（Google Driveの反映待ち）
                            st.success("🎉 固定費管理シートの作成と連携が完了しました！")
                            st.markdown(f"**🔗 [こちらのリンクから新しく作成されたスプレッドシートを開いて編集してください]({data.get('url')})**")
                            st.info("※次回のアクセス時からは、この画面の上部に直接リンクが表示されます。")
                            import time
                            time.sleep(2)
                            st.rerun()
                        else:
                            st.error(f"作成エラー: {data.get('message', '不明なエラー')}")
                    except Exception as e:
                        st.error(f"通信エラーが発生しました: {e}")

    except Exception as e:
        st.error(f"Google API 通信中にエラーが発生しました: {e}")

def show_fixed_cost_master_settings():
    """固定費マスター設定画面"""
    import streamlit as st
    import gspread
    import pandas as pd
    
    st.markdown("#### ⚙️ 固定費マスター設定")
    
    username = st.session_state.get("username", "")
    sheet_name = f"{username}_支払管理"
    
    client = get_gspread_client()
    if client is None:
        st.error("Google Drive / Sheets への接続に失敗しました。")
        return
        
    try:
        def _open_fixed_sheet():
            return client.open(sheet_name)
        sheet = safe_gspread_call(_open_fixed_sheet)
        
        try:
            worksheet = sheet.worksheet("固定費マスター")
            url = f"{sheet.url}#gid={worksheet.id}"
        except gspread.exceptions.WorksheetNotFound:
            # フォールバック: シートが存在しない場合はスプレッドシート全体のURLを使う
            url = sheet.url
            
        st.info("固定費管理シートの「固定費マスター」シートを開いて直接設定を行います。")
        if st.link_button("🌐 「固定費マスター」シートを開く", url=url, type="primary"):
            pass
            
        st.markdown("---")
        st.markdown("##### 📝 「固定費マスター」設定項目説明")
        
        data = [
            {"項目名": "科目１", "説明": "支払手段の大きな分類を設定します。", "入力例・補足": "口座引落、クレジットカード、現金振込 など"},
            {"項目名": "科目２", "説明": "具体的な支払元（銀行名やカード名）を設定します。", "入力例・補足": "りそな銀行、楽天カード、三井住友銀行 など"},
            {"項目名": "有限or無限", "説明": "支払に終わりがあるか（ローン等）継続するかを設定します。", "入力例・補足": "有限: ローン、分割払、契約期間あり\\n無限: 公共料金、サブスク、家賃 など"},
            {"項目名": "科目詳細", "説明": "具体的な支払内容の名称です。支払管理シートとの紐付けに使います。", "入力例・補足": "住宅ローン、ガス代、Netflix、学資保険 など"},
            {"項目名": "支払額", "説明": "1回あたりの支払金額を入力します。", "入力例・補足": "125567、10000（カンマや￥はなしで数値のみ推奨）"},
            {"項目名": "変動or固定", "説明": "金額が毎月一定か、月によって変わるかを選択します。", "入力例・補足": "固定: 家賃、保険料、定額サブスク\\n変動: 電気代、水道代、ガス代"},
            {"項目名": "支払月", "説明": "支払が発生するタイミング（頻度）を設定します。", "入力例・補足": "毎月: 毎月支払\\n数値+月: 「9月」など特定の月のみ（年払など）"},
            {"項目名": "最終月額", "説明": "ローンの最終回など、通常と金額が異なる場合の額を入力します。", "入力例・補足": "125706（端数調整が必要な場合のみ記入）"},
            {"項目名": "振込手数料", "説明": "支払時に別途手数料が発生する場合にその額を入力します。", "入力例・補足": "110、220、440 など"},
            {"項目名": "開始月", "説明": "その支払がいつから始まるかを設定します。", "入力例・補足": "2026年4月（yyyy年m月の形式）"},
            {"項目名": "完済月", "説明": "「有限」の場合のみ、支払が終わる月を設定します。", "入力例・補足": "2029年8月（無限の場合は空欄）"},
        ]
        df = pd.DataFrame(data)
        st.dataframe(df, hide_index=True, use_container_width=True)
        
    except gspread.exceptions.SpreadsheetNotFound:
        st.warning(f"現在、あなた（{username}）専用の固定費管理シートは作成されていません。")
        st.info("先に「支払管理シート新規作成」メニューからシートを作成してください。")
    except Exception as e:
        st.error(f"Google API 通信中にエラーが発生しました: {e}")

def show_payment_master():
    """支払い方法マスター設定画面"""
    st.markdown("#### 💳 支払い方法マスター")
    st.info("レシート登録時に選択できる「支払い方法」を追加・管理します。")
    
    # 現在の支払い方法一覧を表示
    methods = get_payment_methods(st.session_state['username'])
    
    st.write("##### 登録済みの支払い方法")
    if methods:
        # DataFrameにして表示
        df_methods = pd.DataFrame(methods)
        
        # is_credit_card 列がなければ False にしておく
        if "is_credit_card" not in df_methods.columns:
            df_methods["is_credit_card"] = False
        # credit_limit 列がなければ 空白 または NaN
        if "credit_limit" not in df_methods.columns:
            df_methods["credit_limit"] = None
        
        # 表示用カラムを絞る
        display_cols = ["payment_id", "name", "type", "is_credit_card", "closing_date", "payment_month", "payment_date", "credit_limit"]
        df_display = df_methods[[c for c in display_cols if c in df_methods.columns]].copy()
        
        # Boolean を "はい" / "いいえ" にする
        if "is_credit_card" in df_display.columns:
            # 念のため文字列になっている場合を考慮
            df_display["is_credit_card"] = df_display["is_credit_card"].apply(
                lambda x: "はい" if str(x).lower() == "true" or x == True else "いいえ"
            )

        df_display = df_display.rename(columns={
            "name": "支払い方法名", "type": "種類", "is_credit_card": "クレカ判定", 
            "closing_date": "締日", "payment_month": "支払月", "payment_date": "支払日", "credit_limit": "限度額"
        })
        event = st.dataframe(
            df_display, 
            use_container_width=True, 
            hide_index=True,
            selection_mode="single-row",
            on_select="rerun",
            key="payment_master_list"
        )
        
        selected_indices = event.selection.rows
        if selected_indices:
            selected_idx = selected_indices[0]
            selected_payment_id = df_display.iloc[selected_idx]['payment_id']
            
            # Find original dict
            selected_method = next((m for m in methods if m["payment_id"] == selected_payment_id), None)
            
            if selected_method:
                st.markdown(f"**選択されている支払い方法:** {selected_method['name']}")
                
                # Check if we are confirming deletion
                if st.session_state.get("confirm_delete_payment_id") == selected_payment_id:
                    st.warning(f"「{selected_method['name']}」を本当に削除しますか？")
                    col_yes, col_no = st.columns(2)
                    with col_yes:
                        if st.button("はい、削除します", type="primary", use_container_width=True):
                            with st.spinner("削除中..."):
                                success, msg = delete_payment_method(st.session_state['username'], selected_payment_id)
                                if success:
                                    st.success(msg)
                                    del st.session_state["confirm_delete_payment_id"]
                                    time.sleep(1)
                                    st.rerun()
                                else:
                                    st.error(msg)
                    with col_no:
                        if st.button("キャンセル", use_container_width=True):
                            del st.session_state["confirm_delete_payment_id"]
                            st.rerun()
                else:
                    col_e, col_d = st.columns(2)
                    with col_e:
                        if st.button("✏️ 修正する", use_container_width=True):
                            st.session_state.show_add_payment_form = True
                            st.session_state.editing_payment = selected_method
                            st.rerun()
                    with col_d:
                        if st.button("🗑️ 削除する", use_container_width=True):
                            st.session_state.confirm_delete_payment_id = selected_payment_id
                            st.rerun()
    else:
        st.write("登録されている支払い方法はありません。（デフォルトとして「現金」が利用可能です）")
        
    if not st.session_state.get("show_add_payment_form", False):
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➕ 新しく支払い方法を追加する", type="primary"):
            st.session_state.show_add_payment_form = True
            if "editing_payment" in st.session_state:
                del st.session_state["editing_payment"]
            st.rerun()
    else:
        st.markdown("---")
        is_edit_mode = "editing_payment" in st.session_state and st.session_state.editing_payment is not None
        edit_data = st.session_state.get("editing_payment", {})
        
        st.write("##### 支払い方法の修正" if is_edit_mode else "##### 新規支払い方法の追加")
        with st.container(border=True):
            p_col1, p_col2 = st.columns(2)
            with p_col1:
                default_name = edit_data.get("name", "")
                new_name = st.text_input("支払い方法名（例: 楽天カード、現金、PayPay）", value=default_name)
            with p_col2:
                default_type = edit_data.get("type", "クレジットカード")
                type_options = ["クレジットカード", "デビットカード", "電子マネー", "ポイント", "現金", "銀行振込", "その他", "未設定"]
                type_index = type_options.index(default_type) if default_type in type_options else (len(type_options) - 1)
                new_type = st.selectbox("種類", type_options, index=type_index)
                is_credit_card = (new_type == "クレジットカード")
            
            if is_credit_card:
                # クレジットカード専用設定
                st.markdown("###### クレジットカード詳細設定")
                c_col1, c_col2, c_col3 = st.columns(3)
                
                closing_opts = ["", "末日"] + [str(i) for i in range(1, 31)]
                def_closing = str(edit_data.get("closing_date", ""))
                closing_idx = closing_opts.index(def_closing) if def_closing in closing_opts else 0
                
                month_opts = ["", "当月", "翌月", "翌々月"]
                def_month = str(edit_data.get("payment_month", ""))
                month_idx = month_opts.index(def_month) if def_month in month_opts else 0
                
                date_opts = ["", "末日"] + [str(i) for i in range(1, 31)]
                def_date = str(edit_data.get("payment_date", ""))
                date_idx = date_opts.index(def_date) if def_date in date_opts else 0
                
                def_limit = int(edit_data.get("credit_limit", 0)) if edit_data.get("credit_limit") else 0
                
                with c_col1:
                    closing_date = st.selectbox("締日", closing_opts, index=closing_idx)
                with c_col2:
                    payment_month = st.selectbox("支払月", month_opts, index=month_idx)
                with c_col3:
                    payment_date = st.selectbox("支払日", date_opts, index=date_idx)

                credit_limit = st.number_input("限度額（円）", min_value=0, step=10000, value=def_limit, help="0の場合は設定なし")
            else:
                closing_date = ""
                payment_month = ""
                payment_date = ""
                credit_limit = 0
                
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                submit_label = "更新する" if is_edit_mode else "保存する"
                submit_payment = st.button(submit_label, type="primary", use_container_width=True)
            with btn_col2:
                cancel_payment = st.button("キャンセル", use_container_width=True)
                
            if cancel_payment:
                st.session_state.show_add_payment_form = False
                if "editing_payment" in st.session_state:
                    del st.session_state["editing_payment"]
                st.rerun()
                
            if submit_payment:
                if not new_name.strip():
                    st.warning("⚠️ 支払い方法名は必須です。")
                elif is_credit_card and (not closing_date or not payment_month or not payment_date):
                    st.warning("⚠️ クレジットカードフラグがONの場合は、締日・支払月・支払日をすべて設定してください。")
                else:
                    with st.spinner("保存中..."):
                        # クレカ以外は詳細設定をクリアして保存
                        limit_val = credit_limit if credit_limit > 0 else ""
                        if not is_credit_card:
                            closing_date = ""
                            payment_month = ""
                            payment_date = ""
                            limit_val = ""
                            
                        # IDは自動採番、または既存IDを継承
                        generated_id = edit_data.get("payment_id") if is_edit_mode else f"pay_{int(time.time() * 1000)}"
                            
                        data_to_save = {
                            "payment_id": generated_id,
                            "name": new_name.strip(),
                            "type": new_type.strip(),
                            "is_credit_card": is_credit_card,
                            "closing_date": closing_date,
                            "payment_month": payment_month,
                            "payment_date": payment_date,
                            "credit_limit": limit_val
                        }
                        succ, msg = save_payment_method(st.session_state['username'], data_to_save)
                        if succ:
                            st.success(msg)
                            st.session_state.show_add_payment_form = False
                            if "editing_payment" in st.session_state:
                                del st.session_state["editing_payment"]
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error(msg)


def show_yearly_dashboard():
    # ヘッダーを表示するためのプレースホルダー
    header_placeholder = st.empty()
    
    # 年次ナビゲーションを表示
    render_year_navigation()
    
    # メインタイトル表示
    header_placeholder.markdown("#### 📊 ダッシュボード (年次集計)")
    
    selected_year = st.session_state['current_month'].year
    target_date = datetime(selected_year, 1, 1)
    
    with st.spinner(f"{selected_year}年のデータを集計中..."):
        # 年次モードでデータを取得
        df = load_transactions_data(target_date, mode="yearly")
        # 前年比較用に前年データも取得
        prev_year_date = target_date - relativedelta(years=1)
        df_prev = load_transactions_data(prev_year_date, mode="yearly")

    if df.empty:
        st.info(f"※{selected_year}年のデータはまだありません。")
        return

    # 集計用のデータ（内税を除外）
    df_agg = df[df["category"] != "消費税（内税）"] if "category" in df.columns else df
    df_prev_agg = df_prev[df_prev["category"] != "消費税（内税）"] if "category" in df_prev.columns else df_prev

    # --- グラフ表示選択 (月次と同様に2カラムのドロップダウン) ---
    col_a, col_b = st.columns(2)
    with col_a:
        analysis_axis = st.selectbox(
            "分析軸を選択", 
            ["大分類別", "小分類別", "店舗別", "消費税", "支払い方法"], 
            index=0, 
            key="yearly_analysis_axis"
        )
    with col_b:
        graph_type = st.selectbox(
            "グラフを選択",
            ["円グラフ", "棒グラフ", "前年対比"],
            index=0,
            key="yearly_graph_type"
        )

    # 選択に応じて集計対象の列を決定
    group_col = None
    title_label = ""
    if analysis_axis == "大分類別":
        group_col = "category"
        title_label = "大分類別"
    elif analysis_axis == "小分類別":
        for col in ["subcategory", "sub_category", "小分類"]:
            if col in df.columns:
                group_col = col
                break
        title_label = "小分類別"
    elif analysis_axis == "店舗別":
        for col in ["store_name", "store", "店舗"]:
            if col in df.columns:
                group_col = col
                break
        title_label = "店舗別"
    elif analysis_axis == "消費税":
        group_col = "tax_group"
        title_label = "消費税率別"
        # 税率マッピング関数の定義
        def map_tax(subcategory):
            sub = str(subcategory)
            if "10%" in sub: return "外税10%＋内税10%"
            elif "8%" in sub: return "外税8%+内税8%"
            else: return "外税？%+内税？%"
        
        # サブカテゴリ列を探す
        sub_col = "subcategory"
        for col in ["subcategory", "sub_category", "小分類"]:
            if col in df.columns:
                sub_col = col
                break
        
        # 【修正】当年度および前年度のデータを、税金カテゴリのみに限定
        df = df[df["category"].isin(["消費税（外税）", "消費税（内税）"])]
        if not df_prev.empty:
            df_prev = df_prev[df_prev["category"].isin(["消費税（外税）", "消費税（内税）"])]
            
        if sub_col in df.columns:
            df["tax_group"] = df[sub_col].apply(map_tax)
        if not df_prev.empty and sub_col in df_prev.columns:
            df_prev["tax_group"] = df_prev[sub_col].apply(map_tax)
            
        df_agg = df 
        df_prev_agg = df_prev
        
    elif analysis_axis == "支払い方法":
        group_col = "payment_method"
        title_label = "支払い方法別"

    if graph_type == "前年対比":
        # 当年データの月別集計 (年次グラフでは内税を含めて内訳を提示)
        # ※分析軸が店舗別の場合は意図により内税を除外するか検討の余地ありだが、月次は内税込みで傾向を見る
        df_for_yearly = df_agg if analysis_axis == "店舗別" else df
        df_for_yearly['month'] = df_for_yearly['date'].dt.month
        monthly_summary = df_for_yearly.groupby('month', as_index=False)['amount'].sum()
        full_months = pd.DataFrame({'month': range(1, 13)})
        monthly_summary = pd.merge(full_months, monthly_summary, on='month', how='left').fillna(0)
        monthly_summary['month_label'] = monthly_summary['month'].apply(lambda x: f"{x}月")

        # 前年データの月別集計
        df_prev_for_yearly = df_prev_agg if analysis_axis == "店舗別" else df_prev
        df_prev_for_yearly['month'] = df_prev_for_yearly['date'].dt.month
        prev_summary = df_prev_for_yearly.groupby('month', as_index=False)['amount'].sum()
        prev_summary = pd.merge(full_months, prev_summary, on='month', how='left').fillna(0)
        
        comparison_data = pd.DataFrame({
            '月': list(monthly_summary['month_label']) * 2,
            '金額': list(prev_summary['amount']) + list(monthly_summary['amount']),
            '年度': [f'{selected_year-1}年'] * 12 + [f'{selected_year}年'] * 12
        })
        fig = px.bar(comparison_data, x='月', y='金額', color='年度',
                     barmode='group',
                     title=f"{selected_year}年 vs {selected_year-1}年 支出比較 (月次展開)")
        
        # マイナスのデータを黄色にする
        for trace in fig.data:
            if hasattr(trace, 'y') and trace.y is not None:
                orig_color = trace.marker.color
                trace.marker.color = ['#FFFF00' if v is not None and (isinstance(v, (int, float)) and v < 0) else orig_color for v in trace.y]

        fig.update_yaxes(zerolinewidth=2, zerolinecolor='black')
        st.plotly_chart(fig, use_container_width=True)

    elif graph_type == "棒グラフ":
        # 当年の月別推移 (積上げ棒グラフ)
        df_for_bar = df_agg.copy() if analysis_axis == "店舗別" else df.copy()
        df_for_bar['month'] = df_for_bar['date'].dt.month
        df_for_bar['month_label'] = df_for_bar['month'].apply(lambda x: f"{x}月")
        
        if group_col and group_col in df_for_bar.columns:
            yearly_grouped = df_for_bar.groupby(['month', 'month_label', group_col], as_index=False)["amount"].sum()
            # 指定の順番を保つため、全体の合計額順でカテゴリーをソートする (消費税の場合は固定順)
            if analysis_axis == "消費税":
                cat_sum = ["外税8%+内税8%", "外税10%＋内税10%", "外税？%+内税？%"]
            else:
                cat_sum = yearly_grouped.groupby(group_col)["amount"].sum().sort_values(ascending=False).index.tolist()
            
            all_month_labels = [f"{i}月" for i in range(1, 13)]
            fig = px.bar(
                yearly_grouped,
                x='month_label',
                y='amount',
                color=group_col,
                title=f"{selected_year}年 {title_label}月次推移 (積上げ棒グラフ)",
                labels={"amount": "金額", "month_label": "月", group_col: analysis_axis[:-1]},
                category_orders={"month_label": all_month_labels, group_col: cat_sum},
                color_discrete_map=CATEGORY_COLOR_MAP
            )

            # --- トレンド線（カテゴリ別累積推移）の追加 ---
            # 各カテゴリの積上げ高さに合わせるため、下層のカテゴリから順に合算した累積シリーズを維持する
            cumulative_monthly = pd.Series(0.0, index=all_month_labels)
            for cat in reversed(cat_sum): # 下層から積上げるためreversed
                cat_data = yearly_grouped[yearly_grouped[group_col] == cat]
                if not cat_data.empty:
                    # all_month_labelsに合わせて補完し、その月のカテゴリ合計を出す
                    cat_data_full = pd.DataFrame({'month_label': all_month_labels}).merge(cat_data, on='month_label', how='left').fillna({'amount': 0.0})
                    cumulative_monthly += cat_data_full['amount']
                    
                    line_color = CATEGORY_COLOR_MAP.get(cat, None)
                    fig.add_trace(go.Scatter(
                        x=all_month_labels,
                        y=cumulative_monthly,
                        mode='lines+markers',
                        name=f'{cat} (累積推移)',
                        line=dict(color=line_color, width=1),
                        marker=dict(size=4),
                        showlegend=False
                    ))

            if analysis_axis == "消費税":
                fig.update_xaxes(categoryorder='array', categoryarray=cat_sum)
            # マイナスのデータを黄色にする (棒グラフのトレースのみ対象)
            for trace in fig.data:
                if trace.type == 'bar' and hasattr(trace, 'y') and trace.y is not None:
                    orig_color = trace.marker.color
                    trace.marker.color = ['#FFFF00' if v is not None and (isinstance(v, (int, float)) and v < 0) else orig_color for v in trace.y]

            fig.update_yaxes(zerolinewidth=2, zerolinecolor='black')
            st.plotly_chart(fig, use_container_width=True)
                
            # 年間合計は内税除外
            year_total = df_agg["amount"].sum()
            st.metric(f"{selected_year}年 総支出額", f"￥{int(year_total):,}")
        else:
            st.warning(f"分析に必要な列（{analysis_axis[:-1]}）がありません。")

    elif graph_type == "円グラフ": # 円グラフ
        if group_col and group_col in df_agg.columns:
            cat_grouped = df_agg.groupby(group_col, as_index=False)["amount"].sum()
            # 金額が0以下のデータを除外
            grouped_df = cat_grouped[cat_grouped["amount"] > 0]
            grouped_df = grouped_df.sort_values(by="amount", ascending=False)
            
            # 表示順序の固定
            tax_order = ["外税8%+内税8%", "外税10%＋内税10%", "外税？%+内税？%"]
            category_orders = {group_col: tax_order} if analysis_axis == "消費税" else {group_col: grouped_df[group_col].tolist()}

            fig_pie = px.pie(
                grouped_df, 
                values='amount', 
                names=group_col, 
                hole=0.4, 
                title=f'{selected_year}年 {title_label}支出シェア',
                category_orders=category_orders,
                color=group_col,
                color_discrete_map=CATEGORY_COLOR_MAP
            )
            fig_pie.update_traces(
                textposition='inside', 
                textinfo='percent+label'
            )
            st.plotly_chart(fig_pie, use_container_width=True)
            
            year_total = df_agg["amount"].sum()
            st.metric(f"{selected_year}年 総支出額", f"￥{int(year_total):,}")
        else:
            st.warning(f"分析に必要な列（{analysis_axis[:-1]}）がありません。")
    
    st.markdown("---")
    st.markdown("##### カテゴリ別内訳 (年次)")
    render_transaction_breakdown(df, "yearly_dashboard")

def handle_menu_change():
    """サイドバーでのメニュー変更時にURLパラメータをクリアし、必要に応じてダッシュボード表示を月次にリセットする"""
    if "date" in st.query_params:
        del st.query_params["date"]
    if "menu" in st.query_params:
        del st.query_params["menu"]
    
    # セッション内のメニュー選択を確認（on_change時点で st.session_state.menu_selection は更新されている）
    target_menu = st.session_state.get("menu_selection")
    # 仕様：カレンダー、レシート取込、レシート修正を選択した際、ダッシュボード選択状態をリセット
    if target_menu in ["カレンダー", "レシート取込", "レシート修正"]:
        st.session_state["menu_selection_reset_flag"] = True
        
    if target_menu == "👁AI相談":
        st.session_state["refresh_ai_data_flag"] = True
    
    # サイドバーを閉じるフラグ
    st.session_state["collapse_sidebar_flag"] = True


def show_bank_master():
    st.markdown("## 🏦 銀行マスター")
    st.caption("口座引落や銀行振込で利用する銀行名を登録・管理します。")
    username = st.session_state['username']
    
    st.markdown("### 📋 登録済みの銀行一覧")
    st.caption("※左端のチェックボックスを選択すると、そのデータの下に編集フォームが開きます。一番上の「➕ 新規追加」を選ぶと新しく登録できます。")
    import pandas as pd
    banks = get_banks(username)
    
    df = pd.DataFrame(banks) if banks else pd.DataFrame(columns=["bank_id", "bank_name", "balance"])
    if not df.empty:
        df = df[["bank_id", "bank_name", "balance"]].copy()
    else:
        df = pd.DataFrame(columns=["bank_id", "bank_name", "balance"])
    df.rename(columns={"bank_name": "銀行名", "balance": "残高"}, inplace=True)
    
    dummy = pd.DataFrame([{"bank_id": "new", "銀行名": "➕ 新規追加（ここを選択して追加）", "残高": 0}])
    df = pd.concat([dummy, df], ignore_index=True)
    
    event = st.dataframe(
        df[["銀行名", "残高"]],
        use_container_width=True,
        on_select="rerun",
        selection_mode="single-row",
        hide_index=True
    )
    
    selected_rows = event.selection.rows
    if selected_rows:
        selected_idx = selected_rows[0]
        target_id = df.iloc[selected_idx]["bank_id"]
        if target_id == "new":
            edit_target = {"bank_id": "new", "bank_name": "", "balance": 0}
        else:
            edit_target = next(b for b in banks if b["bank_id"] == target_id)
            
    if not selected_rows:
        st.info("上の表の左端のチェックボックスをクリックして、編集または追加を行ってください。")
    elif edit_target["bank_id"] == "new":
        st.markdown("#### 新規登録")
        with st.form("add_bank_form"):
            c1, c2 = st.columns(2)
            with c1:
                new_name = st.text_input("銀行名*", placeholder="例：りそな銀行")
            with c2:
                new_balance = st.number_input("初期残高", value=0, step=1000)
            
            if st.form_submit_button("登録する", type="primary"):
                if new_name:
                    import time
                    bid = f"bank_{int(time.time())}"
                    if add_bank(username, bid, new_name, new_balance):
                        st.success(f"「{new_name}」を登録しました！")
                        time.sleep(1)
                        st.rerun()
                else:
                    st.error("銀行名を入力してください。")
    else:
        target_id = edit_target["bank_id"]
        confirm_key = "confirm_del_bank"
        
        if st.session_state.get(confirm_key) == target_id:
            st.warning(f"「{edit_target['bank_name']}」を本当に削除しますか？この操作は取り消せません。")
            c_yes, c_no = st.columns(2)
            with c_yes:
                if st.button("はい、削除します", type="primary", use_container_width=True):
                    import time
                    if delete_bank(username, target_id):
                        st.success(f"「{edit_target['bank_name']}」を削除しました！")
                        st.session_state[confirm_key] = None
                        time.sleep(1)
                        st.rerun()
            with c_no:
                if st.button("キャンセル", use_container_width=True):
                    st.session_state[confirm_key] = None
                    st.rerun()
        else:
            st.markdown("#### 編集")
            with st.form(f"edit_bank_form_{target_id}"):
                col1, col2 = st.columns(2)
                with col1:
                    edit_name = st.text_input("銀行名*", value=edit_target["bank_name"])
                with col2:
                    edit_balance = st.number_input("残高", value=int(edit_target["balance"]), step=1000)
                
                col_upd, col_del = st.columns([2, 1])
                with col_upd:
                    update_submitted = st.form_submit_button("情報を更新する", type="primary", use_container_width=True)
                with col_del:
                    delete_submitted = st.form_submit_button("削除する", use_container_width=True)
                
                if update_submitted:
                    if edit_name:
                        import time
                        if update_bank(username, target_id, edit_name, edit_balance):
                            st.success(f"「{edit_name}」の情報を更新しました！")
                            time.sleep(1)
                            st.rerun()
                    else:
                        st.error("銀行名を入力してください。")
                if delete_submitted:
                    st.session_state[confirm_key] = target_id
                    st.rerun()

def _render_fixed_cost_form(action_type, username, target_data=None):
    if target_data is None:
        target_data = {
            "fixed_cost_id": "new", "item_name": "", "amount": 0, "major_category": "固定費",
            "payment_1": "口座引落", "payment_2": "", "is_finite": "無限", "fixed_or_variable": "固定",
            "payment_month": "毎月", "final_amount": 0, "transfer_fee": 0, "start_month": "", "completion_month": ""
        }
        
    p1_opts = ["口座引落", "クレジットカード", "銀行振込"]
    idx_p1 = p1_opts.index(target_data["payment_1"]) if target_data["payment_1"] in p1_opts else 0

    col1, col2 = st.columns(2)
    with col1:
        payment_1 = st.selectbox("固定費支払１*", p1_opts, index=idx_p1)
    
    with col2:
        if payment_1 == "クレジットカード":
            methods = get_payment_methods(username)
            cc_names = [m["name"] for m in methods if m.get("is_credit_card") or m.get("type") == "クレジットカード"]
            if not cc_names: cc_names = ["(カード未登録)"]
            idx_p2 = cc_names.index(target_data["payment_2"]) if target_data["payment_2"] in cc_names else 0
            payment_2 = st.selectbox("固定費支払２（カード選択）*", cc_names, index=idx_p2)
        else:
            banks = get_banks(username)
            b_names = [b["bank_name"] for b in banks]
            if not b_names: b_names = ["(銀行未登録)"]
            idx_p2 = b_names.index(target_data["payment_2"]) if target_data["payment_2"] in b_names else 0
            payment_2 = st.selectbox("固定費支払２（銀行選択）*", b_names, index=idx_p2)

    col4, col5 = st.columns(2)
    with col4:
        item_name = st.text_input("科目（支払いの名前）*", value=target_data["item_name"], placeholder="例：住宅ローン")
    with col5:
        try:
            amt_val = int(target_data["amount"])
        except:
            amt_val = 0
        amount = st.number_input("支払額*", value=amt_val, min_value=0, step=100)

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        fv_opts = ["固定", "変動"]
        idx_fv = fv_opts.index(target_data["fixed_or_variable"]) if target_data["fixed_or_variable"] in fv_opts else 0
        fixed_or_variable = st.selectbox("変動or固定*", fv_opts, index=idx_fv)
    with col_b:
        if_opts = ["無限", "有限"]
        idx_if = if_opts.index(target_data["is_finite"]) if target_data["is_finite"] in if_opts else 0
        is_finite = st.selectbox("有限or無限*", if_opts, index=idx_if)
    with col_c:
        pm_opts = ["毎月"] + [f"{i}月" for i in range(1, 13)]
        idx_pm = pm_opts.index(target_data["payment_month"]) if target_data["payment_month"] in pm_opts else 0
        payment_month = st.selectbox("支払月*", pm_opts, index=idx_pm)
        
    final_amount = 0
    transfer_fee = 0
    start_month = ""
    completion_month = ""
    
    if is_finite == "有限" or payment_1 == "銀行振込":
        if is_finite == "有限":
            import datetime
            now = datetime.datetime.now()
            start_month_opts = []
            prev_m = now.month - 1
            prev_y = now.year
            if prev_m == 0:
                prev_m = 12
                prev_y -= 1
            start_month_opts.append(f"{prev_y}年{prev_m}月")
            for i in range(13):
                moff = now.month + i
                mv = (moff - 1) % 12 + 1
                yv = now.year + (moff - 1) // 12
                start_month_opts.append(f"{yv}年{mv}月")
            
            c_start, c_dummy = st.columns(2)
            with c_start:
                s_idx = start_month_opts.index(target_data["start_month"]) if target_data["start_month"] in start_month_opts else 0
                start_month = st.selectbox("開始月", start_month_opts, index=s_idx)
            
            comp_years = [f"{now.year + i}年" for i in range(30)]
            comp_months = [f"{i}月" for i in range(1, 13)]    
            c_cy, c_cm = st.columns(2)
            cur_cy = target_data["completion_month"][:5] if target_data["completion_month"] and "年" in target_data["completion_month"] else comp_years[0]
            cur_cm = target_data["completion_month"][5:] if target_data["completion_month"] and "年" in target_data["completion_month"] else comp_months[0]
            with c_cy:
                cy_idx = comp_years.index(cur_cy) if cur_cy in comp_years else 0
                c_year = st.selectbox("完済年", comp_years, index=cy_idx)
            with c_cm:
                cm_idx = comp_months.index(cur_cm) if cur_cm in comp_months else 0
                c_month = st.selectbox("完済月", comp_months, index=cm_idx)
            completion_month = f"{c_year}{c_month}" 
            
        c_amt, c_fee = st.columns(2)
        with c_amt:
            if is_finite == "有限":
                try: fa_val = int(target_data.get("final_amount", 0))
                except: fa_val = 0
                final_amount = st.number_input("最終月額", value=fa_val, min_value=0, step=100)
        with c_fee:
            if payment_1 == "銀行振込":
                try: tf_val = int(target_data.get("transfer_fee", 0))
                except: tf_val = 0
                transfer_fee = st.number_input("振込手数料", value=tf_val, min_value=0, step=10)

    if action_type == "add":
        if st.button("登録する", type="primary"):
            if item_name and payment_1 and payment_month:
                import time
                fixed_cost_id = f"fc_{int(time.time())}"
                success = add_fixed_cost(username, fixed_cost_id, "固定費", payment_1, payment_2, is_finite, item_name, amount, fixed_or_variable, payment_month, final_amount, transfer_fee, start_month, completion_month)
                if success:
                    st.success(f"「{item_name}」を登録しました！")
                    time.sleep(1)
                    st.rerun()
            else:
                st.error("必須項目を入力してください。")
    elif action_type == "update":
        col_u, col_d = st.columns([2, 1])
        with col_u:
            update_btn = st.button("情報を更新する", type="primary", use_container_width=True)
        with col_d:
            delete_btn = st.button("削除する", use_container_width=True)
            
        if update_btn:
            if item_name and payment_1 and payment_month:
                import time
                success = update_fixed_cost(username, target_data["fixed_cost_id"], "固定費", payment_1, payment_2, is_finite, item_name, amount, fixed_or_variable, payment_month, final_amount, transfer_fee, start_month, completion_month)
                if success:
                    st.success(f"「{item_name}」の情報を更新しました！")
                    time.sleep(1)
                    st.rerun()
            else:
                st.error("必須項目を入力してください。")
        return delete_btn
    return False

def main():
    # --- 初期化 ---
    if 'logged_in' not in st.session_state:
        st.session_state['logged_in'] = False
        
    # 生体認証関連の残骸変数を徹底的にクリア
    biometric_keys = [
        'webauthn_reg_comp', 'biometric_auth_status', 'auth_status', 
        'webauthn_auth_data', 'biometric_user', 'passkey_status'
    ]
    for k in biometric_keys:
        if k in st.session_state:
            del st.session_state[k]

    pass
    
    # ログイン済みの状態
    if st.session_state.get('logged_in', False):

        
        # 自動画面遷移のためのリダイレクト処理
        if st.session_state.get('redirect_to_dashboard'):
            st.session_state['menu_selection'] = "ダッシュボード（月次集計）"
            st.session_state['redirect_to_dashboard'] = False
            
        # サイドバー連動ロジック（自動切り替え）
        # handle_menu_change でセットされたフラグをチェック
        if st.session_state.get("menu_selection_reset_flag"):
            # リセット対象メニュー（カレンダー、レシート系）が選ばれた現在の状態から、
            # 次にダッシュボードに戻った時に「月次集計」になるように内部状態をいじる
            # ただし、現状の radio ボタンの挙動として、「もし次ダッシュボード系を選ぶなら」という制御が必要
            st.session_state["menu_selection_reset_flag"] = False

        # サイドバーメニューの実装
        with st.sidebar:
            st.markdown("""
                <style>
                /* サイドバーのメニュータイトル・選択肢の文字を1pt小さく */
                [data-testid="stSidebar"] .stMarkdown p, 
                [data-testid="stSidebar"] label[data-baseweb="radio"] div {
                    font-size: calc(1rem - 1pt) !important;
                }
                /* サイドバーの大タイトル下の余白を詰める */
                [data-testid="stSidebar"] .stMarkdown {
                    margin-bottom: -15px !important;
                }
                /* メイン画面の表示タイトルを2pt大きく */
                .block-container h2 { font-size: calc(1.75rem + 2pt) !important; }
                .block-container h3 { font-size: calc(1.50rem + 2pt) !important; }
                .block-container h4 { font-size: calc(1.25rem + 2pt) !important; }
                .block-container h5 { font-size: calc(1.00rem + 2pt) !important; }
                </style>
            """, unsafe_allow_html=True)
            st.subheader("マイニー [Ver 4.22.0]")
            st.write(f"🔑 ユーザー: **{st.session_state['username']}**")
            st.markdown("---")
            if 'menu_selection' not in st.session_state:
                st.session_state['menu_selection'] = "カレンダー"
            
            # 別の画面から戻ってきたとき用
            if "last_menu_selection" not in st.session_state:
                st.session_state.last_menu_selection = st.session_state['menu_selection']
                
            def on_menu_change(key):
                new_val = st.session_state.get(key)
                if new_val is not None:
                    st.session_state['menu_selection'] = new_val
                    handle_menu_change()
                    st.session_state.last_menu_selection = st.session_state['menu_selection']

            # Categories and items
            group1_opts = ["ダッシュボード（月次集計）", "ダッシュボード（年次集計）", "カレンダー", "クレジットカード"]
            group2_opts = ["レシート取込", "レシート手入力", "レシート修正"]
            group3_opts = ["マニュアル", "ヘルプ", "AI相談"]
            group4_opts = ["支払方法マスター", "銀行マスター", "プロフィール設定"]
            group5_opts = ["支払管理シート新規作成", "固定費マスター設定", "固定費データ展開", "変動費データ更新", "支払管理シートを確認"]
            group6_opts = []

            if st.session_state.get('username', '').lower() == 'tkouho':
                group6_opts = ["カテゴリマスター", "データチェック"]

            
            current_sel = st.session_state['menu_selection']
            
            # 各radioコンポーネントの値を現在の選択と強制的に一致させる（他グループの選択を解除）
            st.session_state["menu_g1"] = current_sel if current_sel in group1_opts else None
            st.session_state["menu_g2"] = current_sel if current_sel in group2_opts else None
            st.session_state["menu_g3"] = current_sel if current_sel in group3_opts else None
            st.session_state["menu_g4"] = current_sel if current_sel in group4_opts else None
            st.session_state["menu_g5"] = current_sel if current_sel in group5_opts else None
            st.session_state["menu_g6"] = current_sel if current_sel in group6_opts else None
            
            # ＝＝＝ 変動費管理セクション ＝＝＝
            st.markdown("<div style='padding-bottom: 30px;'><h3 style='color: #007bff; border-bottom: 2px solid #007bff; padding-bottom: 5px; margin: 0;'>🔵 変動費管理</h3></div>", unsafe_allow_html=True)
            
            st.markdown("【表示・分析系】")
            st.radio("g1", group1_opts, key="menu_g1", 
                     on_change=on_menu_change, args=("menu_g1",), label_visibility="collapsed")
            
            st.markdown("【レシート管理】")
            st.radio("g2", group2_opts, key="menu_g2", 
                     on_change=on_menu_change, args=("menu_g2",), label_visibility="collapsed")
            
            st.markdown("【相談・サポート】")
            st.radio("g3", group3_opts, key="menu_g3", 
                     on_change=on_menu_change, args=("menu_g3",), label_visibility="collapsed")
            
            st.markdown("【マスター設定】")
            st.radio("g4", group4_opts, key="menu_g4", 
                     on_change=on_menu_change, args=("menu_g4",), label_visibility="collapsed")

            # ＝＝＝ 支払管理セクション ＝＝＝
            st.markdown("<div style='padding-top: 10px; padding-bottom: 30px;'><h3 style='color: #ff6b6b; border-bottom: 2px solid #ff6b6b; padding-bottom: 5px; margin: 0;'>🔴 支払管理</h3></div>", unsafe_allow_html=True)
            
            st.radio("g5", group5_opts, key="menu_g5", 
                     on_change=on_menu_change, args=("menu_g5",), label_visibility="collapsed")

            if st.session_state.get('username', '').lower() == 'tkouho' and group6_opts:
                st.markdown("<div style='padding-top: 10px; padding-bottom: 30px;'><h3 style='color: #006400; border-bottom: 2px solid #006400; padding-bottom: 5px; margin: 0;'>● ツール</h3></div>", unsafe_allow_html=True)
                st.radio("g6", group6_opts, key="menu_g6", 
                         on_change=on_menu_change, args=("menu_g6",), label_visibility="collapsed")

                        
            menu_selection = st.session_state['menu_selection']

            # サイドバー自動折りたたみJS削除
            if st.session_state.get("collapse_sidebar_flag"):
                st.session_state["collapse_sidebar_flag"] = False
                pass
            
            st.markdown("---")
            if st.button("ログアウト", use_container_width=True):
                # ログアウト時にURLパラメータとセッション状態を完全にクリアする
                st.query_params.clear()
                st.session_state.clear()
                st.rerun()

            st.markdown("---")
            
            # ダウンロード状態の初期化
            if 'dl_step' not in st.session_state:
                st.session_state.dl_step = "init"
            if 'dl_format' not in st.session_state:
                st.session_state.dl_format = None

            with st.expander("📥 データのダウンロード", expanded=(st.session_state.dl_step != "init")):
                if st.session_state.dl_step == "init":
                    st.info("集計やバックアップ用にデータをダウンロードできます。形式を選択してください。")
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("📊 Excel形式", use_container_width=True):
                            st.session_state.dl_format = "Excel"
                            st.session_state.dl_step = "confirm"
                            st.rerun()
                    with col2:
                        if st.button("📄 CSV形式", use_container_width=True):
                            st.session_state.dl_format = "CSV"
                            st.session_state.dl_step = "confirm"
                            st.rerun()

                elif st.session_state.dl_step == "confirm":
                    # 選択された形式でデータを生成（スピナー表示）
                    with st.spinner(f"{st.session_state.dl_format}データを準備中..."):
                        if st.session_state.dl_format == "Excel":
                            data = generate_excel_download(st.session_state['username'])
                            ext = "xlsx"
                            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        else:
                            data = generate_csv_download(st.session_state['username'])
                            ext = "csv"
                            mime = "text/csv"
                        
                        filename = f"kakeibo_{st.session_state['username']}_{datetime.now().strftime('%Y%m%d')}.{ext}"

                    st.warning(f"**{st.session_state.dl_format}形式**で出力します。")
                    st.write("📂 ブラウザの「ダウンロード」フォルダ等に保存されます。")
                    
                    # 「実行する」ボタン（実体はダウンロードボタン）
                    st.download_button(
                        label="🚀 実行する (ダウンロード)",
                        data=data,
                        file_name=filename,
                        mime=mime,
                        type="primary",
                        use_container_width=True
                    )
                    
                    if st.button("キャンセル", use_container_width=True):
                        st.session_state.dl_step = "init"
                        st.session_state.dl_format = None
                        st.rerun()

                    # 💡 ポイント: ボタン表示後、あらかじめ状態をinitに戻しておく。
                    # ユーザーがダウンロードボタンを押すと再レンダリングされ、
                    # その時には dl_step="init" なのでエキスパンダーが閉じる。
                    st.session_state.dl_step = "init"
                    st.session_state.dl_format = None

        # メインコンテンツの切り替え
        if menu_selection == "ダッシュボード（月次集計）":
            show_dashboard()
        elif menu_selection == "データチェック":
            show_data_check()
        elif menu_selection == "ダッシュボード（年次集計）":
            show_yearly_dashboard()
        elif menu_selection == "カレンダー":
            st.markdown("#### 📅 カレンダー")
            
            # 共通ナビゲーションの適用
            df = render_month_navigation()
            
            daily_totals = {}
            if not df.empty and "date" in df.columns and "amount" in df.columns:
                df['day'] = df['date'].dt.day
                # 合計用のデータ（内税を除去）
                df_for_calc = df[df["category"] != "消費税（内税）"].copy() if "category" in df.columns else df
                daily_totals = df_for_calc.groupby('day')["amount"].sum().to_dict()
                
            year = st.session_state['current_month'].year
            month = st.session_state['current_month'].month

            # カレンダーの週の開始曜日を日曜日に設定し、その月のカレンダーマトリックスを取得
            calendar.setfirstweekday(calendar.SUNDAY)
            month_days = calendar.monthcalendar(year, month)

            # セッション状態の初期化（選択日の保持）
            if 'selected_date' not in st.session_state:
                # 初期状態では本日の日付を選択状態にする
                st.session_state['selected_date'] = datetime.today().strftime('%Y-%m-%d')

            # CSS定義（リンク方式でのマス目レイアウト）
            st.markdown("""
<style>
.calendar-grid {
    display: grid;
    grid-template-columns: repeat(7, 1fr);
    gap: 4px;
    width: 100%;
    max-width: 400px;
    margin: 0 auto;
}
/* リンクをマス目（枠線付き）として機能させる */
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

/* 日付：左上に配置 */
.cal-date {
    position: absolute;
    top: 4px; left: 8px;
    font-weight: bold;
    color: #333;
}

/* 金額：右下に赤字で配置 */
.cal-amount {
    position: absolute;
    bottom: 2px; right: 2px;
    color: red !important;
    font-size: 11px;
    font-weight: bold;
}

/* 曜日ヘッダー */
.weekday-header { text-align: center; font-weight: bold; padding: 5px 0; font-size: 0.85em; }
.sat-text { color: #3182ce; }
.sun-bg { background-color: #fff5f5; }
.sat-bg { background-color: #ebf8ff; }

/* 祝日名：中央付近に配置 */
.cal-holiday {
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    font-size: 8px;
    color: navy;
    width: 90%;
    text-align: center;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    pointer-events: none;
}
</style>
""", unsafe_allow_html=True)

            # カレンダーの表示
            # カレンダーのHTML構築
            cal_html = '<div class="calendar-grid">'
            
            # ヘッダー（曜日）
            for i, wd in enumerate(["日", "月", "火", "水", "木", "金", "土"]):
                cls = "sun-text" if i == 0 else "sat-text" if i == 6 else ""
                cal_html += f'<div class="weekday-header {cls}">{wd}</div>'

            # 日付の描画
            for week in month_days:
                for i, day in enumerate(week):
                    if day == 0:
                        # 空白のマス目
                        cal_html += '<div></div>'
                    else:
                        amount = daily_totals.get(day, None)
                        if amount is not None:
                            amount_text = f"￥{int(amount):,}" if amount > 0 else "￥0" if amount == 0 else ""
                        else:
                            amount_text = ""
                        date_obj = datetime(year, month, day).date()
                        date_str = date_obj.strftime('%Y-%m-%d')
                        is_selected = st.session_state.get('selected_date') == date_str
                        select_cls = "selected-link" if is_selected else ""
                        current_user = st.session_state.get("username", "")
                        
                        # 曜日および祝日による背景色の判定 (i: 0=日, 6=土)
                        holiday_name = jpholiday.is_holiday_name(date_obj)
                        bg_cls = "sun-bg" if (i == 0 or holiday_name) else "sat-bg" if i == 6 else ""
                        
                        cal_html += f'<a href="/?date={date_str}&user={current_user}&menu=カレンダー" target="_self" class="cal-link {select_cls} {bg_cls} notranslate" translate="no">'
                        cal_html += f'<div class="cal-date notranslate" translate="no">{day}</div>'
                        if holiday_name:
                            cal_html += f'<div class="cal-holiday notranslate" translate="no">{holiday_name}</div>'
                        cal_html += f'<div class="cal-amount notranslate" translate="no">{amount_text}</div>'
                        cal_html += '</a>'
            
            cal_html += '</div>'
            st.markdown(cal_html, unsafe_allow_html=True)

            # --- 対象日の詳細表示 ---
            selected_date = st.session_state.get('selected_date')
            if selected_date:
                # '2026-02-28' のような形式から日(day)を抽出して表示
                try:
                    display_day = int(selected_date.split("-")[-1])
                except:
                    display_day = selected_date

                # 該当日のデータをフィルタリング
                # selected_date (YYYY-MM-DD) と一致するか、current_month内でdayが一致するか
                day_val = int(selected_date.split("-")[-1])
                
                day_df = pd.DataFrame()
                if not df.empty and 'date' in df.columns:
                    day_df = df[df['date'].dt.day == day_val].copy()
                
                # 合計額の計算（内税を除去）
                df_for_calc = day_df[day_df["category"] != "消費税（内税）"] if "category" in day_df.columns else day_df
                day_total = int(df_for_calc['amount'].sum()) if not df_for_calc.empty else 0
                # デザインより翻訳回避を優先し、ネイティブなマークダウンで表示
                st.markdown(f"##### 📋 {display_day}日の支出詳細 (合計: ￥{day_total:,})")
                
                if not day_df.empty:
                    render_transaction_breakdown(day_df, "calendar")
                else:
                    st.info("この日の支出データはありません。")
            
            st.markdown("---")
            
        elif menu_selection == "レシート取込":
            st.markdown("#### 📸 レシート取込")
            
            # 共通ナビゲーションの適用
            _ = render_month_navigation()
            
            st.info("画像ファイルをアップロードしてレシートを解析します。")
            
            if "uploader_key" not in st.session_state:
                st.session_state.uploader_key = 0
                
            if "parsed_results" not in st.session_state:
                st.session_state.parsed_results = None
            
            uploaded_file = None
            
            file_img = st.file_uploader("レシートの画像をアップロード", type=['png', 'jpg', 'jpeg'], accept_multiple_files=False, key=f"uploader_{st.session_state.uploader_key}")
            if file_img:
                uploaded_file = file_img
            
            if uploaded_file is not None:
                # 表示用の向き補正
                display_img = Image.open(uploaded_file)
                display_img = ImageOps.exif_transpose(display_img)
                st.image(display_img, caption="取得したレシート画像", use_container_width=True)
                
                # Streamlitのボタンに色をつける（このページのみに適用される）
                st.markdown("""
                <style>
                    /* 登録ボタン(Primary) を青色に */
                    div.stButton > button[kind="primary"] {
                        background-color: #007bff !important;
                        color: white !important;
                        border-color: #007bff !important;
                    }
                    div.stButton > button[kind="primary"]:hover {
                        background-color: #0056b3 !important;
                        border-color: #0056b3 !important;
                    }

                    /* キャンセルボタン(Secondary) を赤色に */
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
                    # まだ解析していない場合
                    col1, col2 = st.columns(2)
                    with col1:
                        parse_btn = st.button("レシートを解析する", type="primary", use_container_width=True)
                    with col2:
                        cancel_parse_btn = st.button("キャンセル", type="secondary", use_container_width=True, key="cancel_upload")
                        
                    if cancel_parse_btn:
                        st.session_state.uploader_key += 1
                        st.rerun()
                        
                    if parse_btn:
                        try:
                            with st.spinner("画像を解析中... Geminiが読み取っています"):
                                results = parse_receipt_with_gemini(uploaded_file)
                                
                                # 整合性チェックと自動再解析ロジック
                                is_valid, clean_results = verify_receipt_checksum(results)
                                if not is_valid and isinstance(results, list) and len(results) > 0 and "error" not in results[0]:
                                    st.warning("🔄 読み取った金額に矛盾を検知しました。再解析を行っています...")
                                    retry_instruction = "【重要】読み取った数値に矛盾があります。特に『1』と『7』、あるいは小数点やポイント値の誤認がないか、元のレシート画像を再精査して数値を修正してください。"
                                    results_retry = parse_receipt_with_gemini(uploaded_file, additional_instruction=retry_instruction)
                                    is_valid_retry, clean_results_retry = verify_receipt_checksum(results_retry)
                                    
                                    if "error" not in results_retry and len(results_retry) > 0:
                                        is_valid = is_valid_retry
                                        clean_results = clean_results_retry
                                        results = results_retry # エラーハンドリング用にresultsも更新
                                
                            if isinstance(results, list) and len(results) > 0 and isinstance(results[0], dict) and "error" in results[0]:
                                st.error(f"解析に失敗しました: {results[0]['error']}")
                            elif isinstance(results, dict) and "error" in results:
                                st.error(f"解析に失敗しました: {results['error']}")
                            else:
                                st.session_state.ocr_checksum_valid = is_valid
                                st.session_state.parsed_results = clean_results
                                st.rerun()
                        except Exception as e:
                            st.error(f"解析処理中に予期せぬエラーが発生しました: {e}")
                
                else:
                    # 解析完了後、プレビューと確認画面を表示
                    results = st.session_state.parsed_results
                    is_valid = st.session_state.get('ocr_checksum_valid', True)
                    
                    if is_valid:
                        st.success("✅ 読み取り成功: 合計金額の計算が一致しました。")
                    else:
                        st.warning("⚠️ 合計金額の不整合があります。各明細の金額やポイント利用等を確認・修正してください。")
                    
                    if len(results) > 0:
                        preview_date = results[0].get("date", "")
                        preview_store = results[0].get("store_name", "")
                    else:
                        preview_date = ""
                        preview_store = ""
                        
                    # 支出合計の計算 (内税は二重計上防止のため除外)
                    def is_internal_tax(item):
                        cat = item.get("major_category", "その他")
                        return "内税" in cat or cat == "消費税（内税）"

                    def is_any_tax(item):
                        cat = item.get("major_category", "その他")
                        return "消費税" in cat or "内税" in cat or "外税" in cat

                    # AIの解析結果に消費税が含まれていない場合、自動で10%の内税項目を追加する機能
                    if len(results) > 0 and not any(is_any_tax(item) for item in results):
                        total_before_tax = sum(safe_money_int_cast(item.get("amount", 0)) for item in results)
                        tax_amount = int(total_before_tax * 0.1)
                        tax_item = {
                            "date": preview_date,
                            "store_name": preview_store,
                            "item_name": "消費税（内税）10%",
                            "amount": tax_amount,
                            "major_category": "消費税（内税）",
                            "minor_category": "内税10%"
                        }
                        results.append(tax_item)
                        st.session_state.parsed_results = results

                    total_amount = sum(safe_money_int_cast(item.get("amount", 0)) for item in results if not is_internal_tax(item))
                    
                    # 大分類別の内訳を集計
                    category_totals = {}
                    for item in results:
                        cat = item.get("major_category", "その他")
                        # 正規化処理を適用して大分類を揃える
                        majors = list(get_categories().keys())
                        final_major = "その他"
                        for m in majors:
                            if m in cat or cat in m:
                                final_major = m
                                break
                        
                        amt = safe_money_int_cast(item.get("amount", 0))
                        if final_major == "消費税（内税）":
                            # 内税は合計に加算しないが、内訳には実際の税額を表示する
                            category_totals[final_major] = category_totals.get(final_major, 0) + amt
                        else:
                            category_totals[final_major] = category_totals.get(final_major, 0) + amt
                    
                    st.markdown("#### 📋 解析結果の確認")
                    
                    # Convert AI payload date string to proper python object if possible for the calendar
                    parsed_date_val = pd.to_datetime(preview_date).date() if preview_date else datetime.today().date()
                    edited_date = st.date_input("**日付**", value=parsed_date_val, key="receipt_import_date")
                    edited_store = st.text_input("**店舗**", value=preview_store, key="receipt_import_store", placeholder="店舗名を入力してください")
                    st.write(f"**合計金額**: ￥{total_amount:,}")
                    
                    # DataFrameで一覧表示
                    cat_df = pd.DataFrame([
                        {"大分類": k, "金額": f"￥{v:,}"} for k, v in category_totals.items()
                    ])
                    st.dataframe(cat_df, hide_index=True, use_container_width=True)
                    
                    st.markdown("##### 📂 明細の詳細（ブラインド形式）")
                    # 解析結果をブラインド（アコーディオン）形式で表示するために整形
                    results_for_disp = []
                    for item in results:
                        disp_item = item.copy()
                        disp_item["category"] = item.get("major_category", "その他")
                        disp_item["subcategory"] = item.get("minor_category", "その他")
                        try:
                            disp_item["amount"] = safe_money_int_cast(disp_item.get("amount", 0))
                        except ValueError:
                            disp_item["amount"] = 0
                        results_for_disp.append(disp_item)
                        
                    if results_for_disp:
                        disp_df = pd.DataFrame(results_for_disp)
                        display_categories_as_html(disp_df)
                    
                    st.markdown("---")
                    
                    # 🎯 支払い方法のUIを追加
                    methods = get_payment_methods(st.session_state['username'])
                    method_options = [m["name"] for m in methods] if methods else ["現金"]
                    
                    if "未設定" not in method_options:
                        method_options = ["未設定"] + method_options
                    else:
                        # 確実に未設定を先頭あるいは初期選択位置に持っていくために、インデックスを取得
                        pass
                        
                    default_idx = method_options.index("未設定") if "未設定" in method_options else 0
                    selected_payment = st.selectbox("支払い方法", options=method_options, index=default_idx)
                    
                    st.write("この内容で登録しますか？")
                    
                    if st.session_state.get('confirm_unset_payment', False) and selected_payment == "未設定":
                        st.warning("⚠️ 支払い方法が「未設定」のままですが、よろしいですか？ よろしければ再度「登録」ボタンを押してください。")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        submit_btn = st.button("登録", type="primary", use_container_width=True)
                    with col2:
                        cancel_btn = st.button("キャンセル", type="secondary", use_container_width=True)
                        
                    if cancel_btn:
                        st.session_state.parsed_results = None
                        st.session_state.uploader_key += 1
                        st.session_state.confirm_unset_payment = False
                        st.rerun()
                        
                    if submit_btn:
                        if selected_payment == "未設定" and not st.session_state.get('confirm_unset_payment', False):
                            st.session_state.confirm_unset_payment = True
                            st.rerun()
                        else:
                            st.session_state.confirm_unset_payment = False
                            edited_date_str = str(edited_date).strip() if edited_date else ""
                        edited_store_str = str(edited_store).strip() if edited_store else ""
                        
                        if not edited_date_str:
                            st.error("エラー: 日付を入力してください。")
                        elif not edited_store_str:
                            st.error("エラー: 店舗名を入力してください。")
                        else:
                            try:
                                with st.spinner("保存中..."):
                                    sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                    init_transactions_sheet(sheet)
                                    
                                    rows_to_append = []
                                    new_receipt_id = datetime.now().strftime("%Y%m%d%H%M%S")
                                    current_user = st.session_state.get('username')
                                    
                                    if not current_user:
                                        st.error("🚨 ログインセッションが切れました。再度ログインしてください。")
                                        st.stop()
                                        
                                    for item in results:
                                        # -----------------------------------------------------------
                                        # セッション整合性チェック (ユーザー名の漏洩防止)
                                        target_username = str(current_user).lower().strip()
                                        # -----------------------------------------------------------
                                        # カテゴリの正規化（14カテゴリ体系に強制）
                                        majors = list(get_categories().keys())
                                        major = str(item.get("major_category", "その他"))
                                        final_major = "その他"
                                        for m in majors:
                                            if m in major or major in m:
                                                final_major = m
                                                break
                                                
                                        minors = get_categories().get(final_major, get_categories()["その他"])
                                        minor = str(item.get("minor_category", "❓その他"))
                                        final_minor = minors[-1] if minors else "❓その他"
                                        for m in minors:
                                            text_only = "".join([c for c in m if c.isalnum() or c in "類物食品未分類その他%"])
                                            if text_only and (text_only in minor or minor in text_only) and len(minor) > 0:
                                                final_minor = m
                                                break
                                                
                                        store_name = str(edited_store).strip()
                                        item_name = str(item.get("item_name", ""))
                                        amt = safe_money_int_cast(item.get("amount", 0))
                                        
                                        # 日付を yyyy-mm-dd に整形
                                        formatted_date = edited_date.strftime("%Y-%m-%d") if edited_date else ""
    
                                        p_type, p_close, p_month, p_date = get_payment_details_for_transaction(target_username, selected_payment)
                                        row_data = [
                                            target_username,
                                            formatted_date,
                                            str(store_name),
                                            str(item_name),
                                            str(final_major),
                                            str(final_minor),
                                            amt,
                                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                            str(selected_payment),
                                            str(p_type),
                                            str(p_close),
                                            str(p_month),
                                            str(p_date),
                                            new_receipt_id,
                                            str(target_username) # memo column: Set account only for receipt OCR
                                        ]
                                        rows_to_append.append(row_data)
                                    
                                    if rows_to_append:
                                        safe_gspread_call(sheet.append_rows, rows_to_append)
                                        written_count = len(rows_to_append)
                                    
                                    st.session_state.flash_message = f"✅ 解析が完了し、{written_count}件のデータを保存しました！"
                                    
                                    time.sleep(1)
                                    
                                    st.session_state.parsed_results = None
                                    st.session_state.uploader_key += 1
                                    st.rerun()
                                    
                            except Exception as e:
                                st.error(f"保存エラー: {e}")

        elif menu_selection == "レシート手入力":
            st.markdown("#### 📝 レシート手入力")
            
            # セッション状態で入力を管理
            if 'manual_input_form_id' not in st.session_state:
                st.session_state.manual_input_form_id = 0
            if 'manual_input_items' not in st.session_state:
                st.session_state.manual_input_items = [{"id": int(time.time() * 1000), "name": "", "amount": 0}]
            if 'manual_input_date' not in st.session_state:
                st.session_state.manual_input_date = datetime.today()
            if 'manual_input_store' not in st.session_state:
                st.session_state.manual_input_store = ""

            # IME制御とnumber_inputのボタン隠し用CSS
            st.markdown("""
                <style>
                    /* 金額入力欄（number_input）の +/- ボタンを非表示にする */
                    div[data-testid="stNumberInput"] button {
                        display: none !important;
                    }
                    div[data-testid="stNumberInput"] input {
                        ime-mode: disabled !important;
                    }
                    /* ブラウザ標準のスピンボタンも非表示にする */
                    input[type=number]::-webkit-inner-spin-button, 
                    input[type=number]::-webkit-outer-spin-button { 
                        -webkit-appearance: none; 
                        margin: 0; 
                    }
                    input[type=number] {
                        -moz-appearance: textfield;
                    }
                    /* 入力欄全体を極限までコンパクトにする (1/3程度に) */
                    div[data-testid="stTextInput"] input, 
                    div[data-testid="stNumberInput"] input, 
                    div[data-testid="stDateInput"] input {
                        padding: 2px 8px !important;
                        min-height: 28px !important; /* 通常の約1/3を目標に */
                        font-size: 14px !important;
                        line-height: 1.2 !important;
                        border: 1px solid #ccc !important; /* 罫線を追加 */
                        border-radius: 4px !important;
                    }
                    /* 各入力項目のラベルの余白も削る */
                    div[data-testid="stWidgetLabel"] p {
                        font-size: 13px !important;
                        margin-bottom: 2px !important;
                    }
                    /* 1. アプリ全体とフォームの背景を白、文字を黒に強制 */
                    .stApp, [data-testid="stForm"] {
                        background-color: #ffffff !important;
                        color: #000000 !important;
                    }
                    /* 2. スマホ画面（768px以下）でのグリッド強制 */
                    @media (max-width: 768px) {
                        /* 日付・店舗名の段 */
                        div[data-testid="stForm"] > div:nth-child(2) div[data-testid="stHorizontalBlock"] {
                            display: grid !important;
                            grid-template-columns: 4fr 6fr !important;
                            gap: 5px !important;
                        }

                        /* 商品名・金額・削除ボタンの段 */
                        div[data-testid="stForm"] .row-widget.stHorizontalBlock {
                            display: grid !important;
                            grid-template-columns: 5fr 3fr 1fr !important; /* ボタンを最小限に */
                            gap: 4px !important;
                            width: 100% !important;
                        }
                        /* 3. 入力欄の最小幅を強制解除して画面内に収める */
                        div[data-baseweb="input"], div[data-baseweb="base-input"] {
                            min-width: 0 !important;
                            width: 100% !important;
                        }
                        input {
                            padding: 6px 4px !important;
                            font-size: 14px !important;
                        }

                        /* 4. ラベル（日付、店舗名など）の文字も小さく */
                        label p {
                            font-size: 12px !important;
                        }
                    }
                </style>
            """, unsafe_allow_html=True)

            fid = st.session_state.manual_input_form_id
            # st.form をコンテナに変更してリアクティブな挙動を可能にする
            with st.container():
                col_d, col_s = st.columns([4, 6])
                with col_d:
                    # keyを追加してリセット可能にする
                    input_date = st.date_input("日付", value=st.session_state.manual_input_date, key=f"mi_d_{fid}")
                with col_s:
                    # keyを追加してリセット可能にする
                    input_store = st.text_input("店舗名", value=st.session_state.manual_input_store, key=f"mi_s_{fid}", placeholder="店舗名")
                
                st.write("---")
                st.write("**明細入力**")
                
                updated_items = []
                for i, item in enumerate(st.session_state.manual_input_items):
                    c1, c2, c3 = st.columns([5, 3, 1.5])
                    row_id = item.get("id", i) # 互換性のため
                    with c1:
                        iname = st.text_input(f"商品名 {i+1}", value=item["name"], key=f"mi_n_{row_id}_{fid}", label_visibility="collapsed", placeholder="商品名")
                    with c2:
                        # 金額入力時に自動で次の行を追加するコールバック用
                        def add_empty_row_if_last(idx=i):
                            if idx == len(st.session_state.manual_input_items) - 1:
                                # 金額が入力されたら新しい行を追加（IDを付与）
                                new_id = int(time.time() * 1000) + len(st.session_state.manual_input_items)
                                st.session_state.manual_input_items.append({"id": new_id, "name": "", "amount": 0})

                        iamount = st.number_input(f"金額 {i+1}", value=int(item["amount"]), step=1, key=f"mi_a_{row_id}_{fid}", label_visibility="collapsed", on_change=add_empty_row_if_last)
                    with c3:
                        # 削除ボタンに確認フェーズを追加
                        with st.popover("🗑️" if len(st.session_state.manual_input_items) > 1 else "×", disabled=len(st.session_state.manual_input_items) <= 1):
                            st.write("この行を削除しますか？")
                            if st.button("削除実行", key=f"mi_del_manual_{row_id}_{fid}"): 
                                # IDを元に削除対象を特定して削除
                                st.session_state.manual_input_items = [itm for itm in st.session_state.manual_input_items if itm.get("id") != row_id]
                                st.rerun()
                    updated_items.append({"id": row_id, "name": iname, "amount": iamount})
                
                st.session_state.manual_input_items = updated_items
                
                # 🎯 支払い方法のUIを追加
                st.write("---")
                methods = get_payment_methods(st.session_state['username'])
                method_options = [m["name"] for m in methods] if methods else ["現金"]
                
                # 「未設定」を確実に追加し、デフォルトとして選択させる
                if "未設定" not in method_options:
                    method_options = ["未設定"] + method_options
                
                default_manual_idx = method_options.index("未設定") if "未設定" in method_options else 0
                selected_payment_manual = st.selectbox("支払い方法", options=method_options, index=default_manual_idx, key=f"mi_pay_{fid}")
                
                # ボタン類
                st.markdown("<br>", unsafe_allow_html=True) 
                col_btn_l, col_btn_r = st.columns(2)
                with col_btn_l:
                    submit_manual = st.button("登録", use_container_width=True, type="primary", key="submit_manual_input")
                with col_btn_r:
                    cancel_manual = st.button("キャンセル", use_container_width=True, key="cancel_manual_input")
                
                if cancel_manual:
                    # フォームIDを更新して初期状態に戻す
                    st.session_state.manual_input_form_id += 1
                    st.session_state.manual_input_items = [{"id": int(time.time() * 1000), "name": "", "amount": 0}]
                    st.session_state.manual_input_store = ""
                    st.session_state.manual_input_date = datetime.today()
                    st.rerun()

                if submit_manual:
                    # 入力されているデータのみを抽出（商品名が入力されているもの。0円のレシートも許容）
                    valid_items = [itm for itm in st.session_state.manual_input_items if itm["name"].strip() != ""]

                    if not input_store:
                        st.error("店舗名を入力してください。")
                    elif not valid_items:
                        st.error("少なくとも1件以上の有効なデータを入力してください。")
                    else:
                        with st.spinner("AIがカテゴリを判定中..."):
                            # 登録ボタン押下時に有効な明細のみ解析を実行
                            item_names = [itm["name"] for itm in valid_items]
                            categories = categorize_items_with_ai(item_names, input_store)
                            
                            try:
                                sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                init_transactions_sheet(sheet)
                                
                                rows_to_append = []
                                new_receipt_id = datetime.now().strftime("%Y%m%d%H%M%S")
                                current_user = st.session_state.get('username')
                                if not current_user:
                                    st.error("🚨 ログインセッションが切れました。再度ログインしてください。")
                                    st.stop()
                                
                                target_username = str(current_user).lower().strip()
                                
                                for itm, cat in zip(valid_items, categories):
                                    major = cat.get("major_category", "その他")
                                    minor = cat.get("minor_category", "📁未分類")
                                    amt = safe_money_int_cast(itm.get("amount", 0))
                                    
                                    p_type, p_close, p_month, p_date = get_payment_details_for_transaction(target_username, selected_payment_manual)
                                    row_data = [
                                        target_username,
                                        str(input_date.strftime('%Y-%m-%d')),
                                        str(input_store),
                                        str(itm["name"]),
                                        str(major),
                                        str(minor),
                                        amt,
                                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        str(selected_payment_manual),
                                        str(p_type),
                                        str(p_close),
                                        str(p_month),
                                        str(p_date),
                                        new_receipt_id,
                                        str(target_username) # memo column: Set account for manual input
                                    ]
                                    rows_to_append.append(row_data)
                                
                                if rows_to_append:
                                    safe_gspread_call(sheet.append_rows, rows_to_append)
                                
                                st.success(f"✅ {len(rows_to_append)}件のデータを登録しました！")
                                # フォームIDを更新して全ウィジェットを強制リセット
                                st.session_state.manual_input_form_id += 1
                                st.session_state.manual_input_items = [{"id": int(time.time() * 1000), "name": "", "amount": 0}]
                                st.session_state.manual_input_store = ""
                                st.session_state.manual_input_date = datetime.today()
                                
                                time.sleep(1.5)
                                st.rerun()
                            except Exception as e:
                                st.error(f"登録エラー: {e}")
        elif menu_selection == "レシート修正":
            st.markdown("#### ⚙️ レシート修正")
            
            # 共通ナビゲーションの適用
            df = render_month_navigation()
                
            if df.empty:
                st.info("※この月のデータはありません。")
            else:
                # 店舗名と商品名のカラムを動的に判定
                store_col = "store_name" if "store_name" in df.columns else "store" if "store" in df.columns else None
                item_col = "item_name" if "item_name" in df.columns else "item" if "item" in df.columns else "items" if "items" in df.columns else None
                
                if not store_col:
                    st.warning("スプレッドシートに店舗名（'store_name' または 'store'）の列が見つかりません。")
                else:
                    if "receipt_id" not in df.columns:
                        df["receipt_id"] = "不明_" + df["date"].dt.strftime('%Y%m%d') + "_" + df[store_col].astype(str)
                    empty_mask = df["receipt_id"].isna() | (df["receipt_id"] == "")
                    if empty_mask.any():
                        fallback_ids = "不明_" + df["date"].dt.strftime('%Y%m%d') + "_" + df[store_col].astype(str)
                        if "payment_method" in df.columns:
                            fallback_ids += "_" + df["payment_method"].astype(str)
                        df.loc[empty_mask, "receipt_id"] = fallback_ids[empty_mask]

                    # レシート単位に集約（内税を金額から除外して集計）
                    df_agg = df.copy()
                    if "category" in df_agg.columns:
                        df_agg.loc[df_agg["category"] == "消費税（内税）", "amount"] = 0
                        
                    if "payment_method" not in df_agg.columns:
                        df_agg["payment_method"] = "未設定"
                    else:
                        df_agg["payment_method"] = df_agg["payment_method"].fillna("未設定")
                    
                    receipts_df = df_agg.groupby(["receipt_id"], as_index=False).agg(
                        date=("date", "first"),
                        store_name=(store_col, "first"),
                        payment_method=("payment_method", "first"),
                        amount=("amount", "sum"),
                        明細数=("amount", "count")
                    )
                    receipts_df.columns = ["receipt_id", "日付", "店舗名", "支払い方法", "金額", "明細数"]
                    # 店舗名が空欄の場合は「店舗不明」とする
                    receipts_df["店舗名"] = receipts_df["店舗名"].replace("", "店舗不明")
                    receipts_df["日付"] = receipts_df["日付"].dt.strftime('%Y-%m-%d')
                    receipts_df["金額"] = receipts_df["金額"].apply(lambda x: int(x))
                    receipts_df = receipts_df.sort_values(by="日付", ascending=False).reset_index(drop=True)
                    
                    if "receipt_list_version" not in st.session_state:
                        st.session_state.receipt_list_version = 0

                    st.markdown("<p style='font-size: 0.85rem; font-weight: bold; margin-bottom: 5px;'>レシート一覧表（対象レシートを選択してください）</p>", unsafe_allow_html=True)
                    
                    # dataframe 選択
                    event = st.dataframe(
                        receipts_df, 
                        use_container_width=True, 
                        hide_index=True, 
                        selection_mode="single-row",
                        on_select="rerun",
                        column_config={
                            "receipt_id": None, # IDは非表示
                            "店舗名": st.column_config.TextColumn("店舗名", width="medium"),
                        },
                        key=f"receipt_list_df_{st.session_state.receipt_list_version}"
                    )
                    
                    if len(event.selection.rows) > 0:
                        selected_idx = event.selection.rows[0]
                        sel_rec = receipts_df.iloc[selected_idx]
                        if sel_rec["日付"] != "総合計":
                            # 選択されたレシート情報をセッションに保存して永続化
                            st.session_state['selected_receipt_info'] = {
                                "receipt_id": sel_rec["receipt_id"],
                                "date": sel_rec["日付"],
                                "store": sel_rec["店舗名"]
                            }
                    
                    # セッションに保存された情報に基づいて詳細を表示（表の選択が消えても維持）
                    receipt_info = st.session_state.get('selected_receipt_info')
                    if receipt_info:
                        # 表示中のリストにまだ存在するか確認（削除対策）
                        target_receipt_id = receipt_info.get("receipt_id", "")
                        
                        selected_receipt_matches = receipts_df[
                            (receipts_df["receipt_id"] == target_receipt_id)
                        ]
                        
                        if not selected_receipt_matches.empty:
                            selected_receipt = selected_receipt_matches.iloc[0]
                            target_date = pd.to_datetime(selected_receipt["日付"])
                            target_store = selected_receipt["店舗名"]
                            
                            st.markdown("---")
                            st.write(f"##### 対象レシート明細： {selected_receipt['日付']} - {target_store}")
                            
                            # 該当レシートの明細を取得
                            details = df[df["receipt_id"] == target_receipt_id].copy()
                            
                            receipt_key = f"{target_receipt_id}"
                            if st.session_state.get('current_receipt_key') != receipt_key or st.session_state.get('edit_data') is None:
                                st.session_state['current_receipt_key'] = receipt_key
                                st.session_state['edit_data'] = {}
                                first_payment = details.iloc[0].get("payment_method", "現金") if not details.empty else "現金"
                                st.session_state['edit_header'] = {
                                    "receipt_id": target_receipt_id,
                                    "date": target_date.date(),
                                    "store": target_store,
                                    "payment_method": first_payment
                                }
                                st.session_state['editing_gs_idx'] = None
                                st.session_state['new_row_count'] = 0
                                # 追加: レシート切り替え時に個別項目の編集状態もリセット
                                if "item_list_version" not in st.session_state:
                                    st.session_state.item_list_version = 0
                                st.session_state.item_list_version += 1
                            
                            for idx, row in details.iterrows():
                                row_index_gs = row["_row_index"]
                                if row_index_gs not in st.session_state['edit_data']:
                                    major = row.get("category", "その他")
                                    sub_cols = [c for c in ["subcategory", "sub_category", "小分類"] if c in df.columns]
                                    sub = row.get(sub_cols[0], "❓その他") if sub_cols else "❓その他"
                                    payment_m = row.get("payment_method", "現金")
                                    st.session_state['edit_data'][row_index_gs] = {
                                        "name": row.get(item_col, "不明な商品") if item_col else "不明な商品",
                                        "amount": safe_money_int_cast(row.get("amount", 0)),
                                        "major": major,
                                        "minor": sub,
                                        "payment_method": payment_m
                                    }

                            # 編集中の状態を事前に取得してヘッダーをロックするか判定
                            # --- レシートヘッダー（日付・店舗名）の修正エリア用プレースホルダー ---
                            header_placeholder = st.container()

                            # -- 明細データ構築 --
                            action_placeholder = st.container()

                            st.markdown("<br>", unsafe_allow_html=True)

                            # --- 修正用のDataFrameを作成 (閲覧・選択用) ---
                            edit_items_list = []
                            for row_id, data in st.session_state['edit_data'].items():
                                edit_items_list.append({
                                    "大分類": data["major"],
                                    "小分類": data["minor"],
                                    "商品名": data["name"],
                                    "金額": data["amount"],
                                    "支払い方法": data.get("payment_method", "現金"),
                                    "_id": row_id
                                })
                            edit_df_display = pd.DataFrame(edit_items_list)
                            # 指定された順序でソート（大分類、小分類、商品名）
                            edit_df_display = edit_df_display.sort_values(by=["大分類", "小分類", "商品名"]).reset_index(drop=True)

                            # --- 明細一覧表の表示 (選択用) ---
                            st.write("##### 明細一覧（修正行を選択して下さい）")
                            item_event = st.dataframe(
                                edit_df_display.drop(columns=["_id"]),
                                use_container_width=True,
                                hide_index=True,
                                selection_mode="single-row",
                                column_config={
                                    "大分類": st.column_config.TextColumn(width="small"),
                                    "小分類": st.column_config.TextColumn(width="small"),
                                    "商品名": st.column_config.TextColumn(width="medium"),
                                    "金額": st.column_config.NumberColumn(width="small", format="￥%d"),
                                    "支払い方法": st.column_config.TextColumn(width="small")
                                },
                                on_select="rerun",
                                key=f"item_edit_df_{st.session_state.item_list_version}"
                            )

                            # 選択された行のIDを特定
                            current_editing_id = st.session_state.get('editing_gs_idx')
                            
                            # データフレームでの選択を優先
                            if len(item_event.selection.rows) > 0:
                                row_idx = item_event.selection.rows[0]
                                current_editing_id = edit_df_display.iloc[row_idx]["_id"]
                                st.session_state['editing_gs_idx'] = current_editing_id

                            # 削除済みIDのチェック
                            if current_editing_id and current_editing_id not in st.session_state['edit_data']:
                                current_editing_id = None
                                st.session_state['editing_gs_idx'] = None

                            st.markdown("<br>", unsafe_allow_html=True)
                            
                            # --- 行追加ボタン ---
                            if st.button("➕ 明細を追加する", use_container_width=True):
                                st.session_state['new_row_count'] += 1
                                new_id = f"new_{st.session_state['new_row_count']}"
                                st.session_state['edit_data'][new_id] = {
                                    "name": "",
                                    "amount": 0,
                                    "major": "その他",
                                    "minor": "❓その他",
                                    "payment_method": "現金"
                                }
                                st.session_state['editing_gs_idx'] = new_id
                                st.session_state.item_list_version += 1
                                st.rerun()

                            # -----------------------------------------------------------
                            # ここで最新の current_editing_id 状態を使って is_item_editing を判定
                            is_item_editing = bool(current_editing_id)

                            # 遅延描画したプレースホルダーにヘッダーとボタンを出力
                            with header_placeholder:
                                st.write(f"##### レシート修正（金額：￥{int(selected_receipt['金額']):,}）")
                                with st.container(border=True):
                                    # 1行目：日付
                                    c1, c2, c3 = st.columns([1.5, 8.5, 2])
                                    with c1:
                                        st.markdown("<p style='margin-top: 8px; font-weight: bold;'>日付</p>", unsafe_allow_html=True)
                                    with c2:
                                        new_date = st.date_input("日付", value=st.session_state['edit_header']['date'], key=f"edit_header_date_{receipt_key}", label_visibility="collapsed", disabled=is_item_editing)
                                    
                                    # 2行目：店舗名
                                    c1, c2, c3 = st.columns([1.5, 8.5, 2])
                                    with c1:
                                        st.markdown("<p style='margin-top: 8px; font-weight: bold;'>店舗名</p>", unsafe_allow_html=True)
                                    with c2:
                                        new_store = st.text_input("店舗名", value=st.session_state['edit_header']['store'], key=f"edit_header_store_{receipt_key}", label_visibility="collapsed", disabled=is_item_editing)
                                    
                                    # 3行目：支払い方法
                                    c1, c2, c3 = st.columns([1.5, 8.5, 2])
                                    with c1:
                                        st.markdown("<p style='margin-top: 8px; font-weight: bold;'>支払い方法</p>", unsafe_allow_html=True)
                                    with c2:
                                        methods = get_payment_methods(st.session_state['username'])
                                        method_options = [m["name"] for m in methods] if methods else ["現金"]
                                        
                                        # 「未設定」を先頭に追加（重複を避ける）
                                        if "未設定" not in method_options:
                                            method_options = ["未設定"] + method_options

                                        current_payment = str(st.session_state['edit_header'].get('payment_method', '未設定')).strip()
                                        if not current_payment or str(current_payment).lower() == 'nan':
                                            current_payment = '未設定'
                                            
                                        if current_payment not in method_options:
                                            method_options.append(current_payment)
                                            
                                        payment_idx = method_options.index(current_payment)
                                        new_payment = st.selectbox("支払い方法", method_options, index=payment_idx, key=f"edit_header_payment_{receipt_key}", label_visibility="collapsed", disabled=is_item_editing)
                                    
                                    # ヘッダー情報を更新
                                    st.session_state['edit_header']['date'] = new_date
                                    st.session_state['edit_header']['store'] = new_store
                                    st.session_state['edit_header']['payment_method'] = new_payment

                            with action_placeholder:
                                # --- アクションボタンエリア（上部） ---
                                action_col1, action_col2 = st.columns(2)
                                
                                with action_col1:
                                    if st.button("日付・店舗名・支払更新", use_container_width=True, type="primary", disabled=is_item_editing):
                                        try:
                                            with st.spinner("一括更新中..."):
                                                sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                                target_date_str = st.session_state['edit_header']['date'].strftime("%Y-%m-%d")
                                                target_store = st.session_state['edit_header']['store']
                                                target_payment = st.session_state['edit_header']['payment_method']
                                                
                                                # 既存の全明細行をループして日付と店舗を更新
                                                existing_indices = [int(k) for k in st.session_state['edit_data'].keys() if not str(k).startswith("new_")]
                                                current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                
                                                p_type, p_close, p_month, p_date = get_payment_details_for_transaction(st.session_state['username'], target_payment)
                                                batch_updates = []
                                                for r_idx in existing_indices:
                                                    batch_updates.append({"range": f"B{r_idx}:C{r_idx}", "values": [[target_date_str, target_store]]})
                                                    batch_updates.append({"range": f"H{r_idx}:M{r_idx}", "values": [[current_time, target_payment, str(p_type), str(p_close), str(p_month), str(p_date)]]})
                                                    
                                                if batch_updates:
                                                    safe_gspread_call(sheet.batch_update, batch_updates)
                                                    
                                                st.success("✅ レシート情報を一括更新しました")
                                                st.session_state.receipt_list_version += 1
                                                time.sleep(1)
                                                st.rerun()
                                        except Exception as e:
                                            st.error(f"更新エラー: {e}")
                                
                                with action_col2:
                                    with st.popover("このレシートを全削除", use_container_width=True, disabled=is_item_editing):
                                        st.warning("このレシート（全明細）を完全に削除します。")
                                        if st.button("レシート削除を実行", use_container_width=True, type="primary"):
                                            try:
                                                with st.spinner("削除中..."):
                                                    sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                                    user_name = st.session_state['username']
                                                    existing_indices = [int(k) for k in st.session_state['edit_data'].keys() if not str(k).startswith("new_")]
                                                    
                                                    # 削除実行
                                                    for r_idx in sorted(existing_indices, reverse=True):
                                                        # 削除直前の安全チェック
                                                        current_row_values = sheet.row_values(r_idx)
                                                        if len(current_row_values) < 1 or current_row_values[0].lower() != user_name.lower():
                                                            st.error(f"🚨 エラー: 行 {r_idx} の削除中に不整合を検知しました。処理を中断します。リロードしてください。")
                                                            st.stop()
                                                        sheet.delete_rows(r_idx)
                                                    st.success("✅ レシートを削除しました")
                                                    st.session_state.receipt_list_version += 1
                                                    st.session_state['edit_data'] = None
                                                    st.session_state['selected_receipt_info'] = None # レシートごと消えたのでクリア
                                                    st.session_state['editing_gs_idx'] = None
                                                    time.sleep(1)
                                                    st.rerun()
                                            except Exception as e:
                                                st.error(f"削除エラー: {e}")
                            # -----------------------------------------------------------

                            # --- 個別修正フォーム (選択されている場合のみ表示) ---
                            if current_editing_id:
                                st.markdown("---")
                                st.write("##### 選択中の明細を修正")
                                target_item = st.session_state['edit_data'][current_editing_id]
                                
                                with st.container(border=True):
                                    categories = get_categories()
                                    major_cats = list(categories.keys())
                                    
                                    edit_name = st.text_input("商品名", value=target_item["name"], key=f"edit_name_{receipt_key}_{current_editing_id}")
                                    edit_amount = st.number_input("金額", value=int(target_item["amount"]), step=1, key=f"edit_amount_{receipt_key}_{current_editing_id}")
                                    
                                    # --- 支払い方法のUIを削除し、大分類と小分類のみに変更 ---
                                    
                                    # 大分類
                                    current_major = target_item["major"]
                                    if current_major not in major_cats: current_major = "その他"
                                    edit_major = st.selectbox("大分類", options=major_cats, index=major_cats.index(current_major), key=f"edit_major_{receipt_key}_{current_editing_id}")
                                    
                                    # 小分類
                                    minor_cats = categories.get(edit_major, ["❓その他"])
                                    current_minor = target_item["minor"]
                                    if current_minor not in minor_cats: current_minor = minor_cats[0]
                                    edit_minor = st.selectbox("小分類", options=minor_cats, index=minor_cats.index(current_minor), key=f"edit_minor_{receipt_key}_{current_editing_id}")
                                    
                                    st.markdown("<br>", unsafe_allow_html=True)
                                    
                                    b_col1, b_col2, b_col3 = st.columns(3)
                                    with b_col1:
                                        if st.button("登録実行", use_container_width=True, type="primary", key=f"save_btn_{current_editing_id}"):
                                            if not edit_name.strip():
                                                st.warning("⚠️ 商品名を入力してください。")
                                            # 0円の場合の警告（elif edit_amount == 0:）を削除し、0円を許容する
                                            else:
                                                try:
                                                    with st.spinner("保存中..."):
                                                        sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                                        user_name = st.session_state['username']
                                                        # 修正されたヘッダー情報を使用
                                                        target_date_str = st.session_state['edit_header']['date'].strftime("%Y-%m-%d")
                                                        target_store = st.session_state['edit_header']['store']
                                                        
                                                        target_payment = st.session_state['edit_header']['payment_method']
                                                        p_type, p_close, p_month, p_date = get_payment_details_for_transaction(user_name, target_payment)
                                                        if str(current_editing_id).startswith("new_"):
                                                            # 新規追加
                                                            new_row = [user_name, target_date_str, target_store, edit_name, edit_major, edit_minor, edit_amount, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), target_payment, str(p_type), str(p_close), str(p_month), str(p_date), st.session_state['edit_header'].get('receipt_id', ''), str(user_name)]
                                                            sheet.append_row(new_row)
                                                        else:
                                                            # 既存更新: 安全装置（行データの検証）
                                                            r_idx = int(current_editing_id)
                                                            
                                                            # 書き込み直前に、その行が本当に正しいか確認する
                                                            # (スプレッドシートの行がずれている可能性があるため)
                                                            current_row_values = sheet.row_values(r_idx)
                                                            # ヘッダーを除いたデータ行(2行目以降)であることを確認しつつ、ユーザー名が一致するかチェック
                                                            if len(current_row_values) < 1 or current_row_values[0].lower() != user_name.lower():
                                                                st.error("🚨 エラー: スプレッドシートの行が同期されていません。一度画面をリロードしてやり直してください。")
                                                                st.stop()
                                                            
                                                            # バッチ更新（1回のAPI呼び出しで範囲を更新）
                                                            # A:username, B:date, C:store, D:item, E:major, F:minor, G:amount, H:update, I:payment_method, J:type, K:close_date, L:pay_month, M:pay_date, N:receipt_id
                                                            # 更新範囲: B (Col 2) から N (Col 14)
                                                            update_range = f"B{r_idx}:N{r_idx}"
                                                            update_values = [[target_date_str, target_store, edit_name, edit_major, edit_minor, edit_amount, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), target_payment, str(p_type), str(p_close), str(p_month), str(p_date), st.session_state['edit_header'].get('receipt_id', '')]]
                                                            sheet.update(range_name=update_range, values=update_values)
                                                        
                                                        st.success("✅ 修正を登録しました")
                                                        st.session_state['editing_gs_idx'] = None
                                                        st.session_state['edit_data'] = None
                                                        st.session_state.item_list_version += 1
                                                        st.session_state.receipt_list_version += 1
                                                        time.sleep(1)
                                                        st.rerun()
                                                except Exception as e:
                                                    st.error(f"エラー: {e}")
                                    
                                    with b_col2:
                                        with st.popover("明細を削除", use_container_width=True):
                                            st.warning("この明細を完全に削除します。よろしいですか？")
                                            if st.button("明細を削除", use_container_width=True, type="primary", key=f"del_item_{current_editing_id}"):
                                                try:
                                                    if not str(current_editing_id).startswith("new_"):
                                                        sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                                        user_name = st.session_state['username']
                                                        r_idx = int(current_editing_id)
                                                        
                                                        # 削除直前の安全チェック
                                                        current_row_values = sheet.row_values(r_idx)
                                                        if len(current_row_values) < 1 or current_row_values[0].lower() != user_name.lower():
                                                            st.error("🚨 エラー: 削除対象の行を特定できませんでした。データが移動している可能性があります。リロード後に再度お試しください。")
                                                            st.stop()
                                                        
                                                        sheet.delete_rows(r_idx)
                                                    
                                                    del st.session_state['edit_data'][current_editing_id]
                                                    st.success("🗑️ 明細を削除しました")
                                                    st.session_state['editing_gs_idx'] = None
                                                    st.session_state['edit_data'] = None
                                                    st.session_state.item_list_version += 1
                                                    st.session_state.receipt_list_version += 1
                                                    time.sleep(1)
                                                    st.rerun()
                                                except Exception as e:
                                                    st.error(f"削除エラー: {e}")
                                    
                                    with b_col3:
                                        if st.button("キャンセル", use_container_width=True, key=f"cancel_btn_{current_editing_id}"):
                                            st.session_state['editing_gs_idx'] = None
                                            st.session_state.item_list_version += 1
                                            st.rerun()

                            # CSSの代わりにJSを使ってより確実にボタンの色を変更
                            # JSによるボタン色変更を削除
                            pass

                            
        elif menu_selection == "AI相談":
            st.markdown("#### 👁AI相談（専属ファイナンシャルプランナー）")
            st.info("あなたの家計簿データに基づいて、AIが分析やアドバイスを行います。")
            
            # --- データの準備（全期間からログインユーザー分のみ抽出） ---
            @st.cache_data(ttl=300)
            def get_user_data_csv_for_ai(username):
                # ユーザーに関連する3つのシートのデータをすべて取得して結合する
                result_parts = []
                
                # ヘルパー: 指定シートからログインユーザーの全行・全列をCSVテキストで取得
                def fetch_sheet_csv(sheet_name, title):
                    try:
                        sheet = get_sheet(sheet_name, create_if_not_found=True)
                        if sheet_name == USER_MASTER_WORKSHEET_NAME:
                            init_user_master_sheet(sheet)
                        elif sheet_name == PAYMENT_MASTER_WORKSHEET_NAME:
                            init_payment_master_sheet(sheet)
                        else:
                            init_transactions_sheet(sheet)
                            
                        values = safe_gspread_call(sheet.get_all_values)
                        if not values or len(values) < 2:
                            return f"--- {title} ---\nデータなし\n"
                            
                        headers = [h.strip() if h.strip() else f"empty_{i}" for i, h in enumerate(values[0])]
                        df_all = pd.DataFrame(values[1:])
                        if df_all.shape[1] > len(headers):
                            headers += [f"extra_{i}" for i in range(len(headers), df_all.shape[1])]
                        df_all.columns = headers[:df_all.shape[1]]
                        
                        if "username" in df_all.columns:
                            df_user = df_all[df_all["username"].astype(str).str.lower() == username.lower()].copy()
                        else:
                            df_user = pd.DataFrame()
                            
                        if df_user.empty:
                            return f"--- {title} ---\nデータなし\n"
                            
                        # 不要なパスワードハッシュ等が含まれている場合は念のため除外（User_Master等には通常ないが念の為）
                        if "password_hash" in df_user.columns:
                            df_user = df_user.drop(columns=["password_hash"])
                            
                        return f"--- {title} ---\n{df_user.to_csv(index=False)}\n"
                    except Exception as e:
                        return f"--- {title} ---\nデータ取得エラー: {e}\n"

                result_parts.append(fetch_sheet_csv(TRANSACTIONS_WORKSHEET_NAME, "支出データ"))
                result_parts.append(fetch_sheet_csv(USER_MASTER_WORKSHEET_NAME, "ユーザープロフィール"))
                result_parts.append(fetch_sheet_csv(PAYMENT_MASTER_WORKSHEET_NAME, "支払方法マスター"))
                
                return "\n".join(result_parts)
                
            # メニュー切り替え直後、または手動更新時はキャッシュをクリア
            if st.session_state.get("refresh_ai_data_flag"):
                get_user_data_csv_for_ai.clear()
                st.session_state["refresh_ai_data_flag"] = False

            csv_data_string = get_user_data_csv_for_ai(st.session_state['username'])
            if not csv_data_string:
                csv_data_string = "現在、参照できる家計簿データはありません。"

            # --- チャットセッションとメッセージ履歴の初期化 ---
            if "ai_consult_messages" not in st.session_state:
                st.session_state.ai_consult_messages = []
            
            # --- チャット履歴の初期化 (SDKが期待するリスト形式) ---
            if "ai_consult_chat_history" not in st.session_state:
                st.session_state.ai_consult_chat_history = []

            # 最初のメッセージを追加（履歴が空の場合）
            if not st.session_state.ai_consult_messages:
                st.session_state.ai_consult_messages.append({
                    "role": "assistant", 
                    "content": f"こんにちは、{st.session_state['username']}さん！あなたの専属FPです。全期間のデータを読み込みました。何でも相談してくださいね。"
                })

            # 履歴の表示
            for i, msg in enumerate(st.session_state.ai_consult_messages):
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
                    if msg["role"] == "assistant":
                        render_speech_synthesis_button(msg["content"], f"ai_{i}")
            
            # 音声入力ボタン表示
            render_voice_input_button("ai_consult")

            # ユーザー入力
            if user_input := st.chat_input("質問を入力してください..."):
                st.session_state.ai_consult_messages.append({"role": "user", "content": user_input})
                with st.chat_message("user"):
                    st.markdown(user_input)
                    
                client = st.session_state.get('genai_client')
                if not client:
                    with st.chat_message("assistant"):
                        st.error("APIキーが設定されていないため、相談を開始できません。")
                else:
                    with st.chat_message("assistant"):
                        message_placeholder = st.empty()
                        message_placeholder.markdown("分析中...")
                        
                        try:
                            # ユーザーマスターからプロファイルを取得
                            user_profile = get_user_master_data(st.session_state['username'])
                            profile_prompt = ""
                            if user_profile:
                                name = user_profile.get("name", "ユーザー")
                                gender = user_profile.get("gender", "未設定")
                                birthdate = user_profile.get("birthdate", "未設定")
                                mbti = user_profile.get("mbti", "未設定")
                                occupation = user_profile.get("occupation", "未設定")
                                hobbies = user_profile.get("hobbies", "未設定")
                                life_stance = user_profile.get("life_stance", "未設定")
                                ai_base_instruction = user_profile.get("ai_base_instruction", "")
                                
                                profile_prompt = f"""
【ユーザーのプロフィール情報】
名前: {name}
性別: {gender}
生年月日: {birthdate}
MBTI: {mbti}
職業: {occupation}
趣味: {hobbies}
ライフスタンス（大切にしていること）: {life_stance}

【AI相談への基本指示】
{ai_base_instruction if ai_base_instruction else "（特になし）"}

あなたは、ユーザー（{name}さん）の属性や趣味、職業的背景を理解した上で、単なる数字の増減だけでなく、その人の人生の質を上げるための家計アドバイスを行う親身なコンサルタントです。
"""
                            else:
                                profile_prompt = f"""
あなたはユーザー専属の優秀なファイナンシャルプランナーです。
"""

                            # システムプロンプトを都度構築（最新データを反映させるため）
                            system_prompt = f"""{profile_prompt}
以下のCSVデータは、このユーザー（{st.session_state['username']}）個人の家計簿データです。
このデータには「商品名」も含まれており、いつ、どこで、何を買ったかを詳細に把握できます。
ユーザーからの「特定の商品の購入時期（例：鶏肉ナンコツはいつ買った？）」や「商品の価格推移」などの質問に対し、正確かつ親身に答えてください。
データに存在しない推測は避け、無駄遣いの指摘や節約のアドバイスなども積極的に行ってください。

【ユーザーの家計簿データ】
{csv_data_string}
"""
                            
                            # 送信直前でチャットオブジェクトを「履歴付き」で作成
                            chat = client.chats.create(
                                model='gemini-2.5-flash',
                                config=types.GenerateContentConfig(system_instruction=system_prompt),
                                history=st.session_state.ai_consult_chat_history
                            )
                            
                            # 429等のエラーハンドリングを日本語化
                            def _send():
                                return chat.send_message(user_input)
                            
                            try:
                                response = safe_gemini_call(_send)
                                response_text = response.text
                                message_placeholder.markdown(response_text)
                                
                                # 履歴を更新（画面表示用）
                                st.session_state.ai_consult_messages.append({"role": "assistant", "content": response_text})
                                
                                # 最新の回答にも読み上げボタンを表示
                                render_speech_synthesis_button(response_text, "ai_latest")
                                
                                # SDKの履歴をセッションに同期
                                st.session_state.ai_consult_chat_history = chat.get_history()
                                
                            except Exception as e:
                                err_msg = str(e)
                                if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                                    friendly_err = "現在AIの通信が混み合っています。数十秒待ってから再度送信してください。"
                                    st.warning(friendly_err)
                                else:
                                    st.error(f"エラーが発生しました: {e}")

                        except Exception as e:
                            st.error(f"予期せぬエラーが発生しました: {e}")

        elif menu_selection == "プロフィール設定":
            st.markdown("#### ⚙️ プロフィール設定")
            st.info("AI相談でよりパーソナライズされたアドバイスを受けるための基本情報を設定します。")
            
            # 既存データの取得
            user_profile = get_user_master_data(st.session_state['username'])
            if user_profile is None:
                user_profile = {}
                
            with st.form("profile_settings_form"):
                st.write("各項目を入力・編集して「保存する」ボタンを押してください。")
                
                col1, col2 = st.columns(2)
                with col1:
                    name_input = st.text_input("氏名", value=user_profile.get("name", ""))
                    
                    gender_options = ["未設定", "男性", "女性", "その他"]
                    current_gender = user_profile.get("gender", "未設定")
                    if current_gender not in gender_options: current_gender = "未設定"
                    gender_idx = gender_options.index(current_gender)
                    gender_input = st.selectbox("性別", options=gender_options, index=gender_idx)
                    
                    mbti_options = ["未設定", "INTJ", "INTP", "ENTJ", "ENTP", "INFJ", "INFP", "ENFJ", "ENFP", "ISTJ", "ISFJ", "ESTJ", "ESFJ", "ISTP", "ISFP", "ESTP", "ESFP"]
                    current_mbti = user_profile.get("mbti", "未設定")
                    if current_mbti not in mbti_options: current_mbti = "未設定"
                    mbti_idx = mbti_options.index(current_mbti)
                    mbti_input = st.selectbox("MBTI", options=mbti_options, index=mbti_idx)
                
                with col2:
                    try:
                        if user_profile.get("birthdate"):
                            birthdate_val = datetime.strptime(user_profile.get("birthdate", "1990-01-01"), "%Y-%m-%d").date()
                        else:
                            birthdate_val = datetime(1990, 1, 1).date()
                    except:
                        birthdate_val = datetime(1990, 1, 1).date()
                        
                    birthdate_input = st.date_input("生年月日", value=birthdate_val, min_value=datetime(1900, 1, 1), max_value=datetime.today())
                    occupation_input = st.text_input("職業", value=user_profile.get("occupation", ""))
                
                st.markdown("---")
                hobbies_input = st.text_area("趣味リスト", value=user_profile.get("hobbies", ""), placeholder="例：映画鑑賞、ドライブ、カフェ巡り...")
                life_stance_input = st.text_area("ライフスタンス（大切にしていること）", value=user_profile.get("life_stance", ""), placeholder="例：自己投資を惜しまない、健康第一、家族との時間を大切にする...", height=300)
                ai_base_instruction_input = st.text_area("AI相談の基本指示", value=user_profile.get("ai_base_instruction", ""), placeholder="例：回答はまず簡潔な結論から述べて。語尾に「だワン」をつけて。節約には厳しめにアドバイスして。", height=300)
                
                submit_button = st.form_submit_button("保存する", type="primary")
                
                if submit_button:
                    with st.spinner("保存中..."):
                        profile_data_to_save = {
                            "name": name_input,
                            "gender": gender_input,
                            "birthdate": birthdate_input.strftime("%Y-%m-%d") if birthdate_input else "",
                            "mbti": mbti_input,
                            "occupation": occupation_input,
                            "hobbies": hobbies_input,
                            "life_stance": life_stance_input,
                            "ai_base_instruction": ai_base_instruction_input
                        }
                        
                        success, message = save_user_master_data(st.session_state['username'], profile_data_to_save)
                        
                        if success:
                            # AIが新しいプロフィールを使うようキャッシュクリア
                            get_user_master_data.clear()
                            st.success(message)
                        else:
                            st.error(message)

            st.markdown("---")
            st.markdown("#### 🔐 パスワードの変更")
            with st.expander("パスワードを変更する", expanded=False):
                with st.form("change_password_form"):
                    old_pwd = st.text_input("現在のパスワード", type="password")
                    new_pwd = st.text_input("新しいパスワード", type="password")
                    new_pwd_conf = st.text_input("新しいパスワード（確認用）", type="password")
                    
                    pwd_submit = st.form_submit_button("パスワードを変更", type="primary")
                    if pwd_submit:
                        if not old_pwd or not new_pwd or not new_pwd_conf:
                            st.error("全ての項目を入力してください。")
                        elif new_pwd != new_pwd_conf:
                            st.error("新しいパスワードと確認用パスワードが一致しません。")
                        elif len(new_pwd) < 4:
                            st.error("パスワードは4文字以上で設定してください。")
                        else:
                            with st.spinner("変更中..."):
                                success, msg = change_user_password(st.session_state['username'], old_pwd, new_pwd)
                                if success:
                                    st.success(msg)
                                else:
                                    st.error(msg)
            
            # 支払い方法マスターは独立したメニューに移動しました。

        elif menu_selection == "ヘルプ":
            st.markdown("#### 💡 ヘルプ・サポート")
            
            st.info("アプリの機能や使い方、データの保存先などについて何でも聞いてください！")
            
            # --- チャット履歴の初期化 ---
            if "help_chat_history" not in st.session_state:
                st.session_state.help_chat_history = []
            
            if "help_messages" not in st.session_state:
                st.session_state.help_messages = [
                    {"role": "assistant", "content": "こんにちは！AI家計簿アプリのサポートAIです。\n機能の使い方や、データがどこに保存されているかなど、質問があればどうぞ！"}
                ]
                
            # メッセージ履歴の表示
            for i, msg in enumerate(st.session_state.help_messages):
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
                    if msg["role"] == "assistant":
                        render_speech_synthesis_button(msg["content"], f"help_{i}")
            
            # 音声入力ボタン表示
            render_voice_input_button("help")

            # ユーザー入力
            if user_input := st.chat_input("質問を入力してください...（例: レシートはどうやって登録するの？）"):
                # ユーザーのメッセージを表示して履歴に追加
                st.session_state.help_messages.append({"role": "user", "content": user_input})
                with st.chat_message("user"):
                    st.markdown(user_input)
                    
                # APIクライアントの取得
                client = st.session_state.get('genai_client')
                if not client:
                    with st.chat_message("assistant"):
                        st.error("APIキーが設定されていないため、回答できません。secrets.toml を確認してください。")
                else:
                    with st.chat_message("assistant"):
                        message_placeholder = st.empty()
                        message_placeholder.markdown("回答を生成中...")
                        
                        try:
                            # 1. サポートAI用のシステムプロンプト（取扱説明書）
                            app_manual = """
あなたは、この高機能家計簿アプリ「マイニー」の公式サポートAIです。
ユーザーから機能の質問や操作方法を聞かれたら、以下の情報を元に、親切かつ分かりやすく案内してください。
現在の左側サイドバーのメニュー構成は以下の4つの大分類に分かれています。

【表示・分析系】（家計の状況を確認・分析するメニュー）
・ダッシュボード（月次集計）：月間の総支出、予算の残り、日別の支出推移をグラフで確認できます。内訳は「店舗別」「大分類別」「支払い方法別」に切り替え可能で、最大で商品（明細）レベルまで掘り下げることができます。
・ダッシュボード（年次集計）：選択した年の支出を月ごとに集計・表示し、前年対比棒グラフなどを確認できます。
・カレンダー：初期表示で「本日の日付」が自動選択され、本日の支出明細がすぐに見られます。月間カレンダー上の日付クリックでも明細が表示されます。
・クレジットカード：登録したカードの利用状況を「当月支払」「次回支払額」「次回以降支払額」の3つの期間に分けて表示します。それぞれの期間の明細は、日付＋店舗名＞大分類＞小分類＞商品名の4階層のドリルダウンで詳細を確認できます。未払い金額に対して現在の利用率や残高も表示されます。
・固定費管理：毎月の固定費と変動費のシミュレーションを行い、支払い方法別の小計や総合計を確認できます。ワンクリックで支払いスプレッドシートへの連携（GAS連携）も可能です。

【レシート管理】（支出データを登録・修正するメニュー）
・レシート取込：写真をアップロードし、AIで自動解析します。「Gift」や「ギフト」のキーワードがあれば自動で「割引・ポイント利用」のマイナス金額として抽出します。解析結果はアコーディオン形式（大分類＞小分類＞商品）で詳細を確認できます。支払い方法が「未設定」のまま登録ボタンを押すと警告が出ますが、そのまま再度押すことで未設定のまま登録も可能です。
・レシート手入力：キーボード操作で画面上の表に高速連続入力が可能です。こちらも支払い方法の初期値は「未設定」です。
・レシート修正：過去データの検索・修正・削除、対象レシートの一括更新が行えます。個別の明細を修正中（選択中）のときは、誤操作を防ぐためにレシート全体の日付や店舗名などのヘッダーがロックされる仕組みがあります。

【相談・サポート】（使い方や家計の悩みを解決するメニュー）
・マニュアル：このアプリの全機能と使い方の一覧です。（※各項目にはテキスト読み上げ機能がついています）
・ヘルプチャット：あなた（AIアシスタント）にアプリの操作方法などを直接質問できる機能です。
・AI相談（専属FP）：ユーザーの実際の家計データを元に、AIがFPとして個別アドバイスを行います。

【マスター設定】（アプリの基本設定を行うメニュー）
・支払方法マスター：クレジットカードや現金などの支払い手段を登録、修正、削除します。新規アカウント登録時には自動で「未設定」「現金」「PayPay」の3件が登録されます。
・カテゴリマスター（オーナー限定）：アプリの表示で使う大分類と小分類を追加・管理できます。使用済みのデータは変更・削除できない保護機能つきです。
・プロフィール設定：AI相談用の情報設定のほか、自分のアカウントのログインパスワードをいつでも変更できます。

【その他の便利機能】
・データのダウンロード：サイドバー下部の「データのダウンロード」から、全データのExcel/CSV出力が可能です。

回答のコツ：
・各機能への移動は、画面左側の「サイドバー（メニュー）」の大カテゴリから行えることを案内してください。
・「クレジットカード」の正しい表示方法を聞かれたら、「マスター設定の支払方法マスター」での設定案内を行ってください。
・専門用語は控え、明るく親身なトーンで答えてください。
"""
                            
                            # 2. 送信直前でチャットオブジェクトを「履歴付き」で作成
                            # 初回メッセージを擬似的に履歴に含める
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
                                
                                # 履歴を更新（画面表示用）
                                st.session_state.help_messages.append({"role": "assistant", "content": full_response})
                                
                                # 最新の回答にも読み上げボタンを表示
                                render_speech_synthesis_button(full_response, "help_latest")
                                
                                # SDKの履歴をセッションに同期
                                st.session_state.help_chat_history = chat.get_history()
                                
                            except Exception as e:
                                error_msg = f"エラーが発生しました: {e}"
                                message_placeholder.error(error_msg)
                                st.session_state.help_messages.append({"role": "assistant", "content": error_msg})
                        except Exception as e:
                            st.error(f"予期せぬエラーが発生しました: {e}")
            
        elif menu_selection == "マニュアル":
            st.markdown("### 📗 マイニー公式マニュアル")
            st.info("家計簿アプリ「マイニー」の全機能とメインメニュー（サイドバー）の構成に基づいて使い方を案内します。")
            
            st.markdown("<h4 style='color: navy; margin-top: 30px;'>【表示・分析系】</h4>", unsafe_allow_html=True)
            st.caption("入力された家計データを様々な角度から確認・分析するためのメニューです。")
            with st.expander("📊 ダッシュボード（月次集計）", expanded=True):
                text_dash_month = """
                **概要**: 月間の総支出、予算、日別の推移をグラフで可視化します。
                **操作手順**:
                1. サイドバーから「ダッシュボード（月次集計）」を選択します。
                2. 画面上部の月選択（前月・翌月ボタン等）で確認したい月を選びます。
                3. 画面中央のボタン群（「店舗別」「大分類別」「支払い方法別」）をクリックして分析軸を切り替えます。
                4. 各項目のアコーディオン（詳細枠）をクリックすると、最大で「商品（明細）」レベルまで深い階層で内訳を追跡できます。
                - **カラー同期**: 円グラフと積上げ棒グラフで同じカテゴリには同じ色が適用されます。
                """
                st.markdown(text_dash_month)
                render_speech_synthesis_button(text_dash_month.replace("**", "").replace("-", ""), "sp_dash_mon")

            with st.expander("📊 ダッシュボード（年次集計）"):
                text_dash_year = """
                **概要**: 選択した「年」全体の支出データを集計・分析します。
                **操作手順**:
                1. サイドバーから「ダッシュボード（年次集計）」を選択します。
                2. 画面上部の年選択（前年・翌年ボタン等）で集計対象の年を切り替えます。
                3. 上部の「前年対比棒グラフ」や「年次大分類別シェア円グラフ」で傾向を視覚的に把握します。
                4. 画面下部に並ぶ各カテゴリ別内訳のアコーディオンを展開すると、年間を通した支出の詳細を追跡できます。
                """
                st.markdown(text_dash_year)
                render_speech_synthesis_button(text_dash_year.replace("**", "").replace("-", ""), "sp_dash_yr")
                
            with st.expander("📅 カレンダー"):
                text_calendar = """
                **概要**: 日付ごとの支出額をカレンダー形式で一覧できます。
                **操作手順**:
                1. サイドバーから「カレンダー」を選択します。初期状態では本日の明細が自動的に下部に表示されます。
                2. カレンダー上の任意の日付マス（青は土曜、赤は日曜・祝日）をクリックします。
                3. カレンダー下部にその日の「支出明細」が一覧表示されます。
                4. 「店舗別」「大分類別」「支払い方法別」の切り替えボタンを使い、最大5階層（店舗名 ＞ 支払い方法 ＞ 大分類 ＞ 小分類 ＞ 商品）でデータを掘り下げて確認します。
                """
                st.markdown(text_calendar)
                render_speech_synthesis_button(text_calendar.replace("**", "").replace("-", ""), "sp_cal")
                
            with st.expander("💳 クレジットカード"):
                text_cc = """
                **概要**: 登録したクレジットカードの利用状況や引き落としスケジュールを管理します。
                **操作手順**:
                1. サイドバーから「クレジットカード」を選択します。
                2. 画面上部のセレクトボックスから、確認したい特定のカード（または全体）を選びます。
                3. 画面中央のラジオボタンで「当月支払」「次回支払額」「次回以降支払額」の3つの期間を切り替えます。
                4. 表示された明細のアコーディオン（日付＋店舗名 ＞ 大分類 ＞ 小分類 ＞ 商品名）を開いて、利用内容や限度額に対する現在の利用割合（プログレスバー）を確認します。
                """
                st.markdown(text_cc)
                render_speech_synthesis_button(text_cc.replace("**", "").replace("-", ""), "sp_cc")

            st.markdown("<h4 style='color: navy; margin-top: 30px;'>【支払管理】</h4>", unsafe_allow_html=True)
            st.caption("毎月の固定費のシミュレーションと支払い情報を管理するためのメニューです。")
            with st.expander("📝 支払管理の基本操作"):
                text_fixed = """
                ・**支払管理シート新規作成**: 初回のみ実行します。ご自身のGoogleドライブ上に専用のスプレッドシート（支払管理・固定費マスター）を自動生成します。
                ・**固定費マスター設定**: 毎月発生する固定費や変動費の基本ルールを、「固定費マスター」シートへ登録・編集します。
                ・**固定費データ展開**: マスターの設定内容をもとに、2036年までのタイムラインへ月ごとの支払額を自動計算して展開します。（罫線や背景色の自動フォーマット付き）
                ・**変動費データ更新**: 変動費データを更新します。
                ・**支払管理シートを確認**: 毎月の支払予定が一覧になったスプレッドシートを確認・編集できます。口座引落日などの条件に合わせて「完了フラグ」が自動更新されるため、支払漏れを防止し、収支見通しを立てるのに役立ちます。
                """
                st.markdown(text_fixed)
                render_speech_synthesis_button(text_fixed.replace("**", "").replace("・", ""), "sp_fixed")

            st.markdown("<h4 style='color: navy; margin-top: 30px;'>【レシート管理】</h4>", unsafe_allow_html=True)
            st.caption("日々の買い物や支出の記録を追加・修正するためのメニューです。")
            with st.expander("📸 レシート取込（AI解析）"):
                text_ai_receipt = """
                **概要**: レシートの写真を撮ってアップロードするだけで、AIが内容を読み取ります。
                **操作手順**:
                1. サイドバーから「レシート取込」を選択します。
                2. 画面上のボタン（Browse files 等）からレシートの画像を選択し、「レシートを解析する」をクリックします。
                3. AIが店舗名・商品名・金額などを自動解析します（Gift等の割引や消費税も自動処理されます）。
                4. 解析結果が「大分類 ＞ 小分類 ＞ 商品」の階層形式で表示されるので、内容を確認・修正します。
                5. 問題なければ「登録」をクリックします（支払い方法未設定時は確認メッセージが出ます）。
                """
                st.markdown(text_ai_receipt)
                render_speech_synthesis_button(text_ai_receipt.replace("**", "").replace("-", ""), "sp_ai_rec")

            with st.expander("⌨️ レシート手入力（高速入力）"):
                text_manual_receipt = """
                **概要**: キーボード操作で素早く支出を入力できます。
                **操作手順**:
                1. サイドバーから「レシート手入力」を選択します。
                2. 上部の入力欄に「日付」と「店舗名」を設定します。
                3. 下部の表で「商品名」「金額」「大分類」「小分類」を入力します（金額入力後Enterキー等で自動で次の行が追加されます）。
                4. 入力が終わったら、下部で「支払い方法」を選んで「一括登録」ボタンをクリックして保存します。
                """
                st.markdown(text_manual_receipt)
                render_speech_synthesis_button(text_manual_receipt.replace("**", "").replace("-", ""), "sp_man_rec")

            with st.expander("✏️ レシート修正・履歴管理"):
                text_edit_receipt = """
                **概要**: 過去に登録した全てのデータを一覧・検索・編集できます。
                **操作手順**:
                1. サイドバーから「レシート修正」を選択します。
                2. 対象月を選択するか、検索ボックスからキーワードを入力して対象のレシートを探します。
                3. 画面左側のリストからレシートを選択すると、右側にその詳細（明細一覧）が表示されます。
                4. 日付、店舗名を追加・修正する場合は上部から、個別の明細を修正する場合は表から直接書き換えます。
                5. 修正が終わったら「変更を保存」ボタンを押します（※個別明細の編集中はレシート全体の操作がロックされます）。
                """
                st.markdown(text_edit_receipt)
                render_speech_synthesis_button(text_edit_receipt.replace("**", "").replace("-", ""), "sp_edit_rec")

            st.markdown("<h4 style='color: navy; margin-top: 30px;'>【相談・サポート】</h4>", unsafe_allow_html=True)
            st.caption("アプリの使い方に困った時や、家計改善のアドバイスが欲しい時のメニューです。")
            with st.expander("📗 マニュアル"):
                text_manual = """
                **概要**: 今ご覧いただいているこの画面です。全機能の概要と使い方を確認できます。
                **操作手順**:
                1. サイドバーから「マニュアル」を選択します。
                2. 確認したい説明項目（アコーディオン）をクリックして展開します。
                3. 音声で使い方を聞きたい場合は、文章の下にある「音声で読み上げる」ボタンを押してください。
                """
                st.markdown(text_manual)
                render_speech_synthesis_button(text_manual.replace("**", ""), "sp_manual")
                
            with st.expander("❓ ヘルプチャット"):
                text_help = """
                **概要**: アプリの使い方で困ったら、チャットで何でも質問できます。
                **操作手順**:
                1. サイドバーから「ヘルプチャット」を選択します。
                2. 画面下部の入力欄に質問内容（例：「レシートはどう登録するの？」）を入力し、エンターキーを押します。
                3. AIアシスタントが回答を生成し表示します。右下のマイクボタンから音声で質問することも可能です。
                """
                st.markdown(text_help)
                render_speech_synthesis_button(text_help.replace("**", "").replace("-", ""), "sp_help")
                
            with st.expander("🤖 AI相談（専属FP）"):
                text_ai_fp = """
                **概要**: あなたの実際の支出データを基に、AIがプロのFPとして分析やアドバイスを行います。
                本アプリで最も活用していただきたい、パーソナライズされたコンサルティング機能です。
                **操作手順**:
                1. サイドバーから「AI相談（専属FP）」を選択します。
                2. 下部のチャット入力欄に相談内容を入力します（例：「先月より食費が増えた理由は？」）。
                3. 入力欄右側のマイクボタンを押すと、声による相談も可能です。
                4. AIがあなたの登録プロフィールや実際の支出データを読み解き、パーソナライズされた回答を提示します。
                """
                st.markdown(text_ai_fp)
                render_speech_synthesis_button(text_ai_fp.replace("**", "").replace("-", ""), "sp_ai_fp")
                
            st.markdown("<h4 style='color: navy; margin-top: 30px;'>【マスター設定】</h4>", unsafe_allow_html=True)
            st.caption("アプリ全体の基本設定や、あなたに合わせたカスタマイズを行うメニューです。")
            with st.expander("💳 支払方法マスター"):
                text_pay_master = """
                **概要**: アプリ全体で利用する「支払い方法」を管理します。新規登録時には「未設定」「現金」「PayPay」が自動で作成されます。
                **操作手順**:
                1. サイドバーから「支払方法マスター」を選択します。
                2. 【新規登録の場合】: 左側のフォームに支払い方法の名称等を入力し「登録」ボタンを押します。クレジットカードの場合は締日や支払日なども設定できます。
                3. 【修正・削除の場合】: 右側の一覧表から対象の名称を選び、「更新」または「削除」を行います。
                """
                st.markdown(text_pay_master)
                render_speech_synthesis_button(text_pay_master.replace("**", "").replace("-", ""), "sp_pay_master")

            if st.session_state.get('username', '').lower() == 'tkouho':
                with st.expander("📂 カテゴリマスター（オーナー専用）"):
                    text_cat_master = """
                    **概要**: アプリ全体の「大分類」と「小分類」の構成を直感的に設定します。
                    **操作手順**:
                    1. サイドバーから「カテゴリマスター」を選択します。
                    2. 画面左側のリストから編集したい「大分類」を選択します。
                    3. 画面右側にその大分類に属する「小分類」のリストが表示されるので、追加・変更・削除を行い、「保存」ボタンを押します。
                    ※既に使用されているカテゴリは変更・削除が自動ブロックされます。
                    """
                    st.markdown(text_cat_master)
                    render_speech_synthesis_button(text_cat_master.replace("**", "").replace("-", ""), "sp_cat_master")

            with st.expander("⚙️ プロフィール設定"):
                text_profile = """
                **概要**: パーソナライズ設定とアカウント管理を行うメニューです。
                **操作手順**:
                1. サイドバーから「プロフィール設定」を選択します。
                2. AI向けのプロフィール（年齢、職業、趣味、目標等）を入力・確認し「プロフィールを保存」を押します。
                3. パスワードを変更する場合は、下部の変更フォームに現在のパスワードと新パスワードを入力し「パスワードを変更する」を押します。
                """
                st.markdown(text_profile)
                render_speech_synthesis_button(text_profile.replace("**", "").replace("-", ""), "sp_profile")
                
            st.markdown("<h4 style='color: navy; margin-top: 30px;'>その他の便利機能</h4>", unsafe_allow_html=True)
            with st.expander("📥 データのダウンロード"):
                text_download = """
                **概要**: 登録したすべての支出データを、自分の端末に保存できます。
                **操作手順**:
                1. サイドバーを一番下までスクロールします。
                2. 「データのダウンロード」セクションで、「エクセル形式(.xlsx)」または「CSV形式(.csv)」のファイルダウンロードボタンをクリックします。
                3. お使いの端末にファイルが保存されます。
                """
                st.markdown(text_download)
                render_speech_synthesis_button(text_download.replace("**", "").replace("-", ""), "sp_dl")

        elif menu_selection == "クレジットカード":
            show_credit_card_dashboard()

        elif menu_selection == "支払方法マスター":
            show_payment_master()
        elif menu_selection == "銀行マスター":
            show_bank_master()
        elif menu_selection == "カテゴリマスター":
            show_category_master()

        elif menu_selection == "プロフィール設定":
            show_profile_settings()
            
        elif menu_selection == "支払管理シート新規作成":
            show_fixed_cost_management()
            
        elif menu_selection == "固定費マスター設定":
            show_fixed_cost_master_settings()
            
        elif menu_selection == "固定費データ展開":
            show_fixed_cost_data_expansion()
            
        elif menu_selection == "変動費データ更新":
            show_variable_cost_update()
            
        elif menu_selection == "支払管理シートを確認":
            show_open_management_sheet()
            
        st.caption("マイニー Ver 4.22.0 - ユーザー: %s" % st.session_state['username'])
            
    # 未ログインの状態 (ログイン・登録画面)
    else:
        st.title("AI家計簿アプリ")
        
        tab1, tab2 = st.tabs(["ログイン", "新規ユーザー登録"])
        
        with tab1:
            st.subheader("ログイン")
            

            with st.form("login_form"):
                login_username = st.text_input("ユーザー名", key="login_username_input_v2")
                login_password = st.text_input("パスワード", type="password", key="login_password_input_v2")
                remember_me = st.checkbox("ログイン状態を保持する", value=True, key="remember_me_input_v2")
                
                submitted = st.form_submit_button("ログイン", use_container_width=True, type="primary")
            
            if submitted:
                if login_username and login_password:
                    with st.spinner("認証中..."):
                        if authenticate_user(login_username, login_password):
                            st.session_state['logged_in'] = True
                            st.session_state['username'] = login_username.strip().lower()
                            if remember_me:
                                st.query_params['user'] = st.session_state['username']
                            st.rerun()
                        else:
                            st.error("ユーザー名またはパスワードが間違っています。")
                else:
                    st.warning("ユーザー名とパスワードを入力してください。")
                        
        with tab2:
            st.subheader("新規ユーザー登録")
            
            with st.form("register_form"):
                reg_username = st.text_input("新しいユーザー名", key="reg_username_v2")
                reg_password = st.text_input("新しいパスワード", type="password", key="reg_password_v2")
                reg_password_confirm = st.text_input("パスワード（確認用）", type="password", key="reg_password_confirm_v2")
                
                reg_submitted = st.form_submit_button("登録する", use_container_width=True, type="primary")
            
            if reg_submitted:
                if reg_username and reg_password and reg_password_confirm:
                    if reg_password != reg_password_confirm:
                        st.error("パスワードが一致しません。")
                    else:
                        with st.spinner("登録中..."):
                            success, message = register_user(reg_username, reg_password)
                            if success:
                                st.success(message)
                            else:
                                st.error(message)
                else:
                    st.warning("すべてのフィールドを入力してください。")

if __name__ == "__main__":
    main()
