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
from datetime import datetime
from dateutil.relativedelta import relativedelta
from PIL import Image, ImageOps
from google import genai
from google.genai import types
import time
import re
import base64
import xlsxwriter


# ---------- 構成設定 ----------
from urllib.parse import urlparse

SPREADSHEET_NAME = "Kakeibo_Data"
WORKSHEET_NAME = "users"
TRANSACTIONS_WORKSHEET_NAME = "transactions"
USER_MASTER_WORKSHEET_NAME = "User_Master"
PAYMENT_MASTER_WORKSHEET_NAME = "Payment_Master"

# ---------- カテゴリ定義 ----------
# AI判別やセレクトボックスで利用するための大分類・小分類の親子関係定義
EXPENSE_CATEGORIES = {
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
    "その他": ["📁未分類"]
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
    "割引・ポイント利用": "#FFFF00" # 黄色
}

def get_categories():
    return EXPENSE_CATEGORIES

def get_categories_prompt_text():
    """AI（Gemini等）のプロンプトに埋め込むためのカテゴリ定義文字列を生成"""
    text = "【カテゴリシステム: 大分類と小分類のリスト】\n"
    for major, minors in EXPENSE_CATEGORIES.items():
        text += f"- {major}: {', '.join(minors)}\n"
    text += "\n※ 必ず上記の大分類と小分類の組み合わせに従ってください。"
    return text

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
            
    if "user" in params and not st.session_state.get('logged_in'):
        u_val = params["user"]
        if isinstance(u_val, list): u_val = u_val[0]
        st.session_state['username'] = u_val
        st.session_state['logged_in'] = True
        need_rerun = True
        
    if "menu" in params:
        m_val = params["menu"]
        if isinstance(m_val, list): m_val = m_val[0]
        st.session_state['menu_selection'] = m_val
        need_rerun = True

    st.session_state['_init_done'] = True
    
    if need_rerun:
        st.query_params.clear()
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

def safe_gspread_call(func, *args, max_retries=3, delay=2, **kwargs):
    """API呼び出しをリトライする関数"""
    last_error = None
    for i in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            last_error = e
            # 一時的な接続エラーの場合にリトライ
            if "RemoteDisconnected" in str(e) or "Connection aborted" in str(e) or "TimeoutError" in str(e):
                time.sleep(delay * (i + 1)) # 指数バックオフ的に待機
                continue
            else:
                # 致命的なエラー（認証等）はすぐに上げる
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
            # 429 RESOURCE_EXHAUSTED または 500/503 系エラーの場合にリトライ
            if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg or "500" in err_msg or "503" in err_msg:
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

def init_users_sheet(sheet):
    """初期セットアップ：ヘッダーがない場合に作成する"""
    try:
        headers = sheet.row_values(1)
        if not headers or headers[0] != "username":
            sheet.insert_row(["username", "password_hash"], 1)
    except Exception:
        # シートが空の場合に例外が発生する可能性があるため、その場合はヘッダーを追加
        sheet.insert_row(["username", "password_hash"], 1)

def init_transactions_sheet(sheet):
    """初期セットアップ：取引シートのヘッダーがない場合に作成する"""
    expected_headers = ["username", "date", "store_name", "item_name", "category", "subcategory", "amount", "update", "payment_method"]
    try:
        headers = sheet.row_values(1)
        if not headers or headers[0] != "username":
            sheet.insert_row(expected_headers, 1)
        elif len(headers) < len(expected_headers):
            # 不足しているヘッダーを追記する
            for i in range(len(headers), len(expected_headers)):
                sheet.update_cell(1, i + 1, expected_headers[i])
    except Exception:
        sheet.insert_row(expected_headers, 1)

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
    except Exception:
        safe_gspread_call(sheet.insert_row, expected_headers, 1)

def init_payment_master_sheet(sheet):
    """初期セットアップ：Payment_Masterシートのヘッダーがない場合に作成する"""
    expected_headers = ["username", "payment_id", "name", "type", "closing_date", "payment_month", "payment_date"]
    try:
        headers = safe_gspread_call(sheet.row_values, 1)
        if not headers or headers[0] != "username":
            safe_gspread_call(sheet.insert_row, expected_headers, 1)
        elif len(headers) < len(expected_headers):
            for i in range(len(headers), len(expected_headers)):
                safe_gspread_call(sheet.update_cell, 1, i + 1, expected_headers[i])
    except Exception:
        safe_gspread_call(sheet.insert_row, expected_headers, 1)

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
                    methods.append(row)
        return methods
    except Exception as e:
        st.error(f"支払い方法取得エラー: {e}")
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
            payment_data.get("payment_date", "")
        ]
        
        if row_idx:
            # 既存の行を更新
            update_range = f"A{row_idx}:G{row_idx}"
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
        "payment_method": "payment_method"
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
    else:  # yearly
        df = df[df["date"].dt.year == target_date.year]
    
    # 金額を数値に変換
    if "amount" in df.columns:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    else:
        df["amount"] = 0
        
    # --- カテゴリの正規化（集計時やセレクトボックス等で指定外が出ないようにする） ---
    # 大分類の正規化
    if "category" in df.columns:
        valid_majors = list(EXPENSE_CATEGORIES.keys())
        # 定義にない大分類は「その他」にまとめる
        df["category"] = df["category"].apply(lambda x: x if x in valid_majors else "その他")
        
    # 小分類の正規化
    sub_cols = [c for c in ["subcategory", "sub_category", "小分類"] if c in df.columns]
    if sub_cols:
        sub_col = sub_cols[0]
        def normalize_sub(row):
            major = row.get("category", "その他")
            sub = str(row.get(sub_col, "")).strip()
            valid_subs = EXPENSE_CATEGORIES.get(major, sorted(EXPENSE_CATEGORIES["その他"]))
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
            try:
                reported_total = int(item.get("amount", 0))
            except ValueError:
                reported_total = None
            continue
            
        clean_results.append(item)
        
        try:
            amt = int(item.get("amount", 0))
        except ValueError:
            amt = 0
            
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
    view_pattern = st.radio("表示パターン", ["店舗別", "大分類別"], horizontal=True, key=f"{key_prefix}_view_pattern")
    
    if view_pattern == "店舗別":
        store_col = "store_name" if "store_name" in df.columns else "store" if "store" in df.columns else None
        if store_col:
            store_grouped = df_agg.groupby(store_col, as_index=False)["amount"].sum()
            store_grouped = store_grouped.sort_values(by="amount", ascending=False)
            
            for _, row in store_grouped.iterrows():
                store = row[store_col]
                total_amt_str = f"￥{int(row['amount']):,}"
                
                with st.expander(f"{store}：{total_amt_str}"):
                    store_df = df[df[store_col] == store].copy()
                    # 内訳（大分類・小分類）は内税を含めて集計する
                    cat_grouped = store_df.groupby("category", as_index=False)["amount"].sum()
                    cat_grouped = cat_grouped.sort_values(by="amount", ascending=False)
                    
                    for _, cat_row in cat_grouped.iterrows():
                        cat = cat_row["category"]
                        cat_amt_str = f"￥{int(cat_row['amount']):,}"
                        
                        # 2段階目：大分類アコーディオン
                        sub_df = store_df[store_df["category"] == cat].copy()
                        sub_col = None
                        for col_name in ["subcategory", "sub_category", "小分類"]:
                            if col_name in sub_df.columns:
                                sub_col = col_name
                                break
                        
                        if sub_col:
                            sub_grouped = sub_df.groupby(sub_col, as_index=False)["amount"].sum()
                            sub_grouped = sub_grouped.sort_values(by="amount", ascending=False)
                            
                            with st.expander(f"  └ {cat}：{cat_amt_str}"):
                                if key_prefix == "calendar":
                                    # カレンダー詳細（店舗別）のみ 3階層目以降をカスタムHTMLで極薄表示
                                    for _, sub_row in sub_grouped.iterrows():
                                        sub_name = sub_row[sub_col]
                                        sub_amt_str = f"￥{int(sub_row['amount']):,}"
                                        
                                        # 3階層目（小分類）と4階層目（商品名）を一つのdetailsタグにまとめる
                                        # インデントがあるとMarkdownのコードブロックと誤認されるため、左詰めにする
                                        html_str = f'<details style="margin: 1px 0;">'
                                        html_str += f'<summary style="background-color: #f0f2f6; padding: 2px 8px; margin: 0; border-left: 5px solid #007bff; font-size: 0.9rem; line-height: 1.2; list-style: none; cursor: pointer;">'
                                        html_str += f'L {sub_name}：{sub_amt_str}</summary>'
                                        html_str += f'<div style="padding-left: 10px;">'
                                        
                                        # 4階層目：商品名（詳細）
                                        item_df = sub_df[sub_df[sub_col] == sub_name].copy()
                                        item_col = "item_name" if "item_name" in item_df.columns else "item" if "item" in item_df.columns else None
                                        
                                        if item_col:
                                            item_grouped = item_df.groupby(item_col, as_index=False)["amount"].sum()
                                            item_grouped = item_grouped.sort_values(by="amount", ascending=False)
                                            for _, i_row in item_grouped.iterrows():
                                                i_name = i_row[item_col]
                                                i_amt = f"￥{int(i_row['amount']):,}"
                                                html_str += f'<div style="padding-left: 10px; font-size: 0.85rem; line-height: 1.1; margin: 0; color: #555;">└ {i_name}：{i_amt}</div>'
                                        else:
                                            for _, i_row in item_df.iterrows():
                                                i_amt = f"￥{int(i_row['amount']):,}"
                                                html_str += f'<div style="padding-left: 10px; font-size: 0.85rem; line-height: 1.1; margin: 0; color: #555;">└ {i_amt}</div>'
                                        
                                        html_str += "</div></details>"
                                        st.markdown(html_str, unsafe_allow_html=True)
                                else:
                                    # その他（ダッシュボード等）は 3階層のまま（大分類 > 小分類リスト）
                                    sub_grouped_disp = sub_grouped.copy()
                                    sub_grouped_disp["amount"] = sub_grouped_disp["amount"].apply(lambda x: f"￥{int(x):,}")
                                    sub_grouped_disp.columns = ["小分類", "金額"]
                                    st.dataframe(sub_grouped_disp, use_container_width=True, hide_index=True)
                        else:
                            # 小分類がない場合は明細
                            item_cols = [c for c in ["item_name", "item", "amount"] if c in sub_df.columns]
                            display_items = sub_df[item_cols].copy()
                            display_items["amount"] = display_items["amount"].apply(lambda x: f"￥{int(x):,}")
                            
                            with st.expander(f"  └ {cat}：{cat_amt_str}"):
                                st.dataframe(display_items, use_container_width=True, hide_index=True)
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
        for major, minors in EXPENSE_CATEGORIES.items():
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
            ["大分類別", "小分類別", "店舗別", "消費税"], 
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
        df["tax_group"] = df[sub_col].apply(map_tax)
        df_agg = df # 消費税軸の場合はこのフィルタ済みデータを全表示

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
            ["大分類別", "小分類別", "店舗別", "消費税"], 
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
            
        df["tax_group"] = df[sub_col].apply(map_tax)
        df_prev["tax_group"] = df_prev[sub_col].apply(map_tax) if not df_prev.empty else None
        df_agg = df 
        df_prev_agg = df_prev

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
            st.subheader("マイニー [Ver 4.1.6]")
            st.write(f"🔑 ユーザー: **{st.session_state['username']}**")
            st.markdown("---")
            if 'menu_selection' not in st.session_state:
                st.session_state['menu_selection'] = "カレンダー"
            
            # 既存の "ダッシュボード" を "ダッシュボード（月次集計）" に置換し、年次を追加
            menu_options = [
                "ダッシュボード（月次集計）", 
                "ダッシュボード（年次集計）", 
                "カレンダー", 
                "レシート取込", 
                "レシート手入力", 
                "レシート修正", 
                "マニュアル", 
                "ヘルプ", 
                "AI相談", 
                "プロフィール設定"
            ]
            
            # メニューのリセット処理（別の画面から戻ってきたとき用）
            # もしカレンダー等から「ダッシュボード系以外」を経由して戻ってきた場合、
            # 次にダッシュボードをクリックしたときに「月次」にしたいという要件。
            # 直前の値を保持しておき、遷移を検知する
            if "last_menu_selection" not in st.session_state:
                st.session_state.last_menu_selection = st.session_state['menu_selection']
            
            # サイドバーの仕切り線（カレンダーとレシート取込の間）
            st.markdown("""
                <style>
                /* radioボタン全体の中で、3番目の項目（カレンダー）の直後に線を引く */
                div[data-testid="stSidebar"] div[role="radiogroup"] > div:nth-of-type(3) {
                    border-bottom: 2px solid #ddd;
                    margin-bottom: 15px;
                    padding-bottom: 10px;
                }
                </style>
            """, unsafe_allow_html=True)
            
            menu_selection = st.radio(
                "機能を選択",
                menu_options,
                key="menu_selection",
                on_change=handle_menu_change
            )
            st.session_state.last_menu_selection = menu_selection

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
                st.session_state['selected_date'] = None

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
                        total_before_tax = sum(int(item.get("amount", 0)) for item in results)
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

                    total_amount = sum(int(item.get("amount", 0)) for item in results if not is_internal_tax(item))
                    
                    # 大分類別の内訳を集計
                    category_totals = {}
                    for item in results:
                        cat = item.get("major_category", "その他")
                        # 正規化処理を適用して大分類を揃える
                        majors = list(EXPENSE_CATEGORIES.keys())
                        final_major = "その他"
                        for m in majors:
                            if m in cat or cat in m:
                                final_major = m
                                break
                        
                        amt = int(item.get("amount", 0))
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
                    
                    st.markdown("---")
                    
                    # 🎯 支払い方法のUIを追加
                    methods = get_payment_methods(st.session_state['username'])
                    method_options = [m["name"] for m in methods] if methods else ["現金"]
                    selected_payment = st.selectbox("支払い方法", options=method_options)
                    
                    st.write("この内容で登録しますか？")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        submit_btn = st.button("登録", type="primary", use_container_width=True)
                    with col2:
                        cancel_btn = st.button("キャンセル", type="secondary", use_container_width=True)
                        
                    if cancel_btn:
                        st.session_state.parsed_results = None
                        st.session_state.uploader_key += 1
                        st.rerun()
                        
                    if submit_btn:
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
                                    
                                    written_count = 0
                                    for item in results:
                                        # カテゴリの正規化（14カテゴリ体系に強制）
                                        majors = list(EXPENSE_CATEGORIES.keys())
                                        major = str(item.get("major_category", "その他"))
                                        final_major = "その他"
                                        for m in majors:
                                            if m in major or major in m:
                                                final_major = m
                                                break
                                                
                                        minors = EXPENSE_CATEGORIES.get(final_major, EXPENSE_CATEGORIES["その他"])
                                        minor = str(item.get("minor_category", "❓その他"))
                                        final_minor = minors[-1] if minors else "❓その他"
                                        for m in minors:
                                            text_only = "".join([c for c in m if c.isalnum() or c in "類物食品未分類その他%"])
                                            if text_only and (text_only in minor or minor in text_only) and len(minor) > 0:
                                                final_minor = m
                                                break
                                                
                                        store_name = str(edited_store).strip()
                                        item_name = str(item.get("item_name", ""))
                                        
                                        # 日付を yyyy-mm-dd に整形
                                        formatted_date = edited_date.strftime("%Y-%m-%d") if edited_date else ""
    
                                        row_data = [
                                            str(st.session_state['username']),
                                            formatted_date,
                                            str(store_name),
                                            str(item_name),
                                            str(final_major),
                                            str(final_minor),
                                            int(item.get("amount", 0)),
                                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                            str(selected_payment)
                                        ]
                                        sheet.append_row(row_data)
                                        written_count += 1
                                    
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
                selected_payment_manual = st.selectbox("支払い方法", options=method_options, key=f"mi_pay_{fid}")
                
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
                                
                                for itm, cat in zip(valid_items, categories):
                                    major = cat.get("major_category", "その他")
                                    minor = cat.get("minor_category", "📁未分類")
                                    
                                    row_data = [
                                        str(st.session_state['username']),
                                        str(input_date.strftime('%Y-%m-%d')),
                                        str(input_store),
                                        str(itm["name"]),
                                        str(major),
                                        str(minor),
                                        int(itm["amount"]),
                                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        str(selected_payment_manual)
                                    ]
                                    safe_gspread_call(sheet.append_row, row_data)
                                
                                st.success(f"✅ {len(st.session_state.manual_input_items)}件のデータを登録しました！")
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
                    # レシート単位に集約（内税を金額から除外して集計）
                    df_agg = df.copy()
                    if "category" in df_agg.columns:
                        df_agg.loc[df_agg["category"] == "消費税（内税）", "amount"] = 0
                    
                    receipts_df = df_agg.groupby(["date", store_col], as_index=False).agg(
                        amount=("amount", "sum"),
                        明細数=("amount", "count")
                    )
                    receipts_df.columns = ["日付", "店舗名", "金額合計", "明細数"]
                    # 店舗名が空欄の場合は「店舗不明」とする
                    receipts_df["店舗名"] = receipts_df["店舗名"].replace("", "店舗不明")
                    receipts_df["日付"] = receipts_df["日付"].dt.strftime('%Y-%m-%d')
                    receipts_df["金額合計"] = receipts_df["金額合計"].apply(lambda x: int(x))
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
                        key=f"receipt_list_df_{st.session_state.receipt_list_version}"
                    )
                    
                    if len(event.selection.rows) > 0:
                        selected_idx = event.selection.rows[0]
                        sel_rec = receipts_df.iloc[selected_idx]
                        if sel_rec["日付"] != "総合計":
                            # 選択されたレシート情報をセッションに保存して永続化
                            st.session_state['selected_receipt_info'] = {
                                "date": sel_rec["日付"],
                                "store": sel_rec["店舗名"]
                            }
                    
                    # セッションに保存された情報に基づいて詳細を表示（表の選択が消えても維持）
                    receipt_info = st.session_state.get('selected_receipt_info')
                    if receipt_info:
                        # 表示中のリストにまだ存在するか確認（削除対策）
                        target_date_str = receipt_info["date"]
                        target_store = receipt_info["store"]
                        
                        selected_receipt_matches = receipts_df[
                            (receipts_df["日付"] == target_date_str) & 
                            (receipts_df["店舗名"] == target_store)
                        ]
                        
                        if not selected_receipt_matches.empty:
                            selected_receipt = selected_receipt_matches.iloc[0]
                            target_date = pd.to_datetime(selected_receipt["日付"])
                            target_store = selected_receipt["店舗名"]
                            
                            st.markdown("---")
                            st.write(f"##### 対象レシート明細： {selected_receipt['日付']} - {target_store}")
                            
                            # 該当レシートの明細を取得
                            details = df[(df["date"] == target_date) & (df[store_col] == target_store)].copy()
                            
                            receipt_key = f"{selected_receipt['日付']}_{target_store}"
                            if st.session_state.get('current_receipt_key') != receipt_key or st.session_state.get('edit_data') is None:
                                st.session_state['current_receipt_key'] = receipt_key
                                st.session_state['edit_data'] = {}
                                st.session_state['edit_header'] = {
                                    "date": target_date.date(),
                                    "store": target_store
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
                                        "amount": int(row.get("amount", 0)),
                                        "major": major,
                                        "minor": sub,
                                        "payment_method": payment_m
                                    }

                            # --- レシートヘッダー（日付・店舗名）の修正エリア ---
                            st.write(f"##### レシート修正（金額：￥{int(selected_receipt['金額合計']):,}）")
                            with st.container(border=True):
                                h_col1, h_col2 = st.columns(2)
                                with h_col1:
                                    new_date = st.date_input("日付", value=st.session_state['edit_header']['date'], key=f"edit_header_date_{receipt_key}")
                                with h_col2:
                                    new_store = st.text_input("店舗名", value=st.session_state['edit_header']['store'], key=f"edit_header_store_{receipt_key}")
                                
                                # ヘッダー情報を更新
                                st.session_state['edit_header']['date'] = new_date
                                st.session_state['edit_header']['store'] = new_store

                            # --- アクションボタンエリア（上部） ---
                            action_col1, action_col2 = st.columns(2)
                            
                            with action_col1:
                                if st.button("日付・店舗名更新", use_container_width=True, type="primary"):
                                    try:
                                        with st.spinner("一括更新中..."):
                                            sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                            target_date_str = st.session_state['edit_header']['date'].strftime("%Y-%m-%d")
                                            target_store = st.session_state['edit_header']['store']
                                            
                                            # 既存の全明細行をループして日付と店舗を更新
                                            existing_indices = [int(k) for k in st.session_state['edit_data'].keys() if not str(k).startswith("new_")]
                                            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                            for r_idx in existing_indices:
                                                sheet.update_cell(r_idx, 2, target_date_str)
                                                sheet.update_cell(r_idx, 3, target_store)
                                                sheet.update_cell(r_idx, 8, current_time)
                                                
                                            st.success("✅ レシート情報を一括更新しました")
                                            st.session_state.receipt_list_version += 1
                                            time.sleep(1)
                                            st.rerun()
                                    except Exception as e:
                                        st.error(f"更新エラー: {e}")
                            
                            with action_col2:
                                with st.popover("このレシートを全削除", use_container_width=True):
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
                                    
                                    # 🎯 支払い方法のUIを追加
                                    methods = get_payment_methods(st.session_state['username'])
                                    method_options = [m["name"] for m in methods] if methods else ["現金"]
                                    
                                    current_payment = target_item.get("payment_method", "現金")
                                    if current_payment not in method_options:
                                        if method_options:
                                            current_payment = method_options[0]
                                        else:
                                            # 現金だけの場合は必ずそれが選ばれる
                                            current_payment = "現金"
                                            
                                    edit_payment = st.selectbox("支払い方法", options=method_options, index=method_options.index(current_payment) if current_payment in method_options else 0, key=f"edit_payment_{receipt_key}_{current_editing_id}")
                                    
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
                                                        
                                                        if str(current_editing_id).startswith("new_"):
                                                            # 新規追加
                                                            new_row = [user_name, target_date_str, target_store, edit_name, edit_major, edit_minor, edit_amount, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), edit_payment]
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
                                                            # A:username, B:date, C:store, D:item, E:major, F:minor, G:amount, H:update, I:payment_method
                                                            # 更新範囲: B (Col 2) から I (Col 9)
                                                            update_range = f"B{r_idx}:I{r_idx}"
                                                            update_values = [[target_date_str, target_store, edit_name, edit_major, edit_minor, edit_amount, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), edit_payment]]
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
                # 全データを取得（load_transactions_dataを流用せず、全期間を対象にするため直接取得）
                sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                init_transactions_sheet(sheet)
                records = safe_gspread_call(sheet.get_all_records)
                
                if not records:
                    return ""
                
                df_all = pd.DataFrame(records)
                # セキュリティの最重要要件：現在ログインしているユーザーのデータのみにフィルタリング
                if "username" in df_all.columns:
                    df_user = df_all[df_all["username"].astype(str).str.lower() == username.lower()].copy()
                else:
                    return ""

                if df_user.empty:
                    return ""

                # 必要なカラムのみ抽出・整形
                # 「対象年月、日付、店舗名、大分類、小分類、金額」
                df_user["date"] = pd.to_datetime(df_user["date"], errors="coerce")
                df_user = df_user.dropna(subset=["date"])
                df_user["対象年月"] = df_user["date"].dt.strftime('%Y-%m')
                df_user["日付"] = df_user["date"].dt.strftime('%Y-%m-%d')
                
                # 表示用カラムのリネーム
                rename_map = {
                    "store_name": "店舗名",
                    "item_name": "商品名",
                    "category": "大分類",
                    "subcategory": "小分類",
                    "amount": "金額"
                }
                # 存在するカラムのみマッピング
                actual_rename = {k: v for k, v in rename_map.items() if k in df_user.columns}
                df_user = df_user.rename(columns=actual_rename)
                
                target_cols = ["対象年月", "日付", "店舗名", "商品名", "大分類", "小分類", "金額"]
                available_cols = [c for c in target_cols if c in df_user.columns]
                
                return df_user[available_cols].to_csv(index=False)
                
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
            st.markdown("#### 💳 支払い方法マスター")
            st.info("レシート登録時に選択できる「支払い方法」を追加・管理します。")
            
            # 現在の支払い方法一覧を表示
            methods = get_payment_methods(st.session_state['username'])
            
            st.write("##### 登録済みの支払い方法")
            if methods:
                # DataFrameにして表示
                df_methods = pd.DataFrame(methods)
                # 表示用カラムを絞る
                display_cols = ["payment_id", "name", "type", "closing_date", "payment_month", "payment_date"]
                df_display = df_methods[[c for c in display_cols if c in df_methods.columns]].copy()
                df_display = df_display.rename(columns={
                    "name": "支払い方法名", "type": "種類", "closing_date": "締日", 
                    "payment_month": "支払月", "payment_date": "支払日"
                })
                st.dataframe(df_display, use_container_width=True, hide_index=True)
                
                # 削除機能
                with st.expander("🗑️ 支払い方法を削除する"):
                    del_id = st.selectbox("削除する支払い方法（ID）を選択", options=[m["payment_id"] for m in methods])
                    if st.button("削除実行"):
                        with st.spinner("削除中..."):
                            success, msg = delete_payment_method(st.session_state['username'], del_id)
                            if success:
                                st.success(msg)
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error(msg)
            else:
                st.write("登録されている支払い方法はありません。（デフォルトとして「現金」が利用可能です）")
                
            st.write("##### 新規追加 / 編集")
            with st.form("payment_method_form"):
                p_col1, p_col2 = st.columns(2)
                with p_col1:
                    new_id = st.text_input("ID（半角英数、例: card_rakuten）", help="既存のIDを指定すると上書き更新されます。")
                    new_name = st.text_input("支払い方法名（例: 楽天カード、現金）")
                with p_col2:
                    type_options = ["現金", "クレジットカード", "電子マネー", "QR・バーコード決済", "その他"]
                    new_type = st.selectbox("種類", options=type_options)
                
                # クレジットカード専用設定
                st.markdown("###### クレジットカード詳細設定（種類が「クレジットカード」の場合のみ有効）")
                c_col1, c_col2, c_col3 = st.columns(3)
                with c_col1:
                    closing_date = st.selectbox("締日", ["", "末日"] + [str(i) for i in range(1, 31)])
                with c_col2:
                    payment_month = st.selectbox("支払月", ["", "当月", "翌月", "翌々月"])
                with c_col3:
                    payment_date = st.selectbox("支払日", ["", "末日"] + [str(i) for i in range(1, 31)])
                    
                submit_payment = st.form_submit_button("保存する", type="primary")
                
                if submit_payment:
                    if not new_id or not new_name:
                        st.warning("⚠️ IDと支払い方法名は必須です。")
                    elif new_type == "クレジットカード" and (not closing_date or not payment_month or not payment_date):
                        st.warning("⚠️ クレジットカードの場合は、締日・支払月・支払日をすべて設定してください。")
                    else:
                        with st.spinner("保存中..."):
                            # クレカ以外は詳細設定をクリアして保存
                            if new_type != "クレジットカード":
                                closing_date = ""
                                payment_month = ""
                                payment_date = ""
                                
                            data_to_save = {
                                "payment_id": new_id,
                                "name": new_name,
                                "type": new_type,
                                "closing_date": closing_date,
                                "payment_month": payment_month,
                                "payment_date": payment_date
                            }
                            succ, msg = save_payment_method(st.session_state['username'], data_to_save)
                            if succ:
                                st.success(msg)
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error(msg)
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

【1. ダッシュボード（月次集計）】
・概要：月間の総支出、予算の残り、日別の支出推移をグラフで確認できます。
・カテゴリ別内訳：画面中央のラジオボタンで「店舗別」「大分類別」「小分類別」の3パターンに切り替え可能です。
・2段階表示：例えば「店舗別」を選ぶと店舗が並び、クリックするとその店舗で買った大分類が表示されます。
・並び替え：支出の多い順（降順）に自動で並ぶため、どこに一番お金を使っているか一目でわかります。
・配色の同期：円グラフと棒グラフでカテゴリごとの色が統一されているため、直感的にデータを比較できます。

【1-2. ダッシュボード（年次集計）】
・概要：選択した「年」全体の支出を月ごとに集計して表示します。
・年次推移（前年対比）：前年との支出比較を月ごとの棒グラフで確認できます。
・年次大分類別シェア：1年間の支出を大分類ごとの円グラフで表示します。
・年次カテゴリ別内訳：1年間の全データに対して、2段階アコーディオン方式（大分類→小分類など）で詳細を確認できます。

【2. カレンダー機能】
・概要：カレンダー上で日々の支出額を一覧できます。祝日は赤く色付けされます。
・詳細確認：日付をクリックすると、その日の「支出明細」がアコーディオン形式で表示されます。
・3つの表示：カレンダー内でも「店舗別」「大分類別」「小分類別」をボタン一つで切り替えて分析できます。

【3. レシート取込（OCR）】
・操作：カメラで撮ったレシート画像をアップロードすると、AIが「店舗名」「商品名」「金額」「カテゴリ」を瞬時に解析します。
・自動補正：画像が横向きや逆さまでも、AIが正しい向き（縦向き）に自動で調整して表示します。
・自動消費税：AIによる解析結果に消費税が含まれていない場合、システムが自動的に内税10%の消費税項目を計算して追加します。
・確認と修正：解析結果の「日付」はカレンダーから、「店舗名」はテキストボックスで直接編集して登録できます。空欄の場合はエラー表示で登録を防ぎます。

【4. レシート手入力】
・操作：1行目の金額を入力中に「Enterキー」または「Tabキー」を押すと、自動的に次の行が追加されます。
・削除：各行の「✕」ボタンで、行を個別に削除できます。データがずれることはありません。
・登録：空行があっても、入力されているデータのみを正確に登録します。

【5. レシート修正・履歴】
・操作：「レシート修正」メニューから、過去に登録した全てのデータを表形式で確認できます。
・UI：明細一覧の指示が「（修正する行を選択して下さい）」となり、直感的に操作できます。
・編集：内容（日付・店舗名、個別明細）を書き換えて「更新」または「登録実行」ボタンを押すだけで修正完了です。
・安全性（データ保護）：書き込み直前にデータの不整合（行のずれなど）を検知する安全装置が搭載されました。不整合時はエラーを出して停止するため、間違ったデータを消すことはありません。
・画面遷移の改善：明細の修正や削除後も、選択中のレシートが保持されます。また、行の選択状態のみがクリアされるため、続けて次の明細をスムーズに修正できます。
・リロード機能：個別明細の修正・削除後は、レシート全体の合計金額などが即座に自動更新されます。
・安全な削除：削除ボタンを押すと再確認（ポップオーバー）が表示されるため、誤操作を防げます。

【6. AI相談（専属FP）】
・概要：あなたの実際の支出データを元に、AIがプロのファイナンシャルプランナーとして分析や節約のアドバイスを行います。
・音声入力：チャット入力欄の右側にある「🎤（マイク）ボタン」を押すと、声で直接相談内容を入力することができます。
・スマート回答：AIはまず簡潔な結論を述べ、「詳細内容をご覧になりますか？」と確認してから詳細を提示する対話型ロジックを採用しています。

【7. プロフィール設定】
・概要：AI相談をよりパーソナライズするための基本情報（氏名、性別、生年月日、MBTI、職業、趣味、ライフスタンス、AI相談の基本指示）を設定できます。
・効果：ここで設定した情報をAIが読み込み、あなた個人の価値観やライフスタイル、さらに「基本指示」で指定した好みの振る舞いに沿った、より質の高いアドバイスを提供します。

【8. データのダウンロード】
・概要：登録された家計データをファイルとして保存できます。
・操作：左側のサイドバー下部にある「📥 データのダウンロード」をクリックし、Excel形式またはCSV形式を選択、「🚀 実行する」を押してください。
・活用：PCでの詳細な分析や、データのバックアップとして利用可能です。

回答のコツ：
・各機能への移動は、画面左側の「サイドバー（メニュー）」から行えることを案内してください。
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
            st.info("家計簿アプリ「マイニー」の全機能と操作方法をこちらで確認できます。")
            
            with st.expander("📊 ダッシュボード（月次集計）", expanded=True):
                st.markdown("""
                **概要**: 月間の総支出、予算、日別の推移をグラフで可視化します。
                - **3つの分析パターン**: 画面中央のボタンで「店舗別」「大分類別」「小分類別」を切り替え可能です。
                - **2段階表示**: 項目をクリックすると、さらに詳細な内訳が表示されます。
                - **並び替え**: 常に「金額の高い順」に並ぶため、節約ポイントがすぐに見つかります。
                - **カラー同期**: 円グラフと積上げ棒グラフで同じカテゴリには同じ色が適用されます。
                - **絞り込み**: 月次ナビゲーションで過去のデータも簡単に振り返れます。
                """)

            with st.expander("📊 ダッシュボード（年次集計）"):
                st.markdown("""
                **概要**: 選択した「年」全体の支出データを集計・分析します。
                - **前年対比棒グラフ**: 今年度と前年度の支出を月ごとに並べて、支出の増減を視覚的に把握できます。
                - **年次大分類別シェア**: 1年間の総支出における各カテゴリの割合を円グラフで確認できます。
                - **リンク形式の年選択**: 「◀ 前年」「翌年 ▶」のリンクで、簡単に集計対象の年を切り替えられます。
                - **年次カテゴリ別内訳**: 年間を通した支出の詳細を、月次と同様のアコーディオン形式で追跡できます。
                """)
                
            with st.expander("📅 カレンダー機能"):
                st.markdown("""
                **概要**: 日付ごとの支出額をカレンダー形式で一覧できます。
                - **詳細確認**: 日付をクリックすると、その日の「支出明細」が下に表示されます。
                - **多角的な分析**: カレンダー内でも「店舗別」「大分類別」「小分類別」の切り替えが可能です。
                - **カラー表示**: 土曜日は青、日曜・祝日は赤で表示され、視認性を高めています。
                """)

            with st.expander("📸 レシート取込（AI解析）"):
                st.markdown("""
                **概要**: レシートの写真を撮ってアップロードするだけで、AIが内容を読み取ります。
                - **自動向き補正**: アップロードされた画像の向きをEXIF情報に基づいて自動的に正しく（縦向きに）調整します。
                - **自動解析**: 店舗名、商品名、金額、カテゴリをAIが自動で推測して入力します。
                - **自動消費税追加**: 解析結果に消費税が含まれていない場合、システムが自動的に内税10%の消費税項目を計算して追加します。
                - **編集と登録**: 解析完了後の確認画面で、「日付」をカレンダーから、「店舗名」をテキスト入力で直感的に修正できます。未入力での誤登録を防ぐチェック機能も搭載しています。
                """)

            with st.expander("⌨️ レシート手入力（高速入力）"):
                st.markdown("""
                **概要**: キーボード操作で素早く支出を入力できます。
                - **自動行追加**: 金額を入力して `Enter` または `Tab` キーを押すと、自動で次の行が作成されます。
                - **柔軟な登録**: 空白の行があっても、入力済みのデータのみを正確に登録します。
                - **行削除**: 右端の `✕` ボタンで、特定の行だけを削除できます。
                """)

            with st.expander("✏️ レシート修正・履歴管理"):
                st.markdown("""
                **概要**: 過去に登録した全てのデータを一覧・検索・編集できます。
                - **一括管理**: 全ての支出データが時系列で表示されます。
                - **かんたん修正**: 修正内容を入力して「更新」または「登録実行」を押すだけ。
                - **遷移改善**: 明細の修正・削除後も対象レシートの選択が維持され、続けて次の修正が行えます。
                - **自動リロード**: 個別明細の操作後、一覧の合計表示などが自動的に最新の状態に更新されます。
                - **UI改善**: 「日付・店舗名更新」ボタンの名称変更や、選択指示の明確化（「選択して下さい」）により、使いやすさが向上しました。
                - **安全な削除**: 削除時は再確認が出るため、誤操作を防げます。
                """)

            with st.expander("🤖 AI相談（専属FP）"):
                st.markdown("""
                **概要**: あなたの実際の支出データを基に、AIがプロのファイナンシャルプランナーとして分析やアドバイスを行います。
                - **音声入力**: 入力欄右側のマイクボタンで、タイピング不要で声による相談が可能です。
                本アプリで最も活用していただきたい、パーソナライズされたコンサルティング機能です。

                - **✨ あなたのデータを深く理解**:
                    - 「先月と比べて外食費が増えた理由は？」といった分析。
                    - 「今のペースで使うと、今月の残予算はどうなる？」といった予測。
                    - 「どこを削れば、もっと趣味にお金を回せる？」といった具体的な改善提案。
                
                - **👤 プロフィール連動型の回答**:
                    - 設定した「職業」「趣味」「ライフスタンス」に加えて、新機能の「AI相談の基本指示」をAIが常に把握しています。
                    - 一般論ではなく、「あなたならこうすべき」という、背景とユーザーの好みに寄り添ったアドバイスを提供します。
                
                - **🤖 スマート・メッセージング**:
                    - 忙しい時でもすぐに内容を把握できるよう、AIはまず簡潔な結論から話します。
                    - 詳細なデータを見たい場合は「はい」と答えることで、深掘りした分析結果が表示されます。
                
                - **💡 使いこなしのコツ**:
                    - **具体的に聞く**: 「1万円節約したい」など具体的であればあるほど、AIは正確なプランを提示できます。
                    - **雑談もOK**: 「最近の物価高についてどう思う？」など、家計にまつわる気軽な相談も歓迎です。
                
                - **🎤 音声入力にも対応**:
                    - 画面下のマイクボタンを使って、スマホからでも手軽に話しかけることができます。
                """)

            with st.expander("❓ ヘルプチャット"):
                st.markdown("""
                **概要**: アプリの使い方で困ったら、チャットで何でも質問できます。
                - **操作相談**: 「レシートの修正はどうやるの？」など、操作に関する疑問を解決します。
                """)

            with st.expander("⚙️ プロフィール設定"):
                st.markdown("""
                **概要**: AI相談のアドバイスをよりパーソナライズするための情報を登録・管理します。
                - **パーソナライズ**: 入力した情報（職業、趣味、大切にしていること等）をAIが事前に把握し、一般的なアドバイスではなく「あなたのため」の親身なコンサルティングを実現します。
                - **AI相談の基本指示**: AIの話し方（語尾、厳しさ等）や回答スタイルを直接指定できるようになりました。
                """)

            with st.expander("📥 データのダウンロード"):
                st.markdown("""
                **概要**: 登録したすべての支出データを、自分の端末に保存できます。
                - **2つの形式**: 「Excel（.xlsx）」または「CSV（.csv）」から選択可能です。
                - **保存場所**: サイドバーの一番下に専用のボタンがあります。
                - **活用方法**: 表計算ソフトでの詳細な分析や、万が一のためのバックアップに活用してください。
                """)

            st.markdown("---")
            st.caption("マイニー Ver 4.1.7 - ユーザー: %s" % st.session_state['username'])
            
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
