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
from PIL import Image, ImageOps
from google import genai
from google.genai import types
import streamlit.components.v1 as components
import time
import re

# ---------- 構成設定 ----------
SPREADSHEET_NAME = "Kakeibo_Data" # 実際のGoogleスプレッドシート名に合わせて変更してください
WORKSHEET_NAME = "users"
TRANSACTIONS_WORKSHEET_NAME = "transactions"

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
    "割引・ポイント利用": "#000000" # 黒
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

# ブラウザの自動翻訳の誤作動を防ぐため、HTMLの言語設定と翻訳拒否設定を強化
components.html(
    """
    <script>
        // HTML要素の言語を日本語に固定
        const html = window.parent.document.getElementsByTagName('html')[0];
        html.setAttribute('lang', 'ja');
        html.setAttribute('translate', 'no');
        html.classList.add('notranslate');

        // headにmetaタグを挿入してGoogle翻訳を明示的に拒否
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

# 全体のスタイルとしても翻訳拒否を適用
st.markdown("""
    <style>
        /* アプリ全体で翻訳を無効化 */
        .stApp {
            unicode-bidi: isolate;
        }
        /* classNameを使った拒否（一部ブラウザ向け） */
        .notranslate {
            translate: no !important;
        }
    </style>
""", unsafe_allow_html=True)

# ---------- セッション状態の初期化 ----------
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'username' not in st.session_state:
    st.session_state['username'] = None
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

def get_sheet(worksheet_name):
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
    try:
        headers = sheet.row_values(1)
        if not headers or headers[0] != "username":
            sheet.insert_row(["username", "date", "store_name", "item_name", "category", "subcategory", "amount"], 1)
        # 既存シートで subcategory 列がない場合でも、順次追加で対応可能とする
    except Exception:
        sheet.insert_row(["username", "date", "store_name", "item_name", "category", "subcategory", "amount"], 1)

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
        "小分類": "subcategory"
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
    # レコード取得にリトライを適用
    records = safe_gspread_call(sheet.get_all_records)
    
    # 共通ヘルパーでクレンジングとフィルタリング
    curr_user = st.session_state.get('username', "")
    df = get_clean_df(records, curr_user)
    
    if df.empty:
         return pd.DataFrame()
    
    # 行インデックスの付与（recordsの順番に基づく）
    # recordsのインデックスとdfのインデックスを合わせる必要があるため、クレンジング前のrecords長を使用
    # recordsは全ユーザー分あるが、dfはフィルタ済み。
    # recordsにある元の行番号を保持するためにDataFrame作成時に付与しておく
    df_all_temp = pd.DataFrame(records)
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
    records = safe_gspread_call(sheet.get_all_records)
    
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
            components.html(js, height=0)
            
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
def parse_receipt_with_gemini(image_file):
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

# ---------- 音声機能関連のユーティリティ ----------
def clean_text_for_speech(text):
    """音声読み上げ用にマークダウン記号などをクリーンアップする"""
    if not text:
        return ""
    
    # 1. 太字や斜体のアスタリスクを削除
    clean_text = text.replace("**", "").replace("*", "")
    
    # 2. 見出し記号（#）を削除
    clean_text = re.sub(r'#+\s?', '', clean_text)
    
    # 3. 箇条書き記号（行頭の - または *）を「項目、」に置換、または削除して読点に
    #    行頭の「- 」や「* 」を「項目、」に変える正規表現
    clean_text = re.sub(r'^[ \t]*[-*][ \t]+', '項目、', clean_text, flags=re.MULTILINE)
    
    return clean_text

def render_speech_synthesis_button(text, key):
    """テキストを読み上げるスピーカーボタンを表示する"""
    if not text:
        return
    
    # JavaScriptによる読み上げロジック
    # マークダウン記号のクリーンアップ
    target_text = clean_text_for_speech(text)
    
    # JS用にエスケープと改行の除去
    js_text = target_text.replace("'", "\\'").replace("\n", " ")
    
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
            const uttr = new SpeechSynthesisUtterance('{js_text}');
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
    components.html(html_code, height=45)

def render_voice_input_button(key_prefix):
    """音声入力ボタンを表示し、結果をセッション状態に返す"""
    # Streamlitのセッション状態との橋渡し用hidden field
    input_key = f"{key_prefix}_voice_input_result"
    
    html_code = f"""
    <script>
    (function() {{
        const parentDoc = window.parent.document;
        function injectMic() {{
            // stChatInput コンテナを探す
            const chatInputContainer = parentDoc.querySelector('div[data-testid="stChatInput"]');
            if (!chatInputContainer) return;
            
            // 既に存在するかチェック
            if (parentDoc.getElementById('mic-container-{key_prefix}')) return;

            // コンテナ自体のスタイルを調整（絶対配置の基準にするため）
            chatInputContainer.style.position = 'relative';

            const micContainer = parentDoc.createElement('div');
            micContainer.id = 'mic-container-{key_prefix}';
            micContainer.style.position = 'absolute';
            micContainer.style.top = '6px';
            micContainer.style.right = '52px';
            micContainer.style.zIndex = '999999';
            micContainer.style.display = 'flex';
            micContainer.style.alignItems = 'center';
            micContainer.style.pointerEvents = 'auto';
            
            micContainer.innerHTML = `
                <style>
                    #mic-btn-{key_prefix}:active {{
                        transform: scale(0.9);
                        background-color: #e8eaed !important;
                    }}
                    @keyframes pulse-red {{
                        0% {{ box-shadow: 0 0 0 0 rgba(255, 75, 75, 0.7); }}
                        70% {{ box-shadow: 0 0 0 10px rgba(255, 75, 75, 0); }}
                        100% {{ box-shadow: 0 0 0 0 rgba(255, 75, 75, 0); }}
                    }}
                    .pulsing {{
                        animation: pulse-red 1.5s infinite;
                        background-color: #ff4b4b !important;
                        color: white !important;
                        border-color: #ff4b4b !important;
                    }}
                </style>
                <button id="mic-btn-{key_prefix}" style="
                    background-color: #f8f9fa;
                    border: 1px solid #dadce0;
                    border-radius: 50%;
                    width: 38px;
                    height: 38px;
                    cursor: pointer;
                    font-size: 22px;
                    padding: 0;
                    margin: 0;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    box-shadow: 0 1px 3px rgba(60,64,67,0.3);
                    transition: all 0.1s;
                    -webkit-tap-highlight-color: transparent;
                " title="音声入力">🎤</button>
                <span id="status-{key_prefix}" style="margin-right: 8px; font-size: 11px; color: #ff4b4b; background: rgba(255,255,255,0.9); border-radius: 3px; padding: 0 4px; white-space: nowrap; font-weight: bold; display: none; order: -1; pointer-events: none;">認識中...</span>
            `;
            
            chatInputContainer.appendChild(micContainer);

            const btn = micContainer.querySelector('#mic-btn-{key_prefix}');
            const status = micContainer.querySelector('#status-{key_prefix}');

            const startRecognition = () => {{
                try {{
                    // HTTPS または localhost チェック
                    if (window.parent.location.protocol !== 'https:' && window.parent.location.hostname !== 'localhost') {{
                        alert('音声入力にはHTTPS通信（またはlocalhost）が必要です。現在の接続（' + window.parent.location.protocol + '）では動作しません。');
                        return;
                    }}

                    const SpeechRecognition = window.parent.SpeechRecognition || window.parent.webkitSpeechRecognition;
                    if (!SpeechRecognition) {{
                        alert('お使いのブラウザは音声認識に対応していません。ChromeやSafariの最新版をお試しください。');
                        return;
                    }}

                    const recognition = new SpeechRecognition();
                    recognition.lang = 'ja-JP';
                    recognition.interimResults = false;
                    recognition.maxAlternatives = 1;

                    recognition.onstart = () => {{
                        btn.classList.add('pulsing');
                        status.style.display = 'inline';
                    }};

                    recognition.onresult = (event) => {{
                        const result = event.results[0][0].transcript;
                        
                        // 親ウィンドウのテキストエリアを探す
                        const input = parentDoc.querySelector('textarea[data-testid="stChatInputTextArea"]');
                        if (input) {{
                            const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
                            nativeInputValueSetter.call(input, result);
                            input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                            
                            // 送信
                            setTimeout(() => {{
                                const submitBtn = parentDoc.querySelector('button[data-testid="stChatInputSubmitButton"]');
                                if (submitBtn && !submitBtn.disabled) submitBtn.click();
                            }}, 300);
                        }}
                    }};

                    recognition.onerror = (event) => {{
                        btn.classList.remove('pulsing');
                        status.style.display = 'none';
                        if (event.error === 'not-allowed') {{
                            alert('マイクの使用が許可されていません。ブラウザの設定でマイクを許可してください。');
                        }} else {{
                            console.error('Speech recognition error:', event.error);
                        }}
                    }};

                    recognition.onend = () => {{
                        btn.classList.remove('pulsing');
                        status.style.display = 'none';
                    }};

                    recognition.start();
                }} catch (e) {{
                    alert('エラーが発生しました: ' + e.message);
                }}
            }};

            btn.onclick = (e) => {{
                e.preventDefault();
                startRecognition();
            }};
            
            // スマホ向けのタッチイベント追加
            btn.ontouchstart = (e) => {{
                e.preventDefault(); // クリックイベントとの重複防止
                startRecognition();
            }};
        }}
        
        // 定期的にチェックして再挿入（Streamlitの再レンダリング対策）
        setInterval(injectMic, 1500);
        injectMic();
    }})();
    </script>
    """
    
    # 認識結果を受け取るためのコンポーネント
    # 注意: Streamlit公式のiframeから親への通信は制限があるため、
    # 実際にはURLパラメータや、カスタムコンポーネントライブラリなしでは少し工夫が必要。
    # ここでは、認識されたテキストを一時的に表示し、ユーザーが確認して送信できるスタイルにするか、
    # あるいは直接セッションに書き込むための「隠しボタン」的なアプローチをとる。
    
    components.html(html_code, height=0)
    
    # 結果を受け取るための実験的な仕組み
    # (実際には st.chat_input に自動で流し込むのはJSのセキュリティ制約上難しいため、
    # 音声認識されたテキストを通知として表示し、それを入力欄に反映させるガイドを出すのが現実的)
    return None

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
            ["大分類別", "小分類別", "店舗別"], 
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

    if group_col and group_col in df.columns and "amount" in df.columns:
        if graph_type == "円グラフ":
            # 選択された軸が店舗別なら内税を除外し、大分類・小分類なら含めて表示する（二重計上防止だが内訳は正しく出す）
            df_for_chart = df_agg if analysis_axis == "店舗別" else df
            grouped_df = df_for_chart.groupby(group_col, as_index=False)["amount"].sum()
            # 金額が0以下のデータ（マイナス値や0円）を除外
            grouped_df = grouped_df[grouped_df["amount"] > 0]
            grouped_df = grouped_df.sort_values(by="amount", ascending=False)
            
            fig = px.pie(
                grouped_df, 
                values='amount', 
                names=group_col, 
                hole=0.4, 
                title=title_label,
                category_orders={group_col: grouped_df[group_col].tolist()},
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
            ["大分類別", "小分類別", "店舗別"], 
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
        st.plotly_chart(fig, use_container_width=True)

    elif graph_type == "棒グラフ":
        # 当年の月別推移 (積上げ棒グラフ)
        df_for_bar = df_agg.copy() if analysis_axis == "店舗別" else df.copy()
        df_for_bar['month'] = df_for_bar['date'].dt.month
        df_for_bar['month_label'] = df_for_bar['month'].apply(lambda x: f"{x}月")
        
        if group_col and group_col in df_for_bar.columns:
            yearly_grouped = df_for_bar.groupby(['month', 'month_label', group_col], as_index=False)["amount"].sum()
            cat_sum = yearly_grouped.groupby(group_col)["amount"].sum().sort_values(ascending=False).index.tolist()
            
            fig = px.bar(
                yearly_grouped,
                x='month_label',
                y='amount',
                color=group_col,
                title=f"{selected_year}年 {title_label}推移 (積上げ棒グラフ)",
                labels={"amount": "金額", "month_label": "月", group_col: analysis_axis[:-1]},
                category_orders={"month_label": [f"{i}月" for i in range(1, 13)], group_col: cat_sum},
                color_discrete_map=CATEGORY_COLOR_MAP
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 年間合計は内税除外
            year_total = df_agg["amount"].sum()
            st.metric(f"{selected_year}年 総支出額", f"￥{int(year_total):,}")
        else:
            st.warning(f"分析に必要な列（{analysis_axis[:-1]}）がありません。")

    else: # 円グラフ
        if group_col and group_col in df_agg.columns:
            cat_grouped = df_agg.groupby(group_col, as_index=False)["amount"].sum()
            # 金額が0以下のデータ（マイナス値や0円）を除外
            cat_grouped = cat_grouped[cat_grouped["amount"] > 0]
            cat_grouped = cat_grouped.sort_values(by="amount", ascending=False)
            
            fig_pie = px.pie(cat_grouped, values='amount', names=group_col, hole=0.4,
                             title=f'{selected_year}年 {title_label}支出シェア',
                             category_orders={group_col: cat_grouped[group_col].tolist()},
                             color=group_col,
                             color_discrete_map=CATEGORY_COLOR_MAP)
            fig_pie.update_traces(
                textposition='inside', 
                textinfo='percent+label'
            )
            st.plotly_chart(fig_pie, use_container_width=True)
            
            year_total = cat_grouped["amount"].sum()
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
    
    # サイドバーを閉じるフラグ
    st.session_state["collapse_sidebar_flag"] = True

def main():
    # URLパラメータの同期（セッション維持のため冒頭で行う）
    params = st.query_params
    
    # ログイン状態の復元
    if "user" in params:
        st.session_state["username"] = params["user"]
        st.session_state["logged_in"] = True
        
    # --- 初期化およびエラー防止 ---
    if 'menu_selection' in st.session_state:
        # 古いメニュー名（▶付き）が残っている場合の自動変換
        mapping = {
            "レシート手入力": "レシート手入力",
            "レシート修正": "レシート修正"
        }
        if st.session_state['menu_selection'] in mapping:
            st.session_state['menu_selection'] = mapping[st.session_state['menu_selection']]
            
    selected_date_str = None # UnboundLocalError防止
    if "date" in params:
        # 指定された日付を取得
        selected_date_str = params["date"]
        st.session_state['selected_date'] = selected_date_str
        
        # 明示的にメニューが指定されている場合はそれに従う
        if "menu" in params:
            st.session_state['menu_selection'] = params["menu"]
        else:
            # メニュー指定がない場合（カレンダーの日付クリック等）はカレンダー画面へ
            st.session_state['menu_selection'] = "カレンダー"

        # URLのdateから表示月(current_month)を自動同期
        try:
            dt = datetime.strptime(selected_date_str, '%Y-%m-%d')
            st.session_state['current_month'] = dt.replace(day=1)
        except:
            pass
            
        # URLのパラメータを整理（一度適用したら不必要なリロードを防ぐための配慮が必要な場合もあるが、
        # 現状はリンク方式のため、このままセッションに保持する）

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
            st.subheader("マイニー [Ver 3.2.9]")
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
                "👁AI相談", 
                "ヘルプ", 
                "📗マニュアル"
            ]
            
            # メニューのリセット処理（別の画面から戻ってきたとき用）
            # もしカレンダー等から「ダッシュボード系以外」を経由して戻ってきた場合、
            # 次にダッシュボードをクリックしたときに「月次」にしたいという要件。
            # 直前の値を保持しておき、遷移を検知する
            if "last_menu_selection" not in st.session_state:
                st.session_state.last_menu_selection = st.session_state['menu_selection']
            
            menu_selection = st.radio(
                "機能を選択",
                menu_options,
                key="menu_selection",
                on_change=handle_menu_change
            )
            st.session_state.last_menu_selection = menu_selection

            # サイドバー自動折りたたみJS
            if st.session_state.get("collapse_sidebar_flag"):
                st.session_state["collapse_sidebar_flag"] = False
                st.components.v1.html(
                    """
                    <script>
                    var windowParent = window.parent;
                    var buttons = windowParent.document.querySelectorAll('button');
                    for (var i = 0; i < buttons.length; i++) {
                        if (buttons[i].getAttribute('aria-label') === 'Collapse sidebar') {
                            buttons[i].click();
                            break;
                        }
                    }
                    </script>
                    """,
                    height=0
                )
            
            st.markdown("---")
            if st.button("ログアウト", use_container_width=True):
                # ログアウト時にURLパラメータとセッション状態を完全にクリアする
                st.query_params.clear()
                st.session_state.clear()
                st.rerun()

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
                        amount = daily_totals.get(day, 0)
                        amount_text = f"￥{int(amount):,}" if amount > 0 else ""
                        date_obj = datetime(year, month, day).date()
                        date_str = date_obj.strftime('%Y-%m-%d')
                        is_selected = st.session_state.get('selected_date') == date_str
                        select_cls = "selected-link" if is_selected else ""
                        current_user = st.session_state.get("username", "")
                        
                        # 曜日および祝日による背景色の判定 (i: 0=日, 6=土)
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
                                
                            if isinstance(results, list) and len(results) > 0 and isinstance(results[0], dict) and "error" in results[0]:
                                st.error(f"解析に失敗しました: {results[0]['error']}")
                            elif isinstance(results, dict) and "error" in results:
                                st.error(f"解析に失敗しました: {results['error']}")
                            else:
                                st.session_state.parsed_results = results
                                st.rerun()
                        except Exception as e:
                            st.error(f"解析処理中に予期せぬエラーが発生しました: {e}")
                
                else:
                    # 解析完了後、プレビューと確認画面を表示
                    results = st.session_state.parsed_results
                    
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
                    st.write(f"**日付**: {preview_date}")
                    st.write(f"**店舗**: {preview_store}")
                    st.write(f"**合計金額**: ￥{total_amount:,}")
                    
                    # DataFrameで一覧表示
                    cat_df = pd.DataFrame([
                        {"大分類": k, "金額": f"￥{v:,}"} for k, v in category_totals.items()
                    ])
                    st.dataframe(cat_df, hide_index=True, use_container_width=True)
                    
                    st.markdown("---")
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
                                            
                                    store_name = str(item.get("store_name", ""))
                                    item_name = str(item.get("item_name", ""))
                                    
                                    # 日付を yyyy-mm-dd に整形
                                    raw_date = item.get("date", "")
                                    formatted_date = ""
                                    try:
                                        if isinstance(raw_date, datetime):
                                            formatted_date = raw_date.strftime("%Y-%m-%d")
                                        else:
                                            formatted_date = pd.to_datetime(raw_date).strftime("%Y-%m-%d")
                                    except:
                                        formatted_date = str(raw_date)

                                    row_data = [
                                        str(st.session_state['username']),
                                        formatted_date,
                                        str(store_name),
                                        str(item_name),
                                        str(final_major),
                                        str(final_minor),
                                        int(item.get("amount", 0))
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
                    # 入力されているデータのみを抽出（商品名があり、かつ金額が 0 ではないもの）
                    valid_items = [itm for itm in st.session_state.manual_input_items if itm["name"].strip() != "" and itm["amount"] != 0]

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
                                        int(itm["amount"])
                                    ]
                                    sheet.append_row(row_data)
                                
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
                                    st.session_state['edit_data'][row_index_gs] = {
                                        "name": row.get(item_col, "不明な商品") if item_col else "不明な商品",
                                        "amount": int(row.get("amount", 0)),
                                        "major": major,
                                        "minor": sub
                                    }

                            # --- レシートヘッダー（日付・店舗名）の修正エリア ---
                            st.write("##### レシート情報の修正")
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
                                            for r_idx in existing_indices:
                                                sheet.update_cell(r_idx, 2, target_date_str)
                                                sheet.update_cell(r_idx, 3, target_store)
                                                
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
                                    "金額": st.column_config.NumberColumn(width="small", format="￥%d")
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
                                    "minor": "❓その他"
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
                                            elif edit_amount == 0:
                                                st.warning("⚠️ 金額を入力してください。")
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
                                                            new_row = [user_name, target_date_str, target_store, edit_name, edit_major, edit_minor, edit_amount]
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
                                                            # A:username, B:date, C:store, D:item, E:major, F:minor, G:amount
                                                            # 更新範囲: B (Col 2) から G (Col 7)
                                                            update_range = f"B{r_idx}:G{r_idx}"
                                                            update_values = [[target_date_str, target_store, edit_name, edit_major, edit_minor, edit_amount]]
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
                            components.html("""
                            <script>
                            setInterval(() => {
                                const elements = window.parent.document.querySelectorAll('button');
                                elements.forEach(b => {
                                    const text = b.innerText.trim();
                                    if (text === '修正' || text === '登録') {
                                        b.style.backgroundColor = '#007bff';
                                        b.style.color = 'white';
                                        b.style.borderColor = '#007bff';
                                    }
                                    if (text === '削除' || text === 'キャンセル' || text === '削除実行' || text === 'はい、削除します') {
                                        b.style.backgroundColor = '#dc3545';
                                        b.style.color = 'white';
                                        b.style.borderColor = '#dc3545';
                                    }
                                });
                            }, 500);
                            </script>
                            """, height=0, width=0)

                            
        elif menu_selection == "👁AI相談":
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
            
            # 音声入力ボタン
            render_voice_input_button("ai_consult")
            
            # 音声入力結果をチャット入力欄に反映させるためのJS
            components.html("""
                <script>
                window.parent.document.addEventListener('voiceInput', function(e) {
                    const input = window.parent.document.querySelector('textarea[data-testid="stChatInputTextArea"]');
                    if (input) {
                        // Streamlit(React)の内部状態と同期させるためのハック
                        const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
                        nativeInputValueSetter.call(input, e.detail);
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        
                        // 送信ボタンをクリックして自動送信
                        setTimeout(() => {
                            const btn = window.parent.document.querySelector('button[data-testid="stChatInputSubmitButton"]');
                            if (btn && !btn.disabled) {
                                btn.click();
                            } else {
                                // 送信ボタンが検知できない場合のフォールバック（Enterキーのシミュレート）
                                input.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true}));
                            }
                        }, 500);
                    }
                });
                </script>
            """, height=0)

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
                            # システムプロンプトを都度構築（最新データを反映させるため）
                            system_prompt = f"""あなたはユーザー専属の優秀なファイナンシャルプランナーです。
以下のCSVデータは、このユーザー（{st.session_state['username']}）個人の家計簿データです。
このデータには「商品名」も含まれており、いつ、どこで、何を買ったかを詳細に把握できます。
ユーザーからの「特定の商品の購入時期（例：鶏肉ナンコツはいつ買った？）」や「商品の価格推移」などの質問に対し、正確かつ親身に答えてください。
データに存在しない推測は避け、無駄遣いの指摘や節約のアドバイスなども積極的に行ってください。

【ユーザーの家計簿データ】
{csv_data_string}"""
                            
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
            
            # 音声入力ボタン
            render_voice_input_button("help")

            # 音声入力結果を反映するためのJS（AI相談と同様）
            components.html("""
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
・確認：解析結果を確認・修正して、そのまま家計簿に登録できます。

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
            
        elif menu_selection == "📗マニュアル":
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
                - **編集と登録**: 解析結果を確認・修正し、そのまま家計簿へ登録できます。
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
                **概要**: あなたの支出データに基づき、AIがプロのFPとしてアドバイスします。
                - **パーソナル分析**: 「先月に比べて外食は増えた？」「どこを削ればいい？」など、あなたのデータに沿った会話が可能です。
                """)

            with st.expander("❓ ヘルプチャット"):
                st.markdown("""
                **概要**: アプリの使い方で困ったら、チャットで何でも質問できます。
                - **操作相談**: 「レシートの修正はどうやるの？」など、操作に関する疑問を解決します。
                """)

            st.markdown("---")
            st.caption(f"マイニー Ver 3.2.9 - ユーザー: {st.session_state['username']}")
            
    # 未ログインの状態 (ログイン・登録画面)
    else:
        st.title("AI家計簿アプリ")
        
        # ユーザー名とパスワードを半角英数字のみに制限するJS
        components.html("""
            <script>
            function enforceAlphanumeric() {
                const inputs = window.parent.document.querySelectorAll('input[type="text"], input[type="password"]');
                inputs.forEach(input => {
                    if (!input.dataset.alphanumericEnforced) {
                        input.dataset.alphanumericEnforced = 'true';
                        // スマホ対応：英語キーボードを出しやすくする
                        input.setAttribute('inputmode', 'email');
                        input.setAttribute('autocomplete', 'off');
                        
                        input.addEventListener('input', function(e) {
                            // 全角英数字を半角に変換
                            let val = e.target.value.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
                                return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
                            });
                            // 半角英数字と一部記号以外を削除
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
            // 定期的にチェックして適用
            setInterval(enforceAlphanumeric, 1000);
            </script>
        """, height=0)
        
        tab1, tab2 = st.tabs(["ログイン", "新規ユーザー登録"])
        
        with tab1:
            st.subheader("ログイン")
            with st.form("login_form"):
                login_username = st.text_input("ユーザー名")
                login_password = st.text_input("パスワード", type="password")
                submitted = st.form_submit_button("ログイン")
                
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
                reg_username = st.text_input("新しいユーザー名")
                reg_password = st.text_input("新しいパスワード", type="password")
                reg_password_confirm = st.text_input("パスワード（確認用）", type="password")
                submitted = st.form_submit_button("登録する")
                
                if submitted:
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
