import streamlit as st
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
    "消費税": ["8%", "10%", "❓その他"],
    "その他": ["📁未分類"]
}

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

# ブラウザの自動翻訳の誤作動を防ぐため、大元の言語設定を日本語(ja)に強制上書き
components.html(
    """
    <script>
        const html = window.parent.document.getElementsByTagName('html')[0];
        html.setAttribute('lang', 'ja');
    </script>
    """,
    width=0,
    height=0,
)

# ---------- 翻訳拒否設定（ブラウザの自動翻訳による誤変換防止） ----------
st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)

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
import time

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

# ---------- データ取得機能 ----------
def load_transactions_data(target_month):
    """指定した月・ログインユーザーのデータを取得する"""
    sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
    init_transactions_sheet(sheet)
    # レコード取得にリトライを適用
    records = safe_gspread_call(sheet.get_all_records)
    
    if not records:
         return pd.DataFrame()
         
    for i, r in enumerate(records):
        r['_row_index'] = i + 2  # ヘッダー行を考慮して+2
         
    df = pd.DataFrame(records)
    
    # クレンジング（不要なスペース等削除、データ型変換）
    df.columns = df.columns.str.strip()
    
    # "username"でフィルタ
    if "username" in df.columns:
        df = df[df["username"].astype(str).str.lower() == st.session_state['username']]
    
    if df.empty or "date" not in df.columns:
         return pd.DataFrame()

    # "date"列をdatetime型にするため、エラーは強制的にNaTに
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"])
    
    # 月でフィルタ
    df = df[(df["date"].dt.year == target_month.year) & (df["date"].dt.month == target_month.month)]
    
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

# ---------- レシート解析機能 ----------
def parse_receipt_with_gemini(image_file):
    try:
        img = Image.open(image_file)
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

【消費税の抽出ルール】:
レシート内に「消費税（8%や10%など）」が明細や項目として記載されている場合、その行を1つの明細として抽出し、大分類を "消費税" 、小分類をその税率（"8%" や "10%"など）として設定してください。

【合計金額の整合性ルール】（重要）:
レシートの「合計金額」と、抽出したすべての明細の「金額（amount）」の合計額が、計算上必ず完全に一致するようにしてください。
金額が合わない場合は、明細行・割引や値引（マイナス金額で抽出）・消費税・小計などのいずれかを読み飛ばしているか誤読している可能性があります。読み飛ばしがないよう、すべての金額要素を漏れなく抽出してください。

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

# ---------- ページUIの実装 ----------
def show_dashboard():
    # ヘッダーを表示するためのプレースホルダー（コンテナ）を先に準備
    header_placeholder = st.empty()

    # 月の切替UI
    col1, col2, col3, col4, col5 = st.columns([3, 1, 2, 1, 3])
    with col2:
        # vertical alignment adjustment if needed, but standard is fine
        if st.button("◀ 前月", use_container_width=True):
            st.session_state['current_month'] -= relativedelta(months=1)
            st.rerun()
    with col3:
        # 中央に現在選択中の年月を大きく表示
        st.markdown(f"<h3 style='text-align: center; margin: 0;'>{st.session_state['current_month'].strftime('%Y年%m月')}</h3>", unsafe_allow_html=True)
    with col4:
        if st.button("翌月 ▶", use_container_width=True):
            st.session_state['current_month'] += relativedelta(months=1)
            st.rerun()
            
    # 月の切り替え操作が行われた「後」の最新の状態でヘッダーを更新する
    header_placeholder.markdown("#### ダッシュボード (月別集計)")


    st.markdown("---")

    # データの読み込み
    with st.spinner("データを読み込み中..."):
        df = load_transactions_data(st.session_state['current_month'])

    if df.empty:
        st.info("※この月のデータはまだありません。Google SpreadSheet の `transactions` シートにサンプルデータ（自分のusernameを含める）を追加して確認してください。")
        return

    # グラフと表の表示エリア
    
    # カテゴリ("category")ごとに金額("amount")を合計
    if "category" in df.columns and "amount" in df.columns:
        grouped_df = df.groupby("category", as_index=False)["amount"].sum()
        grouped_df = grouped_df.sort_values(by="amount", ascending=False)
        
        # 円グラフ（ドーナツ型）
        fig = px.pie(
            grouped_df, 
            values='amount', 
            names='category', 
            hole=0.4, 
            title='大分類別金額シェア'
        )
        fig.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)
        
        total_amount = grouped_df["amount"].sum()
        st.metric("総支出額", f"￥{int(total_amount):,}")
        st.markdown("---")
        
        st.markdown("##### カテゴリ別内訳")
        
        # 大分類ごとの一覧をアコーディオン形式（st.expander）で表示
        for _, row in grouped_df.iterrows():
            cat = row['category']
            total_amt_str = f"￥{int(row['amount']):,}"
            
            with st.expander(f"{cat}：{total_amt_str}"):
                # 該当カテゴリのデータを抽出
                cat_df = df[df["category"] == cat].copy()
                
                # 小分類を判別するための列名を探す
                sub_col = None
                for col_name in ["subcategory", "sub_category", "小分類"]:
                    if col_name in cat_df.columns:
                        sub_col = col_name
                        break
                
                if sub_col:
                    sub_grouped = cat_df.groupby(sub_col, as_index=False)["amount"].sum()
                    sub_grouped = sub_grouped.sort_values(by="amount", ascending=False)
                    sub_grouped["amount"] = sub_grouped["amount"].apply(lambda x: f"￥{int(x):,}")
                    sub_grouped.columns = ["小分類", "金額"]
                    st.dataframe(sub_grouped, use_container_width=True, hide_index=True)
                else:
                    # 小分類カラムがない場合は、明細レベルで内訳を表示する
                    if "store_name" in cat_df.columns and "item_name" in cat_df.columns:
                        display_df = cat_df[["date", "store_name", "item_name", "amount"]].copy()
                        if "date" in display_df.columns:
                            display_df["date"] = display_df["date"].dt.strftime('%m/%d').fillna("")
                        display_df["amount"] = display_df["amount"].apply(lambda x: f"￥{int(x):,}")
                        display_df.columns = ["日付", "店舗名", "商品名", "金額"]
                        st.dataframe(display_df, use_container_width=True, hide_index=True)
                    elif "memo" in cat_df.columns:
                        display_df = cat_df[["date", "memo", "amount"]].copy()
                        if "date" in display_df.columns:
                            display_df["date"] = display_df["date"].dt.strftime('%m/%d').fillna("")
                        display_df["amount"] = display_df["amount"].apply(lambda x: f"￥{int(x):,}")
                        display_df.columns = ["日付", "内容", "金額"]
                        st.dataframe(display_df, use_container_width=True, hide_index=True)
                    else:
                        display_df = cat_df[["amount"]].copy()
                        if "date" in cat_df.columns:
                            display_df.insert(0, "date", cat_df["date"].dt.strftime('%m/%d').fillna(""))
                        display_df["amount"] = display_df["amount"].apply(lambda x: f"￥{int(x):,}")
                        st.dataframe(display_df, use_container_width=True, hide_index=True)
    else:
        st.warning("シートに 'category' または 'amount' 列がありません。")

def main():
    # URLパラメータの同期（セッション維持のため冒頭で行う）
    params = st.query_params
    
    # ログイン状態の復元
    if "user" in params:
        st.session_state["username"] = params["user"]
        st.session_state["logged_in"] = True
        
    # 日付選択の同期
    if "date" in params:
        # サイドバーで既に「カレンダー」以外が選択されている場合は、URLのパラメータをクリアしてリロードする
        # これにより、日付選択状態を維持したまま他のメニューへ移動できない問題を解消
        current_menu = st.session_state.get('menu_selection')
        if current_menu not in [None, "カレンダー"]:
            del st.query_params["date"]
            st.rerun()
        else:
            st.session_state['selected_date'] = params["date"]
            st.session_state['menu_selection'] = "カレンダー"

    # ログイン済みの状態
    if st.session_state.get('logged_in', False):
        
        # 自動画面遷移のためのリダイレクト処理
        if st.session_state.get('redirect_to_dashboard'):
            st.session_state['menu_selection'] = "ダッシュボード"
            st.session_state['redirect_to_dashboard'] = False
            
        # サイドバーメニューの実装
        with st.sidebar:
            st.subheader("メインメニュー [Ver 1.7.1 [Rendering Fix]]")
            st.write(f"🔑 ユーザー: **{st.session_state['username']}**")
            st.markdown("---")
            if 'menu_selection' not in st.session_state:
                st.session_state['menu_selection'] = "ダッシュボード"
                
            menu_selection = st.radio(
                "機能を選択",
                ["ダッシュボード", "カレンダー", "レシート取込", "レシート手入力", "レシート修正", "🤖 AI相談", "ヘルプ"],
                key="menu_selection"
            )
            st.markdown("---")
            if st.button("ログアウト", use_container_width=True):
                st.session_state['logged_in'] = False
                st.session_state['username'] = None
                st.rerun()

        # メインコンテンツの切り替え
        if menu_selection == "ダッシュボード":
            show_dashboard()
        elif menu_selection == "カレンダー":
            st.markdown("#### 📅 カレンダー")
            
            # 月選択UI
            col1, col2, col3, col4, col5 = st.columns([3, 1, 2, 1, 3])
            with col2:
                if st.button("◀ 前月", use_container_width=True, key="cal_prev"):
                    st.session_state['current_month'] -= relativedelta(months=1)
                    st.rerun()
            with col3:
                st.markdown(f"<h3 style='text-align: center; margin: 0;'>{st.session_state['current_month'].strftime('%Y年%m月')}</h3>", unsafe_allow_html=True)
            with col4:
                if st.button("翌月 ▶", use_container_width=True, key="cal_next"):
                    st.session_state['current_month'] += relativedelta(months=1)
                    st.rerun()
                    
            st.markdown("---")
            
            # データ取得
            with st.spinner("データを読み込み中..."):
                df = load_transactions_data(st.session_state['current_month'])
                
            daily_totals = {}
            if not df.empty and "date" in df.columns and "amount" in df.columns:
                df['day'] = df['date'].dt.day
                daily_totals = df.groupby('day')["amount"].sum().to_dict()
                
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
    max-width: 250px;
    margin: 0 auto;
}
/* リンクをマス目（枠線付き）として機能させる */
.cal-link {
    display: block;
    position: relative;
    height: 60px;
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
.sun-text { color: #e53e3e; }
.sat-text { color: #3182ce; }
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
                        date_str = f"{year}-{month:02d}-{day:02d}"
                        is_selected = st.session_state.get('selected_date') == date_str
                        select_cls = "selected-link" if is_selected else ""
                        current_user = st.session_state.get("username", "")
                        
                        cal_html += f'<a href="/?date={date_str}&user={current_user}" target="_self" class="cal-link {select_cls} notranslate" translate="no">'
                        cal_html += f'<div class="cal-date notranslate" translate="no">{day}</div>'
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
                day_df = df[df['date'].dt.day == day_val].copy()
                
                # 合計額の計算
                day_total = int(day_df['amount'].sum()) if not day_df.empty else 0
                # 合計額の計算
                day_total = int(day_df['amount'].sum()) if not day_df.empty else 0
                # デザインより翻訳回避を優先し、ネイティブなマークダウンで表示
                st.markdown(f"##### 📋 {display_day}日の支出詳細 (合計: ￥{day_total:,})")
                
                if not day_df.empty:
                    # 翻訳回避のため Canvas 描画である st.dataframe を使用
                    # 表示用に列名を整理
                    display_df = day_df.copy()
                    
                    # 列名のマッピング（存在するものを優先）
                    col_map = {
                        "store_name": "店舗名",
                        "store": "店舗名",
                        "category": "大分類",
                        "major_category": "大分類",
                        "subcategory": "小分類",
                        "minor_category": "小分類",
                        "item_name": "商品名",
                        "amount": "金額"
                    }
                    
                    rename_dict = {}
                    for old_col, new_col in col_map.items():
                        if old_col in display_df.columns:
                            rename_dict[old_col] = new_col
                    
                    display_df = display_df.rename(columns=rename_dict)
                    
                    # フィルタリング機能の追加
                    if "大分類" in display_df.columns:
                        major_cats = sorted(display_df["大分類"].unique().tolist())
                        selected_cats = st.multiselect("大分類で絞り込み (未選択で全表示)", options=major_cats)
                        if selected_cats:
                            display_df = display_df[display_df["大分類"].isin(selected_cats)]

                    # 集計処理（大分類、小分類、店舗名でグループ化して金額を合計）
                    group_cols = [c for c in ["大分類", "小分類", "店舗名"] if c in display_df.columns]
                    if group_cols and "金額" in display_df.columns:
                        display_df = display_df.groupby(group_cols, as_index=False)["金額"].sum()
                    
                    # 表示する列の選択と順序
                    target_cols = ["大分類", "小分類", "店舗名", "金額"]
                    final_cols = [c for c in target_cols if c in display_df.columns]
                    
                    # 大分類、小分類、店舗名の順でソート（昇順）
                    sort_cols = [c for c in ["大分類", "小分類", "店舗名"] if c in display_df.columns]
                    if sort_cols:
                        display_df = display_df.sort_values(by=sort_cols, ascending=True)
                    
                    # 合計行の追加 (フィルタ後の合計を表示)
                    if not display_df.empty and "金額" in display_df.columns:
                        total_amount = int(display_df["金額"].sum())
                        total_row = pd.DataFrame([{
                            "大分類": "---",
                            "小分類": "---",
                            "店舗名": "合計",
                            "金額": total_amount
                        }])
                        display_df = pd.concat([display_df, total_row], ignore_index=True)
                    
                    st.dataframe(
                        display_df[final_cols],
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("この日の支出データはありません。")
            
            st.markdown("---")
            
        elif menu_selection == "レシート取込":
            st.markdown("#### 📸 レシート取込")
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
                st.image(uploaded_file, caption="取得したレシート画像", width=300)
                
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
                        
                    total_amount = sum(int(item.get("amount", 0)) for item in results)
                    
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
                                    
                                    row_data = [
                                        st.session_state['username'],
                                        str(item.get("date", "")),
                                        store_name,
                                        item_name,
                                        final_major,
                                        final_minor,
                                        int(item.get("amount", 0))
                                    ]
                                    sheet.append_row(row_data)
                                    written_count += 1
                                
                                st.session_state.flash_message = f"✅ 解析が完了し、{written_count}件のデータを保存しました！"
                                
                                import time
                                time.sleep(1)
                                
                                st.session_state.parsed_results = None
                                st.session_state.uploader_key += 1
                                st.rerun()
                                
                        except Exception as e:
                            st.error(f"保存エラー: {e}")

        elif menu_selection == "レシート手入力":
            st.markdown("#### レシート手入力")
            st.info("準備中: 手動でのレシート入力機能は今後のフェーズで実装されます。")
        elif menu_selection == "レシート修正":
            st.markdown("#### ⚙️ レシート修正")
            
            # 月の切替UI
            col1, col2, col3, col4, col5 = st.columns([3, 1, 2, 1, 3])
            with col2:
                if st.button("◀ 前月", key="prev_mod_btn", use_container_width=True):
                    st.session_state['current_month'] -= relativedelta(months=1)
                    st.rerun()
            with col3:
                st.markdown(f"<h3 style='text-align: center; margin: 0;'>{st.session_state['current_month'].strftime('%Y年%m月')}</h3>", unsafe_allow_html=True)
            with col4:
                if st.button("翌月 ▶", key="next_mod_btn", use_container_width=True):
                    st.session_state['current_month'] += relativedelta(months=1)
                    st.rerun()
                    
            st.markdown("---")
            
            with st.spinner("データを読み込み中..."):
                df = load_transactions_data(st.session_state['current_month'])
                
            if df.empty:
                st.info("※この月のデータはありません。")
            else:
                # 店舗名と商品名のカラムを動的に判定
                store_col = "store_name" if "store_name" in df.columns else "store" if "store" in df.columns else None
                item_col = "item_name" if "item_name" in df.columns else "item" if "item" in df.columns else "items" if "items" in df.columns else None
                
                if not store_col:
                    st.warning("スプレッドシートに店舗名（'store_name' または 'store'）の列が見つかりません。")
                else:
                    # レシート単位に集約（日付と店舗名が同じものを同一レシートとみなす）
                    receipts_df = df.groupby(["date", store_col], as_index=False).agg(
                        amount=("amount", "sum"),
                        明細数=("amount", "count")
                    )
                    receipts_df.columns = ["日付", "店舗名", "金額合計", "明細数"]
                    receipts_df["日付"] = receipts_df["日付"].dt.strftime('%Y-%m-%d')
                    receipts_df["金額合計"] = receipts_df["金額合計"].apply(lambda x: int(x))
                    receipts_df = receipts_df.sort_values(by="日付", ascending=False).reset_index(drop=True)
                    
                    # 総合計行を追加
                    total_amount = receipts_df["金額合計"].sum()
                    total_items = receipts_df["明細数"].sum()
                    total_row = pd.DataFrame([{
                        "日付": "総合計",
                        "店舗名": "",
                        "金額合計": int(total_amount),
                        "明細数": int(total_items)
                    }])
                    receipts_df = pd.concat([receipts_df, total_row], ignore_index=True)
                    
                    st.write("##### ▶レシート一覧表（対象レシートを選択してください）")
                    
                    # dataframe 選択
                    event = st.dataframe(
                        receipts_df, 
                        use_container_width=True, 
                        hide_index=True, 
                        selection_mode="single-row",
                        on_select="rerun"
                    )
                    
                    if len(event.selection.rows) > 0:
                        selected_idx = event.selection.rows[0]
                        selected_receipt = receipts_df.iloc[selected_idx]
                        
                        # 総合計行が選択された場合は詳細表示しない
                        if selected_receipt["日付"] != "総合計":
                            target_date = pd.to_datetime(selected_receipt["日付"])
                            target_store = selected_receipt["店舗名"]
                            
                            st.markdown("---")
                            st.write(f"##### 対象レシート明細： {selected_receipt['日付']} - {target_store}")
                            
                            # 該当レシートの明細を取得
                            details = df[(df["date"] == target_date) & (df[store_col] == target_store)].copy()
                            
                            # session_state 上の変更状態を初期化（対象が変わった場合用）
                            receipt_key = f"{selected_receipt['日付']}_{target_store}"
                            if st.session_state.get('current_receipt_key') != receipt_key:
                                st.session_state['current_receipt_key'] = receipt_key
                                st.session_state['edit_data'] = {}
                            
                            for idx, row in details.iterrows():
                                row_index_gs = row["_row_index"]
                                
                                if row_index_gs not in st.session_state['edit_data']:
                                    major = row.get("category", "その他")
                                    # subcategoryカラムの特定
                                    sub_cols = [c for c in ["subcategory", "sub_category", "小分類"] if c in df.columns]
                                    sub = row.get(sub_cols[0], "❓その他") if sub_cols else "❓その他"
                                    
                                    st.session_state['edit_data'][row_index_gs] = {
                                        "amount": int(row.get("amount", 0)),
                                        "major": major,
                                        "minor": sub
                                    }
                                    
                            st.write("##### 明細一覧")
                            
                            # 閲覧モードと修正モードを管理するState
                            # レシートが切り替わった時に状態をリセットするためのキー制御
                            mode_key = f"mode_{receipt_key}"
                            if mode_key not in st.session_state:
                                st.session_state[mode_key] = False # 初期設定は閲覧モード(False)
                                
                            edit_mode = st.session_state[mode_key]
                            
                            if edit_mode:
                                # 【修正モード】のレイアウト
                                
                                # 明細行を1行のインラインテキストのように表示させるためのCSS
                                st.markdown("""
                                <style>
                                    /* ウィジェット下マージンを消去して余白を完全削除 */
                                    div[data-testid="stVerticalBlock"]:has(span#receipt-table-target):not(:has(div[data-testid="stVerticalBlock"] span#receipt-table-target)) div.stMarkdown,
                                    div[data-testid="stVerticalBlock"]:has(span#receipt-table-target):not(:has(div[data-testid="stVerticalBlock"] span#receipt-table-target)) div.stPopover {
                                        margin-bottom: 0 !important;
                                    }
                                    
                                    /* ポップオーバー（大分類・小分類ボタン）の表示を極力コンパクトに */
                                    div[data-testid="stPopover"] > button {
                                        padding: 0px 4px !important;
                                        font-size: 0.8em !important;
                                        min-height: 24px !important;
                                        height: 26px !important;
                                        width: auto !important;
                                    }
                                </style>
                                """, unsafe_allow_html=True)
                                
                                with st.container():
                                    st.markdown('<span id="receipt-table-target"></span>', unsafe_allow_html=True)
                                    
                                    modified = False
                                    
                                    for idx, row in details.iterrows():
                                        row_index_gs = row["_row_index"]
                                        item_name = row.get(item_col, "不明な商品") if item_col else "不明な商品"
                                        display_item_name = item_name[:10] + "…" if len(item_name) > 10 else item_name
                                        
                                        edit_vals = st.session_state['edit_data'].get(row_index_gs)
                                        if not edit_vals:
                                            continue
                                        
                                        disp_amount = edit_vals['amount']
                                        disp_major = edit_vals['major']
                                        disp_minor = edit_vals['minor']
                                        
                                        row_col1, row_col2, row_col3, row_col4 = st.columns(4)
                                        
                                        with row_col1:
                                            st.markdown(f"<div style='font-size: 0.85em; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;' title='{item_name}'>{display_item_name}</div>", unsafe_allow_html=True)
                                        with row_col2:
                                            st.markdown(f"<div style='font-size: 0.85em;'>¥{disp_amount:,}</div>", unsafe_allow_html=True)
                                            new_amount = disp_amount 
                                        with row_col3:
                                            majors = list(EXPENSE_CATEGORIES.keys())
                                            default_major_idx = majors.index(disp_major) if disp_major in majors else majors.index("その他")
                                            with st.popover(disp_major):
                                                new_major = st.radio("大分類", majors, index=default_major_idx, key=f"maj_{row_index_gs}", label_visibility="collapsed")
                                        with row_col4:
                                            minors = EXPENSE_CATEGORIES.get(new_major, EXPENSE_CATEGORIES["その他"])
                                            default_minor_idx = minors.index(disp_minor) if disp_minor in minors else len(minors)-1
                                            with st.popover(disp_minor):
                                                new_minor = st.radio("小分類", minors, index=default_minor_idx, key=f"min_{row_index_gs}", label_visibility="collapsed")
                                        
                                        if new_amount != disp_amount or new_major != disp_major or new_minor != disp_minor:
                                            st.session_state['edit_data'][row_index_gs]["amount"] = new_amount
                                            st.session_state['edit_data'][row_index_gs]["major"] = new_major
                                            st.session_state['edit_data'][row_index_gs]["minor"] = new_minor
                                            modified = True
                                        
                                if modified:
                                    st.rerun()
                                    
                                st.markdown("---")
                                
                                # 修正用ボタン（登録 / キャンセル）
                                col1, col2 = st.columns(2)
                                with col1:
                                    if st.button("登録", use_container_width=True, key="save_receipt"):
                                        try:
                                            with st.spinner("保存中..."):
                                                sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                                headers = sheet.row_values(1)
                                                amount_col_idx = headers.index("amount") + 1 if "amount" in headers else None
                                                category_col_idx = headers.index("category") + 1 if "category" in headers else None
                                                
                                                sub_col_idx = None
                                                for c in ["subcategory", "sub_category", "小分類"]:
                                                    if c in headers:
                                                        sub_col_idx = headers.index(c) + 1
                                                        break
                                                
                                                updates = []
                                                for r_idx_gs, vals in st.session_state['edit_data'].items():
                                                    if amount_col_idx: updates.append(gspread.Cell(row=r_idx_gs, col=amount_col_idx, value=vals["amount"]))
                                                    if category_col_idx: updates.append(gspread.Cell(row=r_idx_gs, col=category_col_idx, value=vals["major"]))
                                                    if sub_col_idx: updates.append(gspread.Cell(row=r_idx_gs, col=sub_col_idx, value=vals["minor"]))
                                                    
                                                if updates:
                                                    sheet.update_cells(updates)
                                                    
                                                st.success("✅ レシート明細を更新しました")
                                                st.session_state['edit_data'] = {} # リセット
                                                st.session_state[mode_key] = False # 閲覧モードに戻す
                                                import time; time.sleep(1)
                                                st.rerun()
                                        except Exception as e:
                                            st.error(f"エラー: {e}")
                                            
                                with col2:
                                    if st.button("キャンセル", use_container_width=True, key="cancel_receipt_edit"):
                                        # 状態をリセットし閲覧モードに戻る
                                        st.session_state['current_receipt_key'] = "" # キーを空にして初期化処理を無理やり再実行させる
                                        st.session_state[mode_key] = False
                                        st.rerun()
                                
                            else:
                                # 【閲覧モード】のレイアウト
                                
                                st.markdown("")
                                total_amount = 0
                                
                                # Markdownのテーブルヘッダー構築
                                table_md = "| No | 商品名 | 金額 | 大分類 | 小分類 |\n"
                                table_md += "|---|---|---:|---|---|\n"
                                
                                for i, (idx, row) in enumerate(details.iterrows(), 1):
                                    item_name = row.get(item_col, "不明な商品") if item_col else "不明な商品"
                                    # 商品名を全角10文字までに切り詰め
                                    display_item_name = item_name[:10] + "…" if len(item_name) > 10 else item_name
                                    
                                    major = row.get("category", "その他")
                                    sub_cols = [c for c in ["subcategory", "sub_category", "小分類"] if c in df.columns]
                                    sub = row.get(sub_cols[0], "❓その他") if sub_cols else "❓その他"
                                    amount = int(row.get("amount", 0))
                                    total_amount += amount
                                    
                                    # 各行のデータを追加 (金額の円表示は不要)
                                    table_md += f"| {i} | {display_item_name} | {amount:,} | {major} | {sub} |\n"
                                
                                # 合計行の追加
                                table_md += f"| | **合計** | **{total_amount:,}** | | |\n"
                                
                                # テーブルの描画
                                st.markdown(table_md)
                                st.markdown("---")
                                
                                # 閲覧用アクションボタン（修正 / 削除）
                                col1, col2 = st.columns(2)
                                with col1:
                                    if st.button("修正", use_container_width=True, key="edit_receipt"):
                                        st.session_state[mode_key] = True
                                        st.rerun()
                                        
                                with col2:
                                    if st.button("削除", use_container_width=True, key="delete_receipt"):
                                        try:
                                            with st.spinner("削除中..."):
                                                sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                                # 下から順に削除する（インデックスがずれないように）
                                                rows_to_delete = sorted(list(st.session_state['edit_data'].keys()), reverse=True)
                                                for r_idx in rows_to_delete:
                                                    sheet.delete_rows(r_idx)
                                                    
                                                st.success("✅ レシートを削除しました")
                                                st.session_state['edit_data'] = {}
                                                st.session_state[mode_key] = False
                                                import time; time.sleep(1)
                                                st.rerun()
                                        except Exception as e:
                                            st.error(f"エラー: {e}")
                            
                            # CSSの代わりにJSを使ってより確実にボタンの色を変更
                            import streamlit.components.v1 as components
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
                                    if (text === '削除' || text === 'キャンセル') {
                                        b.style.backgroundColor = '#dc3545';
                                        b.style.color = 'white';
                                        b.style.borderColor = '#dc3545';
                                    }
                                });
                            }, 500);
                            </script>
                            """, height=0, width=0)

                            
        elif menu_selection == "🤖 AI相談":
            st.markdown("#### 🤖 AI相談（専属ファイナンシャルプランナー）")
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
                    "category": "大分類",
                    "subcategory": "小分類",
                    "amount": "金額"
                }
                # 存在するカラムのみマッピング
                actual_rename = {k: v for k, v in rename_map.items() if k in df_user.columns}
                df_user = df_user.rename(columns=actual_rename)
                
                target_cols = ["対象年月", "日付", "店舗名", "大分類", "小分類", "金額"]
                available_cols = [c for c in target_cols if c in df_user.columns]
                
                return df_user[available_cols].to_csv(index=False)
                
            csv_data_string = get_user_data_csv_for_ai(st.session_state['username'])
            if not csv_data_string:
                csv_data_string = "現在、参照できる家計簿データはありません。"

            # --- チャットセッションとメッセージ履歴の初期化 ---
            if "ai_consult_messages" not in st.session_state:
                st.session_state.ai_consult_messages = []
            
            if "gemini_chat_session" not in st.session_state:
                client = st.session_state.get('genai_client')
                if client:
                    # システムプロンプトの構築
                    system_prompt = f"""あなたはユーザー専属の優秀なファイナンシャルプランナーです。
以下のCSVデータは、このユーザー（{st.session_state['username']}）個人の家計簿データです。このデータに基づいて、ユーザーの質問に正確かつ親身に答えてください。
データに存在しない推測は避け、無駄遣いの指摘や節約のアドバイスなども積極的に行ってください。

【ユーザーの家計簿データ】
{csv_data_string}"""
                    
                    # セッション開始（SDKの仕様に基づき、モデルからチャットを作成）
                    st.session_state.gemini_chat_session = client.chats.create(
                        model='gemini-2.5-flash',
                        config=types.GenerateContentConfig(system_instruction=system_prompt)
                    )
                else:
                    st.session_state.gemini_chat_session = None

            # 最初のメッセージを追加（履歴が空の場合）
            if not st.session_state.ai_consult_messages:
                st.session_state.ai_consult_messages.append({
                    "role": "assistant", 
                    "content": f"こんにちは、{st.session_state['username']}さん！あなたの専属FPです。全期間のデータを読み込みました。何でも相談してくださいね。"
                })

            # 履歴の表示
            for msg in st.session_state.ai_consult_messages:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
                    
            # ユーザー入力
            if user_input := st.chat_input("質問を入力してください..."):
                st.session_state.ai_consult_messages.append({"role": "user", "content": user_input})
                with st.chat_message("user"):
                    st.markdown(user_input)
                    
                chat = st.session_state.get('gemini_chat_session')
                if not chat:
                    with st.chat_message("assistant"):
                        st.error("APIキーが設定されていないため、相談を開始できません。")
                else:
                    with st.chat_message("assistant"):
                        message_placeholder = st.empty()
                        message_placeholder.markdown("分析中...")
                        
                        try:
                            # チャットセッションでのメッセージ送信
                            # 429等のエラーハンドリングを日本語化
                            def _send():
                                return chat.send_message(user_input)
                            
                            try:
                                response = safe_gemini_call(_send)
                                response_text = response.text
                                message_placeholder.markdown(response_text)
                                st.session_state.ai_consult_messages.append({"role": "assistant", "content": response_text})
                            except Exception as e:
                                err_msg = str(e)
                                if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                                    friendly_err = "現在AIの通信が混み合っています。数十秒待ってから再度送信してください。"
                                    st.warning(friendly_err)
                                    # 履歴には残さないか、エラーとして残す
                                else:
                                    st.error(f"エラーが発生しました: {e}")

                        except Exception as e:
                            st.error(f"予期せぬエラーが発生しました: {e}")

        elif menu_selection == "ヘルプ":
            st.markdown("#### 💡 ヘルプ・サポート")
            st.info("アプリの機能や使い方、データの保存先などについて何でも聞いてください！")
            
            # セッション状態の初期化
            if "help_messages" not in st.session_state:
                st.session_state.help_messages = [
                    {"role": "assistant", "content": "こんにちは！AI家計簿アプリのサポートAIです。\n機能の使い方や、データがどこに保存されているかなど、質問があればどうぞ！"}
                ]
                
            # メッセージ履歴の表示
            for msg in st.session_state.help_messages:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
                    
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
                            # システムプロンプトの構築（アプリの仕様とデータの管理場所）
                            system_prompt = """あなたは「AI家計簿アプリ」のサポート用チャットボットです。
ユーザーからの質問に対して、以下のアプリ仕様に基づいて丁寧かつ分かりやすく、日本語で回答してください。

【アプリの機能一覧と操作方法】
1. ダッシュボード
   - 月ごとの支出を円グラフ（大分類別）と一覧表（大分類からの展開で小分類、または明細）で確認できます。
   - 画面上部の「◀ 前月」「翌月 ▶」ボタンで表示する月を切り替えられます。
2. レシート取込
   - スマホやPCからレシートや領収書の画像をアップロードし、「レシートを解析する」ボタンを押すとAI（Gemini）が自動的に内容（日付、店舗名、商品名、金額、大分類、小分類）を読み取ります。
   - 解析後、内容を確認する画面が表示されます。確認して問題なければ青い「登録」ボタンで保存します。やり直す場合は赤い「キャンセル」ボタンを押してください。
3. レシート手入力
   - （現在準備中）手動で直接レシート情報を入力する機能です。今後のフェーズで実装予定です。
4. レシート修正
   - 過去に登録したレシート明細を月ごとに確認し、金額や分類（大分類・小分類）を修正できます。
   - 対象のレシート（店舗名と日付でグループ化されています）を選択し、内容を変更した後「登録」ボタンで上書き保存するか、「削除」ボタンでそのレシートを一件まるごと削除できます。
5. カレンダー
   - （現在準備中）日別の支出をカレンダー形式で確認できる機能です。今後のフェーズで実装予定です。
6. ヘルプ
   - 今あなたが使っているこのチャット機能です。

【データ管理とセキュリティについて】
- ユーザー認証：ユーザー名とパスワードでログインします。パスワードは暗号化（bcrypt）されて安全に保存されます。
- データの保存先（重要）：すべてのデータ（ユーザー情報、家計簿の明細データ）は、開発者が管理する「Google スプレッドシート」に保存されています。
  - `users` シート：ユーザー名と暗号化されたパスワードを保存。
  - `transactions` シート：各ユーザーの家計簿データ（レシート情報など）を保存。「username」列によってデータが区別されるため、他のユーザーのデータが混ざって表示されることはありません。
- 認証基盤：Google API（GCPのサービスアカウント、または credentials.json）を利用して、アプリからスプレッドシートへ安全にアクセスしています。

【回答のガイドライン】
- アプリの仕様に関すること以外を聞かれた場合は、「私はAI家計簿アプリのサポートボットですので、それについてはお答えできません」と優しく断ってください。
- 回答は長すぎず、箇条書きなどを活用して見やすくしてください。
"""
                            # チャット履歴をGeminiAPIの形式に変換
                            # system_instruction を使うか、プロンプトの先頭にシステムプロンプトを入れる
                            # ここでは安全に、各やり取りのコンテキストとして user プロンプトの先頭に入れる簡易方式をとるか、
                            # genai の chat セッション機能を使うことができます。
                            # 複雑さを避けるため、generate_content に system instruction として渡します
                            
                            
                            prompt_parts = [{"text": system_prompt}]
                            for m in st.session_state.help_messages:
                                prefix = "ユーザー: " if m["role"] == "user" else "AI: "
                                prompt_parts.append({"text": f"{prefix}{m['content']}"})
                                
                            prompt_parts.append({"text": "以上の会話を踏まえて、最後のユーザーの質問に返答してください。"})

                            response = safe_gemini_call(
                                client.models.generate_content,
                                model='gemini-1.5-flash',
                                contents=prompt_parts
                            )
                            
                            full_response = response.text
                            message_placeholder.markdown(full_response)
                            
                            st.session_state.help_messages.append({"role": "assistant", "content": full_response})
                            
                        except Exception as e:
                            error_msg = f"エラーが発生しました: {e}"
                            message_placeholder.error(error_msg)
                            st.session_state.help_messages.append({"role": "assistant", "content": error_msg})
            
    # 未ログインの状態 (ログイン・登録画面)
    else:
        st.title("AI家計簿アプリ")
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
