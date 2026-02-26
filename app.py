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

def get_sheet(worksheet_name):
    client = get_gspread_client()
    if client is None:
        st.stop()
        
    try:
        # スプレッドシートと指定ワークシートに接続
        sheet = client.open(SPREADSHEET_NAME).worksheet(worksheet_name)
        return sheet
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"エラー: スプレッドシート '{SPREADSHEET_NAME}' が見つかりません。クレデンシャルのメールアドレス ({client.auth.signer_email}) とスプレッドシートを共有してください。")
        st.stop()
    except gspread.exceptions.WorksheetNotFound:
        # 特定のシートがない場合
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
    records = sheet.get_all_records()
    
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

JSONの出力形式は以下を厳守してください。マークダウンの ```json などは含めず、純粋なJSON文字列（オブジェクトの配列）のみを返してください。
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
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[prompt, img]
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
    header_placeholder.header("ダッシュボード (月別集計)")


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
    # ログイン済みの状態
    if st.session_state.get('logged_in', False):
        
        # 自動画面遷移のためのリダイレクト処理
        if st.session_state.get('redirect_to_dashboard'):
            st.session_state['menu_selection'] = "ダッシュボード"
            st.session_state['redirect_to_dashboard'] = False
            
        # サイドバーメニューの実装
        with st.sidebar:
            st.title("メインメニュー [Ver 1.0.3]")
            st.write(f"🔑 ユーザー: **{st.session_state['username']}**")
            st.markdown("---")
            if 'menu_selection' not in st.session_state:
                st.session_state['menu_selection'] = "ダッシュボード"
                
            menu_selection = st.radio(
                "機能を選択",
                ["ダッシュボード", "レシート取込", "レシート手入力", "レシート修正", "カレンダー", "ヘルプ"],
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
        elif menu_selection == "レシート取込":
            st.header("📸 レシート取込")
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
                    
                    st.markdown("### 📋 解析結果の確認")
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
            st.header("レシート手入力")
            st.info("準備中: 手動でのレシート入力機能は今後のフェーズで実装されます。")
        elif menu_selection == "レシート修正":
            st.header("⚙️ レシート修正")
            
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
                            
                            # ヘッダー行と明細行を横スクロールさせるためのCSS
                            st.markdown("""
                            <style>
                                /* 明細ブロック全体を包むコンテナ（横スクロール用） */
                                .scrollable-container {
                                    overflow-x: auto;
                                    white-space: nowrap;
                                    padding-bottom: 10px;
                                    -webkit-overflow-scrolling: touch;
                                    width: 100%;
                                }
                                
                                /* 明細行の折り返しを無効化し、モバイルでも強制的に1行に収める */
                                .scrollable-container [data-testid="stHorizontalBlock"] {
                                    display: flex !important;       /* モバイルでblockになるのを防ぐ */
                                    flex-direction: row !important; /* モバイル時の強制縦並び(column)を解除 */
                                    flex-wrap: nowrap !important;   /* 絶対に折り返さない・最重要 */
                                    white-space: nowrap !important; /* 改行禁止 */
                                    width: max-content !important;  /* 無駄に広がらない */
                                    justify-content: flex-start !important; /* 左寄せ */
                                    gap: 0.2rem !important;         /* カラム間の隙間を極力狭く */
                                }
                                .scrollable-container [data-testid="column"] {
                                    flex: 0 0 auto !important;      /* 自動伸長をオフ */
                                    width: auto !important;         /* モバイル時の強制100%幅を解除 */
                                    max-width: none !important;     /* 最大幅制限も解除 */
                                    padding-left: 0.1rem !important;
                                    padding-right: 0.1rem !important;
                                }
                                /* 各項目の幅を文字数に合わせて設定する（min-widthで確実に潰れないようにする） */
                                .scrollable-container [data-testid="stHorizontalBlock"] > [data-testid="column"]:nth-child(1) { min-width: 120px !important; width: 140px !important; } /* 商品名 */
                                .scrollable-container [data-testid="stHorizontalBlock"] > [data-testid="column"]:nth-child(2) { min-width: 60px !important; width: 60px !important; }  /* 金額 */
                                .scrollable-container [data-testid="stHorizontalBlock"] > [data-testid="column"]:nth-child(3) { min-width: 80px !important; width: max-content !important; }  /* 大分類 */
                                .scrollable-container [data-testid="stHorizontalBlock"] > [data-testid="column"]:nth-child(4) { min-width: 80px !important; width: max-content !important; }  /* 小分類 */
                                
                                /* ポップオーバー（大分類・小分類ボタン）の表示を極力コンパクトに */
                                div[data-testid="stPopover"] > button {
                                    padding: 2px 8px !important;
                                    font-size: 0.8em !important;
                                    min-height: 28px !important;
                                }
                            </style>
                            """, unsafe_allow_html=True)
                            
                            # 全体をスクロール可能にするために、HTMLのdivでラップ（Streamlitのマークダウン内包機能は限定的だが、CSSのmin-widthで担保するアプローチに変更）
                            
                            st.markdown('<div class="scrollable-container">', unsafe_allow_html=True)
                            
                            # ヘッダー
                            h_col1, h_col2, h_col3, h_col4 = st.columns([4, 1.5, 2.25, 2.25])
                            h_col1.markdown("<div style='font-size: 0.85em; font-weight: bold; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;'>商品名</div>", unsafe_allow_html=True)
                            h_col2.markdown("<div style='font-size: 0.85em; font-weight: bold;'>金額</div>", unsafe_allow_html=True)
                            h_col3.markdown("<div style='font-size: 0.85em; font-weight: bold;'>大分類</div>", unsafe_allow_html=True)
                            h_col4.markdown("<div style='font-size: 0.85em; font-weight: bold;'>小分類</div>", unsafe_allow_html=True)
                            
                            modified = False
                            
                            for idx, row in details.iterrows():
                                row_index_gs = row["_row_index"]
                                item_name = row.get(item_col, "不明な商品") if item_col else "不明な商品"
                                # 商品名を全角10文字までに切り詰め
                                display_item_name = item_name[:10] + "…" if len(item_name) > 10 else item_name
                                
                                edit_vals = st.session_state['edit_data'].get(row_index_gs)
                                if not edit_vals:
                                    continue
                                
                                disp_amount = edit_vals['amount']
                                disp_major = edit_vals['major']
                                disp_minor = edit_vals['minor']
                                
                                row_col1, row_col2, row_col3, row_col4 = st.columns([4, 1.5, 2.25, 2.25])
                                
                                # 商品名
                                with row_col1:
                                    st.markdown(f"<div style='margin-top: 8px; font-size: 0.85em; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;' title='{item_name}'>{display_item_name}</div>", unsafe_allow_html=True)
                                    
                                # 金額 (表示のみにする)
                                with row_col2:
                                    st.markdown(f"<div style='margin-top: 8px; font-size: 0.85em;'>¥{disp_amount:,}</div>", unsafe_allow_html=True)
                                    new_amount = disp_amount # 変更不可なため保持
                                    
                                # 大分類 (ポップオーバーリストに変更)
                                with row_col3:
                                    majors = list(EXPENSE_CATEGORIES.keys())
                                    default_major_idx = majors.index(disp_major) if disp_major in majors else majors.index("その他")
                                    with st.popover(disp_major):
                                        new_major = st.radio("大分類", majors, index=default_major_idx, key=f"maj_{row_index_gs}", label_visibility="collapsed")
                                    
                                # 小分類 (ポップオーバーリストに変更)
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
                                    
                            st.markdown('</div>', unsafe_allow_html=True)
                                    
                            if modified:
                                st.rerun()
                                
                            st.markdown("---")
                            
                            # ボタン（戻すボタンを削除し、2列に変更）
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.button("登録", use_container_width=True):
                                    try:
                                        with st.spinner("保存中..."):
                                            sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                            headers = sheet.row_values(1)
                                            amount_col_idx = headers.index("amount") + 1 if "amount" in headers else None
                                            category_col_idx = headers.index("category") + 1 if "category" in headers else None
                                            
                                            # subcategoryの列名を特定
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
                                            import time; time.sleep(1)
                                            st.rerun()
                                    except Exception as e:
                                        st.error(f"エラー: {e}")
                                        
                            with col2:
                                # 削除ボタン
                                if st.button("削除", use_container_width=True):
                                    try:
                                        with st.spinner("削除中..."):
                                            sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                            # 下から順に削除する（インデックスがずれないように）
                                            rows_to_delete = sorted(list(st.session_state['edit_data'].keys()), reverse=True)
                                            for r_idx in rows_to_delete:
                                                sheet.delete_rows(r_idx)
                                                
                                            st.success("✅ レシートを削除しました")
                                            st.session_state['edit_data'] = {}
                                            import time; time.sleep(1)
                                            st.rerun()
                                    except Exception as e:
                                        st.error(f"エラー: {e}")
                            
                            # CSSの代わりにJSを使ってより確実にボタンの色を変更
                            import streamlit.components.v1 as components
                            components.html("""
                            <script>
                            const elements = window.parent.document.querySelectorAll('button');
                            elements.forEach(b => {
                                const text = b.innerText.trim();
                                if (text === '登録') {
                                    b.style.backgroundColor = '#007bff';
                                    b.style.color = 'white';
                                    b.style.borderColor = '#007bff';
                                }
                                if (text === '削除') {
                                    b.style.backgroundColor = '#ff4b4b';
                                    b.style.color = 'white';
                                    b.style.borderColor = '#ff4b4b';
                                }
                            });
                            </script>
                            """, height=0, width=0)

        elif menu_selection == "カレンダー":
            st.header("📅 カレンダー")
            
            # カレンダー用のセッション状態を初期化
            if 'selected_day' not in st.session_state:
                st.session_state['selected_day'] = None
                
            # 表示月が変更された場合は選択日をリセット
            if 'last_cal_month' not in st.session_state or st.session_state['last_cal_month'] != st.session_state['current_month']:
                st.session_state['selected_day'] = None
                st.session_state['last_cal_month'] = st.session_state['current_month']
                
            # ダッシュボードと共通の月選択UI
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
            
            # データ取得と日次集計
            with st.spinner("データを読み込み中..."):
                df = load_transactions_data(st.session_state['current_month'])
                
            daily_totals = {}
            if not df.empty and "date" in df.columns and "amount" in df.columns:
                df['day'] = df['date'].dt.day
                daily_totals = df.groupby('day')["amount"].sum().to_dict()
                
            # カレンダーの描画（日曜始まりに設定）
            calendar.setfirstweekday(calendar.SUNDAY)
            year = st.session_state['current_month'].year
            month = st.session_state['current_month'].month
            cal = calendar.monthcalendar(year, month)
            
            # 曜日ヘッダー (日曜始まり)
            weekdays = ["日", "月", "火", "水", "木", "金", "土"]
            cols = st.columns(7)
            for i, wd in enumerate(weekdays):
                color = "#ff4b4b" if wd == "日" else "#1f77b4" if wd == "土" else "black"
                cols[i].markdown(f"<div style='text-align: center; font-weight: bold; font-size: 0.9em; color: {color};'>{wd}</div>", unsafe_allow_html=True)
                
            # カレンダー全体にのみ影響を与えるためのCSS
            st.markdown("""
            <style>
            /* 枠線の内側要素をターゲットにして、ボタンを浮かせ、透明にする */
            div[data-testid="column"] div[data-testid="stVerticalBlockBorderWrapper"] {
                position: relative !important;
                min-height: 70px !important;
                padding: 0 !important;
            }
            div[data-testid="column"] div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stButton"] {
                position: absolute !important;
                top: 0 !important;
                left: 0 !important;
                width: 100% !important;
                height: 100% !important;
                z-index: 10 !important;
            }
            div[data-testid="column"] div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stButton"] button {
                width: 100% !important;
                height: 100% !important;
                opacity: 0.001 !important;
                border: none !important;
                background: transparent !important;
                margin: 0 !important;
                padding: 0 !important;
            }
            div[data-testid="column"] div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stButton"] button:hover,
            div[data-testid="column"] div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stButton"] button:focus {
                background: transparent !important;
                border: none !important;
                box-shadow: none !important;
                outline: none !important;
            }
            </style>
            """, unsafe_allow_html=True)
            
            # カレンダーグリッド描画
            for week in cal:
                cols = st.columns(7)
                for i, day in enumerate(week):
                    with cols[i]:
                        if day == 0:
                            # 空セル
                            st.markdown("<div style='min-height: 70px;'></div>", unsafe_allow_html=True)
                        else:
                            total = daily_totals.get(day, 0)
                            # 選択状態の背景レイヤー
                            bg_div = f'<div style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; background-color: #e6f7ff; z-index: 0; border-radius: 0.4rem;"></div>' if st.session_state['selected_day'] == day else ''
                            
                            with st.container(border=True):
                                # CSSラップとセルのレイアウト描画
                                st.markdown(f'''
                                {bg_div}
                                <div style="position: absolute; top: 4px; left: 6px; color: black; font-weight: bold; font-size: 1.1rem; z-index: 1;">{day}</div>
                                <div style="position: absolute; bottom: 4px; right: 6px; color: #ff4b4b; font-weight: bold; font-size: 0.85em; z-index: 1;">
                                    {'￥{:,}'.format(int(total)) if total > 0 else ''}
                                </div>
                                <div style="min-height: 55px; opacity: 0;">0</div>
                                ''', unsafe_allow_html=True)
                                
                                # 透明なボタンを配置してクリックを検知
                                if st.button(" ", key=f"cal_btn_{day}", use_container_width=True):
                                    st.session_state['selected_day'] = day
                                    st.rerun()
                                    
            # 指定日の明細一覧表示（ブラインド表示・アコーディオン形式）
            if st.session_state['selected_day']:
                sel_day = st.session_state['selected_day']
                st.markdown("---")
                
                if df.empty or sel_day not in daily_totals:
                    st.info(f"{month}月{sel_day}日の明細はありません。")
                else:
                    day_df = df[df["day"] == sel_day].copy()
                    
                    st.markdown(f"#### {month}月{sel_day}日の明細一覧")
                    
                    # 店舗名を確実に取得
                    # カレンダー以外の画面でも同じロジックが使用されているため統一して正確に元のデータを取得する
                    # df作成時の正規化済みのカラムを使うのが安全
                    store_col = next((c for c in ["store_name", "store", "店舗名", "店舗"] if c in day_df.columns), None)
                    
                    if store_col:
                        day_df["_display_store"] = day_df[store_col].apply(
                            lambda x: str(x).strip() if pd.notna(x) and str(x).strip() != "" else "店舗名不明"
                        )
                    else:
                        day_df["_display_store"] = "店舗名不明"
                        
                    # 取引明細としての順番を保持しつつ集計する
                    store_groups = day_df.groupby("_display_store", sort=False)
                    
                    for store_name, group in store_groups:
                        store_total = int(group["amount"].sum())
                        disp_store = str(store_name)
                        
                        # st.expander() がブラインド表示（アコーディオン形式）になります
                        with st.expander(f"🛒 **{disp_store}**　　（合計: ￥{store_total:,}）", expanded=False):
                            
                            # そのレシートの大分類を金額が多い順に一覧表示する
                            cat_groups = group.groupby("category", dropna=False)["amount"].sum().sort_values(ascending=False)
                            for cat, cat_amount in cat_groups.items():
                                disp_cat = cat if pd.notna(cat) else "その他"
                                st.markdown(f"""
                                <div style='display: flex; justify-content: space-between; border-bottom: 1px dotted #ccc; padding: 6px 0; font-size: 0.9em;'>
                                    <div>{disp_cat}</div>
                                    <div style='text-align: right;'>￥{int(cat_amount):,}</div>
                                </div>
                                """, unsafe_allow_html=True)
                                
                            # 最終行に合計金額
                            st.markdown(f"""
                            <div style='display: flex; justify-content: space-between; padding-top: 10px; font-weight: bold;'>
                                <div>合計</div>
                                <div style='color: #ff4b4b; text-align: right;'>￥{store_total:,}</div>
                            </div>
                            """, unsafe_allow_html=True)
        elif menu_selection == "ヘルプ":
            st.header("🤖 ヘルプ・サポート")
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

                            response = client.models.generate_content(
                                model='gemini-2.5-flash',
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
