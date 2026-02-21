import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import bcrypt
import os
import json
import io
from datetime import datetime
from dateutil.relativedelta import relativedelta
from PIL import Image
import google.generativeai as genai

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
    "その他": ["📁未分類"]
}

def get_categories_prompt_text():
    """AI（Gemini等）のプロンプトに埋め込むためのカテゴリ定義文字列を生成"""
    text = "【カテゴリシステム: 大分類と小分類のリスト】\n"
    for major, minors in EXPENSE_CATEGORIES.items():
        text += f"- {major}: {', '.join(minors)}\n"
    text += "\n※ 必ず上記の大分類と小分類の組み合わせに従ってください。"
    return text

# APIキー設定（Gemini用）
if "general" in st.secrets and "gemini_api_key" in st.secrets["general"]:
    genai.configure(api_key=st.secrets["general"]["gemini_api_key"])

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
        # 1. まずStreamlit Cloudの st.secrets から取得を試みる
        if "gcp_service_account" in st.secrets:
            # st.secrets の内容（文字列）を辞書型に変換
            credentials_dict = json.loads(st.secrets["gcp_service_account"])

            from google.oauth2.service_account import Credentials
            
            # gspreadが必要とするスコープを指定
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            creds = Credentials.from_service_account_info(credentials_dict, scopes=scopes)
            return gspread.authorize(creds)
            
        # 2. ローカル環境用のフォールバック (credentials.json)
        elif os.path.exists("credentials.json"):
            return gspread.service_account(filename="credentials.json")
            
        else:
            st.error("Google APIの無効な環境設定: st.secrets も credentials.json も見つかりません。")
            return None
            
    except Exception as e:
        st.error(f"Google APIの認証エラーが発生しました: {e}")
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
                text_only = "".join([c for c in v_sub if c.isalpha() or c in "類物食品未分類その他"]) 
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
        
        model = genai.GenerativeModel('gemini-2.5-flash')
        
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

それ以外の場合は、以下の14のカテゴリ体系に厳密に従って、明細ごとに適切に分類してください。
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
        response = model.generate_content([
            prompt,
            {"mime_type": "image/jpeg", "data": img_byte_arr}
        ])
        
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
    if st.session_state['logged_in']:
        
        # サイドバーメニューの実装
        with st.sidebar:
            st.title("メインメニュー")
            st.write(f"🔑 ユーザー: **{st.session_state['username']}**")
            st.markdown("---")
            menu_selection = st.radio(
                "機能を選択",
                ["ダッシュボード (月別集計)", "レシート取込", "レシート手入力", "レシート修正", "カレンダー", "設定・ヘルプ"]
            )
            st.markdown("---")
            if st.button("ログアウト", use_container_width=True):
                st.session_state['logged_in'] = False
                st.session_state['username'] = None
                st.rerun()

        # メインコンテンツの切り替え
        if menu_selection == "ダッシュボード (月別集計)":
            show_dashboard()
        elif menu_selection == "レシート取込":
            st.header("レシート取込")
            st.info("カメラ撮影または画像ファイルからレシートを読み取り、自動で入力します。")
            
            uploaded_file = st.file_uploader("レシートの画像をアップロード（またはカメラで撮影）", type=['png', 'jpg', 'jpeg'], accept_multiple_files=False)
            
            if uploaded_file is not None:
                st.image(uploaded_file, caption="アップロードされたレシート", width=300)
                
                if st.button("画像を解析して保存する", type="primary"):
                    with st.spinner("画像を解析中... Geminiが読み取っています"):
                        results = parse_receipt_with_gemini(uploaded_file)
                        
                        if isinstance(results, list) and len(results) > 0 and "error" in results[0]:
                            st.error(f"解析に失敗しました: {results[0]['error']}")
                        elif isinstance(results, dict) and "error" in results:
                            st.error(f"解析に失敗しました: {results['error']}")
                        else:
                            try:
                                sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
                                init_transactions_sheet(sheet)
                                
                                written_data = []
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
                                        text_only = "".join([c for c in m if c.isalpha() or c in "類物食品未分類その他"])
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
                                    
                                    written_data.append({
                                        "日付": str(item.get("date", "")),
                                        "店舗名": store_name,
                                        "商品名": item_name,
                                        "金額": int(item.get("amount", 0)),
                                        "大分類": final_major,
                                        "小分類": final_minor
                                    })
                                
                                st.success(f"解析が完了し、{len(written_data)}件のデータをスプレッドシートに保存しました！")
                                st.markdown("### 登録されたデータ")
                                st.dataframe(written_data, use_container_width=True)
                                
                            except Exception as e:
                                st.error(f"保存エラー: {e}")

        elif menu_selection == "レシート手入力":
            st.header("レシート手入力")
            st.info("準備中: 手動でのレシート入力機能は今後のフェーズで実装されます。")
        elif menu_selection == "レシート修正":
            st.header("レシート修正")
            st.info("準備中: 登録したデータの修正機能は今後のフェーズで実装されます。")
        elif menu_selection == "カレンダー":
            st.header("カレンダー")
            st.info("準備中: カレンダーUIは今後のフェーズで実装されます。")
        elif menu_selection == "設定・ヘルプ":
            st.header("設定・ヘルプ")
            st.info("準備中: 設定およびチャットボット機能は今後のフェーズで実装されます。")
            
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
