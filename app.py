import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import bcrypt
import os
import json
from datetime import datetime
from dateutil.relativedelta import relativedelta

# ---------- 構成設定 ----------
SPREADSHEET_NAME = "Kakeibo_Data" # 実際のGoogleスプレッドシート名に合わせて変更してください
WORKSHEET_NAME = "users"
TRANSACTIONS_WORKSHEET_NAME = "transactions"

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
            sheet.insert_row(["username", "date", "category", "amount", "memo"], 1)
    except Exception:
        sheet.insert_row(["username", "date", "category", "amount", "memo"], 1)

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
        
    return df

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
        
        col_chart, col_table = st.columns(2)
        
        with col_chart:
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
            
        with col_table:
            # データフレーム表
            st.markdown("##### カテゴリ別合計金額表")
            
            # 見た目を少し整える(フォーマット等)
            display_df = grouped_df.copy()
            display_df.columns = ["カテゴリ (大分類)", "合計金額 (円)"]
            
            # 枠線付きのデータフレームを表示
            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "合計金額 (円)": st.column_config.NumberColumn(
                        "合計金額 (円)",
                        format="%d ¥"
                    )
                }
            )
            
            total_amount = grouped_df["amount"].sum()
            st.metric("総支出額", f"{int(total_amount):,} 円")
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
                ["ダッシュボード (月別集計)", "レシート入力", "カレンダー", "設定・ヘルプ"]
            )
            st.markdown("---")
            if st.button("ログアウト", use_container_width=True):
                st.session_state['logged_in'] = False
                st.session_state['username'] = None
                st.rerun()

        # メインコンテンツの切り替え
        if menu_selection == "ダッシュボード (月別集計)":
            show_dashboard()
        elif menu_selection == "レシート入力":
            st.header("レシート入力")
            st.info("準備中: 画像アップロード機能は今後のフェーズで実装されます。")
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
