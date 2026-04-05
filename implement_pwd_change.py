import sys

with open(r'c:\Users\t_kou\Kakeibo_Final_v3\app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Insert change_user_password function after authenticate_user
old_auth_func = '''def authenticate_user(username, password):
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
    return False'''

new_auth_func = '''def authenticate_user(username, password):
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
        return False, f"パスワード変更エラー: {e}"'''

if old_auth_func in content:
    content = content.replace(old_auth_func, new_auth_func)
    print("Added change_user_password()")
else:
    print("Could not find authenticate_user block.")

# 2. Insert UI in "プロフィール設定"
old_ui_block = '''            st.markdown("---")
            # 支払い方法マスターは独立したメニューに移動しました。

        elif menu_selection == "ヘルプ":'''

new_ui_block = '''            st.markdown("---")
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

        elif menu_selection == "ヘルプ":'''

if old_ui_block in content:
    content = content.replace(old_ui_block, new_ui_block)
    print("Added Password Change UI")
else:
    print("Could not find UI block.")

with open(r'c:\Users\t_kou\Kakeibo_Final_v3\app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Saved app.py.")
