import streamlit as st

st.title("Test")
tab1, tab2 = st.tabs(["ログイン", "新規ユーザー登録"])
with tab1:
    with st.form("login_form_v2"):
        login_username = st.text_input("ユーザー名", key="login_username_input_v2")
        login_password = st.text_input("パスワード", type="password", key="login_password_input_v2")
        remember_me = st.checkbox("ログイン状態を保持する", value=True, key="remember_me_input_v2")
        submitted = st.form_submit_button("ログイン", use_container_width=True)
