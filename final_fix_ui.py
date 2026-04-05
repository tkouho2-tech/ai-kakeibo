import os
import re

target_file = 'fixed_cost_expansion.py'

# Read with appropriate encoding
try:
    with open(target_file, 'rb') as f:
        content = f.read().decode('utf-8', errors='replace')
except:
    with open(target_file, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()

# New UI code for show_fixed_cost_data_expansion
new_ui_code = r'''def show_fixed_cost_data_expansion():
    st.title("⚙️ 固定費データ展開")
    username = st.session_state.get("username", "")

    st.markdown("### ボタンA：【新規 or 全期間再展開】")
    st.info("マスターの「開始月」から全てのデータを展開し直すモードです。")
    if st.button("新規 or 全期間再展開を実行", key="btn_a_init", use_container_width=True):
        st.session_state.confirm_a = True

    if st.session_state.get("confirm_a"):
        st.warning("⚠️ データ展開済みの場合は、既存のデータが初期化（上書き）されます。過去の入力内容も消去されますが、本当によろしいですか？")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("はい、実行します", key="do_a_exec", type="primary", use_container_width=True):
                with st.spinner("展開中..."):
                    success, msg = execute_expansion(username) # start_ym=None is default
                    if success: st.success(msg)
                    else: st.error(msg)
                st.session_state.confirm_a = False
        with col2:
            if st.button("キャンセル (A)", key="cancel_a", use_container_width=True):
                st.session_state.confirm_a = False
                st.rerun()

    st.markdown("---")
    st.markdown("### ボタンB：【期間指定展開（当月以降）】")
    st.info("過去のデータを保護し、指定した月から将来に向かって展開するモードです。")
    
    today = datetime.now(JST)
    m1 = today.strftime("%Y.%m月")
    m2 = (today + relativedelta(months=1)).strftime("%Y.%m月")
    m3 = (today + relativedelta(months=2)).strftime("%Y.%m月")
    
    sel_month = st.selectbox("展開を開始する月を選択してください", [m1, m2, m3], key="sel_month_b")
    
    if st.button("期間指定展開を実行", key="btn_b_init", use_container_width=True, type="primary"):
        st.session_state.confirm_b = True
        st.session_state.target_month_b = sel_month

    if st.session_state.get("confirm_b"):
        st.warning(f"🛈 「{st.session_state.target_month_b}」以降のデータを更新します。よろしいですか？")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("はい、更新します", key="do_b_exec", type="primary", use_container_width=True):
                with st.spinner("展開中..."):
                    parts = st.session_state.target_month_b.replace("月", "").split(".")
                    y_val, m_val = int(parts[0]), int(parts[1])
                    success, msg = execute_expansion(username, mode="PARTIAL", start_ym=(y_val, m_val))
                    if success: st.success(msg)
                    else: st.error(msg)
                st.session_state.confirm_b = False
        with col2:
            if st.button("キャンセル (B)", key="cancel_b", use_container_width=True):
                st.session_state.confirm_b = False
                st.rerun()
'''

# Find the function in content and replace it
# Use a broad regex to find show_fixed_cost_data_expansion until its end or next function
pattern = re.compile(r'def show_fixed_cost_data_expansion\(\):.*?if success: st\.success\(msg\).*?else: st\.error\(msg\)', re.DOTALL)
content = pattern.sub(new_ui_code, content)

# Also fix the corruption in show_variable_cost_update and other titles if possible
content = content.replace('st.title("諁E螟牙虚雋E繝繝ｼ繧E譖ｴ譁E")', 'st.title("📈 変動費データ更新")')
content = content.replace('st.title("屏EE 蝗ｺ螳夊ｲE繝繝ｼ繧E螻暮幁E")', 'st.title("⚙️ 固定費データ展開")')

with open(target_file, 'w', encoding='utf-8') as f:
    f.write(content)

print("UI Replacement complete.")
