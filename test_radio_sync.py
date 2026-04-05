import streamlit as st

if "menu_selection" not in st.session_state:
    st.session_state.menu_selection = "カレンダー"

def handle_radio_change(key):
    # If the user clicks a radio button, its value will be updated in session_state
    new_val = st.session_state[key]
    if new_val is not None:
        st.session_state.menu_selection = new_val

st.markdown("### Test Menu Sync")

opts1 = ["ダッシュボード（月次集計）", "ダッシュボード（年次集計）", "カレンダー", "クレジットカード"]
opts2 = ["レシート取込", "レシート手入力", "レシート修正"]
opts3 = ["マニュアル", "ヘルプ", "AI相談"]
opts4 = ["支払方法マスター", "プロフィール設定"]

# For each radio, if the global selection is in its options, select it, else None
idx1 = opts1.index(st.session_state.menu_selection) if st.session_state.menu_selection in opts1 else None
idx2 = opts2.index(st.session_state.menu_selection) if st.session_state.menu_selection in opts2 else None
idx3 = opts3.index(st.session_state.menu_selection) if st.session_state.menu_selection in opts3 else None
idx4 = opts4.index(st.session_state.menu_selection) if st.session_state.menu_selection in opts4 else None

st.markdown("##### 【表示・分析系】")
st.radio(" ", opts1, index=idx1, key="r1", on_change=handle_radio_change, args=("r1",), label_visibility="collapsed")

st.markdown("##### 【レシート管理】")
st.radio(" ", opts2, index=idx2, key="r2", on_change=handle_radio_change, args=("r2",), label_visibility="collapsed")

st.markdown("##### 【相談・サポート】")
st.radio(" ", opts3, index=idx3, key="r3", on_change=handle_radio_change, args=("r3",), label_visibility="collapsed")

st.markdown("##### 【マスター設定】")
st.radio(" ", opts4, index=idx4, key="r4", on_change=handle_radio_change, args=("r4",), label_visibility="collapsed")

st.write("Current selection:", st.session_state.menu_selection)
