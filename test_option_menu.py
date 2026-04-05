import streamlit as st
from streamlit_option_menu import option_menu

if "active" not in st.session_state:
    st.session_state.active = "ダッシュボード"

def on_change(key):
    st.session_state.active = st.session_state[key]

opts1 = ["ダッシュボード", "カレンダー", "クレジットカード"]
idx1 = opts1.index(st.session_state.active) if st.session_state.active in opts1 else None

opts2 = ["レシート取込", "手入力", "レシート修正"]
idx2 = opts2.index(st.session_state.active) if st.session_state.active in opts2 else None

with st.sidebar:
    st.write("Current:", st.session_state.active)
    
    m1 = option_menu(
        "【1. 表示・分析系】", opts1, icons=["graph-up", "calendar3", "credit-card"], 
        manual_select=idx1, key="m1", on_change=on_change
    )
    
    m2 = option_menu(
        "【2. レシート操作】", opts2, icons=["camera", "pencil-square", "tools"], 
        manual_select=idx2, key="m2", on_change=on_change
    )
