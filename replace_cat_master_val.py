import sys

with open(r'c:\Users\t_kou\Kakeibo_Final_v3\app.py', 'r', encoding='utf-8') as f:
    content = f.read()

old_func_start = 'def show_category_master():'
old_func_end = 'def show_payment_master():'

if old_func_start in content and old_func_end in content:
    start_idx = content.find('def save_categories_to_sheet')
    end_idx = content.find(old_func_end)
    
    new_func_block = '''@st.cache_data(ttl=60)
def get_used_categories():
    try:
        sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
        records = safe_gspread_call(sheet.get_all_records)
        used_majors = set()
        used_minors = {}
        if not records:
            return used_majors, used_minors
        for row in records:
            major = str(row.get("category", "")).strip()
            sub = ""
            for c in ["subcategory", "sub_category", "小分類"]:
                if c in row:
                    sub = str(row[c]).strip()
                    break
                    
            if major:
                used_majors.add(major)
                if major not in used_minors:
                    used_minors[major] = set()
                if sub:
                    used_minors[major].add(sub)
        return used_majors, used_minors
    except Exception as e:
        print(f"Error checking used categories: {e}")
        return set(), {}

def save_categories_to_sheet(cats_dict):
    import streamlit as st
    try:
        sheet = get_sheet(CATEGORY_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        safe_gspread_call(sheet.clear)
        
        new_data = [["major_category", "minor_category"]]
        for major, minors in cats_dict.items():
            if minors:
                for minor in minors:
                    new_data.append([major, minor])
            else:
                new_data.append([major, ""])
                
        safe_gspread_call(sheet.update, range_name="A1", values=new_data)
        get_categories.clear() # Cache clear
        return True
    except Exception as e:
        st.error(f"カテゴリ保存エラー: {e}")
        return False

def show_category_master():
    """カテゴリマスター設定画面 (オーナー専用)"""
    import streamlit as st
    st.markdown("#### 📂 カテゴリマスター （大分類・小分類設定）")
    st.info("家計簿全体で使用される「大分類」と「小分類」の設定を行います。この画面はオーナー専用です。\\n\\n※ 既にレシートデータで登録済みのカテゴリは**変更・削除できません**。どうしても変更が必要な場合は先にレシート修正画面から修正してください。\\n※ 新しく追加したカテゴリグラフの色は自動で割り当てられます。")
    
    current_cats = get_categories()
    used_majors, used_minors = get_used_categories()
    
    col1, padding, col2 = st.columns([4, 1, 6])
    
    with col1:
        st.markdown("##### 📁 【大分類の一覧】")
        st.caption("※ 編集・確認する大分類を以下から選択してください。")
        major_cats = list(current_cats.keys())
        
        if 'selected_major_cat' not in st.session_state:
            st.session_state.selected_major_cat = major_cats[0] if major_cats else None
            
        selected_major = st.radio("大分類選択", options=major_cats, key="cb_major_cat", label_visibility="collapsed")
        st.session_state.selected_major_cat = selected_major
        
        st.markdown("---")
        with st.expander("➕ 新しい大分類を追加する", expanded=False):
            new_major_name = st.text_input("大分類名（例: 日用品費）", key="new_major_input")
            if st.button("大分類を追加", use_container_width=True):
                new_major = new_major_name.strip()
                if new_major:
                    if new_major in current_cats:
                        st.warning("その大分類は既に存在します。")
                    else:
                        current_cats[new_major] = []
                        if save_categories_to_sheet(current_cats):
                            st.success(f"「{new_major}」を追加しました！")
                            st.session_state.selected_major_cat = new_major
                            import time
                            time.sleep(1)
                            st.rerun()
                else:
                    st.warning("大分類名を入力してください。")
                    
        with st.expander("⚠️ 大分類名を変更・削除する", expanded=False):
            if selected_major:
                st.markdown(f"**対象:** 「{selected_major}」")
                changed_major_name = st.text_input("新しい名前", value=selected_major, key="change_major_input")
                
                if st.button("名前を変更する", type="primary", use_container_width=True):
                    new_m_name = changed_major_name.strip()
                    if new_m_name and new_m_name != selected_major:
                        if new_m_name in current_cats:
                            st.warning("その名前は既に使用されています。")
                        elif selected_major in used_majors:
                            st.error(f"「{selected_major}」は既にレシートデータで登録済みの為、名前の変更はできません。")
                        else:
                            new_cats = {}
                            for k, v in current_cats.items():
                                if k == selected_major:
                                    new_cats[new_m_name] = v
                                else:
                                    new_cats[k] = v
                            if save_categories_to_sheet(new_cats):
                                st.success("名前を変更しました！")
                                st.session_state.selected_major_cat = new_m_name
                                import time
                                time.sleep(1)
                                st.rerun()
                
                st.markdown("<br>", unsafe_allow_html=True)
                st.warning(f"「{selected_major}」を削除しますか？紐づく小分類もすべて消去されます。")
                if st.button(f"「{selected_major}」を完全に削除する", use_container_width=True):
                    if selected_major in used_majors:
                        st.error(f"「{selected_major}」は既にレシートデータで登録済みの為、削除はできません。")
                    else:
                        del current_cats[selected_major]
                        if save_categories_to_sheet(current_cats):
                            st.success(f"「{selected_major}」を削除しました！")
                            current_majors = list(current_cats.keys())
                            st.session_state.selected_major_cat = current_majors[0] if current_majors else None
                            import time
                            time.sleep(1)
                            st.rerun()
                        
    with col2:
        if selected_major:
            st.markdown(f"##### 📄 【小分類の編集】")
            st.write(f"**現在の対象大分類:** 📁 {selected_major}")
            st.caption("※ 下の表を直接クリックして小分類名を追加・修正・削除できます。最下部の `+` マークで行を追加できます。")
            
            minors = current_cats.get(selected_major, [])
            import pandas as pd
            if minors:
                df_minors = pd.DataFrame([{"小分類": m} for m in minors])
            else:
                df_minors = pd.DataFrame(columns=["小分類"])
            
            edited_minors_df = st.data_editor(
                df_minors,
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_{selected_major}",
                column_config={
                    "小分類": st.column_config.TextColumn("小分類", required=True)
                }
            )
            
            if st.button(f"「{selected_major}」の小分類の変更を保存", type="primary", use_container_width=True):
                old_minors = set(minors)
                new_minors = []
                for _, row in edited_minors_df.iterrows():
                    minor_val = str(row.get("小分類", "")).strip()
                    if minor_val and minor_val != "nan":
                        if minor_val not in new_minors:
                            new_minors.append(minor_val)
                
                # Identify deleted or renamed categories (which appear as deletions)
                removed_minors = old_minors - set(new_minors)
                used_sub_for_major = used_minors.get(selected_major, set())
                in_use_removed = removed_minors.intersection(used_sub_for_major)
                
                if in_use_removed:
                    err_msg = ", ".join(in_use_removed)
                    st.error(f"以下の小分類は既にレシートデータで登録済みのため、変更・削除できません: {err_msg}")
                else:
                    with st.spinner("保存中..."):
                        current_cats[selected_major] = new_minors
                        if save_categories_to_sheet(current_cats):
                            st.success(f"「{selected_major}」の小分類を保存しました！設定をアプリに反映するにはページをリロードしてください。")
                            import time
                            time.sleep(1)
                            st.rerun()
        else:
            st.info("大分類が登録されていません。左側から大分類を追加してください。")

'''
    
    content = content[:start_idx] + new_func_block + content[end_idx:]
    with open(r'c:\Users\t_kou\Kakeibo_Final_v3\app.py', 'w', encoding='utf-8') as f:
        f.write(content)
    print("Replaced show_category_master with validation")
else:
    print("Could not find target block")
