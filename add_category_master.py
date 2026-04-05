import sys
import re

with open(r'c:\Users\t_kou\Kakeibo_Final_v3\app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Add WORKSHEET_NAME
if 'CATEGORY_MASTER_WORKSHEET_NAME' not in content:
    content = content.replace(
        'PAYMENT_MASTER_WORKSHEET_NAME = "Payment_Master"',
        'PAYMENT_MASTER_WORKSHEET_NAME = "Payment_Master"\nCATEGORY_MASTER_WORKSHEET_NAME = "Category_Master"'
    )
    print("Added CATEGORY_MASTER_WORKSHEET_NAME")

# 2. Rename EXPENSE_CATEGORIES to DEFAULT_EXPENSE_CATEGORIES
content = content.replace('EXPENSE_CATEGORIES = {', 'DEFAULT_EXPENSE_CATEGORIES = {')
print("Renamed EXPENSE_CATEGORIES to DEFAULT_EXPENSE_CATEGORIES")

# 3. Modify get_categories()
old_get_cat = '''def get_categories():
    return EXPENSE_CATEGORIES'''

new_get_cat = '''@st.cache_data(ttl=600)
def get_categories():
    try:
        sheet = get_sheet(CATEGORY_MASTER_WORKSHEET_NAME, create_if_not_found=True)
        records = safe_gspread_call(sheet.get_all_records)
        
        if not records:
            headers = ["major_category", "minor_category"]
            safe_gspread_call(sheet.insert_row, headers, 1)
            rows = []
            for major, minors in DEFAULT_EXPENSE_CATEGORIES.items():
                if minors:
                    for minor in minors:
                        rows.append([major, minor])
                else:
                    rows.append([major, ""])
            if rows:
                safe_gspread_call(sheet.append_rows, rows)
            return DEFAULT_EXPENSE_CATEGORIES
            
        new_categories = {}
        for row in records:
            major = str(row.get("major_category", "")).strip()
            minor = str(row.get("minor_category", "")).strip()
            if major:
                if major not in new_categories:
                    new_categories[major] = []
                if minor and minor not in new_categories[major]:
                    new_categories[major].append(minor)
                    
        return new_categories if new_categories else DEFAULT_EXPENSE_CATEGORIES
    except Exception as e:
        print(f"Error loading categories: {e}")
        return DEFAULT_EXPENSE_CATEGORIES'''

if old_get_cat in content:
    content = content.replace(old_get_cat, new_get_cat)
    print("Replaced get_categories()")
else:
    print("old_get_cat not found! Let's check regex replacement later")

# 4. Replace remaining EXPENSE_CATEGORIES with get_categories()
# First replace the prompt text loop manually
content = content.replace('for major, minors in EXPENSE_CATEGORIES.items():', 'for major, minors in get_categories().items():')

# Then regex replace the rest
content = re.sub(r'(?<!DEFAULT_)EXPENSE_CATEGORIES', 'get_categories()', content)
print("Replaced remaining EXPENSE_CATEGORIES")

# 5. Add show_category_master() function completely.
new_func = '''def show_category_master():
    """カテゴリマスター設定画面 (オーナー専用)"""
    st.markdown("#### 📂 カテゴリマスター （大分類・小分類設定）")
    st.info("家計簿全体で使用される「大分類」と「小分類」の設定を行います。この画面はオーナー専用です。\\n\\n※ 新しく追加したカテゴリグラフの色は自動で割り当てられます。")
    
    current_cats = get_categories()
    
    rows = []
    for major, minors in current_cats.items():
        if minors:
            for minor in minors:
                rows.append({"大分類": major, "小分類": minor})
        else:
            rows.append({"大分類": major, "小分類": ""})
            
    import pandas as pd
    import time
    df = pd.DataFrame(rows)
    
    st.markdown("※ 以下の表を直接編集して行の追加・削除が可能です。「大分類」は必須項目です。")
    
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
    )
    
    if st.button("設定を保存", type="primary"):
        with st.spinner("保存中..."):
            try:
                sheet = get_sheet(CATEGORY_MASTER_WORKSHEET_NAME, create_if_not_found=True)
                safe_gspread_call(sheet.clear)
                
                new_data = [["major_category", "minor_category"]]
                for _, row in edited_df.iterrows():
                    major = str(row.get("大分類", "")).strip()
                    minor = str(row.get("小分類", "")).strip()
                    if major and major != "nan":
                         new_data.append([major, minor if minor != "nan" else ""])
                
                safe_gspread_call(sheet.update, range_name="A1", values=new_data)
                
                st.success("カテゴリを保存しました！設定を反映するにはページをリロードしてください。")
                get_categories.clear() # Cache clear
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"保存に失敗しました: {e}")

def show_payment_master():'''

if 'def show_payment_master():' in content and 'def show_category_master():' not in content:
    content = content.replace('def show_payment_master():', new_func)
    print("Added show_category_master()")
else:
    print("Could not add show_category_master()")

# 6. Sidebar logic
sidebar_old = '''            group4_opts = ["支払方法マスター", "プロフィール設定"]
            
            current_sel = st.session_state['menu_selection']'''

sidebar_new = '''            group4_opts = ["支払方法マスター", "プロフィール設定"]
            if st.session_state.get('username', '').lower() == 'tkouho':
                group4_opts.append("カテゴリマスター")
            
            current_sel = st.session_state['menu_selection']'''

if sidebar_old in content:
    content = content.replace(sidebar_old, sidebar_new)
    print("Updated sidebar")
else:
    print("Could not update sidebar")

# 7. Menu routing logic (main loop)
route_old = '''        elif menu_selection == "支払方法マスター":
            show_payment_master()'''

route_new = '''        elif menu_selection == "支払方法マスター":
            show_payment_master()
        elif menu_selection == "カテゴリマスター":
            show_category_master()'''

if route_old in content:
    content = content.replace(route_old, route_new)
    print("Updated routes")
else:
    print("Could not update routes")

with open(r'c:\Users\t_kou\Kakeibo_Final_v3\app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Saved app.py")
