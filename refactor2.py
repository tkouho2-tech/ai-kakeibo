import os

def rewrite_code():
    target_file = 'fixed_cost_expansion.py'
    with open(target_file, 'rb') as f:
        content = f.read().decode('utf-8', 'replace')
    
    lines = content.splitlines()

    def find_bounds(func_name):
        s_idx = -1
        for i, l in enumerate(lines):
            if l.startswith(f'def {func_name}('):
                s_idx = i
                break
        if s_idx == -1: return -1, -1
        e_idx = len(lines)
        for i in range(s_idx + 1, len(lines)):
            if lines[i].startswith('def ') or lines[i].startswith('class '):
                e_idx = i
                break
        return s_idx, e_idx

    e_s, e_e = find_bounds('execute_expansion')
    ui_s, ui_e = find_bounds('show_fixed_cost_data_expansion')
    var_s, var_e = find_bounds('execute_variable_cost_update')
    
    print(f"execute_expansion: {e_s} to {e_e}")
    print(f"show_fixed_cost_data_expansion: {ui_s} to {ui_e}")
    print(f"execute_variable_cost_update: {var_s} to {var_e}")
    
    new_show_ui = [
        'def show_fixed_cost_data_expansion():',
        '    import streamlit as st',
        '    st.markdown("## 🛠️ 固定費データ展開")',
        '    st.info("「固定費マスター」の情報をもとに、「支払管理」シートに月別のデータを「値のみ」で安全に展開・追加します。")',
        '    username = st.session_state.get("username", "")',
        '    from app import get_gspread_client, safe_gspread_call',
        '    client = get_gspread_client()',
        '    if not client:',
        '        st.error("Google Drive APIに接続できません。")',
        '        return',
        '    ss = None',
        '    try:',
        '        ss = client.open(f"{username}_支払管理")',
        '    except: pass',
        '    ',
        '    if st.button("🚀 固定費データ展開・更新", type="primary", use_container_width=True):',
        '        with st.spinner("データ展開中...（数秒〜数十秒かかります）"):',
        '            success, msg = execute_expansion(username)',
        '            if success:',
        '                with st.spinner("続けて変動費（クレジットカード利用等）を集計中..."):',
        '                    v_success, v_msg = execute_variable_cost_update(username, skip_backup=True)',
        '                    if v_success:',
        '                        st.success("固定費データの展開と変動費の集計が完了しました！")',
        '                    else:',
        '                        st.warning(f"固定費展開は完了しましたが、変動費更新でエラーが発生しました: {v_msg}")',
        '                if ss:',
        '                    st.markdown(f"**🔗 [支払管理シートを確認]({ss.url})**")',
        '            else:',
        '                st.error(msg)'
    ]

    new_exec = [
        'def execute_expansion(username, mode="NEW", start_ym=None):',
        '    from app import get_gspread_client, safe_gspread_call',
        '    from fixed_cost_expansion import _find_val, _clean_val, _normalize, _get_year_month',
        '    import re',
        '    ',
        '    client = get_gspread_client()',
        '    if not client: return False, "Google Drive APIに接続できません。"',
        '    ',
        '    try:',
        '        ss = client.open(f"{username}_支払管理")',
        '        ws_master = ss.worksheet("固定費マスター")',
        '        ws_pay = next((ws for ws in ss.worksheets() if ws.title.strip() == "支払管理"), None)',
        '        if not ws_pay: return False, "支払管理シートが見つかりません。"',
        '    except Exception as e:',
        '        return False, f"シート読み込みエラー: {e}"',
        '        ',
        '    master_raw = safe_gspread_call(ws_master.get_all_values)',
        '    if not master_raw or len(master_raw) < 2: return False, "固定費マスターにデータがありません。"',
        '    headers = master_raw[0]',
        '    master_data = [dict(zip(headers, r + [""]*(len(headers)-len(r)))) for r in master_raw[1:]]',
        '    ',
        '    # GET PAYMENT SHEET DATA',
        '    pay_formatted = safe_gspread_call(ws_pay.get_all_values, value_render_option="FORMATTED_VALUE")',
        '    ',
        '    h_row_idx = -1',
        '    for i_r, r_v in enumerate(pay_formatted):',
        '        if r_v and str(r_v[0]).strip().lower() in ["id", "key"]:',
        '            h_row_idx = i_r',
        '            break',
        '    if h_row_idx == -1: h_row_idx = 6',
        '    ',
        '    actual_headers = pay_formatted[h_row_idx]',
        '    ',
        '    # GET COLUMNS BY MONTH',
        '    month_cols = []',
        '    for i, h in enumerate(actual_headers):',
        '        h_clean = _clean_val(h).strip()',
        '        y_m = _get_year_month(h_clean)',
        '        if y_m != (9999, 12):',
        '            month_cols.append({"col_idx": i, "year": y_m[0], "month": y_m[1], "ym": y_m[0]*100 + y_m[1]})',
        '            ',
        '    # FIND DETAIL COLUMN',
        '    det_idx = next((i for i, h in enumerate(actual_headers) if "詳細" in h or "明細" in h), -1)',
        '    if det_idx == -1: return False, "支払管理シートに「科目明細」列が見つかりません。"',
        '    ',
        '    # PREPARE BATCH UPDATE',
        '    requests = []',
        '    ',
        '    for m_rec in master_data:',
        '        m_detail = _clean_val(_find_val(m_rec, ["詳細", "明細"])).strip()',
        '        if not m_detail: continue',
        '        ',
        '        # FIND TARGET ROW IN PAY SHEET',
        '        target_r_idx = -1',
        '        for r_i, r_v in enumerate(pay_formatted):',
        '            if r_i <= h_row_idx: continue',
        '            if len(r_v) > det_idx and _clean_val(r_v[det_idx]).strip() == m_detail:',
        '                target_r_idx = r_i',
        '                break',
        '                ',
        '        if target_r_idx == -1:',
        '            continue # SKIP TARGETS NOT IN SPREADSHEET (NO ROW INSERTIONS)',
        '            ',
        '        amt_str = str(_find_val(m_rec, ["支払額", "金額"], exclude=["最終月額", "最終"])).replace(",", "").replace("¥", "").replace("￥", "")',
        '        amt = amt_str.strip()',
        '        final_amt_str = str(_find_val(m_rec, ["最終月額"])).replace(",", "").replace("¥", "").replace("￥", "").strip()',
        '        final_amt = final_amt_str if final_amt_str else amt',
        '        ',
        '        start_m_str = str(_find_val(m_rec, ["開始"])).strip()',
        '        end_m_str = str(_find_val(m_rec, ["完済", "終了", "完了"])).strip()',
        '        is_finite_str = str(_find_val(m_rec, ["有限", "無限"]))',
        '        pay_month_freq = str(_find_val(m_rec, ["支払月", "頻度"])).strip()',
        '        ',
        '        is_finite = "有限" in is_finite_str',
        '        sy, sm = _get_year_month(start_m_str) if start_m_str else (0,0)',
        '        ey, em = _get_year_month(end_m_str) if (is_finite and end_m_str) else (9999,12)',
        '        start_ym_val = sy * 100 + sm',
        '        end_ym_val = ey * 100 + em',
        '        ',
        '        for m_col in month_cols:',
        '            c_idx = m_col["col_idx"]',
        '            col_ym = m_col["ym"]',
        '            c_m = m_col["month"]',
        '            ',
        '            val_to_set = ""',
        '            if col_ym >= start_ym_val and (not is_finite or col_ym <= end_ym_val):',
        '                months_targeted = []',
        '                if "偶数" in pay_month_freq:',
        '                    months_targeted = [2, 4, 6, 8, 10, 12]',
        '                elif "奇数" in pay_month_freq:',
        '                    months_targeted = [1, 3, 5, 7, 9, 11]',
        '                else:',
        '                    mm = re.findall(r"\d+", pay_month_freq)',
        '                    if mm:',
        '                        months_targeted = [int(x) for x in mm]',
        '                    else:',
        '                        months_targeted = [c_m]',
        '                ',
        '                if c_m in months_targeted:',
        '                    if is_finite and col_ym == end_ym_val:',
        '                        val_to_set = final_amt',
        '                    else:',
        '                        val_to_set = amt',
        '            ',
        '            current_val = pay_formatted[target_r_idx][c_idx] if c_idx < len(pay_formatted[target_r_idx]) else ""',
        '            if val_to_set and val_to_set != str(current_val).replace(",", ""):',
        '                col_letter = chr(ord("A") + c_idx) if c_idx < 26 else chr(ord("A") + c_idx//26 - 1) + chr(ord("A") + c_idx%26)',
        '                cell_name = f"{col_letter}{target_r_idx + 1}"',
        '                requests.append({',
        '                    "range": f"支払管理!{cell_name}",',
        '                    "values": [[int(val_to_set) if val_to_set.isdigit() else val_to_set]]',
        '                })',
        '                ',
        '    if requests:',
        '        safe_gspread_call(ss.values_batch_update, {"valueInputOption": "USER_ENTERED", "data": requests})',
        '        ',
        '    # Call the variable cost update explicitly as required in the old flow',
        '    return True, "データ展開に成功しました！"'
    ]

    # Reassemble safely using indices (we assume order: e_s < ui_s < var_s)
    # The original AST shows: 169 (execute), 1198 (open), 1227 (show_fixed), 1382 (var)
    # Since we only replace `execute` and `show_fixed`, we leave the rest untouched.
    final_lines = []
    
    # 0 to start of execute_expansion
    final_lines.extend(lines[:e_s])
    # newly replaced execute_expansion
    final_lines.extend(new_exec)
    # end of execute_expansion to start of show_fixed_cost_data_expansion
    final_lines.extend(lines[e_e:ui_s])
    # newly replaced show_fixed_cost_data_expansion
    final_lines.extend(new_show_ui)
    # end of show_fixed_cost_data_expansion to start of execute_variable_cost_update
    final_lines.extend(lines[ui_e:var_s])
    
    # Let's fix the variable cost update directly
    var_lines = lines[var_s:var_e]
    for i, vl in enumerate(var_lines):
        if 'add_rows' in vl or 'insert_rows' in vl or 'append_row' in vl:
            var_lines[i] = '        pass # FIXED_FORMAT_REMOVED_NO_INSERTIONS'
            
    final_lines.extend(var_lines)
    # Everything else
    final_lines.extend(lines[var_e:])
    
    with open(target_file, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(final_lines))

if __name__ == "__main__":
    rewrite_code()
    print("Done")
