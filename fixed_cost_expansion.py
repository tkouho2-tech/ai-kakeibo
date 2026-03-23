import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import gspread

import re

def _get_year_month(ym_str):
    s = str(ym_str).strip()
    if not s:
        return (9999, 12)
    m = re.search(r"(\d{4})[年/\.\-](\d{1,2})", s)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    m2 = re.search(r"(\d{2})[年/\.\-](\d{1,2})", s)
    if m2:
        return (2000 + int(m2.group(1)), int(m2.group(2)))
    return (9999, 12)

def _generate_target_months():
    # 2026.1月 to 2036.12月
    months = []
    for y in range(2026, 2037):
        for m in range(1, 13):
            months.append(f"{y}.{m}月")
    return months

def execute_expansion(username, mode="NEW"):
    """
    mode: "NEW", "RE_EXECUTE", "NEXT_MONTH"
    """
    from app import get_gspread_client, safe_gspread_call
    client = get_gspread_client()
    if not client:
        return False, "Google Docsへの接続に失敗しました。"
        
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
    except Exception as e:
        return False, f"支払管理シート({sheet_name})が見つかりません。"
        
    try:
        ws_master = ss.worksheet("固定費マスター")
        ws_pay = ss.worksheet("支払管理")
    except Exception as e:
        return False, f"「固定費マスター」または「支払管理」シートが見つかりません。"
        
    try:
        master_data = safe_gspread_call(ws_master.get_all_records)
    except Exception as e:
        # Fallback to get_all_values if header is weird
        master_raw = safe_gspread_call(ws_master.get_all_values)
        if not master_raw or len(master_raw) < 2:
            return False, "固定費マスターにデータがありません。"
        headers = master_raw[0]
        master_data = []
        for row in master_raw[1:]:
            row_dict = {}
            for i, h in enumerate(headers):
                if i < len(row):
                    row_dict[h] = row[i]
                else:
                    row_dict[h] = ""
            master_data.append(row_dict)

    pay_raw = safe_gspread_call(ws_pay.get_all_values)
    
    # Extract headers from pay sheet (row 7 -> index 6)
    if len(pay_raw) < 7:
        return False, "「支払管理」のフォーマットが正しくありません（7行目にヘッダーが必要です）。"
        
    pay_headers = pay_raw[6]
    month_cols = _generate_target_months()
    
    # Base columns before months
    base_cols = ["大分類", "変動or固定", "有限or無限", "科目１", "科目２", "Sno", "科目詳細"]
    # We will build matching rows.
    
    # Prepare old data if mode == NEXT_MONTH
    old_data_map = {}
    if mode == "NEXT_MONTH":
        now = datetime.now()
        current_ym_str = f"{now.year}.{now.month}月"
        # Find index of current month in month_cols
        try:
            current_month_idx = month_cols.index(current_ym_str)
        except:
            current_month_idx = 0
            
        old_month_cols = month_cols[:current_month_idx+1] # up to current month (inclusive)
        
        # Read old rows from row 8 onwards
        for row in pay_raw[7:]:
            # Ignore empty 科目1 or "計" or "合計"
            if len(row) < 7: continue
            
            clean_h = [str(x).replace("\n", "").replace(" ", "").strip() for x in pay_headers]
            k1_idx = next((i for i, h in enumerate(clean_h) if "科目1" in h or "科目１" in h or "固定支払1" in h or "固定支払１" in h), 3)
            det_idx = next((i for i, h in enumerate(clean_h) if "詳細" in h or "明細" in h), 6)
            
            k1 = row[k1_idx] if k1_idx < len(row) else ""
            k_detail = row[det_idx] if det_idx < len(row) else ""
            
            if not k1 or "計" in k1:
                continue
                
            key = f"{k1}_{k_detail}"
            old_vals = {}
            for mc in old_month_cols:
                try:
                    c_idx = pay_headers.index(mc)
                    old_vals[mc] = row[c_idx] if c_idx < len(row) else ""
                except:
                    old_vals[mc] = ""
            old_data_map[key] = old_vals

    new_rows_data = []
    
    # Process master data
    category_groups = {"クレジットカード": [], "口座引落": [], "銀行振込": []}
    
    sno = 1
    def _find_val(d, keywords, exclude=[]):
        for k, v in d.items():
            clean_k = str(k).replace("\n", "").replace(" ", "").replace("　", "").strip()
            for kw in keywords:
                if kw in clean_k:
                    if any(ex in clean_k for ex in exclude):
                        continue
                    return v
        return ""
        
    for m_rec in master_data:
        # Robust key matching
        k1 = str(_find_val(m_rec, ["科目1", "科目１", "固定支払1", "固定支払１"])).strip()
        if not k1:
            k1 = str(m_rec.get("科目１", m_rec.get("固定支払１", ""))).strip()
            
        k2 = str(_find_val(m_rec, ["科目2", "科目２", "固定支払2", "固定支払２"])).strip()
        is_finite_str = str(_find_val(m_rec, ["有限", "無限"])).strip()
        detail = str(_find_val(m_rec, ["詳細", "明細"])).strip()
        
        amt_str = str(_find_val(m_rec, ["支払額", "金額"], exclude=["最終月額", "最終"])).replace(",", "").replace("¥", "").replace("￥", "")
        amt = int(amt_str) if amt_str.isdigit() else 0
        
        fixed_var = str(_find_val(m_rec, ["変動", "固定"], exclude=["支払"])).strip()
        pay_month_freq = str(_find_val(m_rec, ["支払月", "頻度"])).strip()
        
        final_amt_str = str(_find_val(m_rec, ["最終月額"])).replace(",", "").replace("¥", "").replace("￥", "").strip()
        try:
            final_amt = int(final_amt_str) if final_amt_str else amt
        except:
            final_amt = amt
        
        fee_str = str(_find_val(m_rec, ["振込手数料", "手数料"])).replace(",", "").replace("¥", "").replace("￥", "")
        fee = int(fee_str) if fee_str.isdigit() else 0
        
        start_m_str = str(_find_val(m_rec, ["開始"])).strip()
        end_m_str = str(_find_val(m_rec, ["完済", "終了", "完了"])).strip()
        
        if not k1:
            continue
            
        is_finite = ("有限" in is_finite_str)
        sy, sm = _get_year_month(start_m_str) if start_m_str else (0,0)
        ey, em = _get_year_month(end_m_str) if is_finite and end_m_str else (9999,12)
        
        # Build row
        row_dict = {
            "大分類": "固定費",
            "変動or固定": fixed_var,
            "有限or無限": is_finite_str,
            "科目１": k1,
            "科目２": k2,
            "科目詳細": detail
        }
        
        key = f"{k1}_{detail}"
        
        # Calculate for each month
        for mc in month_cols:
            my, mm = [int(x.replace("月","")) for x in mc.split(".")]
            
            # For NEXT_MONTH mode, pull from old_data if < current_month
            if mode == "NEXT_MONTH" and key in old_data_map and mc in old_data_map[key]:
                row_dict[mc] = old_data_map[key][mc]
                continue
                
            val = ""
            # Check Active duration
            if (my > sy) or (my == sy and mm >= sm):
                if (my < ey) or (my == ey and mm <= em):
                    # It's active
                    
                    # Check Frequency
                    is_pay_month = False
                    if pay_month_freq == "毎月":
                        is_pay_month = True
                    else:
                        # expected "9月" or similar
                        if str(mm) in pay_month_freq:
                            is_pay_month = True
                            
                    if is_pay_month:
                        # Is it the very last month?
                        if is_finite and my == ey and mm == em:
                            v = final_amt
                        else:
                            v = amt
                            
                        # Add fee
                        v += fee
                        row_dict[mc] = v
                        
            if mc not in row_dict:
                row_dict[mc] = val
                
        # Group by 科目1
        if k1 not in category_groups:
            category_groups[k1] = []
        category_groups[k1].append(row_dict)

    # Now assemble the final sheet data
    # Ensure standard order: クレジットカード -> 口座引落 -> 銀行振込
    target_k1_order = ["クレジットカード", "口座引落", "銀行振込"]
    
    # Build array for batch update
    final_sheet_array = []
    
    # Retain the first 7 rows (headers, titles etc) from original sheet
    for i in range(7):
        if i < len(pay_raw):
            final_sheet_array.append(pay_raw[i])
        else:
            final_sheet_array.append([""] * len(pay_headers))
            
    header_len = len(pay_headers)
    
    def dict_to_row(d):
        r = []
        for h in pay_headers:
            clean_h = str(h).replace("\n", "").replace(" ", "").strip()
            if "科目1" in clean_h or "科目１" in clean_h or "固定支払1" in clean_h or "固定支払１" in clean_h: val = d.get("科目１", "")
            elif "科目2" in clean_h or "科目２" in clean_h or "固定支払2" in clean_h or "固定支払２" in clean_h: val = d.get("科目２", "")
            elif "変動" in clean_h or ("固定" in clean_h and "支払" not in clean_h): val = d.get("変動or固定", "")
            elif "有限" in clean_h or "無限" in clean_h: val = d.get("有限or無限", "")
            elif "Sno" in clean_h or "seq" in clean_h.lower(): val = d.get("Sno", "")
            elif "詳細" in clean_h or "明細" in clean_h: val = d.get("科目詳細", "")
            elif "大分類" in clean_h: val = d.get("大分類", "")
            else:
                val = d.get(h, "")
            r.append(val)
        return r
        
    start_row_idx = 8 # 1-based, after row 7 headers
    current_row_num = start_row_idx
    
    # Formula ranges memory
    group_ranges = []
    
    # Track rows for '変動' highlighting
    variable_rows = []
    
    # Format requests for borders
    sheet_id = ws_pay.id
    format_requests = []
    
    # 1. Clear previous borders from row 8 downwards
    format_requests.append({
        "updateBorders": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 7,
                "endRowIndex": 500,
                "startColumnIndex": 0,
                "endColumnIndex": header_len
            },
            "top": {"style": "NONE"},
            "bottom": {"style": "NONE"},
            "left": {"style": "NONE"},
            "right": {"style": "NONE"},
            "innerHorizontal": {"style": "NONE"},
            "innerVertical": {"style": "NONE"}
        }
    })
    
    for tk in target_k1_order:
        group_rows = category_groups.get(tk, [])
        g_start = current_row_num
        sno = 1
        
        last_var = ""
        last_inf = ""
        last_k2 = ""
        
        for gr in group_rows:
            gr["Sno"] = str(sno)
            sno += 1
            last_var = gr["変動or固定"]
            last_inf = gr["有限or無限"]
            last_k2 = gr["科目２"]
            
            if "変動" in str(last_var):
                variable_rows.append(current_row_num - 1)
                
            final_sheet_array.append(dict_to_row(gr))
            current_row_num += 1
            
        # Ensure EXACTLY 3 empty rows
        for _ in range(3):
            empty_dict = {
                "大分類": "固定費",
                "変動or固定": last_var,
                "有限or無限": last_inf,
                "科目１": tk,
                "科目２": last_k2,
                "Sno": str(sno)
            }
            sno += 1
            
            if "変動" in str(last_var):
                variable_rows.append(current_row_num - 1)
                
            r = dict_to_row(empty_dict)
            final_sheet_array.append(r)
            current_row_num += 1
            
        g_end = current_row_num - 1
        
        # Add 計 row
        subtotal_row = [""] * header_len
        try:
            k1_idx = pay_headers.index("科目１")
            subtotal_row[k1_idx] = f"【{tk} 計】"
        except: pass
        
        # Add formulas for months
        for mc in month_cols:
            try:
                c_idx = pay_headers.index(mc)
                col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                if g_start <= g_end:
                    formula = f"=SUM({col_letter}{g_start}:{col_letter}{g_end})"
                    subtotal_row[c_idx] = formula
            except: pass
            
        final_sheet_array.append(subtotal_row)
        
        # Add thick border format request for this group
        format_requests.append({
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": g_start - 1, # 0-based
                    "endRowIndex": current_row_num, # 0-based exclusive (includes the subtotal row)
                    "startColumnIndex": 0,
                    "endColumnIndex": header_len
                },
                "top": {"style": "SOLID_MEDIUM"},
                "bottom": {"style": "SOLID_MEDIUM"},
                "left": {"style": "SOLID_MEDIUM"},
                "right": {"style": "SOLID_MEDIUM"}
            }
        })
        
        group_ranges.append(current_row_num) # save the row number of the subtotal
        
        # Add background color and bold for subtotal row
        format_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": current_row_num - 1, # 0-based
                    "endRowIndex": current_row_num,
                    "startColumnIndex": 0,
                    "endColumnIndex": header_len
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 0.95,
                            "blue": 0.8
                        },
                        "textFormat": {
                            "bold": True
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat.bold)"
            }
        })
        
        current_row_num += 1
        
    # Add Grand Total Row
    grand_total_row = [""] * header_len
    try:
        k1_idx = pay_headers.index("科目１")
        grand_total_row[k1_idx] = "【合計】"
    except: pass
    
    for mc in month_cols:
        try:
            c_idx = pay_headers.index(mc)
            col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
            if group_ranges:
                cells = [f"{col_letter}{r}" for r in group_ranges]
                formula = f"=SUM({','.join(cells)})"
                grand_total_row[c_idx] = formula
        except: pass
        
    final_sheet_array.append(grand_total_row)
    
    # Add background color and bold for grand total row
    format_requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": current_row_num - 1,
                "endRowIndex": current_row_num,
                "startColumnIndex": 0,
                "endColumnIndex": header_len
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {
                        "red": 0.85,
                        "green": 0.95,
                        "blue": 0.85
                    },
                    "textFormat": {
                        "bold": True
                    }
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat.bold)"
        }
    })
    total_data_end = current_row_num
    
    # Base borders request for ALL data (insert at index 1 so it runs after ANY clears and before THICK borders)
    base_border_request = {
        "updateBorders": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 6, # Also cover header for normal borders just in case
                "endRowIndex": total_data_end,
                "startColumnIndex": 0,
                "endColumnIndex": header_len
            },
            "top": {"style": "SOLID"},
            "bottom": {"style": "SOLID"},
            "left": {"style": "SOLID"},
            "right": {"style": "SOLID"},
            "innerHorizontal": {"style": "SOLID"},
            "innerVertical": {"style": "SOLID"}
        }
    }
    format_requests.insert(1, base_border_request)
    
    # Add thick borders for each year (12 month columns)
    try:
        years = sorted(list(set([mc.split(".")[0] for mc in month_cols])))
        for y in years:
            start_m = f"{y}.1月"
            end_m = f"{y}.12月"
            if start_m in pay_headers and end_m in pay_headers:
                sc = pay_headers.index(start_m)
                ec = pay_headers.index(end_m)
                format_requests.append({
                    "updateBorders": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 6, # include header
                            "endRowIndex": total_data_end,
                            "startColumnIndex": sc,
                            "endColumnIndex": ec + 1 # exclusive
                        },
                        "top": {"style": "SOLID_MEDIUM"},
                        "bottom": {"style": "SOLID_MEDIUM"},
                        "left": {"style": "SOLID_MEDIUM"},
                        "right": {"style": "SOLID_MEDIUM"}
                    }
                })
    except:
        pass
        
    # Format numbers with commas (#,##0)
    try:
        first_month = month_cols[0]
        if first_month in pay_headers:
            start_m_col = pay_headers.index(first_month)
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 7, # data only
                        "endRowIndex": total_data_end,
                        "startColumnIndex": start_m_col,
                        "endColumnIndex": header_len
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "NUMBER",
                                "pattern": "#,##0"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            })
    except:
        pass
        
    # Format rows with "変動" to have yellow background in cols B to G (index 1 to 7)
    for r_idx in variable_rows:
        format_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": r_idx,
                    "endRowIndex": r_idx + 1,
                    "startColumnIndex": 1,
                    "endColumnIndex": 7
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 1.0,
                            "blue": 0.4  # Yellow
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        })
    
    # Write back to sheet and apply formats
    try:
        safe_gspread_call(ws_pay.clear)
        safe_gspread_call(ws_pay.update, "A1", final_sheet_array, value_input_option='USER_ENTERED')
        # Re-apply formatting
        safe_gspread_call(ss.batch_update, {"requests": format_requests})

        return True, "データ展開に成功しました！"
    except Exception as e:
        return False, f"書き込みエラー: {e}"

def show_open_management_sheet():
    """管理シートを開く/削除する UI"""
    st.markdown("## 📊 管理シートの確認と削除")
    
    username = st.session_state.get("username", "")
    from app import get_gspread_client, safe_gspread_call
    client = get_gspread_client()
    
    if not client:
        st.error("Google Drive APIに接続できません。")
        return
        
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
        try:
            ws = ss.worksheet("支払管理")
            target_url = getattr(ws, 'url', f"{ss.url}#gid={ws.id}")
        except:
            target_url = ss.url
    except Exception as e:
        st.warning(f"現在、あなた（{username}）専用の支払管理シートは見つかりません。")
        st.info("「管理シート新規作成」メニューからシートを発行してください。")
        return
        
    st.info(f"あなたの固定費管理シート「{sheet_name}」が連携されています。")
    
    col1, col2 = st.columns(2)
    with col1:
        st.link_button("🌐 開く", url=target_url, type="primary", use_container_width=True)
    with col2:
        if st.button("🗑️ 削除", use_container_width=True):
            st.session_state["confirm_delete_sheet"] = True
            st.rerun()
            
    if st.session_state.get("confirm_delete_sheet", False):
        st.error("⚠️ 全ての固定費管理データが削除されてしまいます。\n一度開いて内容をご確認頂き、問題ないかどうかを先にご確認ください。")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("削除する", type="primary", use_container_width=True):
                try:
                    with st.spinner("削除中..."):
                        safe_gspread_call(client.del_spreadsheet, ss.id)
                    st.success(f"固定費管理シート「{sheet_name}」を削除しました。")
                    st.session_state["confirm_delete_sheet"] = False
                    import time
                    time.sleep(2)
                    st.rerun()
                except Exception as e:
                    if "403" in str(e) or "permissions" in str(e).lower():
                        st.error(f"⚠️ Googleドライブの権限制限により、アプリから直接シートを削除できませんでした（シート作成者があなた自身のため）。\n\n誠にお手数ですが、上記の「開く」ボタンからスプレッドシートを開き、Googleドライブの画面から手動でゴミ箱へ移動しファイル削除をお願いいたします。")
                        st.info("削除後、再度「管理シート新規作成」メニューを実行すると、空の新しいシートを生成できるようになります。")
                    else:
                        st.error(f"削除中にエラーが発生しました: {e}")
        with c2:
            if st.button("キャンセル", use_container_width=True):
                st.session_state["confirm_delete_sheet"] = False
                st.rerun()

def show_fixed_cost_data_expansion():
    """固定費データ展開 UI"""
    st.markdown("## 🛠️ 固定費データ展開")
    st.info("「固定費マスター」の情報をもとに、「支払管理」シートに月別のデータ（2036年12月まで）を展開・更新します。")
    
    username = st.session_state.get("username", "")
    
    from app import get_gspread_client, safe_gspread_call
    
    # 状態確認
    client = get_gspread_client()
    if not client:
        st.error("Google Drive APIに接続できません。")
        return
        
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
        ws_pay = ss.worksheet("支払管理")
        pay_raw = safe_gspread_call(ws_pay.get_all_values)
    except Exception as e:
        st.warning(f"現在、あなた（{username}）専用の支払管理シート、または必要なシートが見つかりません。")
        st.info("「管理シート新規作成」メニューからシートを発行してください。")
        return
        
    # 展開済みか判定
    # 8行目以降に何らかのデータがあるか
    is_expanded = False
    if len(pay_raw) > 7:
        # Check if rows 8+ have actual data
        for row in pay_raw[7:]:
            # check any cell has value
            if any(cell.strip() for cell in row):
                is_expanded = True
                break
                
    if is_expanded:
        st.warning(f"「{sheet_name}」は既にデータ展開済です。下記をご確認下さい。")
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("再実行", type="primary", use_container_width=True):
                st.session_state["fce_action"] = "RE_EXECUTE"
        with col2:
            if st.button("翌月以降", use_container_width=True):
                st.session_state["fce_action"] = "NEXT_MONTH"
        with col3:
            if st.button("キャンセル", use_container_width=True):
                st.session_state["fce_action"] = "CANCEL"
                st.rerun()
                
        action = st.session_state.get("fce_action")
        if action in ["RE_EXECUTE", "NEXT_MONTH"]:
            with st.spinner("データ展開中...（数秒〜数十秒かかります）"):
                success, msg = execute_expansion(username, mode=action)
                if success:
                    st.success(msg)
                    st.markdown(f"**🔗 [支払管理シートを開く]({ss.url})**")
                    st.session_state["fce_action"] = None
                else:
                    st.error(msg)
                    
    else:
        if st.button("新規データ展開", type="primary"):
            with st.spinner("データ展開中...（数秒〜数十秒かかります）"):
                success, msg = execute_expansion(username, mode="NEW")
                if success:
                    st.success(msg)
                    st.markdown(f"**🔗 [支払管理シートを開く]({ss.url})**")
                else:
                    st.error(msg)
