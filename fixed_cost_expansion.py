import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import gspread

import re

def safe_money_int_cast(val):
    """
    金額文字列（カンマ、￥、小数点あり）を安全に整数に変換する。
    """
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return int(val)
    try:
        # 文字列クリーンアップ
        s = str(val).replace(",", "").replace("￥", "").replace("¥", "").strip()
        if not s:
            return 0
        # 小数点が含まれる場合は一旦floatにしてからintにする
        if "." in s:
            return int(float(s))
        return int(s)
    except Exception:
        return 0

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

    # 数式を維持するために FORMULA レンダリングオプションで取得
    pay_raw = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMULA')
    
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
            
            clean_h = [str(x).strip() for x in pay_headers]
            # 数式形式のヘッダーを考慮
            def _clean_v(v):
                s = str(v).strip()
                if s.startswith("="):
                    s = s[1:].strip()
                    if s.startswith('"') and s.endswith('"'):
                        s = s[1:-1].strip()
                return s
            
            clean_h_ids = [_clean_v(x).replace("\n", "").replace(" ", "").strip() for x in pay_headers]
            k1_idx = next((i for i, h in enumerate(clean_h_ids) if "科目1" in h or "科目１" in h or "固定支払1" in h or "固定支払１" in h), 3)
            det_idx = next((i for i, h in enumerate(clean_h_ids) if "詳細" in h or "明細" in h), 6)
            
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
        amt = safe_money_int_cast(amt_str)
        
        fixed_var = str(_find_val(m_rec, ["変動", "固定"], exclude=["支払"])).strip()
        pay_month_freq = str(_find_val(m_rec, ["支払月", "頻度"])).strip()
        
        final_amt_str = str(_find_val(m_rec, ["最終月額"])).replace(",", "").replace("¥", "").replace("￥", "").strip()
        final_amt = safe_money_int_cast(final_amt_str) if final_amt_str else amt
        
        fee_str = str(_find_val(m_rec, ["振込手数料", "手数料"])).replace(",", "").replace("¥", "").replace("￥", "")
        fee = safe_money_int_cast(fee_str)
        
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
                f_idx = c_idx + 1 # 完了F
                col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                if g_start <= g_end:
                    # 完了Fが空の行のみを合計
                    formula = f"=SUMIFS({col_letter}{g_start}:{col_letter}{g_end}, {flag_letter}{g_start}:{flag_letter}{g_end}, \"\")"
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
        # ヘッダー項目をクリーンアップして検索 (="科目１" 等の数式対応)
        def _clean_v(v):
            s = str(v).strip()
            if s.startswith("="):
                s = s[1:].strip()
                if s.startswith('"') and s.endswith('"'):
                    s = s[1:-1].strip()
            return s
        clean_pay_headers = [_clean_v(h) for h in pay_headers]
        k1_idx = clean_pay_headers.index("科目１")
        
        # A列(インデックス0) と 科目1列 の両方に文言を入れる（セル結合されるため）
        grand_total_row[0] = "固定費合計"
        grand_total_row[k1_idx] = "固定費合計"
        
        # 科目明細列（G列付近）にも念のため文言を入れる
        k_detail_idx = -1
        for h_i, h_val in enumerate(pay_headers):
            h_clean = _clean_v(h_val).strip()
            if "科目明細" in h_clean or "詳細" in h_clean: 
                k_detail_idx = h_i
                break
        if k_detail_idx != -1:
            grand_total_row[k_detail_idx] = "固定費合計"
            
    except: pass
    
    for mc in month_cols:
        try:
            c_idx = pay_headers.index(mc)
            f_idx = c_idx + 1
            col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
            flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
            if group_ranges:
                # 各サブグループの「計」行そのものがすでにSUMIFSになっているので、
                # グランドトータルは単純にそれ。を合計しても良いが、完了Fが1の「計」行はないはず。
                # ユーザーの意図を汲み取り、ここもSUMIFS(計の行, その右, "")にする。
                cells = [f"IF({flag_letter}{r}=\"\", {col_letter}{r}, 0)" for r in group_ranges]
                formula = f"=SUM({','.join(cells)})"
                grand_total_row[c_idx] = formula
        except: pass
        
    # 合計行のA〜G列（またはheader_len）を結合して中央揃え
    try:
        k_detail_idx = -1
        for h_i, h_val in enumerate(pay_headers):
            h_clean = _clean_v(h_val).strip()
            if "科目明細" in h_clean or "詳細" in h_clean: k_detail_idx = h_i
        
        # もし科目明細が見つからなければ、G列（インデックス6）までを対象とする
        merge_end = k_detail_idx + 1 if k_detail_idx != -1 else 7
        
        # 既存の結合を解除（エラー防止）
        format_requests.append({
            "unmergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": current_row_num - 1,
                    "endRowIndex": current_row_num,
                    "startColumnIndex": 0,
                    "endColumnIndex": header_len
                }
            }
        })
        
        format_requests.append({
            "mergeCells": {
                "mergeType": "MERGE_ALL",
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": current_row_num - 1,
                    "endRowIndex": current_row_num,
                    "startColumnIndex": 0,
                    "endColumnIndex": merge_end
                }
            }
        })
        # 中央揃え
        format_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": current_row_num - 1,
                    "endRowIndex": current_row_num,
                    "startColumnIndex": 0,
                    "endColumnIndex": merge_end
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
            }
        })
    except: pass
    
    final_sheet_array.append(grand_total_row)
    
    # Add background color and bold for grand total row (Fixed Cost Total - Blue Reverse)
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
                        "red": 0.0,
                        "green": 0.0,
                        "blue": 1.0
                    },
                    "textFormat": {
                        "bold": True,
                        "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}
                    }
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat.bold,textFormat.foregroundColor)"
        }
    })
    total_data_end = current_row_num
    
    # Base borders request for ALL data (insert at index 1 so it runs after ANY clears and before THICK borders)
    base_border_request = {
        "updateBorders": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 7, # 8行目から
                "endRowIndex": 1000, # 常に1000行目まで書式を維持（空行でも数式や書式が入るように）
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
                        "endRowIndex": 1000, # Extend to cover potential future rows
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
        
    # 完了Fの左側を破線にする
    try:
        for c_idx, h_name in enumerate(pay_headers):
            clean_h = str(h_name).replace("\n", "").replace(" ", "").strip()
            if "完了" in clean_h:
                format_requests.append({
                    "updateBorders": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 7, 
                            "endRowIndex": 1000,
                            "startColumnIndex": c_idx,
                            "endColumnIndex": c_idx + 1
                        },
                        "left": {"style": "DASHED"}
                    }
                })
    except: pass
        
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
        # 不要な全クリアは避け、データが入る可能性のある範囲の値のみをクリアして更新
        # これにより、シート自体の設定やヘッダー外の書式が安定します。
        safe_gspread_call(ws_pay.batch_clear, [f"A8:ZZ1000"]) 
        safe_gspread_call(ws_pay.update, "A1", final_sheet_array, value_input_option='USER_ENTERED')
        # 書式を一括適用
        safe_gspread_call(ss.batch_update, {"requests": format_requests})

        return True, "データ展開に成功しました！"
    except Exception as e:
        return False, f"書き込みエラー: {e}"

def show_open_management_sheet():
    """支払管理シートを開く/削除する UI"""
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
        st.info("「支払管理シート新規作成」メニューからシートを発行してください。")
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
                        st.info("削除後、再度「支払管理シート新規作成」メニューを実行すると、空の新しいシートを生成できるようになります。")
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
        st.info("「支払管理シート新規作成」メニューからシートを発行してください。")
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

def execute_variable_cost_update(username):
    from app import get_gspread_client, safe_gspread_call, get_payment_methods, get_sheet, TRANSACTIONS_WORKSHEET_NAME
    import calendar
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    
    client = get_gspread_client()
    if not client:
        return False, "Google Docsへの接続に失敗しました。"
        
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
        ws_pay = ss.worksheet("支払管理")
    except Exception as e:
        return False, f"支払管理シート({sheet_name})が見つかりません。先に「支払管理シート新規作成」を実行してください。"
        
    # 数式を維持するために FORMULA レンダリングオプションで取得
    pay_raw = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMULA')
    if len(pay_raw) < 7:
        return False, "「支払管理」のフォーマットが正しくありません。"
        
    pay_headers = pay_raw[6]
    
    # 文字列が数式形式 (="値") の場合に中身を取り出す補助関数
    def _clean_val(v):
        s = str(v).strip()
        if s.startswith("="):
            s = s[1:].strip()
            if s.startswith('"') and s.endswith('"'):
                s = s[1:-1].strip()
        return s

    # 科目１列のインデックスを探す
    k1_idx = -1
    for i, h in enumerate(pay_headers):
        clean_h = _clean_val(h).replace("\n", "").replace(" ", "").strip()
        if "科目1" in clean_h or "科目１" in clean_h or "固定支払1" in clean_h or "固定支払１" in clean_h:
            k1_idx = i
            break
            
    if k1_idx == -1:
        return False, "ヘッダーから「科目１」列が見つかりません。"
        
    # 行をスキャンして「固定費合計」または「【合計】」を探す (大分類～科目詳細のどこにあっても見つける)
    total_row_idx = -1
    for i, row in enumerate(pay_raw):
        if i < 7: continue
        row_str = "".join([str(c) for c in row[:7]]) # 最初の数列を結合して検索
        if "固定費合計" in row_str or "【合計】" in row_str:
            total_row_idx = i
            break
            
    # 見つからない場合のフォールバック：上から順に SUM(IF... 数式がある最初の行を探す (ラベルが消えている場合への対策)
    if total_row_idx == -1:
        for i in range(7, len(pay_raw)):
            row = pay_raw[i]
            # 月カラム（通常はインデックス 7 以降）のどこかにグランド合計用の数式が入っているかチェック
            row_content = "".join([str(c) for c in row])
            if "=SUM(IF(" in row_content or "=SUMIFS(" in row_content:
                # 明細行のサブ計ではなく、複数の SUM かつ IF( が含まれるグランド合計っぽいもの
                if row_content.count("SUM(") >= 1 and row_content.count("IF(") > 1:
                    total_row_idx = i
                    break

    if total_row_idx == -1:
        return False, "「固定費合計」または「【合計】」行が見つかりません。先に「固定費データ展開」を実行してください。"
        
    # 既存の固定費エリアのサブ合計行を収集
    # (row_num, group_start_row, group_end_row)
    fixed_subtotals = []
    group_start = 8
    for i, row in enumerate(pay_raw[:total_row_idx + 1]):
        if i < 7: continue
        r_k1 = str(row[k1_idx]).strip() if k1_idx < len(row) else ""
        if ("【" in r_k1 and "計】" in r_k1) or "【合計】" in r_k1 or "固定費合計" in r_k1:
            fixed_subtotals.append((i + 1, group_start, i))
            group_start = i + 2
            
    from fixed_cost_expansion import _generate_target_months
    # 対象月カラムの抽出
    month_cols = _generate_target_months()
    
    # 新しい挿入開始行 (スプレッドシートの行番号は 1-based)
    start_row_num = total_row_idx + 2 
    
    try:
        methods = safe_gspread_call(get_payment_methods, username)
        cc_methods = [m for m in methods if m.get("is_credit_card", False) or m.get("type") == "クレジットカード"]
    except Exception as e:
        return False, f"支払方法マスターの取得に失敗しました: {e}"
        
    if not cc_methods:
        return False, "クレジットカードが1件も登録されていません。「支払方法マスター」を確認してください。"
        
    try:
        tx_sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
        all_txs = safe_gspread_call(tx_sheet.get_all_records)
        user_txs = [tx for tx in all_txs if str(tx.get("username", "")).lower() == username.lower()]
    except Exception as e:
        return False, f"取引履歴の取得に失敗しました: {e}"
        
    def _dict_to_row(d):
        r = []
        for h in pay_headers:
            clean_h = _clean_val(h).replace("\n", "").replace(" ", "").strip()
            if "科目1" in clean_h or "科目１" in clean_h or "固定支払1" in clean_h or "固定支払１" in clean_h: val = d.get("科目１", "")
            elif "科目2" in clean_h or "科目２" in clean_h or "固定支払2" in clean_h or "固定支払２" in clean_h: val = d.get("科目２", "")
            elif "変動" in clean_h or ("固定" in clean_h and "支払" not in clean_h): val = d.get("変動or固定", "")
            elif "有限" in clean_h or "無限" in clean_h: val = d.get("有限or無限", "")
            elif "Sno" in clean_h or "seq" in clean_h.lower(): val = d.get("Sno", "")
            elif "詳細" in clean_h or "明細" in clean_h: val = d.get("科目詳細", "")
            elif "大分類" in clean_h: val = d.get("大分類", "")
            else:
                # 月カラムのデータ取得 (キーはクリーンアップされたヘッダー名)
                val = d.get(_clean_val(h).strip(), "")
            r.append(val)
        return r

    cc_rows_array = []
    current_row_num = start_row_num
    var_start = current_row_num
    var_sno = 1
    
    # 事前準備：pay_raw内から各クレジットカードごとの固定費行(行番号, 1-based)を抽出
    # ヘッダーより下（7行目以降）から合計行まで
    fc_payment_rows = {cc.get("name", ""): [] for cc in cc_methods}
    # 他の支払情報 (口座引落, 銀行振込) の行を抽出
    other_pay_rows = []
    
    for i, row in enumerate(pay_raw[:total_row_idx]):
        if i < 7: continue
        try:
            # ヘッダー検索もクリーンアップされたインデックスを使用
            k1_h_idx = -1
            k2_h_idx = -1
            for h_i, h_val in enumerate(pay_headers):
                h_clean = _clean_val(h_val).strip()
                if h_clean == "科目１": k1_h_idx = h_i
                if h_clean == "科目２": k2_h_idx = h_i
            
            if k1_h_idx != -1 and k2_h_idx != -1:
                r_k1 = _clean_val(row[k1_h_idx]).strip()
                r_k2 = _clean_val(row[k2_h_idx]).strip()
                if r_k1 == "クレジットカード" and r_k2 in fc_payment_rows:
                    fc_payment_rows[r_k2].append(i + 1)
                elif r_k1 in ["口座引落", "銀行振込"]:
                    if r_k2:
                        other_pay_rows.append(i + 1)
        except: pass
        
    var_cost_rows = {}
    
    for cc in cc_methods:
        cc_name = cc.get("name", "")
        closing_str = str(cc.get("closing_date", ""))
        pay_month_str = str(cc.get("payment_month", ""))
        pay_date_str = str(cc.get("payment_date", ""))
        
        # 支払日の数値抽出と文言（翌月/当月）の判定
        try:
            p_day_match = re.search(r"\d+", str(pay_date_str))
            p_day_val = int(p_day_match.group()) if p_day_match else 27
        except: p_day_val = 27
        timing_label = "翌月" if p_day_val <= 20 else "当月"
        payment_desc = f"支払日は{timing_label}の{pay_date_str}となります。"

        cc_row_dict = {
            "大分類": "変動費",
            "変動or固定": "",
            "有限or無限": "",
            "科目１": "クレジットカード",
            "科目２": cc_name,
            "科目詳細": payment_desc,
            "Sno": str(var_sno)
        }
        var_sno += 1
        
        for mc in month_cols:
            try:
                my, mm = [int(x.replace("月","")) for x in mc.split(".")]
                base_date = datetime(my, mm, 1)
                
                # 表示ルールの適用: 20日基準で表示年月をシフト
                # 20日以前なら：支払月 ＝ 表示月
                # 21日以降なら：支払月 ＝ 表示月 － 1ヶ月 (翌月の21日以降なら翌月の年月に表示、の逆算)
                try:
                    p_day_m = re.search(r"\d+", str(pay_date_str))
                    p_day = int(p_day_m.group()) if p_day_m else 27
                except:
                    p_day = 27

                if p_day <= 20:
                    # 支払月の前月に表示する（例: 2/10払 → 1月カラム） => 支払月 ＝ 表示月 ＋ 1ヶ月
                    target_pay_date = base_date + relativedelta(months=1)
                else:
                    # 支払月と同じ年月に表示する（例: 2/27払 → 2月カラム） => 支払月 ＝ 表示月
                    target_pay_date = base_date
                
                from app import calculate_credit_card_periods
                periods = calculate_credit_card_periods(target_pay_date, closing_str, pay_month_str, pay_date_str)
                
                if periods and len(periods) > 0:
                    s_date = periods[0]["start"]
                    e_date = periods[0]["end"]
                    
                    month_sum = 0
                    for tx in user_txs:
                        if tx.get("payment_method") == cc_name:
                            # 内税は集計に含まない
                            if tx.get("category") == "消費税（内税）":
                                continue
                            tx_date_str = str(tx.get("date", ""))
                            if tx_date_str:
                                try:
                                    tx_date = datetime.strptime(tx_date_str, "%Y-%m-%d").date()
                                    if s_date <= tx_date <= e_date:
                                        amt = str(tx.get("amount", "0")).replace(",", "").replace("¥", "").replace("￥", "")
                                        if amt is not None and str(amt).strip() != "":
                                            month_sum += safe_money_int_cast(amt)
                                except:
                                    pass
                    cc_row_dict[mc] = month_sum if month_sum > 0 else ""
                else:
                    cc_row_dict[mc] = ""
            except Exception as e:
                cc_row_dict[mc] = ""

        var_cost_rows[cc_name] = current_row_num
        row_arr = _dict_to_row(cc_row_dict)
        
        # 月次データとフラグの転記
        for mc in month_cols:
            try:
                c_idx = pay_headers.index(mc)
                f_idx = c_idx + 1
                
                # 金額設定
                amt = cc_row_dict.get(mc, "")
                row_arr[c_idx] = amt
                
                # 支払日の判定と自動フラグ設定 (変動費・カード個別行)
                try:
                    my, mm = [int(x.replace("月","").replace("月","")) for x in mc.split(".")]
                    off = 1
                    if "当月" in str(pay_month_str): off = 0
                    elif "翌々月" in str(pay_month_str): off = 2
                    base_dt = datetime(my, mm, 1) + relativedelta(months=off)
                    d_m = re.search(r"\d+", str(pay_date_str))
                    if d_m:
                        actual_due_date = base_dt + relativedelta(day=int(d_m.group()))
                        if datetime.now().date() >= actual_due_date.date():
                            row_arr[f_idx] = "1"
                except: pass
            except: pass
            
        cc_rows_array.append(row_arr)
        current_row_num += 1
        
    # 変動費合計行を追加 (ピンク)
    var_total_row = [""] * len(pay_headers)
    try:
        k1_idx = pay_headers.index("科目１")
        # 画像に基づき「変動費合計」に変更
        var_total_row[0] = "変動費合計"
        var_total_row[k1_idx] = "変動費合計"
        
        # 明細列にも入れる
        k_detail_idx = -1
        for h_i, h_val in enumerate(pay_headers):
            h_clean = _clean_val(h_val).strip()
            if "詳細" in h_clean or "明細" in h_clean: k_detail_idx = h_i
        if k_detail_idx != -1:
            var_total_row[k_detail_idx] = "変動費合計"
    except: pass
    
    for mc in month_cols:
        try:
            c_idx = pay_headers.index(mc)
            f_idx = c_idx + 1
            col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
            flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
            if var_start <= current_row_num - 1:
                # 完了Fが空の行のみを合計
                formula = f"=SUMIFS({col_letter}{var_start}:{col_letter}{current_row_num - 1}, {flag_letter}{var_start}:{flag_letter}{current_row_num - 1}, \"\")"
                var_total_row[c_idx] = formula
        except: pass
        
    cc_rows_array.append(var_total_row)
    current_row_num += 1
    
    # === クレジットカード総合計ブロック ===
    summary_start_row = current_row_num
    card_total_rows = []
    
    summary_data_start = current_row_num
    
    # 最初の行にだけ「クレジットカード合計」を表示。他は空にする。
    is_first_summary_row = True
    
    for cc in cc_methods:
        cc_name = cc.get("name", "")
        pay_month_str = str(cc.get("payment_month", ""))
        pay_date_str = str(cc.get("payment_date", ""))
        
        # 1. 固定費分
        fc_dict = {
            "大分類": "クレジットカード合計" if is_first_summary_row else "", 
            "科目１": "クレジットカード合計" if is_first_summary_row else "", 
            "科目２": cc_name, 
            "科目詳細": "固定費分"
        }
        is_first_summary_row = False
        r_fc = _dict_to_row(fc_dict)
        for mc in month_cols:
            try:
                c_idx = pay_headers.index(mc)
                f_idx = c_idx + 1
                col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                if fc_payment_rows.get(cc_name):
                    cells = [f"IF({flag_letter}{r}=\"\", {col_letter}{r}, 0)" for r in fc_payment_rows[cc_name]]
                    r_fc[c_idx] = f"=SUM({','.join(cells)})"
                else:
                    r_fc[c_idx] = 0
                
                # 支払日の判定と自動フラグ設定 (固定費分)
                try:
                    y, m = _get_year_month(mc)
                    if y != 9999:
                        month_offset = 1 
                        if "当月" in str(pay_month_str): month_offset = 0
                        elif "翌々月" in str(pay_month_str): month_offset = 2
                        target_base_date = datetime(y, m, 1) + relativedelta(months=month_offset)
                        day_match = re.search(r"\d+", str(pay_date_str))
                        if day_match:
                            day_val = int(day_match.group())
                            actual_due_date = target_base_date + relativedelta(day=day_val)
                            if datetime.now().date() >= actual_due_date.date():
                                r_fc[f_idx] = 1
                except: pass
            except: pass
        cc_rows_array.append(r_fc)
        current_row_num += 1
        
        # 2. 変動費分
        vc_row_idx = var_cost_rows.get(cc_name)
        vc_dict = {"大分類": "", "科目１": "クレジットカード合計", "科目２": cc_name, "科目詳細": "変動費分"}
        r_vc = _dict_to_row(vc_dict)
        for mc in month_cols:
            try:
                c_idx = pay_headers.index(mc)
                f_idx = c_idx + 1
                col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                if vc_row_idx:
                    # 詳細行の完了Fが空の場合のみ表示
                    r_vc[c_idx] = f"=IF({flag_letter}{vc_row_idx}=\"\", {col_letter}{vc_row_idx}, 0)"
                else:
                    r_vc[c_idx] = 0

                # 支払日の判定と自動フラグ設定 (変動費分)
                try:
                    y, m = _get_year_month(mc)
                    if y != 9999:
                        month_offset = 1 
                        if "当月" in str(pay_month_str): month_offset = 0
                        elif "翌々月" in str(pay_month_str): month_offset = 2
                        target_base_date = datetime(y, m, 1) + relativedelta(months=month_offset)
                        day_match = re.search(r"\d+", str(pay_date_str))
                        if day_match:
                            day_val = int(day_match.group())
                            actual_due_date = target_base_date + relativedelta(day=day_val)
                            if datetime.now().date() >= actual_due_date.date():
                                r_vc[f_idx] = 1
                except: pass
            except: pass
        cc_rows_array.append(r_vc)
        current_row_num += 1
        
        # 3. 合計 (固定費分+変動費分)
        # 支払日の文言判定
        try:
            p_day_sum_m = re.search(r"\d+", str(pay_date_str))
            p_day_sum_v = int(p_day_sum_m.group()) if p_day_sum_m else 27
        except: p_day_sum_v = 27
        timing_sum_label = "翌月" if p_day_sum_v <= 20 else "当月"
        payment_sum_desc = f"支払日は{timing_sum_label}の{pay_date_str}となります。"

        sum_dict = {"大分類": "", "科目１": "", "科目２": cc_name, "科目詳細": payment_sum_desc}
        r_sum = _dict_to_row(sum_dict)
        for mc in month_cols:
            try:
                c_idx = pay_headers.index(mc)
                f_idx = c_idx + 1
                col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                
                fc_row = current_row_num - 2
                vc_row = current_row_num - 1
                # 各内訳行（固定費分、変動費分）の完了Fをチェックして合計する形式
                formula = f"SUMIFS({col_letter}{fc_row}:{col_letter}{vc_row}, {flag_letter}{fc_row}:{flag_letter}{vc_row}, \"\")"
                r_sum[c_idx] = "=" + formula
            except: pass
        cc_rows_array.append(r_sum)
        card_total_rows.append(current_row_num)
        current_row_num += 1
        
    # クレジットカード 総合計 (青)
    grand_dict = {"大分類": "", "科目１": "クレジットカード合計", "科目２": "すべて", "科目詳細": "総合計"}
    r_grand = _dict_to_row(grand_dict)
    # A列にも文言を入れる
    r_grand[0] = "クレジットカード合計"
    for mc in month_cols:
        try:
            c_idx = pay_headers.index(mc)
            f_idx = c_idx + 1
            col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
            flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
            if card_total_rows:
                cells = [f"IF({flag_letter}{r}=\"\", {col_letter}{r}, 0)" for r in card_total_rows]
                r_grand[c_idx] = f"=SUM({','.join(cells)})"
            else:
                r_grand[c_idx] = 0
        except: pass
    cc_rows_array.append(r_grand)
    current_row_num += 1
    
    # === 支払合計 (完了Fが空の項目の合計 - 赤) ===
    pay_total_dict = {"大分類": "支払合計 残額", "科目１": "", "科目２": "", "科目詳細": "支払合計 残額"}
    r_pay_total = _dict_to_row(pay_total_dict)
    # A列にも入れる
    r_pay_total[0] = "支払合計 残額"
    
    for mc in month_cols:
        try:
            c_idx = pay_headers.index(mc)
            f_idx = c_idx + 1 # 完了Fのインデックス (通常は月の右隣)
            col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
            flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
            
            formula_parts = []
            
            # 1. 口座引落・銀行振込
            if other_pay_rows:
                cells = [f"IF({flag_letter}{r}=\"\", {col_letter}{r}, 0)" for r in other_pay_rows]
                formula_parts.append(f"SUM({','.join(cells)})")
                
            # 2. クレジットカード合計行
            if card_total_rows:
                cells = [f"IF({flag_letter}{r}=\"\", {col_letter}{r}, 0)" for r in card_total_rows]
                formula_parts.append(f"SUM({','.join(cells)})")
            
            if formula_parts:
                r_pay_total[c_idx] = "=" + "+".join(formula_parts)
            else:
                r_pay_total[c_idx] = 0
        except: pass
    cc_rows_array.append(r_pay_total)
    current_row_num += 1
    
    summary_end_row = current_row_num - 1
    
    header_len = len(pay_headers)
    
    # Append the variable cost area array
    try:
        # 古い変動費（合計より下）の値をクリア（書式は維持）
        safe_gspread_call(ws_pay.batch_clear, [f"A{start_row_num}:ZZ1000"])
        # 新しい変動費データを書き込み
        safe_gspread_call(ws_pay.update, f"A{start_row_num}", cc_rows_array, value_input_option='USER_ENTERED')
        
        # 既存の固定費サブ合計行も新方式の数式（完了F対応）に更新する
        if fixed_subtotals:
            update_data = [] # List of {'range': ..., 'values': [[...]]}
            for row_num, s_row, e_row in fixed_subtotals:
                row_vals = [""] * len(pay_headers)
                for mc in month_cols:
                    if mc in pay_headers:
                        c_idx = pay_headers.index(mc)
                        f_idx = c_idx + 1 # 完了Fのインデックス (通常は月の右隣)
                        col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                        flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                        # SUMIFS(金額範囲, 完了F範囲, "")
                        formula = f"=SUMIFS({col_letter}{s_row}:{col_letter}{e_row}, {flag_letter}{s_row}:{flag_letter}{e_row}, \"\")"
                        row_vals[c_idx] = formula
                
                # 月カラム以外のデータは既存のものを維持
                original_row = pay_raw[row_num - 1]
                for i in range(len(pay_headers)):
                    h_clean = _clean_val(pay_headers[i]).strip()
                    if h_clean not in month_cols:
                        row_vals[i] = original_row[i]
                
                # 【合計】を固定費合計に書き換え (どの列にあっても対応)
                for j in range(7):
                    if j < len(row_vals) and "【合計】" in str(row_vals[j]):
                        row_vals[j] = "固定費合計"
                        break

                update_data.append({
                    'range': f"A{row_num}",
                    'values': [row_vals]
                })

            # --- 追加: 固定費エリアも含めた全シートの自動フラグ更新 ---
            scan_update_data = []
            try:
                # 科目１、科目２、大分類のインデックス特定
                idx_k1 = -1
                idx_k2 = -1
                idx_dai = -1
                for i, h in enumerate(pay_headers):
                    h_c = _clean_val(h).strip()
                    if h_c == "科目１": idx_k1 = i
                    if h_c == "科目２": idx_k2 = i
                    if h_c == "大分類": idx_dai = i

                # 8行目から、変動費エリアの手前までを走査
                for r_idx in range(7, start_row_num - 1):
                    if r_idx >= len(pay_raw): break
                    row = pay_raw[r_idx]
                    
                    # クレジットカード払いか判定
                    is_cc = False
                    current_cc = None
                    if idx_k1 != -1 and "クレジットカード" in str(row[idx_k1]): is_cc = True
                    if not is_cc and idx_dai != -1 and "クレジットカード" in str(row[idx_dai]): is_cc = True
                    
                    # 科目２がマスターにあるか確認
                    if idx_k2 != -1:
                        target_cc_name = str(row[idx_k2]).strip()
                        for cc in cc_methods:
                            if cc.get("name") == target_cc_name:
                                is_cc = True
                                current_cc = cc
                                break
                    
                    if is_cc and current_cc:
                        row_to_update = list(row)
                        row_changed = False
                        p_month = str(current_cc.get("payment_month", ""))
                        p_date = str(current_cc.get("payment_date", ""))
                        
                        for mc in month_cols:
                            try:
                                c_idx = pay_headers.index(mc)
                                f_idx = c_idx + 1
                                if f_idx >= len(row_to_update): continue
                                
                                # 既にフラグがあればスキップ (1以外を尊重する場合)
                                if str(row_to_update[f_idx]).strip() == "1": continue
                                
                                # 支払日判定
                                y, m = _get_year_month(mc)
                                if y != 9999:
                                    off = 1
                                    if "当月" in p_month: off = 0
                                    elif "翌々月" in p_month: off = 2
                                    base_dt = datetime(y, m, 1) + relativedelta(months=off)
                                    d_m = re.search(r"\d+", p_date)
                                    if d_m:
                                        d_v = int(d_m.group())
                                        due_dt = base_dt + relativedelta(day=d_v)
                                        if datetime.now().date() >= due_dt.date():
                                            row_to_update[f_idx] = "1"
                                            row_changed = True
                            except: pass
                        
                        if row_changed:
                            scan_update_data.append({
                                'range': f"A{r_idx + 1}",
                                'values': [row_to_update]
                            })
            except Exception as e:
                print(f"Error in full sheet scan: {e}")

            if scan_update_data:
                update_data.extend(scan_update_data)
            
            if update_data:
                safe_gspread_call(ws_pay.batch_update, update_data, value_input_option='USER_ENTERED')

        # フォーマット適用 (罫線など)
        format_requests = []
        sheet_id = ws_pay.id
        
        # 1. Base borders for the new rows
        format_requests.append({
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row_num - 1, # 0-based
                    "endRowIndex": current_row_num - 1,
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
        })
        
        # 1.5 完了Fの左側を破線にする
        try:
            for c_idx, h_name in enumerate(pay_headers):
                clean_h = str(h_name).replace("\n", "").replace(" ", "").strip()
                if "完了" in clean_h:
                    format_requests.append({
                        "updateBorders": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": 7, # 8行目から
                                "endRowIndex": 1000, # 常に1000行目まで維持
                                "startColumnIndex": c_idx,
                                "endColumnIndex": c_idx + 1
                            },
                            "left": {"style": "DASHED"}
                        }
                    })
        except: pass

        # 2. Outer THICK borders for the variable cost block
        format_requests.append({
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": var_start - 1, # 0-based
                    "endRowIndex": summary_start_row - 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": header_len
                },
                "top": {"style": "SOLID_MEDIUM"},
                "bottom": {"style": "SOLID_MEDIUM"},
                "left": {"style": "SOLID_MEDIUM"},
                "right": {"style": "SOLID_MEDIUM"}
            }
        })
        # 2.5 Outer THICK borders for the Card Summary block
        format_requests.append({
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": summary_data_start - 1, # 0-based
                    "endRowIndex": summary_end_row,
                    "startColumnIndex": 0,
                    "endColumnIndex": header_len
                },
                "top": {"style": "SOLID_MEDIUM"},
                "bottom": {"style": "SOLID_MEDIUM"},
                "left": {"style": "SOLID_MEDIUM"},
                "right": {"style": "SOLID_MEDIUM"}
            }
        })
        
        # 2.6 Merge cell for クレジットカード合計 (A-G相当)
        try:
            # 科目明細(G列)までの範囲を特定
            k_detail_idx = -1
            for h_i, h_val in enumerate(pay_headers):
                h_clean = _clean_val(h_val).strip()
                if "詳細" in h_clean or "明細" in h_clean: k_detail_idx = h_i
            merge_end = k_detail_idx + 1 if k_detail_idx != -1 else 7

            # 既存の結合を解除（エラー防止）- サマリーエリア全体
            format_requests.append({
                "unmergeCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_start_row - 2,
                        "endRowIndex": summary_end_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": header_len
                    }
                }
            })

            # A) 変動費エリア個別の合計行 (Pink)
            format_requests.append({
                "mergeCells": {
                    "mergeType": "MERGE_ALL",
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_start_row - 2, # summary_start_row is right after var total row, so -2 gives the var total row index (0-based)
                        "endRowIndex": summary_start_row - 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": merge_end
                    }
                }
            })
            # セル配置（中央揃え）
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_start_row - 2,
                        "endRowIndex": summary_start_row - 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": merge_end
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE"
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
                }
            })

            # B) クレジットカードサマリ左側結合セル (Cyan)
            if summary_end_row - 2 >= summary_data_start - 1:
                format_requests.append({
                    "mergeCells": {
                        "mergeType": "MERGE_ALL",
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": summary_data_start - 1,
                            "endRowIndex": summary_end_row - 2,
                            "startColumnIndex": 0,
                            "endColumnIndex": 4 # A-D (大分類〜科目2)
                        }
                    }
                })
                # セル配置
                format_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": summary_data_start - 1,
                            "endRowIndex": summary_end_row - 2,
                            "startColumnIndex": 0,
                            "endColumnIndex": 4
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 0.0, "green": 1.0, "blue": 1.0},
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE",
                                "textFormat": { "bold": True }
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat.bold)"
                    }
                })

            # C) クレジットカード 総合計行 (Blue) A-G合併
            format_requests.append({
                "mergeCells": {
                    "mergeType": "MERGE_ALL",
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_end_row - 2,
                        "endRowIndex": summary_end_row - 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": merge_end
                    }
                }
            })
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_end_row - 2,
                        "endRowIndex": summary_end_row - 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": merge_end
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE"
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
                }
            })

            # D) 支払合計 残額行 (Red) A-G合併
            format_requests.append({
                "mergeCells": {
                    "mergeType": "MERGE_ALL",
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_end_row - 1,
                        "endRowIndex": summary_end_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": merge_end
                    }
                }
            })
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_end_row - 1,
                        "endRowIndex": summary_end_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": merge_end
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE"
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
                }
            })
        except: pass
        
        # 3. Add thick borders for each year (12 month columns) matching the fixed cost area
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
                                "startRowIndex": 7, # 8行目から
                                "endRowIndex": 1000, # 常に1000行目まで書式を維持
                                "startColumnIndex": sc,
                                "endColumnIndex": ec + 1
                            },
                            "left": {"style": "SOLID_MEDIUM"},
                            "right": {"style": "SOLID_MEDIUM"}
                        }
                    })
        except: pass
        
        # 4. Format numbers with commas (#,##0)
        try:
            first_month = month_cols[0]
            if first_month in pay_headers:
                start_m_col = pay_headers.index(first_month)
                format_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 7, # 8行目から全範囲
                            "endRowIndex": 1000, # 常に1000行目まで維持
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
        except: pass
        
        # 5. Background color and bold formatting for 変動費エリアの合計行 (Pink with White text - Reverse)
        format_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": summary_start_row - 2, # summary_start_row is right after var total row, so -2 gives the var total row index (0-based)
                    "endRowIndex": summary_start_row - 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": header_len
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 0.0,
                            "blue": 1.0
                        },
                        "textFormat": {
                            "bold": True,
                            "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat.bold,textFormat.foregroundColor)"
            }
        })
        
        # 5.5 Background color and bold formatting for card-specific 合計 rows
        # 添付の配色をできるだけ再現 (Cyan, Green, etc)
        CARD_PALETTE = [
            {"red": 0.0, "green": 1.0, "blue": 1.0}, # Cyan
            {"red": 0.0, "green": 1.0, "blue": 0.0}, # Green
            {"red": 1.0, "green": 1.0, "blue": 0.0}, # Yellow
        ]
        try:
            k2_idx = pay_headers.index("科目２")
            for i, r_num in enumerate(card_total_rows):
                color = CARD_PALETTE[i % len(CARD_PALETTE)]
                format_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": r_num - 1,
                            "endRowIndex": r_num,
                            "startColumnIndex": k2_idx,
                            "endColumnIndex": header_len
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": color,
                                "textFormat": { "bold": True }
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat.bold)"
                    }
                })
        except: pass
            
        # 6. Background color for 総合計 (Blue with White text - Reverse)
        try:
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_end_row - 2, # 総合計の行
                        "endRowIndex": summary_end_row - 1,
                        "startColumnIndex": 0, # A列から
                        "endColumnIndex": header_len
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.0, "green": 0.0, "blue": 1.0},
                            "textFormat": {
                                "bold": True,
                                "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat.bold,textFormat.foregroundColor)"
                }
            })
            
            # 7. Background color for 支払合計残高 (Red with White text - Reverse)
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_end_row - 1, # 最終行
                        "endRowIndex": summary_end_row,
                        "startColumnIndex": 0, # A列から
                        "endColumnIndex": header_len
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0},
                            "textFormat": {
                                "bold": True,
                                "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat.bold,textFormat.foregroundColor)"
                }
            })
        except: pass
        
        if format_requests:
            safe_gspread_call(ss.batch_update, {"requests": format_requests})
            
        return True, "変動費データの更新が完了しました！"
    except Exception as e:
        return False, f"書き込みエラー: {e}"

def show_variable_cost_update():
    """変動費データ更新 UI"""
    import streamlit as st
    st.markdown("## 🔄 変動費データ更新")
    st.info("「支払方法マスター」に登録されたクレジットカードと、各レシート取引の履歴をもとに、変動費エリアを自動集計して「支払管理」シートを更新します。")
    
    username = st.session_state.get("username", "")
    from app import get_gspread_client
    client = get_gspread_client()
    
    if not client:
        st.error("Google Drive APIに接続できません。")
        return
        
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
    except Exception as e:
        st.warning("現在、あなた専用の支払管理シートが見つかりません。先に「支払管理シート新規作成」を行ってください。")
        return
        
    if st.button("更新を実行する", type="primary"):
        with st.spinner("クレジットカードの利用履歴を集計中...（完了まで数秒〜数十秒かかります）"):
            success, msg = execute_variable_cost_update(username)
            if success:
                st.success(msg)
                st.markdown(f"**🔗 [支払管理シートを開く]({ss.url})**")
            else:
                st.error(msg)
