import streamlit as st
import pandas as pd
from datetime import datetime, timezone, timedelta
from dateutil.relativedelta import relativedelta
import gspread
import jpholiday
import re
import unicodedata

JST = timezone(timedelta(hours=+9), 'JST')

def get_next_business_day(d):
    """土日・祝日の場合に翌営業日を返す"""
    while d.weekday() >= 5 or jpholiday.is_holiday(d):
        d += timedelta(days=1)
    return d

def safe_money_int_cast(val):
    if val is None: return 0
    if isinstance(val, (int, float)): return int(val)
    try:
        s = str(val).replace(",", "").replace("円", "").replace("￥", "").replace("¥", "").strip()
        if not s: return 0
        if "." in s: return int(float(s))
        return int(s)
    except Exception: return 0

def _get_year_month(ym_str):
    s = str(ym_str).strip()
    if not s: return (9999, 12)
    m = re.search(r"(\d{4})[年/\.\-](\d{1,2})", s)
    if m: return (int(m.group(1)), int(m.group(2)))
    m2 = re.search(r"(\d{2})[年/\.\-](\d{1,2})", s)
    if m2:
        y_val = int(m2.group(1))
        y = 2000 + y_val if y_val < 50 else 1900 + y_val
        return (y, int(m2.group(2)))
    return (9999, 12)

def _generate_target_months():
    return [f"{y}.{m}月" for y in range(2026, 2037) for m in range(1, 13)]

def _clean_val(v):
    s = str(v).strip()
    if s.startswith("="):
        s = s[1:].strip()
        if s.startswith('"') and s.endswith('"'): s = s[1:-1].strip()
    return s

def _normalize(s):
    if s is None: return ""
    s = str(s).strip()
    try:
        f_val = float(s)
        if f_val == int(f_val): s = str(int(f_val))
    except: pass
    s_norm = unicodedata.normalize('NFKC', s)
    return "".join(s_norm.split())

def _find_val(d, keywords, exclude=[]):
    for k, v in d.items():
        clean_k = str(k).replace("\n", "").replace(" ", "").replace("　", "").strip()
        for kw in keywords:
            if kw in clean_k:
                if any(ex in clean_k for ex in exclude): continue
                return v
    return ""

def execute_expansion(username, mode="NEW", start_ym=None):
    from app import get_gspread_client, safe_gspread_call
    client = get_gspread_client()
    if not client: return False, "Google Docsへの接続に失敗しました。"
    try:
        ss = client.open(f"{username}_支払管理")
        ws_master = ss.worksheet("固定費マスター")
        ws_pay = ss.worksheet("支払管理")
    except: return False, "必要なシートが見つかりません。"

    master_data = safe_gspread_call(ws_master.get_all_records)
    pay_raw = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMULA')
    pay_formatted = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMATTED_VALUE')

    h_row_idx = 6
    for i_r, r_v in enumerate(pay_raw):
        if r_v and str(r_v[0]).strip().lower() in ["id", "key"]:
            h_row_idx = i_r
            break
    pay_headers = pay_raw[h_row_idx]
    actual_headers = pay_formatted[h_row_idx]
    header_len = max(len(pay_headers), len(actual_headers))
    actual_h_ids = [_normalize(_clean_val(actual_headers[i]) if i < len(actual_headers) else "") for i in range(header_len)]
    month_cols = _generate_target_months()

    def dict_to_row(d):
        r = []
        for i in range(header_len):
            ach = actual_h_ids[i]
            if "科目1" in ach: val = d.get("科目1", "")
            elif "科目2" in ach: val = d.get("科目2", "")
            elif any(k in ach for k in ["変動", "固定"]) and "支払" not in ach: val = d.get("変動or固定", "")
            elif any(k in ach for k in ["有限", "無限"]): val = d.get("有限or無限", "")
            elif "sno" in ach.lower(): val = d.get("Sno", "")
            elif any(k in ach for k in ["詳細", "明細"]): val = d.get("科目詳細", "")
            elif "大分類" in ach: val = d.get("大分類", "")
            elif ach in month_cols: val = d.get(ach, "")
            else: val = d.get(ach, "")
            r.append(val)
        return r

    category_groups = {"クレジットカード": [], "口座引落": [], "銀行振込": []}
    budget_row_dict = None
    for m_rec in master_data:
        k1 = _normalize(_clean_val(_find_val(m_rec, ["科目1", "科目１"])))
        k2 = _normalize(_clean_val(_find_val(m_rec, ["科目2", "科目２"])))
        detail = _normalize(_clean_val(_find_val(m_rec, ["詳細", "明細"])))
        if not k1: continue
        if "小遣い" in (k1 + k2 + detail) and k1 != "小遣い予算": continue
        amt = safe_money_int_cast(_find_val(m_rec, ["支払額", "金額"]))
        sy, sm = _get_year_month(_find_val(m_rec, ["開始"]))
        is_finite = "有限" in _normalize(_find_val(m_rec, ["有限", "無限"]))
        ey, em = _get_year_month(_find_val(m_rec, ["完済", "終了", "完了"])) if is_finite else (9999, 12)
        row_dict = {"大分類": "固定費", "変動or固定": _normalize(_find_val(m_rec, ["変動", "固定"])), "有限or無限": "有限" if is_finite else "無限", "科目1": k1, "科目2": k2, "科目詳細": detail}
        for mc in month_cols:
            my, mm = [int(x.replace("月","")) for x in mc.split(".")]
            row_dict[mc] = amt if (my > sy or (my == sy and mm >= sm)) and (my < ey or (my == ey and mm <= em)) else ""
        if k1 == "小遣い予算": budget_row_dict = row_dict
        else:
            for gk in category_groups:
                if gk in k1: category_groups[gk].append(row_dict); break

    row7_pay = safe_gspread_call(ws_pay.row_values, 7)
    pay_month_indices = [(str(val).split("月")[0], idx + 1) for idx, val in enumerate(row7_pay) if idx >= 8 and "." in str(val)]
    today = datetime.now(JST).date()

    configs = [{"cat": "クレジットカード", "start": 8, "limit": 18}, {"cat": "口座引落", "start": 27, "limit": 10}, {"cat": "銀行振込", "start": 38, "limit": 10}]

    for cfg in configs:
        rows = category_groups[cfg["cat"]][:cfg["limit"]]
        update_vals = []
        for i, r_d in enumerate(rows):
            r_d["Sno"] = str(i + 1); full_row = dict_to_row(r_d)[1:]
            # 完了フラグ判定ロジック等を削除しました
            update_vals.append(full_row)
        while len(update_vals) < cfg["limit"]:
            update_vals.append(dict_to_row({"大分類": "固定費", "科目1": cfg["cat"], "Sno": str(len(update_vals)+1)})[1:])
        safe_gspread_call(ws_pay.update, values=update_vals, range_name=f"B{cfg['start']}", value_input_option='USER_ENTERED')

    if budget_row_dict:
        budget_row_dict["科目詳細"] = "小遣い予算"
        safe_gspread_call(ws_pay.update, values=[dict_to_row(budget_row_dict)[1:]], range_name="B78", value_input_option='USER_ENTERED')

    safe_gspread_call(ws_pay.update, values=[[datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")]], range_name="F4", value_input_option='USER_ENTERED')
    execute_variable_cost_update(username, skip_backup=True)
    return True, "固定費データ展開完了"

def execute_variable_cost_update(username, start_ym=None, skip_backup=False):
    from app import get_gspread_client, safe_gspread_call, get_payment_methods, get_sheet, TRANSACT完了ONS_WORKSHEET_NAME
    client = get_gspread_client()
    if not client: return False, "接続失敗"
    ss = client.open(f"{username}_支払管理"); ws_pay = ss.worksheet("支払管理")
    pay_raw = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMULA')
    pay_formatted = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMATTED_VALUE')
    h_row_idx = next((i for i, r in enumerate(pay_raw) if r and str(r[0]).strip().lower() in ["id", "key"]), 6)
    actual_h_ids = [_normalize(_clean_val(pay_formatted[h_row_idx][i]) if i < len(pay_formatted[h_row_idx]) else "") for i in range(max(len(pay_raw[h_row_idx]), len(pay_formatted[h_row_idx])))]
    month_cols = _generate_target_months()

    def _dict_to_row(d, exist=None):
        r = []
        for i in range(len(actual_h_ids)):
            ach = actual_h_ids[i]
            if "科目1" in ach: r.append(exist[i] if exist and i < len(exist) else d.get("科目1", ""))
            elif "科目2" in ach: r.append(d.get("科目2", ""))
            elif "科目明細" in ach: r.append(d.get("科目明細", ""))
            elif "sno" in ach.lower(): r.append(d.get("Sno", ""))
            elif "大分類" in ach: r.append(d.get("大分類", ""))
            elif ach in month_cols: r.append(d.get(ach, ""))
            else: r.append(d.get(ach, ""))
        return r

    try:
        from app import calculate_credit_card_periods, PAYMENT_MASTER_WORKSHEET_NAME
        methods = safe_gspread_call(get_payment_methods, username)
        cc_methods = [m for m in methods if m.get("is_credit_card") or m.get("type") == "クレジットカード"]
        user_txs = [tx for tx in safe_gspread_call(get_sheet(TRANSACT完了ONS_WORKSHEET_NAME).get_all_records) if str(tx.get("username")).lower() == username.lower()]
        row7_pay = safe_gspread_call(ws_pay.row_values, 7)
        pay_month_indices = [(str(val).split("月")[0], idx + 1) for idx, val in enumerate(row7_pay) if idx >= 8 and "." in str(val)]
        today = datetime.now(JST).date()

        rows_50_52 = []
        for i in range(3):
            exist = pay_raw[49+i] if 49+i < len(pay_raw) else None
            if i < len(cc_methods):
                m = cc_methods[i]; p_d = str(m.get("payment_date", "27")); p_day = int(re.search(r"\d+", p_d).group()) if re.search(r"\d+", p_d) else 27
                full = _dict_to_row({"大分類": "変動費", "科目2": m.get("name"), "Sno": str(i+1), "科目明細": f"{m.get('payment_month','')}{p_d}"}, exist)[1:]
                for ym, col in pay_month_indices:
                    try:
                        my, mm = [int(x) for x in ym.split(".")]; target = datetime(my, mm, 1) + relativedelta(months=1) if p_day <= 20 else datetime(my, mm, 1)
                        pers = calculate_credit_card_periods(target, str(m.get("closing_date")), str(m.get("payment_month")), str(m.get("payment_date")))
                        amt = sum(safe_money_int_cast(tx.get("amount",0)) for tx in user_txs if _normalize(tx.get("payment_method")) == _normalize(m.get("name")) and tx.get("category") != "消費税（内税）" and pers and pers[0]["start"] <= datetime.strptime(str(tx.get("date")), "%Y-%m-%d").date() <= pers[0]["end"])
                        if col-2 < len(full): full[col-2] = amt if amt > 0 else ""
                    except: continue
                rows_50_52.append(full)
            else: rows_50_52.append(_dict_to_row({"大分類": "変動費"})[1:])
        safe_gspread_call(ws_pay.update, values=rows_50_52, range_name="B50", value_input_option='USER_ENTERED')

        f54_f62 = []
        for i in range(3):
            n = cc_methods[i].get("name","") if i < len(cc_methods) else ""
            f54_f62.extend([[n],[n],[n]])
        safe_gspread_call(ws_pay.update, values=f54_f62, range_name="F54:F62", value_input_option='USER_ENTERED')

        all_m = [m for m in safe_gspread_call(get_sheet(PAYMENT_MASTER_WORKSHEET_NAME).get_all_records) if str(m.get("username")).lower()==username.lower()]
        rows_79_88 = []
        for i in range(10):
            if i < len(all_m):
                m = all_m[i]; m_name = m.get("name"); is_cc = "クレジットカード" in str(m.get("type")) or m.get("is_credit_card")
                p_day = int(re.search(r"\d+", str(m.get("payment_date","27"))).group()) if re.search(r"\d+", str(m.get("payment_date","27"))) else 27
                full = _dict_to_row({"大分類": "支払集計", "科目明細": m_name})[1:]
                for ym, col in pay_month_indices:
                    try:
                        my, mm = [int(x) for x in ym.split(".")]; amt = 0
                        if is_cc:
                            target = datetime(my, mm, 1) + relativedelta(months=1) if p_day <= 20 else datetime(my, mm, 1)
                            pers = calculate_credit_card_periods(target, str(m.get("closing_date")), str(m.get("payment_month")), str(m.get("payment_date")))
                            amt = sum(safe_money_int_cast(tx.get("amount",0)) for tx in user_txs if _normalize(tx.get("payment_method"))==_normalize(m_name) and tx.get("category")!="消費税（内税）")
                        safe_gspread_call(ws_pay.update, values=rows_79_88, range_name="B79", value_input_option="USER_ENTERED")
    
                if target_row > 0:
                    print(f"Expanding Pocket Money: {k1}/{detail} to Row {target_row}, Amt: {amt}")
                    for ym_str, col_idx in pay_month_indices:
                        my, mm = [int(x) for x in ym_str.split(".")]
                        # --- 物理保護ロジック ---
                        # 78-90行かつF列(6)・G列(7)の書き換えを厳格に禁止（A-H列全体を保護対象へ拡大）
                        if 78 <= target_row <= 90 and col_idx <= 8:
                            # 念のためログを出力
                            print(f"Skipping protection range cell: Row {target_row}, Col {col_idx}")
                            continue

                        # 開始月 <= YM <= 終了月の判定 (固定費展開と同様の仕様)
                        if (my > sy or (my == sy and mm >= sm)) and (my < ey or (my == ey and mm <= em)):
                            upd_pocket_list.append({'range': rowcol_to_a1(target_row, col_idx), 'values': [[amt]]})
                        else:
                            # 期間外はクリア
                            upd_pocket_list.append({'range': rowcol_to_a1(target_row, col_idx), 'values': [[""]]})
            
            if upd_pocket_list:
                safe_gspread_call(ss.values_batch_update, {'value完了nputOption': 'USER_ENTERED', 'data': upd_pocket_list})
        except Exception as e:
            print(f"Pocket money expansion error: {e}")

        safe_gspread_call(ws_pay.update, values=[[datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")]], range_name="F5", value_input_option='USER_ENTERED')
        return True, "更新完了"
    except Exception as e:
        return False, str(e)

def show_open_management_sheet():
    username = st.session_state.get("username", "")
    from app import get_gspread_client, safe_gspread_call
    client = get_gspread_client()
    if not client: return
    try:
        ss = client.open(f"{username}_支払管理")
        try:
            ws = safe_gspread_call(ss.worksheet, "支払管理")
            url = f"{ss.url}#gid={ws.id}"
        except:
            url = ss.url
        st.link_button("🌐 支払管理シートを開く", url=url, type="primary", use_container_width=True)
    except: st.warning("シートが見つかりません。")

def show_fixed_cost_data_expansion():
    st.title("🛠️ 固定費データ展開")
    username = st.session_state.get("username", "")
    if st.button("実行", type="primary"):
        with st.spinner("処理中..."):
            success, msg = execute_expansion(username)
            if success: st.success(msg)
            else: st.error(msg)

def show_variable_cost_update():
    st.title("💳 変動費データ更新")
    username = st.session_state.get("username", "")
    if st.button("更新実行"):
        with st.spinner("処理中..."):
            success, msg = execute_variable_cost_update(username)
            if success: st.success(msg)
            else: st.error(msg)