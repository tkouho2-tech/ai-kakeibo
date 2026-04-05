import gspread
import json
import re
from datetime import datetime, timezone, timedelta
from dateutil.relativedelta import relativedelta
import unicodedata
import jpholiday

JST = timezone(timedelta(hours=+9), 'JST')

def _normalize(s):
    if s is None: return ""
    s = str(s).strip()
    s_norm = unicodedata.normalize('NFKC', s)
    return "".join(s_norm.split())

def _find_val(d, keywords):
    for k, v in d.items():
        clean_k = _normalize(str(k))
        for kw in keywords:
            if kw in clean_k:
                return v
    return ""

def get_next_business_day(d):
    while d.weekday() >= 5 or jpholiday.is_holiday(d):
        d += timedelta(days=1)
    return d

def diag():
    client = gspread.service_account(filename='credentials.json')
    username = "tkouho" # From prompt corpus name
    ss = client.open(f"{username}_支払管理")
    ws_pay = ss.worksheet("支払管理")
    ws_fm = ss.worksheet("固定費マスター")
    
    # Mocking get_payment_methods from app.py
    # From previous context, it seems to come from a '決済手段' worksheet
    try:
        ws_methods = ss.worksheet("決済手段")
        methods = ws_methods.get_all_records()
    except:
        methods = []
        print("WARNING: '決済手段' sheet not found")

    f_master_data = ws_fm.get_all_records()
    row7_pay = ws_pay.row_values(7)
    pay_month_indices = [(str(val).split("月")[0], idx + 1) for idx, val in enumerate(row7_pay) if idx >= 8 and "." in str(val)]
    
    today = datetime.now(JST).date()
    print(f"DEBUG: Today is {today}")
    
    fixed_ranges = [("B8:AJ12", 8)] # Just check first few rows
    for range_str, start_row in fixed_ranges:
        rows = ws_pay.get(range_str)
        for i_r, row in enumerate(rows):
            cur_row = start_row + i_r
            cat1 = _normalize(row[3]) if len(row) > 3 else ""
            cat2 = _normalize(row[4]) if len(row) > 4 else ""
            detail = _normalize(row[6]) if len(row) > 6 else ""
            
            print(f"--- Row {cur_row}: {cat1}/{cat2}/{detail} ---")
            
            if cat1 == "クレジットカード":
                card_info = next((m for m in methods if _normalize(m.get("決済手段名") or m.get("name")) == cat2), {})
                p_day = 27
                if card_info:
                    p_d_s = str(card_info.get("支払日") or card_info.get("payment_date", "27")).strip()
                    m = re.search(r"\d+", p_d_s)
                    if m: p_day = int(m.group())
                    print(f"  Match Card: {cat2}, p_day: {p_day}")
                else:
                    print(f"  Match Card: {cat2} NOT FOUND in methods. Fallback to 27")
                
                for ym, col in pay_month_indices[:1]:
                    my, mm = [int(x) for x in ym.split(".")]
                    p_dt = datetime(my, mm, 1) + relativedelta(months=1, day=p_day)
                    p_date = get_next_business_day(p_dt.date())
                    print(f"  YM {ym}: p_date={p_date}, result={today >= p_date}, flag_cell={gspread.utils.rowcol_to_a1(cur_row, col + 1)}")

            elif cat1 in ["口座引落", "銀行振込", "銀行引落"]:
                m_rec = next((m for m in f_master_data if _normalize(_find_val(m, ["科目2"])) == cat2 and _normalize(_find_val(m, ["詳細", "明細"])) == detail), {})
                p_day = 27
                if m_rec:
                    p_d_s = str(_find_val(m_rec, ["支払日", "引落日"])).strip()
                    m = re.search(r"\d+", p_d_s)
                    if m: p_day = int(m.group())
                    print(f"  Match Master: {detail}, p_day: {p_day}")
                else:
                    print(f"  Match Master: {detail} NOT FOUND. Fallback to 27")

                for ym, col in pay_month_indices[:1]:
                    my, mm = [int(x) for x in ym.split(".")]
                    p_dt = datetime(my, mm, 1) + relativedelta(day=p_day)
                    p_date = get_next_business_day(p_dt.date())
                    print(f"  YM {ym}: p_date={p_date}, result={today >= p_date}, flag_cell={gspread.utils.rowcol_to_a1(cur_row, col + 1)}")

if __name__ == "__main__":
    diag()
