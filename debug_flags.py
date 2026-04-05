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
    # Simplified normalization for debugging
    s_norm = unicodedata.normalize("NFKC", s)
    return "".join(s_norm.split())

def get_next_business_day(d):
    while d.weekday() >= 5 or jpholiday.is_holiday(d):
        d += timedelta(days=1)
    return d

def debug_logic():
    client = gspread.service_account(filename="credentials.json")
    ss = client.open("tkouho_支払管理")
    ws_pay = ss.worksheet("支払管理")
    ws_fm = ss.worksheet("固定費マスター")
    
    f_master_data = ws_fm.get_all_records()
    row7_pay = ws_pay.row_values(7)
    
    pay_month_indices = [(str(val).split("月")[0], idx + 1) for idx, val in enumerate(row7_pay) if idx >= 8 and "." in str(val)]
    
    fixed_rows_raw = ws_pay.get("B8:AJ25")
    today = datetime.now(JST).date()
    
    print(f"DEBUG: Today is {today}")
    print(f"DEBUG: pay_month_indices: {pay_month_indices[:4]}...")
    
    for i, row in enumerate(fixed_rows_raw):
        if i > 5: break # Only debug first few rows
        cat1 = _normalize(row[3]) if len(row) > 3 else ""
        cat2 = _normalize(row[4]) if len(row) > 4 else ""
        detail = _normalize(row[6]) if len(row) > 6 else ""
        
        print(f"--- Row {i+8}: {cat1} / {cat2} / {detail} ---")
        
        if cat1 in ["クレジットカード", "口座引落", "銀行振込", "銀行引落"] and cat2:
            m_rec = next((m for m in f_master_data if _normalize(m.get("科目2")) == cat2 and _normalize(m.get("科目明細")) == detail), {})
            print(f"  Master Match found: {bool(m_rec)}")
            
            p_d_str = str(m_rec.get("支払日など", "")).strip()
            p_day = int(re.search(r"\d+", p_d_str).group()) if re.search(r"\d+", p_d_str) else 27

            for ym, col in pay_month_indices[:3]: # Check first 3 months
                try:
                    my_t, mm_t = [int(x) for x in ym.split(".")]
                    m_base = datetime(my_t, mm_t, 1)
                    if cat1 == "クレジットカード":
                        p_dt = m_base + relativedelta(months=1)
                    else:
                        p_dt = m_base
                    
                    p_date = get_next_business_day((p_dt + relativedelta(day=p_day)).date())
                    result = today >= p_date
                    flag_idx = col + 1 - 2
                    print(f"  Month {ym}: p_date={p_date}, result={result}, flag_idx={flag_idx}")
                except Exception as e:
                    print(f"  Month {ym}: ERROR: {e}")

if __name__ == "__main__":
    debug_logic()
