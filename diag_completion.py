import gspread, json, os, re, unicodedata, calendar
from google.oauth2.service_account import Credentials
from datetime import datetime, date, timedelta, timezone
import jpholiday

JST = timezone(timedelta(hours=+9), 'JST')

def _normalize(s):
    if s is None: return ""
    s = str(s).strip()
    s_norm = unicodedata.normalize('NFKC', s)
    return "".join(s_norm.split())

def _get_year_month(ym_str):
    s = str(ym_str).strip()
    if not s: return (9999, 12)
    m = re.search(r"(\d{4})[年/\.\-](\d{1,2})", s)
    if m: return (int(m.group(1)), int(m.group(2)))
    return (9999, 12)

def get_next_business_day(d):
    while d.weekday() >= 5 or jpholiday.is_holiday(d):
        d += timedelta(days=1)
    return d

info = json.load(open('credentials.json'))
creds = Credentials.from_service_account_info(info, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
client = gspread.authorize(creds)

username = "tkouho"
ss = client.open(f"{username}_支払管理")
ws_pay = ss.worksheet("支払管理")

pay_formatted = ws_pay.get_all_values(value_render_option='FORMATTED_VALUE')
h_row_idx = next((i for i, r in enumerate(pay_formatted) if r and str(r[0]).strip().lower() in ["id", "1", "key"]), 6)
actual_h_ids = [_normalize(pay_formatted[h_row_idx][i]) for i in range(len(pay_formatted[h_row_idx]))]
row7_pay = pay_formatted[6]
pay_month_indices = []
for idx, val in enumerate(row7_pay):
    if idx < 8: continue
    y, m = _get_year_month(val)
    if y != 9999:
        pay_month_indices.append(((y, m), idx + 1))

ss_data = client.open('Kakeibo_Data')
ws_methods = ss_data.worksheet('Payment_Master')
methods = [r for r in ws_methods.get_all_records() if str(r.get('username')).lower() == username.lower()]

ws_fixed = ss_data.worksheet('Fixed_Cost_Master')
user_fixed = [r for r in ws_fixed.get_all_records() if str(r.get("username", "")).lower() == username.lower()]

today = datetime.now(JST).date()
print(f"Today: {today}")

k1_idx = next((i for i, h in enumerate(actual_h_ids) if "科目1" in h), 4)
k2_idx = next((i for i, h in enumerate(actual_h_ids) if "科目2" in h), 5)
k_desc_idx = next((i for i, h in enumerate(actual_h_ids) if "科目明細" in h), 7)
print(f"Indices: k1={k1_idx}, k2={k2_idx}, desc={k_desc_idx}")

for i_row in range(7, 45): # Spreadsheet Rows 8 to 45
    if i_row >= len(pay_formatted): break
    row_vals = pay_formatted[i_row]
    if len(row_vals) <= max(k1_idx, k2_idx): continue
    k1 = _normalize(row_vals[k1_idx])
    k2 = _normalize(row_vals[k2_idx])
    k_desc = _normalize(row_vals[k_desc_idx]) if k_desc_idx < len(row_vals) else ""
    
    if "口座引落" in k1 or "クレジットカード" in k1:
        if i_row + 1 == 27:
            print(f"Row {i_row+1}: k1='{k1}' k2='{k2}' bytes={k2.encode('utf-8').hex()} desc='{k_desc}' bytes={k_desc.encode('utf-8').hex()}")
            for r in user_fixed[:5]:
                r2 = _normalize(r.get("科目2"))
                rd = _normalize(r.get("科目明細"))
                if r2 == k2 or rd == k_desc:
                    print(f"  Check master: k2='{r2}' bytes={r2.encode('utf-8').hex()} desc='{rd}' bytes={rd.encode('utf-8').hex()}")

        p_day = None
        if "クレジットカード" in k1:
            m = next((m for m in methods if _normalize(m.get("name")) == k2), None)
            if m:
                p_d_str = str(m.get("payment_date", "27"))
                p_day_match = re.search(r"\d+", p_d_str)
                p_day = int(p_day_match.group()) if p_day_match else 27
        elif "口座引落" in k1:
            m = next((r for r in user_fixed if _normalize(r.get("科目2")) == k2 and _normalize(r.get("科目明細")) == k_desc), None)
            if m:
                p_d_str = str(m.get("引落日", "27"))
                p_day_match = re.search(r"\d+", p_d_str)
                p_day = int(p_day_match.group()) if p_day_match else (27 if "末日" not in p_d_str else 99)
            else:
                if i_row + 1 < 37:
                    # print(f"Row {i_row+1}: No master match for '{k1}' '{k2}' '{k_desc}'")
                    pass

        if p_day is not None:
            print(f"Row {i_row+1}: matched p_day={p_day}")
            for (my, mm), col_idx in pay_month_indices:
                if my != 2026: continue
                amt_val = row_vals[col_idx - 1] if col_idx - 1 < len(row_vals) else ""
                if not str(amt_val).strip() or str(amt_val).strip() == "0": continue
                try:
                    p_date = date(my, mm, p_day) if p_day != 99 else None
                    if p_date is None: raise ValueError
                except:
                    l_day = calendar.monthrange(my, mm)[1]
                    p_date = date(my, mm, min(p_day, l_day))
                adj_date = get_next_business_day(p_date)
                if today >= adj_date:
                    print(f"  {my}.{mm}: Result=True -> J{i_row+1} range={col_idx+1}")
