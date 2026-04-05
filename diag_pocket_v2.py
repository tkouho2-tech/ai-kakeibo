import gspread
import unicodedata
from app import get_gspread_client

def _normalize(s):
    if s is None: return ""
    s = str(s).strip()
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

def diag():
    client = get_gspread_client()
    username = 'tkouho'
    ss = client.open(f"{username}_支払管理")
    # try both names
    try:
        ws_master = ss.worksheet("固定費マスター")
    except:
        ws_master = ss.worksheet("Fixed_Cost_Master")
    
    master_data = ws_master.get_all_records()
    print(f"Total Master Records: {len(master_data)}")
    for m in master_data:
        k1_raw = _find_val(m, ["科目1", "科目１"])
        k1 = _normalize(k1_raw)
        detail_raw = _find_val(m, ["詳細", "明細"])
        detail = _normalize(detail_raw)
        if "小遣い" in k1 or "小遣い" in detail:
            print(f"Record: raw_k1='{k1_raw}', k1='{k1}', detail='{detail}', Amt={m.get('支払額')}")

if __name__ == "__main__":
    diag()
