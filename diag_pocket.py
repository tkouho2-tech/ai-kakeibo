import gspread
import unicodedata
from app import get_gspread_client, FIXED_COST_MASTER_WORKSHEET_NAME

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
    ws_master = ss.worksheet(FIXED_COST_MASTER_WORKSHEET_NAME)
    master_data = ws_master.get_all_records()
    
    print(f"Total Master Records: {len(master_data)}")
    for m in master_data:
        k1 = _normalize(_find_val(m, ["科目1", "科目１"]))
        detail = _normalize(_find_val(m, ["詳細", "明細"]))
        if "小遣い" in k1 or "小遣い" in detail:
            print(f"Record: k1='{k1}', detail='{detail}', Amt={m.get('支払額')}, Start={m.get('開始')}")

if __name__ == "__main__":
    diag()
