import gspread
import json
from app import get_gspread_client

def get_headers():
    client = get_gspread_client()
    ss = client.open('tkouho_支払管理')
    
    # 決済手段 is index 8 based on previous titles listing
    # 収支計算書, 2026年, 2027年, DB, マスタ, 固定費マスター(5), 支払管理, 収入, 決済手段(8)
    wss = ss.worksheets()
    fm = None
    pm = None
    for ws in wss:
        if "固定費マスター" in ws.title: fm = ws
        if "決済手段" in ws.title: pm = ws
    
    if not fm or not pm:
        print(f"WS Titles: {[ws.title for ws in wss]}")
        return

    res = {
        "Fixed_Cost_Master": fm.row_values(1),
        "Payment_Master": pm.row_values(1)
    }
    
    with open('headers.json', 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    get_headers()
