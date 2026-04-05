import gspread
from app import get_gspread_client
import sys

# Standardize encoding
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def search_by_receipt_id():
    client = get_gspread_client()
    ss = client.open("Kakeibo_Data")
    ws = ss.worksheet("transactions")
    data = ws.get_all_values()
    headers = data[0]
    
    target_rid = "20260325175355"
    print(f"Searching for Receipt ID: {target_rid}")
    for i, row in enumerate(data):
        if i == 0: continue
        if target_rid in row:
            d = dict(zip(headers, row))
            print(f"Row {i+1}: {d}")

if __name__ == "__main__":
    search_by_receipt_id()
