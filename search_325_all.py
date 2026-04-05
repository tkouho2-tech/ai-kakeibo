import gspread
from app import get_gspread_client
import sys

if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def search_by_amount():
    client = get_gspread_client()
    ss = client.open("Kakeibo_Data")
    ws = ss.worksheet("transactions")
    data = ws.get_all_values()
    headers = data[0]
    
    # User to look for: tkouho
    # Date: 2026-03-25
    # Amount: around 1380
    for i, row in enumerate(data):
        if i == 0: continue
        if row[0] == "tkouho" and "2026-03-25" in row[1]:
            d = dict(zip(headers, row))
            print(f"Row {i+1}: {d}")

if __name__ == "__main__":
    search_by_amount()
