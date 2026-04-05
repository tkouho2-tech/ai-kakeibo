import gspread
from app import get_gspread_client
import sys

# Set output encoding to utf-8 if possible
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def inspect_amazon_full():
    client = get_gspread_client()
    ss = client.open("Kakeibo_Data")
    ws = ss.worksheet("transactions")
    data = ws.get_all_values()
    headers = data[0]
    
    print(f"Total Rows: {len(data)}")
    for i, row in enumerate(data):
        if i == 0: continue
        row_str = " ".join([str(x) for x in row]).lower()
        if "amazon" in row_str and "25" in row_str and "3" in row_str:
            d = dict(zip(headers, row))
            print(f"MATCH Row {i+1}: {d}")

if __name__ == "__main__":
    inspect_amazon_full()
