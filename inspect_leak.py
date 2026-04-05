import gspread
from app import get_gspread_client
import sys

if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def check_rows_final():
    client = get_gspread_client()
    ss = client.open("Kakeibo_Data")
    ws = ss.worksheet("transactions")
    data = ws.get_all_values()
    headers = data[0]
    
    print("--- Row 390 to 400 ---")
    for i in range(389, min(400, len(data))):
        row = data[i]
        print(f"Row {i+1}: {row}")
        
    print("\n--- Row 840 to 860 ---")
    for i in range(839, min(860, len(data))):
        row = data[i]
        print(f"Row {i+1}: {row}")

if __name__ == "__main__":
    check_rows_final()
