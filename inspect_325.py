import gspread
from app import get_gspread_client
import sys

# Standardize encoding
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def inspect_tkouho_325():
    client = get_gspread_client()
    ss = client.open("Kakeibo_Data")
    ws = ss.worksheet("transactions")
    data = ws.get_all_values()
    headers = data[0]
    
    print(f"Inspecting rows around 570...")
    for i in range(560, min(580, len(data))):
        row = data[i]
        if "tkouho" in row[0]: # user
            d = dict(zip(headers, row))
            # Just print the whole row nicely
            print(f"Row {i+1}: {d}")

if __name__ == "__main__":
    inspect_tkouho_325()
