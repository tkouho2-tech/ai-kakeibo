import gspread
from app import get_gspread_client
import sys

if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def check_pm_utf8():
    client = get_gspread_client()
    ss = client.open("Kakeibo_Data")
    ws = ss.worksheet("Payment_Master")
    data = ws.get_all_records()
    for row in data:
        if row.get("username") == "tkouho":
            print(f"PM: {row}")

if __name__ == "__main__":
    check_pm_utf8()
