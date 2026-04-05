import gspread
from app import get_gspread_client

def check_payment_master():
    client = get_gspread_client()
    ss = client.open("Kakeibo_Data")
    ws = ss.worksheet("Payment_Master")
    data = ws.get_all_records()
    for row in data:
        if row.get("username") == "tkouho":
            print(f"Payment Method: {row}")

if __name__ == "__main__":
    check_payment_master()
