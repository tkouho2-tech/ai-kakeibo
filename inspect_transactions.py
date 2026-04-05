import gspread
from app import get_gspread_client
import sys

def inspect_transactions_around_570():
    client = get_gspread_client()
    ss = client.open("Kakeibo_Data")
    ws = ss.worksheet("transactions")
    data = ws.get_all_values()
    headers = data[0]
    
    # Range around row 570
    for i in range(540, min(600, len(data))):
        row = data[i]
        if "tkouho" in row[0]:
            try:
                d = dict(zip(headers, row))
                # Only print interesting columns
                subset = {
                    "row": i+1,
                    "date": d.get("date"), 
                    "store": d.get("store_name"), 
                    "item": d.get("item_name"), 
                    "cat": d.get("category"), 
                    "sub": d.get("subcategory"),
                    "amt": d.get("amount"), 
                    "pay_m": d.get("payment_month"),
                    "pay_d": d.get("payment_date")
                }
                # Check for Amazon
                row_str = " ".join([str(x) for x in row]).lower()
                if "amazon" in row_str and "25" in row_str:
                    print(f"MATCH: {subset}")
            except Exception as e:
                pass

if __name__ == "__main__":
    inspect_transactions_around_570()
