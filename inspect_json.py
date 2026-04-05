import gspread
import os
import json

client = None
if os.path.exists("credentials.json"):
    client = gspread.service_account(filename="credentials.json")

if client:
    sheet_name = "tkouho_支払管理"
    try:
        ss = client.open(sheet_name)
        
        ws_pay = ss.worksheet("支払管理")
        pay_data = ws_pay.get_all_values()[:25]
        
        ws_master = ss.worksheet("固定費マスター")
        master_data = ws_master.get_all_values()[:10]
        
        with open("sheet_data.json", "w", encoding="utf-8") as f:
            json.dump({"pay": pay_data, "master": master_data}, f, ensure_ascii=False, indent=2)
            
        print("Data exported to sheet_data.json successfully.")
    except Exception as e:
        print(f"Error: {e}")
