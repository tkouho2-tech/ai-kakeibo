import gspread
import os
import json

client = None
if os.path.exists("credentials.json"):
    client = gspread.service_account(filename="credentials.json")

if client:
    sheet_name = "tkouho_支払管理"  # username is likely tkouho
    try:
        ss = client.open(sheet_name)
        
        # Get "支払管理"
        ws_pay = ss.worksheet("支払管理")
        print("--- 支払管理 (First 25 rows) ---")
        for i, row in enumerate(ws_pay.get_all_values()[:25]):
            print(f"Row {i+1}: {row}")
            
        # Get "固定費マスター"
        ws_master = ss.worksheet("固定費マスター")
        print("\n--- 固定費マスター (First 10 rows) ---")
        for i, row in enumerate(ws_master.get_all_values()[:10]):
            print(f"Row {i+1}: {row}")
            
    except Exception as e:
        print(f"Error: {e}")
else:
    print("Could not load credentials.")
