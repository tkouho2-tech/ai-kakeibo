import gspread
import os
import json
from google.oauth2.service_account import Credentials

def get_client():
    if os.path.exists("credentials.json"):
        return gspread.service_account(filename="credentials.json")
    return None

def diag():
    client = get_client()
    if not client:
        print("Error: credentials.json not found")
        return

    username = "tkouho"
    sheet_name = f"{username}_支払管理"
    
    print(f"Checking for spreadsheet: {sheet_name}")
    try:
        ss = client.open(sheet_name)
        print(f"Success: Opened spreadsheet '{sheet_name}' (ID: {ss.id})")
        
        try:
            ws = ss.worksheet("支払管理")
            print("Success: Found '支払管理' worksheet")
        except Exception as e:
            print(f"Error: '支払管理' worksheet NOT found: {e}")
            
        try:
            ws_master = ss.worksheet("固定費マスター")
            print("Success: Found '固定費マスター' worksheet")
        except Exception as e:
            print(f"Error: '固定費マスター' worksheet NOT found: {e}")
            
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"Error: Spreadsheet '{sheet_name}' NOT found")
        print("Listing available spreadsheets...")
        all_ss = client.openall()
        for s in all_ss:
            print(f" - {s.title} (ID: {s.id})")
    except Exception as e:
        print(f"Unexpected error: {e}")

if __name__ == "__main__":
    diag()
