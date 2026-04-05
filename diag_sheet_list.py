import gspread
import os

def get_client():
    if os.path.exists("credentials.json"):
        return gspread.service_account(filename="credentials.json")
    return None

def diag():
    client = get_client()
    if not client:
        return

    username = "tkouho"
    sheet_name = f"{username}_支払管理"
    
    try:
        ss = client.open(sheet_name)
        print(f"Spreadsheet '{sheet_name}' opened successfully.")
        
        worksheets = ss.worksheets()
        print("Worksheets found:")
        for ws in worksheets:
            print(f" - '{ws.title}' (ID: {ws.id})")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    diag()
