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
        # 「支払管理 」（スペースあり）で開く（現在の実態に合わせる）
        ws = ss.worksheet("支払管理 ")
        
        headers = ws.row_values(7)
        print("Headers (Row 7):")
        for i, h in enumerate(headers):
            print(f" Column {chr(65+i)}: '{h}'")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    diag()
