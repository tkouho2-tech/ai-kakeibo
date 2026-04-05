import gspread
from google.oauth2.service_account import Credentials
import os

def check_pm_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("c:/Users/t_kou/Kakeibo_Final_v3/credentials.json", scopes=scopes)
    client = gspread.authorize(creds)
    
    username = "tkouho2" # Assuming this is the username
    try:
        ss = client.open(f"{username}_支払管理")
        ws = ss.worksheet("支払管理")
        
        # Get rows 1-65 to see everything
        data = ws.get("A1:P65")
        for i, row in enumerate(data):
            print(f"Row {i+1}: {row}")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_pm_sheet()
