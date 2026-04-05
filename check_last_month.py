import os
import sys

# Ensure current directory is in search path
sys.path.append(os.getcwd())

import streamlit as st
import gspread
from app import get_gspread_client

def check_last_month(username):
    client = get_gspread_client()
    if not client:
        print("Failed to get gspread client.")
        return
    
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
        ws = ss.worksheet("支払管理")
        values = ws.get_all_values()
        h_row = -1
        for i, row in enumerate(values):
            if row and str(row[0]).strip().lower() in ["id", "key"]:
                h_row = i
                break
        
        if h_row == -1:
            # Fallback to row 7
            headers = values[6] if len(values) >= 7 else []
        else:
            headers = values[h_row]
            
        import re
        month_pattern = re.compile(r"(\d{4})\.(\d{1,2})月")
        months = []
        for h in headers:
            m = month_pattern.search(str(h))
            if m:
                months.append(f"{m.group(1)}.{m.group(2)}月")
        
        if not months:
            print("No months found in headers.")
        else:
            # Sort naturally
            def ym_key(s):
                y, m = [int(x.replace("月","")) for x in s.split(".")]
                return y * 12 + m
            months.sort(key=ym_key)
            print(f"LAST_MONTH: {months[-1]}")
            
    except Exception as e:
        print(f"Error checking {sheet_name}: {e}")

if __name__ == "__main__":
    check_last_month("tkouho")
