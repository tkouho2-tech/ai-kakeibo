import gspread
import os

SPREADSHEET_NAME = "Kakeibo_Data"
TRANSACTIONS_WORKSHEET_NAME = "transactions"

def fix_headers():
    if os.path.exists("credentials.json"):
        gc = gspread.service_account(filename="credentials.json")
    else:
        print("credentials.json not found")
        return

    try:
        sh = gc.open(SPREADSHEET_NAME)
        ws = sh.worksheet(TRANSACTIONS_WORKSHEET_NAME)
        all_values = ws.get_all_values()
        if not all_values:
            print("Sheet is empty")
            return
        
        headers = all_values[0]
        print(f"Current Headers: {headers}")
        
        # Identify empty columns at the end
        last_real_header_idx = -1
        for i, h in enumerate(headers):
            if h.strip() != "":
                last_real_header_idx = i
        
        if last_real_header_idx < len(headers) - 1:
            cols_to_delete = len(headers) - (last_real_header_idx + 1)
            print(f"Detected {cols_to_delete} empty trailing columns. Deleting...")
            # Delete columns from the end
            # gspread's delete_columns uses 1-based indexing
            start_col = last_real_header_idx + 2
            end_col = len(headers)
            ws.delete_columns(start_col, end_col)
            print("Cleanup complete.")
        else:
            print("No trailing empty columns found in the header row.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    fix_headers()
