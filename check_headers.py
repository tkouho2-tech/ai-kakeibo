import gspread
import os

SPREADSHEET_NAME = "Kakeibo_Data"
TRANSACTIONS_WORKSHEET_NAME = "transactions"

def check_all_headers():
    if os.path.exists("credentials.json"):
        gc = gspread.service_account(filename="credentials.json")
    else:
        print("credentials.json not found")
        return

    try:
        sh = gc.open(SPREADSHEET_NAME)
        for ws in sh.worksheets():
            print(f"\nWorksheet: {ws.title}")
            all_values = ws.get_all_values()
            if not all_values:
                print("  Sheet is empty")
                continue
            
            headers = all_values[0]
            print(f"  Headers: {headers}")
            
            seen = {}
            dupes = []
            for i, h in enumerate(headers):
                if h in seen:
                    dupes.append((i+1, h))
                seen[h] = seen.get(h, 0) + 1
            
            if dupes:
                print(f"  DUPLICATES FOUND: {dupes}")
            else:
                print("  No duplicates found in header row.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_all_headers()

