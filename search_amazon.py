import gspread
from app import get_gspread_client

def search_amazon_data():
    client = get_gspread_client()
    if not client:
        print("Client failed")
        return
    
    # List all spreadsheets
    all_ss = client.openall()
    print(f"Total spreadsheets found: {len(all_ss)}")
    
    for ss in all_ss:
        ss_name = ss.title
        print(f"\n--- Searching in SS: {ss_name} ---")
        try:
            for ws in ss.worksheets():
                print(f"  Checking WS: {ws.title}")
                try:
                    data = ws.get_all_values()
                    for i, row in enumerate(data):
                        row_str = " ".join([str(x) for x in row]).lower()
                        if "amazon" in row_str:
                            # Search for 2026-03-25, 3/25, 3-25, etc.
                            if "25" in row_str and ("3" in row_str or "mar" in row_str):
                                print(f"    [MATCH] Row {i+1}: {row}")
                except Exception as e:
                    print(f"    Error reading {ws.title}: {e}")
        except Exception as e:
            print(f"    Could not access worksheets in {ss_name}: {e}")

if __name__ == "__main__":
    search_amazon_data()
