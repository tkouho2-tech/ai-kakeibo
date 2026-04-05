import gspread
from app import get_gspread_client
import sys

def search_anywhere_amazon():
    client = get_gspread_client()
    all_ss = client.openall()
    
    for ss in all_ss:
        title = ss.title
        print(f"Checking SS: {title}")
        try:
            for ws in ss.worksheets():
                ws_name = ws.title
                try:
                    data = ws.get_all_values()
                    for i, row in enumerate(data):
                        row_str = " ".join([str(x) for x in row]).lower()
                        if "amazon" in row_str:
                            print(f"  [FOUND] SS:{title} | WS:{ws_name} | Row {i+1}: {row}")
                except:
                    pass
        except:
            pass

if __name__ == "__main__":
    search_anywhere_amazon()
