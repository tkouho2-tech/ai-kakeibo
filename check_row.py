import gspread
from app import get_gspread_client

def check_row_570():
    client = get_gspread_client()
    ss = client.open("Kakeibo_Data")
    ws = ss.worksheet("transactions")
    # Row 570 (1-based)
    row = ws.row_values(570)
    print(f"Row 570: {row}")

if __name__ == "__main__":
    check_row_570()
