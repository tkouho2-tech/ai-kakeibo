import sys
import pandas as pd
from datetime import datetime

sys.path.append(r"c:\Users\t_kou\Kakeibo_Final_v3")
import app

class MockSessionState(dict):
    def __getattr__(self, name):
        return self.get(name)
    def __setattr__(self, name, value):
        self[name] = value

app.st.session_state = MockSessionState()
app.st.session_state['username'] = "tkouho"

# Load raw records
sheet = app.get_sheet(app.TRANSACTIONS_WORKSHEET_NAME)
values = app.safe_gspread_call(sheet.get_all_values)

headers = [h.strip() if h.strip() else f"empty_{i}" for i, h in enumerate(values[0])]
records_df = pd.DataFrame(values[1:])
if records_df.shape[1] > len(headers):
    headers += [f"extra_{i}" for i in range(len(headers), records_df.shape[1])]
records_df.columns = headers[:records_df.shape[1]]
records = records_df.to_dict('records')

# Use get_clean_df
df = app.get_clean_df(records, "tkouho")
print("Total rows for tkouho:", len(df))

if not df.empty and "amount" in df.columns:
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    zero_df = df[df["amount"] == 0]
    print(f"Total 0-amount rows for tkouho: {len(zero_df)}")
    
    # Sort by date descending
    zero_df = zero_df.sort_values(by="date", ascending=False)
    
    if not zero_df.empty:
        print("\nAll 0-amount rows for tkouho:")
        for _, row in zero_df.iterrows():
            print(f"Date: {row['date']}, Store: '{row.get('store_name', '')}', Item: '{row.get('item_name', '')}', Category: '{row.get('category', '')}'")

        # How does receipts_df look for these rows?
        store_col = "store_name" if "store_name" in df.columns else "store" if "store" in df.columns else None
        df_agg = df.copy()
        if "category" in df_agg.columns:
            df_agg.loc[df_agg["category"] == "消費税（内税）", "amount"] = 0
            
        receipts_df = df_agg.groupby(["date", store_col], as_index=False).agg(
            amount=("amount", "sum"),
            明細数=("amount", "count")
        )
        receipts_df.columns = ["日付", "店舗名", "金額合計", "明細数"]
        receipts_df["日付"] = receipts_df["日付"].dt.strftime('%Y-%m-%d')
        receipts_df["金額合計"] = receipts_df["金額合計"].apply(lambda x: int(x))
        receipts_df = receipts_df.sort_values(by="日付", ascending=False).reset_index(drop=True)
        
        print("\nreceipts_df 0-amount entries:")
        print(receipts_df[receipts_df["金額合計"] == 0])
