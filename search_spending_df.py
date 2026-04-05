with open(r"c:\Users\t_kou\Kakeibo_Final_v3\app.py", "r", encoding="utf-8") as f:
    for i, line in enumerate(f):
        if "spending_df" in line:
            print(f"Line {i+1}: {line.strip()}")
