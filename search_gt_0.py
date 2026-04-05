with open(r"c:\Users\t_kou\Kakeibo_Final_v3\app.py", "r", encoding="utf-8") as f:
    for i, line in enumerate(f):
        if "> 0" in line or ">0" in line:
            print(f"Line {i+1}: {line.strip()}")
