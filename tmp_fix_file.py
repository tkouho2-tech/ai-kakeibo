import os

target_file = 'fixed_cost_expansion.py'
with open(target_file, 'rb') as f:
    rawdata = f.read()

# Try to decode with utf-8, but handle errors
content = rawdata.decode('utf-8', errors='replace')
lines = content.splitlines()

fixed_lines = []
for i, line in enumerate(lines):
    l_num = i + 1
    # Specific fix for line 225 based on crash report
    if l_num == 225:
        # Reconstruct the line
        # amt = sum(safe_money_int_cast(tx.get("amount",0)) for tx in user_txs if _normalize(tx.get("payment_method"))==_normalize(m_name) and tx.get("category")!="消費税（内税）")
        # safe_gspread_call(ws_pay.update, values=rows_79_88, range_name="B79", value_input_option='USER_ENTERED')
        # Wait, safe_gspread_call should probably be OUTSIDE the sum() but inside the loop?
        # Actually in the dump it look merged.
        line = '                            amt = sum(safe_money_int_cast(tx.get("amount",0)) for tx in user_txs if _normalize(tx.get("payment_method"))==_normalize(m_name) and tx.get("category")!="消費税（内税）")'
        # Check if the next line was merged too
        # In the dump: ... and tx.get("category")!="消費E       safe_gspread_call(ws_pay.update, values=rows_79_88, range_name="B79", value_input_option='USER_ENTERED')
        # So I should add a newline.
        line += '\n                        safe_gspread_call(ws_pay.update, values=rows_79_88, range_name="B79", value_input_option="USER_ENTERED")'
    
    # Fix other common corruptions if possible
    line = line.replace('啪', '大分類')
    line = line.replace('Ȗږ', '科目明細')
    line = line.replace('xWv', '支払集計')
    line = line.replace('NWbgJ[h', 'クレジットカード')
    line = line.replace('xǗ', '支払管理')
    line = line.replace('Œf[^WJ', '固定費データ展開')
    line = line.replace('ϓf[^XV', '変動費データ更新')
    line = line.replace('XV', '更新')
    line = line.replace('Jn', '開始')
    line = line.replace('I', '完了')
    line = line.replace('yEj', '土日・祝日')
    
    fixed_lines.append(line)

with open('fixed_cost_expansion_fixed.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(fixed_lines))
print("Fixed file saved to fixed_cost_expansion_fixed.py")
