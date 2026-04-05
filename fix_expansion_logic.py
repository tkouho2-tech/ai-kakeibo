import os
import re

target_file = 'fixed_cost_expansion.py'

# Read with appropriate encoding
try:
    with open(target_file, 'rb') as f:
        content = f.read().decode('cp932', errors='replace')
except:
    with open(target_file, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()

# Logic to inject start_ym support into execute_expansion
new_expansion_body = r'''    # 実行対象年月のフィルタリング
    valid_indices = pay_month_indices
    is_partial = False
    if start_ym:
        valid_indices = [idx for idx in pay_month_indices if idx[0] >= start_ym]
        is_partial = True

    configs = [{"cat": "クレジットカード", "start": 8, "limit": 18}, {"cat": "口座引落", "start": 27, "limit": 10}, {"cat": "銀行振込", "start": 38, "limit": 10}]
    upd_fixed_list = []

    for cfg in configs:
        rows = category_groups[cfg["cat"]][:cfg["limit"]]
        if not is_partial:
            # 全期間展開：ラベル列含む行全体（B列〜）を上書き更新
            update_vals = []
            for i, r_d in enumerate(rows):
                r_d["Sno"] = str(i + 1)
                update_vals.append(dict_to_row(r_d)[1:])
            while len(update_vals) < cfg["limit"]:
                update_vals.append(dict_to_row({"大分類": "固定費", "科目1": cfg["cat"], "Sno": str(len(update_vals)+1)})[1:])
            safe_gspread_call(ws_pay.update, values=update_vals, range_name=f"B{cfg['start']}", value_input_option='USER_ENTERED')
        else:
            # 期間指定展開：指定月以降の月次金額セルのみをピンポイント更新（ラベル列B〜Hは保護）
            for i in range(cfg["limit"]):
                target_row = cfg["start"] + i
                r_d = rows[i] if i < len(rows) else {"大分類": "固定費", "科目1": cfg["cat"], "Sno": str(i+1)}
                for (my, mm), col_idx in valid_indices:
                    val = r_d.get(f"{my}.{mm}月", "")
                    upd_fixed_list.append({'range': f"'{ws_pay.title}'!{rowcol_to_a1(target_row, col_idx)}", 'values': [[val]]})

    if upd_fixed_list:
        safe_gspread_call(ss.values_batch_update, {'valueInputOption': 'USER_ENTERED', 'data': upd_fixed_list})

    # 78行目（小遣い予算）の更新は execute_variable_cost_update へ移行したため、ここでは行わない
    safe_gspread_call(ws_pay.update, values=[[datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")]], range_name="F4", value_input_option='USER_ENTERED')
    execute_variable_cost_update(username, start_ym=start_ym, skip_backup=True)
    return True, "固定費データ展開完了"'''

# Find the block in execute_expansion and replace it
# We search for the start of configs and the end of execute_expansion (return line)
pattern = re.compile(r'    configs = \[.*?return True, ".*?"', re.DOTALL)
content = pattern.sub(new_expansion_body, content)

# Also update execute_variable_cost_update loops to respect start_ym
# 1. 78行目（小遣い予算）
# 2. 79-88行目（集計）
# 3. 96-100行目（収入連動）

# Add start_ym filtering to row 78
content = content.replace(
    'for (my, mm), col_idx in pay_month_indices:',
    'for (my, mm), col_idx in pay_month_indices:\n                if start_ym and (my, mm) < start_ym: continue'
)

# Add start_ym filtering to row 79-88
content = content.replace(
    'for (my, mm), col in pay_month_indices:',
    'for (my, mm), col in pay_month_indices:\n                    if start_ym and (my, mm) < start_ym: continue'
)

with open(target_file, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement complete.")
