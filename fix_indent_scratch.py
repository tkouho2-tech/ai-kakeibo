import re

target_file = 'fixed_cost_expansion.py'

with open(target_file, 'r', encoding='utf-8', errors='replace') as f:
    content = f.read()

def fix_logic(match):
    # m.group(1): Indentation of the 'for' line
    # m.group(2): the 'for' statement itself
    # m.group(3): Any newline(s)
    # m.group(4): the 'if' statement
    indent = match.group(1)
    for_stmt = match.group(2)
    newlines = match.group(3)
    if_stmt = match.group(4)
    # Indent the 'if' statement 4 spaces more than the 'for' statement
    return f'{indent}{for_stmt}{newlines}{indent}    {if_stmt}'

# Regex to find unindented 'if start_ym' lines immediately after 'for' loops
pattern = r'^([ \t]*)(for \(my, mm\), col(?:_idx)? in pay_month_indices:)(\n+)([ \t]*if start_ym and \(my, mm\) < start_ym: continue)'
new_content = re.sub(pattern, fix_logic, content, flags=re.MULTILINE)

with open(target_file, 'w', encoding='utf-8') as f:
    f.write(new_content)

print("Indentation fix complete.")
