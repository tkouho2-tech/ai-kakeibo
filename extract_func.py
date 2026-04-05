import os

target_file = 'fixed_cost_expansion.py'
with open(target_file, 'rb') as f:
    content = f.read().decode('utf-8', 'replace')

lines = content.splitlines()

s_idx = -1
e_idx = -1

for i, line in enumerate(lines):
    if line.startswith('def execute_expansion'):
        s_idx = i
        break

if s_idx != -1:
    for i in range(s_idx + 1, len(lines)):
        if lines[i].startswith('def ') or lines[i].startswith('class '):
            e_idx = i
            break
    
    if e_idx == -1:
        e_idx = len(lines)
        
    with open('tmp_exec.py', 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines[s_idx:e_idx]))
    print(f"Extracted execute_expansion from line {s_idx} to {e_idx}")
else:
    print("execute_expansion not found in top-level.")
