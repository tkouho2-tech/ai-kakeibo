import os

target_file = 'fixed_cost_expansion.py'
with open(target_file, 'rb') as f:
    content = f.read().decode('utf-8', 'replace')

lines = content.splitlines()

# Extract lines 1100 to 1600
with open('tmp_target.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines[1100:1600]))
