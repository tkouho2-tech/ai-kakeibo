import sys

target_file = 'fixed_cost_expansion.py'

with open(target_file, 'r', encoding='utf-8', errors='replace') as f:
    lines = f.read().splitlines()

new_lines = []
for i, line in enumerate(lines):
    if 'if start_ym and (my, mm) < start_ym: continue' in line:
        # Find the preceding 'for (my, mm)...' loop
        prev_idx = i - 1
        while prev_idx >= 0 and 'for (my, mm)' not in lines[prev_idx]:
            prev_idx -= 1
        
        if prev_idx >= 0:
            for_line = lines[prev_idx]
            # Get leading whitespace (indentation)
            for_indent = for_line[:len(for_line) - len(for_line.lstrip())]
            # Set the 'if' indentation to for_indent + 4 spaces
            new_lines.append(for_indent + '    ' + line.lstrip())
        else:
            new_lines.append(line)
    else:
        new_lines.append(line)

with open(target_file, 'w', encoding='utf-8') as f:
    f.write('\n'.join(new_lines))

print("Indentation fixed.")
