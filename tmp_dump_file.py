import sys

target_file = 'fixed_cost_expansion.py'
with open(target_file, 'rb') as f:
    rawdata = f.read()

# Try to decode with utf-8, but replace errors with a visible marker
content = rawdata.decode('utf-8', errors='replace')
lines = content.splitlines()

for i, line in enumerate(lines):
    print(f"{i+1:4}: {line}")
