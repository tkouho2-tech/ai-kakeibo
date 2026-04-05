import os

encodings = ['utf-8', 'cp932', 'shift_jis', 'euc-jp', 'latin-1']
target_file = 'fixed_cost_expansion.py'

with open(target_file, 'rb') as f:
    rawdata = f.read()

for enc in encodings:
    try:
        content = rawdata.decode(enc)
        print(f"Decoded successfully with {enc}!")
        output_file = f'/tmp/{target_file}_{enc}.txt'
        with open(output_file, 'w', encoding='utf-8') as f_out:
            f_out.write(content)
        print(f"Saved to {output_file}")
    except Exception as e:
        print(f"Decoding failed for {enc}: {e}")
