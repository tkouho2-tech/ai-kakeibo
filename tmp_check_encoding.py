import chardet

with open('fixed_cost_expansion.py', 'rb') as f:
    rawdata = f.read()
    result = chardet.detect(rawdata)
    encoding = result['encoding']
    print(f"Detected encoding: {encoding}")
    
try:
    content = rawdata.decode(encoding)
    print("Decoded successfully!")
    with open('/tmp/fixed_cost_expansion_decoded.txt', 'w', encoding='utf-8') as f:
        f.write(content)
except Exception as e:
    print(f"Decoding failed: {e}")
