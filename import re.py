import re

a2l_file = r"c:\Users\xiehaoyang\Desktop\生成A2L文件\APSW_Model_Higer_Merged_20250103_0919.a2l"
conversion_methods = {}

with open(a2l_file, "r", encoding="utf-8") as f:
    for idx, line in enumerate(f, 1):
        match = re.search(r'/\* Conversion Method\s+\*/\s+(\S+)', line)
        if match:
            method = match.group(1)
            if method not in conversion_methods:
                conversion_methods[method] = idx

for method, line_num in conversion_methods.items():
    print(f"{method}\t首次出现行号: {line_num}")