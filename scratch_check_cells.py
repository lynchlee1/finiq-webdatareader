import openpyxl

wb = openpyxl.load_workbook(r'opendart-to-excel/example.xlsm', data_only=True)
ws = wb['MAIN']

for r in range(1, 15):
    for c in range(1, 10):
        val = ws.cell(row=r, column=c).value
        if isinstance(val, str):
            codepoints = [ord(char) for char in val]
            hex_pts = [hex(cp) for cp in codepoints]
            print(f"Cell({r},{c}): {repr(val)} -> {hex_pts}")
