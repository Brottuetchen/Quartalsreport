"""Check Excel structure for debugging"""
from zipfile import ZipFile
import xml.etree.ElementTree as ET

def col_letter_to_num(col):
    """Convert column letter to number (A=1, B=2, etc.)"""
    num = 0
    for i, c in enumerate(reversed(col)):
        num += (ord(c.upper()) - ord('A') + 1) * (26 ** i)
    return num

z = ZipFile('BSP/Q3-2025.xlsx')
xml_content = z.read('xl/worksheets/sheet3.xml')
root = ET.fromstring(xml_content)

max_col = 0
max_col_ref = ""
ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

for cell in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
    r = cell.get('r', '')
    col = ''.join(c for c in r if c.isalpha())
    if col:
        col_num = col_letter_to_num(col)
        if col_num > max_col:
            max_col = col_num
            max_col_ref = r

print(f"Max column number: {max_col}")
print(f"Max column reference: {max_col_ref}")
print(f"\nColumns P (16) and Q (17) check:")

# Check for specific problematic cells
for row_num in [62, 63, 64, 65]:
    for col_letter in ['P', 'Q', 'R']:
        cell_ref = f"{col_letter}{row_num}"
        cell = root.find(f".//main:c[@r='{cell_ref}']", ns)
        if cell is not None:
            formula = cell.find('main:f', ns)
            value = cell.find('main:v', ns)
            if formula is not None:
                print(f"  {cell_ref}: Formula = {formula.text[:100] if formula.text else 'None'}")
            elif value is not None:
                print(f"  {cell_ref}: Value = {value.text}")
            else:
                print(f"  {cell_ref}: Empty cell")

z.close()
