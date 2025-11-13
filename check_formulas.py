"""Check critical formulas in employee sheets"""
from zipfile import ZipFile
import xml.etree.ElementTree as ET

z = ZipFile('BSP/Q3-2025.xlsx')
xml_content = z.read('xl/worksheets/sheet3.xml')
root = ET.fromstring(xml_content)

ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

print("Checking critical formulas in sheet3 (first employee):\n")

# Check a few rows
for row_num in [7, 8, 62, 63]:
    print(f"Row {row_num}:")

    # Column L (Stundensatz)
    cell_ref = f"L{row_num}"
    cell = root.find(f".//main:c[@r='{cell_ref}']", ns)
    if cell is not None:
        formula = cell.find('main:f', ns)
        if formula is not None and formula.text:
            # Truncate long formulas
            formula_text = formula.text
            if len(formula_text) > 150:
                formula_text = formula_text[:150] + "..."
            print(f"  L (Stundensatz): {formula_text}")

    # Column N (Budget Gesamt)
    cell_ref = f"N{row_num}"
    cell = root.find(f".//main:c[@r='{cell_ref}']", ns)
    if cell is not None:
        formula = cell.find('main:f', ns)
        if formula is not None and formula.text:
            print(f"  N (Budget Gesamt): {formula.text}")

    print()

z.close()
