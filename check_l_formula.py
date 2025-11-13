"""Check actual L formula structure"""
from zipfile import ZipFile
import xml.etree.ElementTree as ET

z = ZipFile('BSP/Q3-2025.xlsx')
xml_content = z.read('xl/worksheets/sheet3.xml')
root = ET.fromstring(xml_content)
ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

# Get L7 formula (should have lookup)
cell = root.find(".//main:c[@r='L7']", ns)
if cell is not None:
    formula = cell.find('main:f', ns)
    if formula is not None and formula.text:
        print("L7 formula:")
        print(formula.text)
        print("\nIFERROR count:", formula.text.count('IFERROR'))

z.close()
