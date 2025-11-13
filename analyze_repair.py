"""Analyze what Excel is repairing"""
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import re

def check_formula_errors(xml_content, sheet_name):
    """Check for problematic formulas in a sheet"""
    root = ET.fromstring(xml_content)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    problems = []

    for cell in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
        cell_ref = cell.get('r', 'unknown')
        formula_elem = cell.find('main:f', ns)

        if formula_elem is not None and formula_elem.text:
            formula = formula_elem.text

            # Check for empty string literals
            if '""' in formula:
                problems.append(f"{sheet_name}!{cell_ref}: Contains empty string literal \"\"")

            # Check for malformed INDEX/MATCH
            if 'INDEX' in formula and 'MATCH' in formula:
                # Check if MATCH has invalid parameters
                if 'MATCH(0,' in formula or 'MATCH("",' in formula:
                    problems.append(f"{sheet_name}!{cell_ref}: MATCH with invalid parameter (0 or empty string)")

            # Check for nested IFERROR with empty result
            if formula.count('IFERROR') > 3:
                problems.append(f"{sheet_name}!{cell_ref}: Too many nested IFERROR ({formula.count('IFERROR')})")

            # Check for invalid IF conditions
            if 'IF(0=' in formula or 'IF(""=' in formula:
                problems.append(f"{sheet_name}!{cell_ref}: IF with constant false condition")

    return problems

z = ZipFile('BSP/Q3-2025.xlsx')

print("Analyzing Excel file for repair issues...\n")

# Check first few employee sheets
for sheet_num in [3, 4, 5]:
    sheet_file = f'xl/worksheets/sheet{sheet_num}.xml'
    xml_content = z.read(sheet_file)

    problems = check_formula_errors(xml_content, f"sheet{sheet_num}")

    if problems:
        print(f"\nProblems in {sheet_file}:")
        for p in problems[:10]:  # Show first 10
            print(f"  - {p}")
        if len(problems) > 10:
            print(f"  ... and {len(problems) - 10} more")
    else:
        print(f"\n{sheet_file}: No obvious formula problems detected")

z.close()

print("\n" + "="*60)
print("Checking for specific problematic patterns...")

# Let's look at actual formulas in detail
z = ZipFile('BSP/Q3-2025.xlsx')
xml_content = z.read('xl/worksheets/sheet3.xml')
root = ET.fromstring(xml_content)
ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

print("\nSample formulas from row 8 (known issue - no budget):")
for col_letter in ['L', 'M', 'N']:
    cell_ref = f"{col_letter}8"
    cell = root.find(f".//main:c[@r='{cell_ref}']", ns)
    if cell is not None:
        formula = cell.find('main:f', ns)
        if formula is not None and formula.text:
            print(f"\n{cell_ref}:")
            print(f"  {formula.text[:200]}")

z.close()
