"""Test script to generate quarterly report"""
from pathlib import Path
from webapp.report_generator import generate_quarterly_report

csv_path = Path("BSP/gesamt.csv")
xml_path = Path("BSP/LUD.xml")
output_dir = Path("BSP")

print("Generating report...")
try:
    result_path = generate_quarterly_report(
        csv_path=csv_path,
        xml_path=xml_path,
        output_dir=output_dir,
    )
    print(f"Report generated: {result_path}")
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
