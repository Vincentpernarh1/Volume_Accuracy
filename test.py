from openpyxl import load_workbook
import os

# Paths to the files
parten_path = r"C:\Users\perna\Desktop\STALLANTIS\Volume_Accuracy\Map Opeperation\Relatorio_61_PARTEN.xlsx"
relatorio_path = r"C:\Users\perna\Desktop\STALLANTIS\Volume_Accuracy\Relatorio_61.xlsx"

def inspect_excel(file_path, sheet_name=None):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return
    wb = load_workbook(file_path, read_only=True, data_only=True)
    print(f"\n--- Inspecting: {os.path.basename(file_path)} ---")
    print(f"Sheets: {wb.sheetnames}")
    if sheet_name and sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        data = list(sheet.iter_rows(values_only=True))
        if data:
            print(f"Columns (first row): {data[0]}")
            print(f"First 5 rows:")
            for row in data[1:6]:  # Skip header
                print(row)
        else:
            print("No data in sheet.")
    else:
        for name in wb.sheetnames[:3]:  # Inspect first 3 sheets
            sheet = wb[name]
            data = list(sheet.iter_rows(values_only=True))
            print(f"\nSheet: {name}")
            if data:
                print(f"Columns: {data[0]}")
                print(f"First 3 rows: {data[1:4]}")
            else:
                print("No data.")

# Inspect parten file (focus on "parten" sheet)
inspect_excel(parten_path, "parten")

# Inspect relatorio_61 file (all sheets)
inspect_excel(relatorio_path)