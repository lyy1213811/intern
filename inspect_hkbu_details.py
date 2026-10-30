import os, sys, openpyxl
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

dir_path = r'c:\Users\rtx\Desktop\HKBU_步骤一至四'
wb = openpyxl.load_workbook(os.path.join(dir_path, '香港浸会大学_数据核对与差异反馈表.xlsx'), data_only=True)

print("Sheet names:", wb.sheetnames)
ws1 = wb[wb.sheetnames[0]]
ws2 = wb[wb.sheetnames[1]]
print(f"Sheet 1: {ws1.title} | Rows: {ws1.max_row} | Cols: {ws1.max_column}")
print(f"Sheet 2: {ws2.title} | Rows: {ws2.max_row} | Cols: {ws2.max_column}")

print("\n--- Inspecting all rows of Sheet 2 ---")
for r in range(1, ws2.max_row + 1):
    vals = [ws2.cell(row=r, column=c).value for c in range(1, ws2.max_column + 1)]
    fill = ws2.cell(row=r, column=1).fill
    fill_color = fill.start_color.rgb if fill and fill.start_color else "None"
    if r == 1:
        print("Header:", vals)
    else:
        print(f"Row {r-1} (Color: {fill_color}): ID={vals[4]} | Type={vals[5]} | Old={vals[6]} | Sugg={vals[7]}")
        print(f"    Quote: {vals[8]}")
        print(f"    Advise: {vals[9]}\n")
