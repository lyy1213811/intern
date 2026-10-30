import os
import shutil
import openpyxl

desktop_file = r"C:\Users\rtx\Desktop\香港理工大学_数据核对与差异反馈表.xlsx"
folder_file = r"C:\Users\rtx\Desktop\PolyU_步骤一至四\香港理工大学_数据核对与差异反馈表.xlsx"

if os.path.exists(desktop_file):
    shutil.copy2(desktop_file, folder_file)
    print("Synced desktop deliverable to folder:", folder_file)

wb = openpyxl.load_workbook(folder_file, read_only=True)
ws2 = wb['数据核对与差异反馈表']
print("Sheet 2 row count:", ws2.max_row)
