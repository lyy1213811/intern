import os, sys, openpyxl
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

dir_path = r'c:\Users\rtx\Desktop\PolyU_步骤一至四'
df_false = pd.read_excel(os.path.join(dir_path, '步骤1附表_数据库失效专业_is_valid_false.xlsx'))

for idx, r in df_false.iterrows():
    print(f"Row {idx+1}: ID={r['数据库ID']} | Name={r['专业英文名']} | CN={r['专业中文名']} | URL={r['官网网址']}")
