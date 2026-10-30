import os, sys, openpyxl
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

dir_path = r'c:\Users\rtx\Desktop\PolyU_步骤一至四'
df_false = pd.read_excel(os.path.join(dir_path, '步骤1附表_数据库失效专业_is_valid_false.xlsx'))

print(f"Total false records: {len(df_false)}")
for idx, r in df_false.iterrows():
    print(f"Idx {idx+1}: ID {r['数据库ID']} | EN: {r['专业英文名']} | CN: {r['专业中文名']}")

# Check why they are false and if any of them resumed
df_slice = pd.read_excel(os.path.join(dir_path, '步骤1_数据库切片_db_slice.xlsx'))
print("\nTotal slice records:", len(df_slice))
print("is_valid counts:\n", df_slice['是否有效(is_valid)'].value_counts())

