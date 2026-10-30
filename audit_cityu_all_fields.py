import os, sys, openpyxl
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

dir_path = r'c:\Users\rtx\Desktop\CityU'
df = pd.read_excel(os.path.join(dir_path, '步骤1_数据库切片_db_slice.xlsx'))

print("=================================================================")
print("  CHECK ROUND 4: DURATION & TUITION AUDIT ACROSS ALL 78 PROGRAMS")
print("=================================================================")

print("Duration value counts:")
print(df['学制'].value_counts())

print("\n--- All Multi-Year Programs (Duration != '1年') ---")
multi_year = df[df['学制'] != '1年']
for idx, r in multi_year.iterrows():
    print(f"ID: {r['数据库ID']} | 校区: {r['校区归属']} | 学制: {r['学制']} | 学费: {r['现有学费']} | 专业: {r['专业英文名']}")

print("\n=================================================================")
print("  CHECK ROUND 5: LANGUAGE REQUIREMENTS AUDIT ACROSS ALL 78 PROGRAMS")
print("=================================================================")
for idx, r in df.iterrows():
    lang = str(r['现有语言要求'])
    pid = r['数据库ID']
    pname = r['专业英文名']
    campus = r['校区归属']
    if 'Law' in pname or 'Juris' in pname or 'LLM' in pname:
        print(f"[LAW] ID: {pid} | {pname} | Lang: {lang}")
    elif 'Dongguan' in campus:
        if '425' not in lang and lang != 'nan':
            print(f"[DG Alert] ID: {pid} | {pname} | Lang: {lang}")
    else:
        if '450' not in lang and '英语' not in lang and lang != 'nan':
            print(f"[General Main Campus Non-450] ID: {pid} | {pname} | Lang: {lang}")

print("\n=================================================================")
print("  CHECK ROUND 6: TUITION CURRENCY & FORMAT AUDIT")
print("=================================================================")
non_standard_tuition = []
for idx, r in df.iterrows():
    t = str(r['现有学费'])
    pid = r['数据库ID']
    if not (t.endswith('港币/年') or t.endswith('人民币/年')):
        non_standard_tuition.append((pid, r['校区归属'], r['专业英文名'], t))

print(f"Non-standard tuition count: {len(non_standard_tuition)}")
for item in non_standard_tuition:
    print("  Alert:", item)

