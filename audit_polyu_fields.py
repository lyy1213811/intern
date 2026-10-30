import os
import sys
import re
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

DATA_DIR = r"C:\Users\rtx\Desktop\intern\data\polyu"
df_slice = pd.read_csv(os.path.join(DATA_DIR, "db_slice.csv"))
df_valid = df_slice[df_slice['is_valid'].astype(str).str.lower() == 'true'].copy().reset_index(drop=True)

print(f"Total active valid programs: {len(df_valid)}")

# Check 1: Missing values in fields
for col in ['program_duration', 'program_tuition', 'application_time', 'language_requirements', 'program_website']:
    null_count = df_valid[col].isna().sum()
    print(f"Column '{col}' nulls: {null_count}")

# Check 2: Programs with zero tuition or abnormal tuition format
print("\n--- Checking abnormal tuitions ---")
for idx, r in df_valid.iterrows():
    tuition = str(r['program_tuition'])
    if tuition == 'nan' or '0' == re.sub(r'\D', '', tuition) or not re.search(r'\d+', tuition):
        print(f"[{r['id']}] {r['program_name_en']} | Tuition: {tuition}")

# Check 3: Check 1.5 year, 2 year, 3 year programs
print("\n--- Distribution of durations ---")
print(df_valid['program_duration'].value_counts())

# Check 4: Check if any 1.5-year / 2-year programs have tuitions matching whole-programme amounts
print("\n--- Inspecting 1.5-year programs ---")
p15 = df_valid[df_valid['program_duration'] == '1.5年']
for idx, r in p15.iterrows():
    num = re.findall(r'\d+', str(r['program_tuition']))
    val = int(num[0]) if num else 0
    # check if val looks like whole programme (e.g. divisible by 9000, 9200, 8100, 7400, etc.)
    # In PolyU, typical total credits are 30-34.
    # 31 credits * 9000 = 279,000 -> 1.5 yr annual is 186,000
    # 31 credits * 9200 = 285,200 -> 1.5 yr annual is 190,133
    # 31 credits * 9300 = 288,300 -> 1.5 yr annual is 192,200
    # 31 credits * 9500 = 294,500 -> 1.5 yr annual is 196,333
    # 31 credits * 10200 = 316,200 -> 1.5 yr annual is 210,800
    # 31 credits * 14900 = 461,900 -> 1.5 yr annual is 307,933
    # 34 credits * 8100 = 275,400 -> 1.5 yr annual is 183,600
    # 31 credits * 8300 = 257,300 -> 1.5 yr annual is 171,533
    # 31 credits * 7400 = 229,400 -> 1.5 yr annual is 152,933
    # 31 credits * 7900 = 244,900 -> 1.5 yr annual is 163,267
    whole_markers = [279000, 285200, 288300, 294500, 316200, 461900, 275400, 257300, 229400, 244900]
    if val in whole_markers:
        print(f"[{r['id']}] {r['program_name_en']}: {val}港币/年 (Unannualized whole programme fee! Should be {int(val/1.5)}/year)")

print("\n--- Inspecting 2-year programs ---")
p2 = df_valid[df_valid['program_duration'] == '2年']
for idx, r in p2.iterrows():
    print(f"[{r['id']}] {r['program_name_en']} | Tuition: {r['program_tuition']}")

print("\n--- Inspecting 3-year programs ---")
p3 = df_valid[df_valid['program_duration'] == '3年']
for idx, r in p3.iterrows():
    print(f"[{r['id']}] {r['program_name_en']} | Tuition: {r['program_tuition']}")
