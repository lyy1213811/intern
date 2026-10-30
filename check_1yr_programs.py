import os
import sys
import re
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

DATA_DIR = r"C:\Users\rtx\Desktop\intern\data\polyu"
df_slice = pd.read_csv(os.path.join(DATA_DIR, "db_slice.csv"))
df_valid = df_slice[df_slice['is_valid'].astype(str).str.lower() == 'true'].copy().reset_index(drop=True)

p1 = df_valid[df_valid['program_duration'] == '1年']
print(f"Total 1-year programs: {len(p1)}")

for idx, r in p1.iterrows():
    p_url = str(r['program_website'])
    if 'speed-polyu.edu.hk' in p_url:
        continue
    num = re.findall(r'\d+', str(r['program_tuition']))
    val = int(num[0]) if num else 0
    print(f"[{r['id']}] {r['program_name_en']} | {r['program_tuition']} | {r['official_code']}")
