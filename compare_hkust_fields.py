import sys
import os
import re
import pandas as pd
from bs4 import BeautifulSoup

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

diff = pd.read_csv('C:/Users/rtx/Desktop/intern/data/hkust/diff_programs.csv')
matched = diff[diff['category'].str.contains('匹配')]
db = pd.read_csv('C:/Users/rtx/Desktop/intern/data/hkust/db_slice.csv').set_index('id')

print(f"Total matched: {len(matched)}")
comparisons = []
for _, row in matched.iterrows():
    db_id = int(float(row['db_id']))
    db_row = db.loc[db_id]
    db_fee = str(db_row.get('program_tuition', '') or '')
    db_app = str(db_row.get('application_time', '') or '')
    
    off_url = str(row['official_url'])
    slug = os.path.basename(off_url.rstrip('/'))
    html_name = f"{db_id}_{slug}.html"
    html_path = os.path.join('C:/Users/rtx/Desktop/intern/data/hkust/htmls', html_name)
    if not os.path.exists(html_path):
        continue
    with open(html_path, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')
    
    off_fee = ''
    off_app = ''
    off_mode = ''
    for r in soup.find_all('div', class_='block-table-row'):
        txt = r.get_text(strip=True)
        if txt.startswith('Program Fee'):
            off_fee = txt[len('Program Fee'):]
        elif txt.startswith('Tuition Fee'):
            off_fee = txt[len('Tuition Fee'):]
        elif txt.startswith('Application Deadlines'):
            off_app = txt[len('Application Deadlines'):]
        elif txt.startswith('Mode of Study'):
            off_mode = txt[len('Mode of Study'):]

    comparisons.append({
        'id': db_id,
        'name': row['db_name_en'],
        'faculty': row['faculty'],
        'mode': off_mode,
        'db_fee': db_fee,
        'off_fee': off_fee,
        'db_app': db_app,
        'off_app': off_app
    })

df_comp = pd.DataFrame(comparisons)
print(f"Loaded comparisons for {len(df_comp)} programs")

# Let's inspect tuition differences
diff_fees = []
for _, r in df_comp.iterrows():
    # extract numbers from db_fee and off_fee
    db_num = re.search(r'\d+', r['db_fee'])
    off_num = re.search(r'HK\$\s*([\d\,]+)', r['off_fee'])
    if db_num and off_num:
        d_val = int(db_num.group(0))
        o_val = int(off_num.group(1).replace(',', ''))
        if d_val != o_val:
            diff_fees.append((r['id'], r['name'], d_val, o_val, r['off_fee']))

print(f"\n--- Tuition differences found: {len(diff_fees)} ---")
for df in diff_fees:
    print(f"[{df[0]}] {df[1]}: DB={df[2]} vs OFF={df[3]} | raw: {df[4][:80]}")
