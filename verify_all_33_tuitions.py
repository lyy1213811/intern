import json
import sys
import re
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

with open('C:/Users/rtx/Desktop/intern/data/hkust/all_official_fields.json', 'r', encoding='utf-8') as f:
    raw_list = json.load(f)

db_df = pd.read_csv('C:/Users/rtx/Desktop/intern/data/hkust/db_slice.csv').set_index('id')

data_map = {}
for d in raw_list:
    if d.get('id'):
        data_map[int(float(d['id']))] = d

# All 33 tuition increases
all_tuition_updates = []

for pid, d in data_map.items():
    if pid not in db_df.index:
        continue
    db_rec = db_df.loc[pid]
    is_valid = str(db_rec.get('is_valid', '')).lower() == 'true'
    if not is_valid:
        continue
        
    db_fee = str(db_rec.get('program_tuition', '') or '')
    off_fee = d.get('off_fee', '')
    dur = str(db_rec.get('program_duration', '') or '')
    off_dur = d.get('off_dur', '')
    pname_en = db_rec['program_name_en']
    pname_cn = db_rec['program_name_cn']
    url = str(db_rec.get('program_website', '') or '')
    
    # parse db numbers
    db_nums = [int(x) for x in re.findall(r'\d+', db_fee)]
    db_val = db_nums[0] if db_nums else 0
    
    # parse off numbers
    off_matches = re.findall(r'HK\$\s*([\d,]+)', off_fee)
    if not off_matches and '$' in off_fee:
        off_matches = re.findall(r'\$\s*([\d,]+)', off_fee)
    off_vals = [int(x.replace(',', '')) for x in off_matches]
    
    if not off_vals:
        continue
        
    # Check if this is an already-matching 2-year program (e.g. 540k/2=270k, 320k/2=160k, 400k/2=200k, 560k/2=280k)
    if ('2年' in dur or '2 years' in off_dur.lower()) and any(db_val * 2 == v for v in off_vals):
        continue
    # Check if 1-year exact match
    if db_val in off_vals and '1.5' not in dur:
        continue
    # Check CIEM & EVEM
    if pid in [3214, 3224] and 210000 in off_vals:
        continue
        
    # Calculate official standard annual fee
    target_annual_fee = None
    official_display_str = ""
    quote_str = off_fee.strip()
    
    if pid in [3183, 3232]: # 1.5 year
        # total 300,000 -> annual 200,000
        total_fee = 300000
        annual_fee = 200000
        diff = annual_fee - db_val
        official_display_str = f"HK$300,000 (1.5年总额，折合 200,000 港币/年)"
        note = f"【学费上涨（1.5年年化）】官网2027/28总学费由原25万/22万统一上调至30万，按1.5年学制折算年化学费为 200,000 港币/年 (上涨 {diff:+,} 港币/年)，建议更新"
    else:
        # 1 year FT
        # pick the FT fee
        # If full-time in off_fee:
        ft_match = re.search(r'Full-time:\s*(?:HK)?\$\s*([\d,]+)', off_fee)
        if ft_match:
            val = int(ft_match.group(1).replace(',', ''))
        else:
            val = off_vals[0]
        diff = val - db_val
        official_display_str = f"HK${val:,}"
        note = f"【学费实质上涨】官网 2027/28 最新公布学费上调至 HK${val:,} (原 {db_val:,} 港币/年 -> 上涨 {diff:+,} 港币)，建议更新"
        
    all_tuition_updates.append({
        'id': pid,
        'pname_en': pname_en,
        'pname_cn': pname_cn,
        'url': url,
        'db_val': db_fee,
        'off_val': official_display_str,
        'quote': quote_str[:120],
        'note': note
    })

print(f"Total confirmed genuine tuition updates: {len(all_tuition_updates)}")
for idx, item in enumerate(all_tuition_updates, 1):
    print(f"{idx}. [{item['id']}] {item['pname_cn']}: {item['db_val']} -> {item['off_val']}")
