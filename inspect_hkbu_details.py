import pandas as pd
import re

df = pd.read_csv(r"C:\Users\rtx\Desktop\intern\data\hkbu\db_slice.csv")

print(f"Total: {len(df)}")
uic_df = df[df['program_website'].str.contains('uic.edu.cn', na=False) | df['program_name_en'].str.contains('BNBU', na=False)]
main_df = df[~df['id'].isin(uic_df['id'])]

print(f"UIC count: {len(uic_df)}")
print(f"Main Campus count: {len(main_df)}")
print(f"Main Campus Valid: {len(main_df[main_df['is_valid'] == True])}")
print(f"Main Campus Invalid: {len(main_df[main_df['is_valid'] == False])}")

print("\n--- Multi-year Programs in Main Campus ---")
for idx, r in main_df[main_df['is_valid'] == True].iterrows():
    dur = str(r['program_duration'])
    if '2' in dur or '1.5' in dur or '3' in dur:
        print(f"ID {r['id']} | Dur: {dur} | Tuition: {r['program_tuition']} | Name: {r['program_name_en']} | URL: {r['program_website']}")

print("\n--- Language Requirements Summary (Main Campus) ---")
for idx, r in main_df[main_df['is_valid'] == True].iterrows():
    lang = str(r['language_requirements'])
    has_cet6 = '六级' in lang or 'CET' in lang or 'CET-6' in lang
    if has_cet6:
        print(f"Accepts CET-6: ID {r['id']} | {r['program_name_en'][:35]} | Lang: {lang}")

print("\n--- Application Times Summary (Main Campus) ---")
old_cycle = []
new_cycle = []
for idx, r in main_df[main_df['is_valid'] == True].iterrows():
    app = str(r['application_time'])
    if '26年' in app or '2026' in app and '27年' not in app:
        old_cycle.append(r['id'])
    else:
        new_cycle.append(r['id'])
print(f"Old 26 cycle count: {len(old_cycle)}, New 27 cycle count: {len(new_cycle)}")
