import pandas as pd

df = pd.read_csv(r"C:\Users\rtx\Desktop\intern\data\cityu\db_slice.csv")
valid_df = df[df['is_valid'] == True].copy()

print(f"Total valid: {len(valid_df)}")

# Check each field for patterns
print("\n--- Language requirements summary ---")
for idx, r in valid_df.iterrows():
    lang = str(r['language_requirements'])
    # check if CET-6 mentioned
    has_cet6 = '六级' in lang or 'CET' in lang or 'CET-6' in lang
    has_ielts = '雅思' in lang or 'IELTS' in lang
    has_toefl = '托福' in lang or 'TOEFL' in lang
    if not (has_ielts and has_toefl):
        print(f"Missing IELTS/TOEFL: ID {r['id']} {r['program_name_en']}: {lang}")

print("\n--- Multi-year and Tuition summary ---")
for idx, r in valid_df.iterrows():
    dur = str(r['program_duration'])
    tui = str(r['program_tuition'])
    name = r['program_name_en']
    if '2' in dur or '3' in dur:
        print(f"Multi-year: ID {r['id']} | {name} | Dur: {dur} | Tuition: {tui} | URL: {r['program_website']}")
