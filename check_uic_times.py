import pandas as pd

df = pd.read_csv(r"C:\Users\rtx\Desktop\intern\data\hkbu\db_slice.csv")
uic_df = df[df['program_website'].str.contains('uic.edu.cn', na=False) | df['program_name_en'].str.contains('BNBU', na=False)]

print("=== UIC Application Times ===")
for idx, r in uic_df.iterrows():
    print(f"ID {r['id']} | {r['program_name_en'][:35]}: {repr(r['application_time'])}")
