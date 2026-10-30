import pandas as pd

df = pd.read_csv(r"C:\Users\rtx\Desktop\intern\data\hkbu\db_slice.csv")
main_df = df[(df['is_valid'] == True) & (~df['program_website'].str.contains('uic.edu.cn', na=False)) & (~df['program_name_en'].str.contains('BNBU', na=False))]

print(f"Total Main Campus valid: {len(main_df)}")
for idx, r in main_df.iterrows():
    print(f"ID {r['id']} ({r['program_name_en'][:30]}): {repr(r['application_time'])}")
