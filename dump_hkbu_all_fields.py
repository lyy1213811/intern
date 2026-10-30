import psycopg2
import pandas as pd
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

conn = psycopg2.connect(
    host='pgm-bp1k176i7sid0e775o.pg.rds.aliyuncs.com',
    port=5432,
    user='data_maintain',
    password='Str0ngP@5s!',
    dbname='tunshuedu_db'
)

sql = """
    SELECT id, school_name_cn, school_name_en, program_name_en,
           program_name_cn, program_code, degree, program_duration,
           program_tuition, application_time, language_requirements,
           program_website, official_code, is_valid
    FROM ai_selection_programs
    WHERE degree = '硕士'
      AND (school_name_cn = '香港浸会大学' OR school_name_en = 'Hong Kong Baptist University')
    ORDER BY id ASC
"""
df = pd.read_sql(sql, conn)
conn.close()

print(f"Total: {len(df)}")
for idx, r in df.iterrows():
    print("="*60)
    print(f"[{idx+1:02d}] ID:{r['id']} | Valid:{r['is_valid']} | Code:{r['program_code']} | OffCode:{r['official_code']}")
    print(f"EN: {r['program_name_en']}")
    print(f"CN: {r['program_name_cn']}")
    print(f"Duration: {r['program_duration']} | Tuition: {r['program_tuition']}")
    print(f"App Time: {r['application_time']}")
    print(f"Lang: {r['language_requirements']}")
    print(f"URL: {r['program_website']}")
