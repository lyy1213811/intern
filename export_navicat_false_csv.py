import psycopg2
import pandas as pd
import sys
import csv

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

DB_CONFIG = {
    'host': 'pgm-bp1k176i7sid0e775o.pg.rds.aliyuncs.com',
    'port': 5432,
    'user': 'data_maintain',
    'password': 'Str0ngP@5s!',
    'dbname': 'tunshuedu_db'
}
conn = psycopg2.connect(**DB_CONFIG)
conn.set_session(readonly=True)
query = """
    SELECT id, school_name_cn, school_name_en, program_name_en, program_name_cn,
           degree, program_duration, program_tuition, application_time, language_requirements, program_website
    FROM ai_selection_programs
    WHERE (school_name_cn LIKE '%香港科技大学%' OR school_name_en ILIKE '%Hong Kong University of Science and Technology%')
      AND is_valid = false
    ORDER BY id ASC
"""
df = pd.read_sql(query, conn)
print(f"Total false: {len(df)}")

out_csv = r"c:\Users\rtx\Desktop\香港科技大学_数据库失效专业清单(is_valid=False_共21项).csv"
with open(out_csv, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.writer(f)
    writer.writerow([
        "序号", "数据库ID", "校区", "院校中英文", "专业中英文 (专业名/学位)",
        "学制", "现有学费", "现有申请时间", "现有语言要求", "官网URL"
    ])

    for idx, r in df.iterrows():
        p_cn = str(r['program_name_cn'] or '')
        p_en = str(r['program_name_en'] or '')
        p_url = str(r['program_website'] or '')
        is_gz = 'gz' in p_url.lower() or '(gz)' in p_en.lower() or '广州' in p_cn
        campus = "广州校区 (HKUST(GZ))" if is_gz else "清水湾本部"
        school = f"{r['school_name_cn']} {r['school_name_en']}"
        prog = f"{p_en} ({p_cn})"
        dur = str(r['program_duration'] or '')
        fee = str(r['program_tuition'] or '')
        app = str(r['application_time'] or '').replace('<br/>', '\n').strip()
        lang = str(r['language_requirements'] or '').replace('<br/>', '\n').strip()
        
        writer.writerow([
            idx + 1,
            r['id'],
            campus,
            school,
            prog,
            dur,
            fee,
            app,
            lang,
            p_url
        ])

print("Saved to:", out_csv)
