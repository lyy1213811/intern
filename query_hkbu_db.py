import os
import sys
import psycopg2
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

DB_CONFIG = {
    "host": os.getenv("DB_HOST", "pgm-bp1k176i7sid0e775o.pg.rds.aliyuncs.com"),
    "port": int(os.getenv("DB_PORT", 5432)),
    "user": os.getenv("DB_USER", "data_maintain"),
    "password": os.getenv("DB_PASSWORD", "Str0ngP@5s!"),
    "dbname": os.getenv("DB_NAME", "tunshuedu_db")
}

conn = psycopg2.connect(**DB_CONFIG)
try:
    conn.set_session(readonly=True)
    sql = """
        SELECT id, school_name_cn, school_name_en, program_name_en,
               program_name_cn, program_code, degree, program_duration,
               program_tuition, application_time, language_requirements,
               program_website, official_code, is_valid
        FROM ai_selection_programs
        WHERE degree = '硕士'
          AND (school_name_cn LIKE '%浸会%' OR school_name_en ILIKE '%Baptist%')
        ORDER BY id ASC
    """
    df = pd.read_sql(sql, conn)
    print(f"Total HKBU Master's records: {len(df)}")
    print(df['is_valid'].value_counts())
    print("\nSchool names cn:", df['school_name_cn'].unique())
    print("School names en:", df['school_name_en'].unique())
finally:
    conn.close()
