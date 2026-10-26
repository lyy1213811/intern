"""
步骤 1：数据库切片脚本 (1_export_db_slice.py)
核心目标：从系统只读数据库拉取指定院校的所有有效专业及核心字段，生成 db_slice.csv。
"""

import os
import sys
import argparse
import psycopg2
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# 数据库只读连接配置
DB_CONFIG = {
    "host": os.getenv("DB_HOST", "pgm-bp1k176i7sid0e775o.pg.rds.aliyuncs.com"),
    "port": int(os.getenv("DB_PORT", 5432)),
    "user": os.getenv("DB_USER", "data_maintain"),
    "password": os.getenv("DB_PASSWORD", "Str0ngP@5s!"),
    "dbname": os.getenv("DB_NAME", "tunshuedu_db")
}

# 院校代码与名称映射
SCHOOL_MAPPING = {
    "hku": {
        "school_name_cn": "香港大学",
        "school_name_en": "The University of Hong Kong",
        "dir": "hku"
    },
    "cuhk": {
        "school_name_cn": "香港中文大学",
        "school_name_en": "The Chinese University of Hong Kong",
        "dir": "cuhk"
    },
    "polyu": {
        "school_name_cn": "香港理工大学",
        "school_name_en": "The Hong Kong Polytechnic University",
        "dir": "polyu"
    },
    "cityu": {
        "school_name_cn": "香港城市大学",
        "school_name_en": "City University of Hong Kong",
        "dir": "cityu"
    },
    "hkust": {
        "school_name_cn": "香港科技大学",
        "school_name_en": "The Hong Kong University of Science and Technology",
        "dir": "hkust"
    }
}

REQUIRED_FIELDS = [
    "id",
    "school_name_cn",
    "school_name_en",
    "program_name_en",
    "program_name_cn",
    "program_code",
    "degree",
    "program_duration",
    "program_tuition",
    "application_time",
    "language_requirements",
    "program_website",
    "official_code",
    "is_valid"
]

def export_db_slice(school_key="hku"):
    school_info = SCHOOL_MAPPING.get(school_key.lower())
    if not school_info:
        # 兼容直接输入中文校名
        school_info = {
            "school_name_cn": school_key,
            "school_name_en": school_key,
            "dir": school_key.lower()
        }

    out_dir = os.path.join("data", school_info["dir"])
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, "db_slice.csv")

    print(f"[*] 正在连接数据库导出 [{school_info['school_name_cn']}] 专业切片...")
    conn = psycopg2.connect(**DB_CONFIG)
    try:
        # 确保只读事务
        conn.set_session(readonly=True)
        
        fields_str = ", ".join(REQUIRED_FIELDS)
        sql = f"""
            SELECT {fields_str}
            FROM ai_selection_programs
            WHERE (school_name_cn = %s OR school_name_en = %s)
              AND degree = '硕士'
            ORDER BY id ASC
        """
        
        with conn.cursor() as cur:
            cur.execute(sql, (school_info["school_name_cn"], school_info["school_name_en"]))
            rows = cur.fetchall()
            columns = [desc[0] for desc in cur.description]
            df = pd.DataFrame(rows, columns=columns)

        # 保存 CSV，统一采用 utf-8-sig 以便 Excel 正确识别中文
        df.to_csv(out_file, index=False, encoding="utf-8-sig")
        print(f"[OK] 成功导出 {len(df)} 条专业记录至: {out_file}")
        
        # 统计有效与失效专业数量
        valid_cnt = (df["is_valid"] == True).sum()
        invalid_cnt = (df["is_valid"] == False).sum()
        print(f"    - 有效专业 (is_valid=True): {valid_cnt} 条")
        print(f"    - 失效/历史专业 (is_valid=False): {invalid_cnt} 条")
        
        return df, out_file
    finally:
        conn.close()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="步骤1：从数据库拉取指定院校切片")
    parser.add_argument("--school", default="hku", help="院校代码或中文名，如 hku, cuhk, polyu")
    args = parser.parse_args()
    
    export_db_slice(args.school)
