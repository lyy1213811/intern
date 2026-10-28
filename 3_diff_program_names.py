import os
import sys
import re
import argparse
import difflib
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

def normalize_name(name):
    if not isinstance(name, str):
        return ""
    # 转小写并去除标点符号与多余空格
    s = name.lower()
    s = re.sub(r'[\(\)\[\]\,\.\-\/\:\'\"]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def diff_program_names(school="hku"):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, "data", school.lower())
    db_file = os.path.join(data_dir, "db_slice.csv")
    official_file = os.path.join(data_dir, "official_list.csv")
    out_file = os.path.join(data_dir, "diff_programs.csv")

    if not os.path.exists(db_file):
        raise FileNotFoundError(f"未找到数据库切片文件: {db_file}，请先运行 1_export_db_slice.py")
    if not os.path.exists(official_file):
        raise FileNotFoundError(f"未找到官网专业列表文件: {official_file}，请先运行 2_fetch_official_list.py")

    df_db = pd.read_csv(db_file)
    df_off = pd.read_csv(official_file)

    # 核心优先级：有效在招专业 (is_valid == True) 必须优先匹配官网名额！
    # 杜绝 1098、1135 等失效小方向提前抢占 44051、44050 等在招大类的官网名额
    df_db['sort_valid'] = df_db['is_valid'].apply(lambda x: 0 if str(x).lower() == 'true' else 1)
    df_db = df_db.sort_values(by=['sort_valid', 'id']).reset_index(drop=True)

    print(f"[*] 读取数据库记录: {len(df_db)} 条")
    print(f"[*] 读取官网记录: {len(df_off)} 条")

    # 预处理英文名称
    df_db['norm_name'] = df_db['program_name_en'].apply(normalize_name)
    df_off['norm_name'] = df_off['official_program_name'].apply(normalize_name)

    matched = []
    unmatched_off = []
    unmatched_db = []

    off_matched_indices = set()
    db_matched_indices = set()

    # 1. 优先通过 URL 或 规范化名称进行精确对齐
    for db_idx, db_row in df_db.iterrows():
        db_url = str(db_row.get('program_website', '')).strip().lower()
        db_norm = db_row['norm_name']
        
        found = False
        for off_idx, off_row in df_off.iterrows():
            if off_idx in off_matched_indices:
                continue
            off_url = str(off_row.get('official_url', '')).strip().lower()
            off_norm = off_row['norm_name']

            # 条件1：URL完全吻合
            if db_url and off_url and (db_url == off_url or db_url in off_url or off_url in db_url):
                matched.append({
                    "category": "完全匹配",
                    "match_score": 1.0,
                    "db_id": db_row['id'],
                    "db_name_en": db_row['program_name_en'],
                    "db_name_cn": db_row['program_name_cn'],
                    "official_name": off_row['official_program_name'],
                    "faculty": off_row['faculty'],
                    "db_url": db_row['program_website'],
                    "official_url": off_row['official_url'],
                    "db_is_valid": db_row.get('is_valid', '')
                })
                off_matched_indices.add(off_idx)
                db_matched_indices.add(db_idx)
                found = True
                break
            
            # 条件2：名称规范化后完全相同
            if db_norm == off_norm:
                matched.append({
                    "category": "完全匹配",
                    "match_score": 1.0,
                    "db_id": db_row['id'],
                    "db_name_en": db_row['program_name_en'],
                    "db_name_cn": db_row['program_name_cn'],
                    "official_name": off_row['official_program_name'],
                    "faculty": off_row['faculty'],
                    "db_url": db_row['program_website'],
                    "official_url": off_row['official_url'],
                    "db_is_valid": db_row.get('is_valid', '')
                })
                off_matched_indices.add(off_idx)
                db_matched_indices.add(db_idx)
                found = True
                break

    # 2. 针对未匹配项，执行模糊匹配 (阈值 0.82)
    for db_idx, db_row in df_db.iterrows():
        if db_idx in db_matched_indices:
            continue
        db_norm = db_row['norm_name']
        best_ratio = 0.0
        best_off_idx = None

        for off_idx, off_row in df_off.iterrows():
            if off_idx in off_matched_indices:
                continue
            ratio = difflib.SequenceMatcher(None, db_norm, off_row['norm_name']).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_off_idx = off_idx

        if best_ratio >= 0.82 and best_off_idx is not None:
            off_row = df_off.loc[best_off_idx]
            matched.append({
                "category": "模糊匹配(疑似更名)",
                "match_score": round(best_ratio, 3),
                "db_id": db_row['id'],
                "db_name_en": db_row['program_name_en'],
                "db_name_cn": db_row['program_name_cn'],
                "official_name": off_row['official_program_name'],
                "faculty": off_row['faculty'],
                "db_url": db_row['program_website'],
                "official_url": off_row['official_url'],
                "db_is_valid": db_row.get('is_valid', '')
            })
            off_matched_indices.add(best_off_idx)
            db_matched_indices.add(db_idx)

    # 3. 收集官网独有集（疑似新增/漏录）
    for off_idx, off_row in df_off.iterrows():
        if off_idx not in off_matched_indices:
            unmatched_off.append({
                "category": "官网独有(疑似新增/漏录)",
                "match_score": 0.0,
                "db_id": "",
                "db_name_en": "",
                "db_name_cn": "",
                "official_name": off_row['official_program_name'],
                "faculty": off_row['faculty'],
                "db_url": "",
                "official_url": off_row['official_url'],
                "db_is_valid": ""
            })

    # 4. 收集数据库独有集（疑似停办/下线）
    for db_idx, db_row in df_db.iterrows():
        if db_idx not in db_matched_indices:
            unmatched_db.append({
                "category": "数据库独有(疑似停办/下线)",
                "match_score": 0.0,
                "db_id": db_row['id'],
                "db_name_en": db_row['program_name_en'],
                "db_name_cn": db_row['program_name_cn'],
                "official_name": "",
                "faculty": "",
                "db_url": db_row['program_website'],
                "official_url": "",
                "db_is_valid": db_row.get('is_valid', '')
            })

    all_records = matched + unmatched_off + unmatched_db
    df_diff = pd.DataFrame(all_records)
    df_diff.to_csv(out_file, index=False, encoding='utf-8-sig')

    print(f"\n[OK] 大盘比对完成，产出文件: {out_file}")
    print(f"    - 成功匹配: {len(matched)} 项")
    print(f"    - 官网独有 (疑似新增/漏录): {len(unmatched_off)} 项")
    print(f"    - 数据库独有 (疑似停办/更名): {len(unmatched_db)} 项")

    return df_diff

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="步骤3：大盘名单差集比对")
    parser.add_argument("--school", default="hku", help="院校代码，如 hku")
    args = parser.parse_args()

    diff_program_names(args.school)
