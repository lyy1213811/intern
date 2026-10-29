import os
import sys
import re
import argparse
import difflib
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from urllib.parse import urlparse

def normalize_name(name):
    if not isinstance(name, str):
        return ""
    # 去除星号与特殊标记
    s = name.replace('*', '').lower()
    # 统一常见学位缩写与变体
    s = re.sub(r'\bm\.sc\.?\b', 'msc', s)
    s = re.sub(r'\bm\.a\.?\b', 'ma', s)
    s = re.sub(r'\bm\.ed\.?\b', 'med', s)
    s = re.sub(r'\bm\.eng\.?\b', 'meng', s)
    s = re.sub(r'[\(\)\[\]\,\.\-\/\:\'\"]', ' ', s)
    s = re.sub(r'\bm\s+sc\b', 'msc', s)
    s = re.sub(r'\bm\s+a\b', 'ma', s)
    s = re.sub(r'\bm\s+eng\b', 'meng', s)
    s = re.sub(r'\bmaster of science and master of engineering\b', 'msc meng', s)
    s = re.sub(r'\bmaster of science\b', 'msc', s)
    s = re.sub(r'\bmaster of arts\b', 'ma', s)
    s = re.sub(r'\bmaster of engineering\b', 'meng', s)
    s = re.sub(r'\bmaster of public policy\b', 'mpp', s)
    s = re.sub(r'\bmaster of public management\b', 'mpm', s)
    s = re.sub(r'\bmaster of business administration\b', 'mba', s)
    s = re.sub(r'\bpostgraduate diploma and\b', '', s)
    s = re.sub(r'\bpostgraduate diploma\b', '', s)
    s = re.sub(r'\bprograms?\b', '', s)
    s = re.sub(r'\bin\b', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def get_url_slug(url):
    if not isinstance(url, str):
        return ''
    p = urlparse(url)
    import urllib.parse
    q = urllib.parse.parse_qs(p.query)
    if 'programme' in q and q['programme']:
        return q['programme'][0].strip()
    if p.fragment:
        return p.fragment.strip()
    base = os.path.basename(p.path.rstrip('/')).strip()
    if base.lower() not in ['programme-details', 'programme', 'programmes', 'admissions', 'index.html', 'index', '']:
        return base
    return ''

def is_independent_campus(school, db_row):
    school_lower = school.lower()
    cn = str(db_row.get('program_name_cn', '') or '')
    en = str(db_row.get('program_name_en', '') or '')
    url = str(db_row.get('program_website', '') or '')
    if school_lower == 'cuhk':
        if ('cuhk-shenzhen' in cn.lower() or '深圳' in cn) or ('cuhk.edu.cn' in url.lower()):
            return "数据库独有(深圳校区/独立招生)"
    elif school_lower == 'hkust':
        if re.search(r'guangzhou|\(gz\)', en, re.I) or '广州' in cn or 'hkust-gz' in url.lower() or 'cft.hkust-gz' in url.lower():
            return "数据库独有(广州校区/独立招生)"
    return None

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
    # 杜绝失效小方向提前抢占在招大类的官网名额
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

    # 1. 优先轮次一：规范化名称完全相同 (仅限 is_valid == True 在招专业参与匹配，并排除异地校区)
    for db_idx, db_row in df_db.iterrows():
        if str(db_row.get('is_valid', '')).lower() != 'true':
            continue
        if is_independent_campus(school, db_row):
            continue
        db_norm = db_row['norm_name']
        for off_idx, off_row in df_off.iterrows():
            if off_idx in off_matched_indices:
                continue
            if db_norm == off_row['norm_name']:
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
                break

    # 2. 优先轮次二：URL完全吻合 或 URL Slug 精确吻合 (仅限 is_valid == True 在招专业参与匹配)
    for db_idx, db_row in df_db.iterrows():
        if db_idx in db_matched_indices:
            continue
        if str(db_row.get('is_valid', '')).lower() != 'true':
            continue
        if is_independent_campus(school, db_row):
            continue
        db_url = str(db_row.get('program_website', '')).strip().lower()
        db_slug = normalize_name(get_url_slug(db_url).replace('-', ' '))
        for off_idx, off_row in df_off.iterrows():
            if off_idx in off_matched_indices:
                continue
            off_url = str(off_row.get('official_url', '')).strip().lower()
            off_slug = normalize_name(get_url_slug(off_url).replace('-', ' '))
            
            url_match = bool(db_url and off_url and (db_url == off_url or db_url in off_url or off_url in db_url))
            slug_match = bool(db_slug and off_slug and len(db_slug) > 3 and (db_slug == off_slug or db_slug in off_slug or off_slug in db_slug))
            
            if url_match or slug_match:
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
                break

    # 3. 优先轮次三：针对剩余未匹配项，执行模糊匹配 (阈值 0.80) (仅限 is_valid == True 在招专业参与匹配)
    for db_idx, db_row in df_db.iterrows():
        if db_idx in db_matched_indices:
            continue
        if str(db_row.get('is_valid', '')).lower() != 'true':
            continue
        if is_independent_campus(school, db_row):
            continue
        db_norm = db_row['norm_name']
        best_ratio = 0.0
        best_off_idx = None

        for off_idx, off_row in df_off.iterrows():
            if off_idx in off_matched_indices:
                continue
            off_norm = off_row['norm_name']
            
            # 避免 MBA 与非 MBA 产生歧义误配
            if ('\bmba\b' in db_norm and '\bmba\b' not in off_norm) or ('\bmba\b' not in db_norm and '\bmba\b' in off_norm):
                continue
                
            ratio = difflib.SequenceMatcher(None, db_norm, off_norm).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_off_idx = off_idx

        if best_ratio >= 0.80 and best_off_idx is not None:
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

    # 4. 收集官网独有集（智能区分兼读制免录、历史休眠重启在招与全日制真正漏录）
    html_dir = os.path.join(data_dir, "htmls")
    for off_idx, off_row in df_off.iterrows():
        if off_idx not in off_matched_indices:
            off_name = off_row['official_program_name']
            off_url = str(off_row.get('official_url', '')).strip()
            slug = get_url_slug(off_url).lower()

            # 1. 检查是否为兼读制 / 在职高管项目 (按导师红线 1：严禁报漏录)
            is_pt = False
            if any(k in off_name.lower() for k in ['part-time', 'emba', 'executive master']) or \
               any(k in off_url.lower() for k in ['part-time', 'pt-', 'emba', 'msc-gfin']):
                is_pt = True
            elif os.path.exists(html_dir):
                # 从已下载快照中核查 Mode of Study
                for fname in os.listdir(html_dir):
                    if slug and slug in fname.lower() and fname.endswith('.html'):
                        try:
                            with open(os.path.join(html_dir, fname), 'r', encoding='utf-8', errors='ignore') as f:
                                h_txt = f.read().lower()
                                if 'mode of studypart-time' in h_txt.replace(' ', '') or 'part-time only' in h_txt:
                                    is_pt = True
                                    break
                        except Exception:
                            pass

            # 2. 检查是否匹配数据库中已失效的历史专业 (休眠重启)
            is_resumed = False
            resumed_db_id = ""
            for _, db_r in df_db[df_db['is_valid'].astype(str).str.lower() == 'false'].iterrows():
                db_u = str(db_r.get('program_website', '')).lower()
                db_s = get_url_slug(db_u).lower()
                if (slug and db_s and slug == db_s) or (slug and slug in db_u):
                    is_resumed = True
                    resumed_db_id = str(db_r['id'])
                    break

            if is_pt:
                cat = "官网独有(兼读制/在职高管免录)"
            elif is_resumed:
                cat = "官网独有(历史休眠专业2027/28重启在招)"
            else:
                cat = "官网独有(疑似新增/漏录)"

            unmatched_off.append({
                "category": cat,
                "match_score": 0.0,
                "db_id": resumed_db_id,
                "db_name_en": "",
                "db_name_cn": "",
                "official_name": off_name,
                "faculty": off_row['faculty'],
                "db_url": "",
                "official_url": off_url,
                "db_is_valid": ""
            })

    # 5. 收集数据库独有集（智能区分异地独立校区与真正失效停办）
    for db_idx, db_row in df_db.iterrows():
        if db_idx not in db_matched_indices:
            is_valid_str = str(db_row.get('is_valid', '')).lower()

            if is_valid_str == 'false':
                cat = "数据库独有(已失效/历史归档)"
            else:
                campus_cat = is_independent_campus(school, db_row)
                if campus_cat:
                    cat = campus_cat
                else:
                    cat = "数据库独有(疑似停办/下线)"

            unmatched_db.append({
                "category": cat,
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
    print(f"    - 数据库独有: {len(unmatched_db)} 项 (含已失效及独立校区)")

    return df_diff

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="步骤3：大盘名单差集比对")
    parser.add_argument("--school", default="hku", help="院校代码，如 hku, cuhk, hkust")
    args = parser.parse_args()

    diff_program_names(args.school)
