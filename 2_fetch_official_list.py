import os
import sys
import shutil
import argparse
from datetime import datetime
import requests
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# 过滤非授课型硕士学位（剔除博士、本科、专科文凭等）
EXCLUDE_ABBRS = {
    'EdD', 'DNurs', 'PsyD(CP)', 'PsyD(EP)', 'DBA', 'DPA',
    'PCLL', 'PGDE', 'PGDE(TCSL)', 'AdvDipOrth', 'PDipID', 'PDipPsyMed'
}

EXCLUDE_NAME_KEYWORDS = [
    'doctor of', 
    'postgraduate diploma', 
    'postgraduate certificate', 
    'advanced diploma'
]

def fetch_hku_official_list(school="hku"):
    out_dir = os.path.join("data", school.lower())
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, "official_list.csv")

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Referer': 'https://portal.hku.hk/tpg-admissions/programme-listing'
    }

    # HKU TPG Portal 官方数据 API 接口
    # 覆盖全日制与非全日制、非本地生与本地生完整盘口
    endpoints = [
        # Full-time non-local & without fee
        'https://portal.hku.hk/SavedQueryService/Execute/any-faculty-programmes/%25/629270000/1/10/hkutgp_feefornonlocalstudent/0/1000000',
        'https://portal.hku.hk/SavedQueryService/Execute/any-faculty-programmes-without-fee/%25/629270000/1/10',
        # Part-time non-local & without fee
        'https://portal.hku.hk/SavedQueryService/Execute/any-faculty-programmes/%25/629270001/1/10/hkutgp_feefornonlocalstudent/0/1000000',
        'https://portal.hku.hk/SavedQueryService/Execute/any-faculty-programmes-without-fee/%25/629270001/1/10',
        # Local queries (确保覆盖仅招本地生的特殊项目)
        'https://portal.hku.hk/SavedQueryService/Execute/any-faculty-programmes/%25/629270000/1/10/hkutgp_feeforlocalstudent/0/1000000',
        'https://portal.hku.hk/SavedQueryService/Execute/any-faculty-programmes/%25/629270001/1/10/hkutgp_feeforlocalstudent/0/1000000',
    ]

    print("[*] 正在从香港大学官方招生系统抓取 2027/28 授课型硕士 (TPg) 列表...")
    raw_items = {}
    for ep in endpoints:
        try:
            r = requests.get(ep, headers=headers, timeout=25)
            r.raise_for_status()
            data = r.json()
            items = data.get('data', []) if isinstance(data, dict) else (data[0].get('data', []) if data else [])
            for item in items:
                raw_items[item['Id']] = item
        except Exception as e:
            print(f"[-] 请求接口异常 ({ep}): {e}")

    print(f"[*] 官网接口返回原始项目总数（含博士/文凭等）: {len(raw_items)}")

    # 规范化并清洗过滤为授课型硕士清单
    records = []
    for prog_id, item in raw_items.items():
        attrs = item.get('Attributes', {})
        formatted = item.get('FormattedValues', {})

        abbr = attrs.get('hkutgp_programmenameabbreviation', '').strip()
        name = attrs.get('hkutgp_name', '').strip()
        faculty = formatted.get('hkutgp_facultyid', '')

        # 剔除博士学位与文凭类
        if abbr in EXCLUDE_ABBRS:
            continue
        name_lower = name.lower()
        if any(kw in name_lower for kw in EXCLUDE_NAME_KEYWORDS):
            continue

        # 提取专业详情页 Slug
        root_val = attrs.get('rootprogramme.hkutgp_name')
        slug = ""
        if isinstance(root_val, dict):
            slug = root_val.get('Value', '')
        elif hasattr(root_val, 'Value'):
            slug = root_val.Value

        official_url = f"https://portal.hku.hk/tpg-admissions/programme-details?programme={slug}" if slug else ""

        records.append({
            "official_program_name": name,
            "faculty": faculty,
            "official_url": official_url
        })

    df = pd.DataFrame(records)
    df = df.drop_duplicates(subset=['official_program_name', 'official_url'])
    df = df.sort_values(by=['faculty', 'official_program_name']).reset_index(drop=True)

    # 自动备份已有文件，防止误覆盖历史数据
    if os.path.exists(out_file):
        backup_dir = os.path.join(out_dir, "backup")
        os.makedirs(backup_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(backup_dir, f"official_list_{timestamp}.csv")
        shutil.copy2(out_file, backup_path)
        print(f"[*] 检测到已存在历史文件，已自动备份至: {backup_path}")

    df.to_csv(out_file, index=False, encoding='utf-8-sig')
    print(f"[OK] 成功抓取官网在招授课型硕士 (TPg) 共 {len(df)} 门专业！")
    print(f"[OK] 产出文件保存至: {out_file}")

    return df, out_file

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="步骤2：抓取指定院校官网在招专业列表")
    parser.add_argument("--school", default="hku", help="院校代码，如 hku")
    args = parser.parse_args()

    if args.school.lower() == "hku":
        fetch_hku_official_list(args.school)
    else:
        print(f"当前暂未适配院校 {args.school}，已默认执行 hku")
        fetch_hku_official_list("hku")
