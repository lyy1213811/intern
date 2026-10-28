r"""
香港中文大学 (CUHK) 步骤五：最终交付件生成脚本
产出：c:\Users\rtx\Desktop\CUHK_数据核对与差异反馈表.xlsx
包含：
Sheet 1: 香港中文大学基准清单(149项TPg)
Sheet 2: 数据核对与差异反馈表 (阶段一漏录 + 阶段二在招字段差异核验)
"""

import os
import re
import sys
import urllib.parse
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data", "cuhk")

db_slice_file = os.path.join(DATA_DIR, "db_slice.csv")
diff_file = os.path.join(DATA_DIR, "diff_programs.csv")
html_dir = os.path.join(DATA_DIR, "htmls")
out_desktop_file = r"c:\Users\rtx\Desktop\CUHK_数据核对与差异反馈表.xlsx"
out_data_file = os.path.join(DATA_DIR, "CUHK_数据核对与差异反馈表.xlsx")

# 样式定义
FONT_FAMILY = "微软雅黑"
HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
HEADER_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
STAGE1_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # 淡橙高亮阶段一（漏录）
STAGE2_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # 淡黄高亮阶段二（字段核验）
INACTIVE_FILL = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")

REGULAR_FONT = Font(name=FONT_FAMILY, size=9)
BOLD_FONT = Font(name=FONT_FAMILY, size=9, bold=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

def get_html_path(url):
    if not url or pd.isna(url): return None
    p = urllib.parse.urlparse(str(url))
    slug = p.fragment if p.fragment else os.path.basename(p.path.rstrip('/'))
    safe_slug = re.sub(r'[\\/*?:"<>|]', '_', slug).strip().lower()
    for f in os.listdir(html_dir):
        if not f.endswith('.html'): continue
        f_stem = f[:-5].lower()
        if f_stem == safe_slug or safe_slug.replace('-', '_') == f_stem.replace('-', '_'):
            return os.path.join(html_dir, f)
    return None

def parse_cuhk_html(fpath):
    if not fpath or not os.path.exists(fpath): return {}
    with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')
    details = {}
    for t_box in soup.find_all('div', class_='title-bg'):
        title = t_box.get_text(strip=True)
        col4 = t_box.find_parent('div', class_='col-12')
        if col4:
            col8 = col4.find_next_sibling('div', class_='col-12')
            if col8:
                details[title] = col8.get_text('\n', strip=True)
    return details

def extract_years_from_duration(duration_str):
    if not duration_str or pd.isna(duration_str):
        return 1.0
    s = str(duration_str).strip()
    m_yr = re.search(r'(\d+(?:\.\d+)?)\s*年', s)
    if m_yr:
        return float(m_yr.group(1))
    m_mo = re.search(r'(\d+)\s*个?月', s)
    if m_mo:
        return float(m_mo.group(1)) / 12.0
    return 1.0

def extract_cuhk_self_financed_fee(tuition_text):
    if not tuition_text or pd.isna(tuition_text):
        return []
    lines = tuition_text.split('\n')
    self_finance_fees = []
    all_fees = []
    for line in lines:
        fees = [int(x.replace(',', '')) for x in re.findall(r'HK\$\s*([\d,]+)', line)]
        if not fees:
            continue
        all_fees.extend(fees)
        line_lower = line.lower()
        if 'local' in line_lower and 'non-local' not in line_lower and 'self-finance' not in line_lower:
            continue
        if any(f in [42100, 44500, 47000, 49500] for f in fees) and 'self-finance' not in line_lower:
            continue
        self_finance_fees.extend(fees)
    return self_finance_fees if self_finance_fees else all_fees

def generate_cuhk_deliverables():
    print("[*] 开始生成香港中文大学 (CUHK) 最终质检成果表...")
    df_slice = pd.read_csv(db_slice_file)
    df_diff = pd.read_csv(diff_file)

    wb = openpyxl.Workbook()

    # ==========================================
    # Sheet 1: 香港中文大学基准清单(149项TPg)
    # ==========================================
    ws1 = wb.active
    ws1.title = "香港中文大学基准清单(149项TPg)"
    ws1.views.sheetView[0].showGridLines = True
    ws1.row_dimensions[1].height = 28

    headers1 = [
        "序号", "数据库ID", "所属学院", "院校中英文", "专业中英文 (专业名/学位)",
        "学制", "数据库现有学费", "数据库现有申请时间", "数据库现有语言要求", "官网URL"
    ]
    ws1.append(headers1)
    for col_idx in range(1, len(headers1) + 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER

    # 从 diff 表提取 faculty 映射
    fac_map = {}
    for _, r in df_diff.iterrows():
        if pd.notna(r.get('db_id')) and r.get('db_id') != '':
            try:
                fac_map[int(r['db_id'])] = str(r.get('faculty', '') or '')
            except:
                pass

    for idx, r in df_slice.iterrows():
        row_num = idx + 1
        db_id = r['id']
        school_str = f"{r['school_name_cn']} {r['school_name_en']}"
        prog_str = f"{r['program_name_en']} ({r['program_name_cn']})"
        faculty = fac_map.get(db_id, "研究生院/深圳校区" if "cuhk-shenzhen" in prog_str.lower() or "深圳" in prog_str else "研究生院/未分区")
        duration = str(r.get('program_duration', '') or '')
        tuition = str(r.get('program_tuition', '') or '')
        app_time = str(r.get('application_time', '') or '').replace('<br/>', '\n').strip()
        lang = str(r.get('language_requirements', '') or '').replace('<br/>', '\n').strip()
        url = str(r.get('program_website', '') or '')
        is_valid = r.get('is_valid', True)

        row_vals = [row_num, db_id, faculty, school_str, prog_str, duration, tuition, app_time, lang, url]
        ws1.append(row_vals)

        current_row = ws1.max_row
        ws1.row_dimensions[current_row].height = 22
        for col_idx in range(1, len(headers1) + 1):
            cell = ws1.cell(row=current_row, column=col_idx)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            if col_idx in [1, 2, 6]:
                cell.alignment = ALIGN_CENTER
            elif col_idx in [3, 4, 5, 7, 8, 9, 10]:
                cell.alignment = ALIGN_LEFT

            if is_valid is False or str(is_valid).lower() == 'false':
                cell.fill = INACTIVE_FILL

    # Sheet 1 列宽
    for col in ws1.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            line_len = sum(2 if ord(c) > 127 else 1 for c in val_str.split('\n')[0])
            max_len = max(max_len, line_len)
        ws1.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 45)
    ws1.auto_filter.ref = f"A1:J{ws1.max_row}"

    # ==========================================
    # Sheet 2: 数据核对与差异反馈表
    # ==========================================
    ws2 = wb.create_sheet(title="数据核对与差异反馈表")
    ws2.views.sheetView[0].showGridLines = True
    ws2.row_dimensions[1].height = 28

    headers2 = [
        "序号", "院校中英文", "专业中英文", "官网URL", "数据库ID",
        "报错类型", "数据库现有值", "官网真实值", "官网原文依据", "备注", "审批意见"
    ]
    ws2.append(headers2)
    for col_idx in range(1, len(headers2) + 1):
        cell = ws2.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER

    feedback_rows = []

    # ----------------------------------------------------
    # 阶段一：全量漏录专业 (来自步骤3官网独有，共35项)
    # ----------------------------------------------------
    unmatched_off = df_diff[df_diff['category'].str.contains('官网独有', na=False)]
    for idx, r in unmatched_off.iterrows():
        off_name = r['official_name']
        off_url = r['official_url']
        faculty = r['faculty']
        
        note_text = "【阶段一确凿漏录】官网 2027/28 正式公布在招课程，系统数据库暂未收录，建议新增建档"
        if off_name in ['Juris Doctor', 'MBA']:
            note_text = f"【核心在招漏录】官网 2027/28 正式在招课程，数据库历史代码误标为失效，建议重新激活建档"
        elif 'Educational Psychology' in off_name or 'Gerontology' in off_name:
            note_text = f"【在招专业漏录】官网 2027/28 正式在招课程，系统库中旧记录已失效且未新建有效档案，建议建档补录"

        feedback_rows.append({
            "stage": 1,
            "school": "香港中文大学 (The Chinese University of Hong Kong)",
            "prog_name": off_name,
            "url": off_url,
            "db_id": "无/新增",
            "err_type": "漏录",
            "db_val": "系统数据库中查无此在招课程",
            "off_val": f"2027/28 官方公布在招授课型硕士 (所属院系: {faculty})",
            "quote": f"CUHK Graduate School 官方招生列表: {off_name}",
            "note": note_text,
            "opinion": "待审批"
        })

    # ----------------------------------------------------
    # 阶段二：在招专业字段核验 (严格遵守导师6大红线)
    # ----------------------------------------------------
    matched_df = df_diff[df_diff['category'].str.contains('匹配', na=False)]
    for idx, r in matched_df.iterrows():
        db_id = int(r['db_id'])
        db_slice_row = df_slice[df_slice['id'] == db_id].iloc[0]
        
        # 红线一：绝对不核查 is_valid == False 的专业
        if str(db_slice_row.get('is_valid', '')).lower() != 'true':
            continue

        url = r['official_url']
        hpath = get_html_path(url)
        details = parse_cuhk_html(hpath)
        if not details:
            continue

        prog_full_name = f"{r['db_name_en']} ({r['db_name_cn']})"
        db_duration = str(db_slice_row.get('program_duration', '') or '')
        db_tuition = str(db_slice_row.get('program_tuition', '') or '')
        db_app_time = str(db_slice_row.get('application_time', '') or '').replace('<br/>', '\n').strip()
        
        off_tuition = details.get('Tuition Fee', '').strip()
        off_deadline = details.get('Application Deadline', '').strip()

        # 1. 核查学费 (严格锁定 Non-local 自费学费，严格排除 UGC 本地生资助与多年制年化折算)
        years = extract_years_from_duration(db_duration)
        db_nums = [int(x.replace(',', '')) for x in re.findall(r'(\d+[\d,]*)', db_tuition)]
        off_self_fees = extract_cuhk_self_financed_fee(off_tuition)

        if off_self_fees and db_nums:
            db_val = db_nums[0]
            match_found = False
            for fee in off_self_fees:
                if fee == db_val:
                    match_found = True; break
                # 年化折算匹配 (红线三)
                if years > 0 and abs(round(fee / years) - db_val) <= 10:
                    match_found = True; break
                if abs(fee * years - db_val) <= 10:
                    match_found = True; break
                if abs(round(fee / 2.0) - db_val) <= 10 or abs(round(fee / 1.5) - db_val) <= 10:
                    match_found = True; break

            if not match_found:
                # 确凿学费差异 (红线二、三严格排查后)
                raw_fee = off_self_fees[0]
                feedback_rows.append({
                    "stage": 2,
                    "school": "香港中文大学 (The Chinese University of Hong Kong)",
                    "prog_name": prog_full_name,
                    "url": url,
                    "db_id": db_id,
                    "err_type": "门槛未更新",
                    "db_val": db_tuition,
                    "off_val": f"HKD {raw_fee:,} (整项目自费学费)",
                    "quote": off_tuition.replace('\n', ' ')[:140],
                    "note": f"【2027学年学费调整】官网最新公布非本地自费学费为 HK${raw_fee:,}，经学制折算与库中现有值仍存在确凿差额",
                    "opinion": "待审批"
                })

        # 2. 核查申请时间 (排查库中停留在过期旧年且官网已公布最新 2027/28 轮次)
        if off_deadline and not any(k in off_deadline.lower() for k in ['to be announced', 'tbc', 'remarks']):
            off_years = set(re.findall(r'202[6-8]', off_deadline))
            db_years = set(re.findall(r'202[6-8]', db_app_time))
            # 仅当官网包含 2027 年且库中完全没有 2027 年(停留在旧年份)时上报 (红线四)
            if '2027' in off_years and '2027' not in db_years and ('27年' not in db_app_time):
                deadlines_clean = ' | '.join([l.strip() for l in off_deadline.split('\n') if l.strip() and not l.strip().startswith('Applications')][:3])
                feedback_rows.append({
                    "stage": 2,
                    "school": "香港中文大学 (The Chinese University of Hong Kong)",
                    "prog_name": prog_full_name,
                    "url": url,
                    "db_id": db_id,
                    "err_type": "门槛未更新",
                    "db_val": db_app_time.replace('\n', ' ')[:100],
                    "off_val": deadlines_clean,
                    "quote": off_deadline.replace('\n', ' ')[:140],
                    "note": "【申请时间停留在旧年】库中记录停留在 2026 或更早旧轮次，官网已公布 2027/28 学年最新截止日",
                    "opinion": "待审批"
                })

    # 填充 Sheet 2
    for seq, r_data in enumerate(feedback_rows, 1):
        row_vals = [
            seq, r_data['school'], r_data['prog_name'], r_data['url'], r_data['db_id'],
            r_data['err_type'], r_data['db_val'], r_data['off_val'], r_data['quote'],
            r_data['note'], r_data['opinion']
        ]
        ws2.append(row_vals)
        cur_row = ws2.max_row
        ws2.row_dimensions[cur_row].height = 24
        fill = STAGE1_FILL if r_data['stage'] == 1 else STAGE2_FILL

        for c_idx in range(1, len(headers2) + 1):
            cell = ws2.cell(row=cur_row, column=c_idx)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            cell.fill = fill
            if c_idx in [1, 5, 6, 11]:
                cell.alignment = ALIGN_CENTER
            else:
                cell.alignment = ALIGN_LEFT

    # Sheet 2 列宽与筛选
    for col in ws2.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            line_len = sum(2 if ord(c) > 127 else 1 for c in val_str.split('\n')[0])
            max_len = max(max_len, line_len)
        ws2.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 45)
    ws2.auto_filter.ref = f"A1:K{ws2.max_row}"

    # 保存
    wb.save(out_desktop_file)
    wb.save(out_data_file)
    print(f"\n[OK] 香港中文大学最终成果表导出完成！")
    print(f"    -> 桌面交付件: {out_desktop_file}")
    print(f"    -> 仓库备份件: {out_data_file}")
    print(f"    - Sheet 1 基准清单记录: {ws1.max_row - 1} 条")
    print(f"    - Sheet 2 质检反馈记录: {ws2.max_row - 1} 条 (阶段一确凿漏录: {len(feedback_rows)} 条)")

if __name__ == "__main__":
    generate_cuhk_deliverables()
