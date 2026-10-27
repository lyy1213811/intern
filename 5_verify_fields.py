"""
步骤 5：核心字段核查脚本 (5_verify_fields.py) - SOP 工业级完整标准版
严格遵循最新发布的《院校数据质检与审核自动化工作流 SOP》：
阶段一：任务 1 —— 专业数量核对（查漏录、标停招）【前置优先执行】
阶段二：任务 2 —— 高价值核心字段核查（门槛未更新、小分有误、其他异常）【深度执行】
报错类型严格限制为：[漏录, 停招, 小分有误, 门槛未更新, 其他异常]
产出物：data/{school}/HKU_数据核对与差异反馈表.xlsx
"""

import os
import re
import sys
import copy
import urllib.parse
from datetime import datetime
import bs4
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# 样式定义
HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
STAGE1_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # 橙色系高亮阶段一（新增/停招）
STAGE2_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # 黄色系高亮阶段二（字段差异）
HIST_FILL = PatternFill(start_color="EDF2F8", end_color="EDF2F8", fill_type="solid")   # 历史批复行

REGULAR_FONT = Font(name="微软雅黑", size=9)
BOLD_FONT = Font(name="微软雅黑", size=9, bold=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

ALLOWED_ERROR_TYPES = ["漏录", "停招", "小分有误", "门槛未更新", "其他异常"]

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()

def get_html_path(url, html_dir):
    if not url or pd.isna(url) or not isinstance(url, str):
        return None
    parsed = urllib.parse.urlparse(url)
    slug = urllib.parse.parse_qs(parsed.query).get('programme', [''])[0]
    if not slug:
        slug = os.path.basename(parsed.path.rstrip('/'))
    safe_slug = sanitize_filename(slug)
    
    direct = os.path.join(html_dir, f"{safe_slug}.html")
    if os.path.exists(direct):
        return direct
        
    for f in os.listdir(html_dir):
        f_lower = f.lower()
        slug_lower = safe_slug.lower()
        if f_lower == f"{slug_lower}.html":
            return os.path.join(html_dir, f)
        if slug_lower.replace('-', '_') in f_lower.replace('-', '_'):
            return os.path.join(html_dir, f)
            
    return None

def extract_from_html(filepath):
    if not filepath or not os.path.exists(filepath):
        return None

    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        html_content = f.read()

    soup = bs4.BeautifulSoup(html_content, 'html.parser')
    data = {
        'duration': '',
        'fees_raw': '',
        'deadlines_raw': '',
        'ielts_raw': '',
        'toefl_raw': ''
    }

    for item in soup.find_all('div', class_='highlights-item'):
        t_el = item.find('div', class_='highlights-item-title')
        d_el = item.find('div', class_='highlights-item-description')
        if not t_el or not d_el: continue
        title = t_el.get_text(strip=True).lower()
        desc = d_el.get_text(' ', strip=True)

        if 'duration' in title and desc and desc != 'To be updated':
            if not data['duration']: data['duration'] = desc
        elif 'fee' in title and desc and desc != 'To be updated':
            if not data['fees_raw']: data['fees_raw'] = desc
        elif 'deadline' in title and desc and desc != 'To be updated':
            if not data['deadlines_raw']: data['deadlines_raw'] = desc

    req_sec = soup.find('div', class_=lambda c: c and 'section-entrance-requirements' in c)
    if not req_sec:
        for h in soup.find_all(['h2', 'h3', 'h4']):
            if 'Entrance Requirements' in h.get_text():
                req_sec = h.find_parent('div')
                break

    if req_sec:
        full_req_text = req_sec.get_text('\n', strip=True)
        ielts_match = re.search(r'(?:IELTS[^\.\;\n]{0,150}(?:band|score|overall)[^\.\;\n]{0,150})', full_req_text, re.I)
        if ielts_match:
            data['ielts_raw'] = ielts_match.group(0).strip()
        toefl_match = re.search(r'(?:TOEFL[^\.\;\n]{0,150}(?:score)?[^\.\;\n]{0,100})', full_req_text, re.I)
        if toefl_match:
            data['toefl_raw'] = toefl_match.group(0).strip()

    return data

def parse_fee_number(fee_str):
    if not fee_str or pd.isna(fee_str): return None
    nums = re.findall(r'(?:HK\$|HKD|RMB|￥)?\s*([0-9]{2,3}(?:,[0-9]{3})+|[0-9]{5,7})', str(fee_str))
    return int(nums[0].replace(',', '')) if nums else None

def parse_duration_years(dur_str):
    if not dur_str or pd.isna(dur_str): return 1.0
    dur_str = str(dur_str).lower()
    m = re.search(r'([0-9]+(?:\.[0-9]+)?)\s*year', dur_str)
    if m: return float(m.group(1))
    m_cn = re.search(r'([0-9]+(?:\.[0-9]+)?)\s*年', dur_str)
    if m_cn: return float(m_cn.group(1))
    return 1.0

def build_sop_feedback_table(school="hku"):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, "data", school.lower())
    html_dir = os.path.join(data_dir, "htmls")
    diff_file = os.path.join(data_dir, "diff_programs.csv")
    slice_file = os.path.join(data_dir, "db_slice.csv")
    desktop_dir = r"c:\Users\rtx\Desktop"
    xlsx_files = [os.path.join(desktop_dir, f) for f in os.listdir(desktop_dir) if f.endswith('.xlsx')]
    if xlsx_files:
        template_file = xlsx_files[0]
    else:
        template_file = os.path.join(data_dir, f"{school.upper()}_数据核对与差异反馈表.xlsx")
    
    desktop_out_file = os.path.join(desktop_dir, f"{school.upper()}_数据核对与差异反馈表.xlsx")
    data_out_file = os.path.join(data_dir, f"{school.upper()}_数据核对与差异反馈表.xlsx")

    print(f"[*] 启动符合最新 SOP 填表标准的质检流水线 [{school.upper()}]...")

    # 读取底表数据
    wb_orig = openpyxl.load_workbook(template_file)
    ws_base = wb_orig['港大在招基准清单(185项TPg)']
    ws_feedback_orig = wb_orig['数据核对与差异反馈表']

    df_slice = pd.read_csv(slice_file) if os.path.exists(slice_file) else pd.DataFrame()
    slice_map = {int(r['id']): r.to_dict() for _, r in df_slice.iterrows()} if not df_slice.empty else {}
    
    df_diff = pd.read_csv(diff_file) if os.path.exists(diff_file) else pd.DataFrame()

    # 创建新的标准输出工作簿
    wb = openpyxl.Workbook()
    # Sheet 1 保留基准底表
    ws_sheet1 = wb.active
    ws_sheet1.title = "港大在招基准清单(185项TPg)"
    for r in range(1, ws_base.max_row + 1):
        for c in range(1, ws_base.max_column + 1):
            val = ws_base.cell(r, c).value
            ws_sheet1.cell(r, c, val)

    # Sheet 2 标准质检反馈表
    ws_sheet2 = wb.create_sheet(title="数据核对与差异反馈表")
    headers = ['序号', '院校中英文', '专业中英文', '官网URL', '数据库ID', '报错类型', '数据库现有值', '官网真实值', '官网原文依据', '备注', '审批意见']
    for c, h in enumerate(headers, 1):
        cell = ws_sheet2.cell(1, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER

    # 1. 保留前 14 行历史审批（并将报错类型归一化至 SOP 标准 5 类）
    type_remap = {
        '申请时间有误': '其他异常',
        '学费有误': '门槛未更新',
        '语言要求有误': '小分有误'
    }
    cur_row = 2
    for r in range(2, 16):
        vals = [ws_feedback_orig.cell(r, c).value for c in range(1, 12)]
        if not vals[0]: continue
        err_type = str(vals[5])
        vals[5] = type_remap.get(err_type, err_type)
        if vals[5] not in ALLOWED_ERROR_TYPES: vals[5] = '其他异常'
        
        for c, v in enumerate(vals, 1):
            cell = ws_sheet2.cell(cur_row, c, v)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            cell.fill = HIST_FILL
            cell.alignment = ALIGN_CENTER if c in [1, 5, 6, 11] else ALIGN_LEFT
        cur_row += 1

    print(f"[*] 已同步保留前 14 条历史样例审批记录。")

    # ---------------- 阶段一：任务 1 —— 专业数量核对（查漏录、标停招）----------------
    print(f"[*] 正在执行【阶段一：任务 1 —— 专业数量核对】...")
    stage1_records = []
    
    if not df_diff.empty:
        # 1. 漏录（官网有，数据库无）
        off_only = df_diff[df_diff['category'].str.contains('官网独有')]
        for _, r in off_only.iterrows():
            stage1_records.append({
                '院校中英文': '香港大学 (The University of Hong Kong)',
                '专业中英文': r['official_name'],
                '官网URL': r['official_url'],
                '数据库ID': '无/新增',
                '报错类型': '漏录',
                '数据库现有值': '系统数据库中查无此在招课程',
                '官网真实值': f"2027/28 官方公布在招授课型硕士 (所属院系: {r.get('faculty', '')})",
                '官网原文依据': f"HKU TPG Portal 官方招生列表: {r['official_name']} ({r['official_url']})",
                '备注': '【阶段一漏录核查】官网 2027/28 学年正式开设的授课型硕士，数据库暂未收录，建议新增建档',
                '审批意见': '待审批'
            })

        # 2. 停招（数据库有，官网无）
        db_only = df_diff[df_diff['category'].str.contains('数据库独有')]
        for _, r in db_only.iterrows():
            db_id_val = int(r['db_id']) if pd.notna(r['db_id']) else ''
            db_name = f"{r['db_name_en']} ({r['db_name_cn']})"
            stage1_records.append({
                '院校中英文': '香港大学 (The University of Hong Kong)',
                '专业中英文': db_name,
                '官网URL': str(r.get('db_url', 'https://portal.hku.hk/tpg-admissions/programme-listing')),
                '数据库ID': db_id_val,
                '报错类型': '停招',
                '数据库现有值': f"数据库仍在案专业 (ID: {db_id_val})",
                '官网真实值': '官网 2027/28 招生列表未收录 / 停办或已并入大类',
                '官网原文依据': 'Official 2027/28 TPG Programme Listing: 项目已下线或未列入招录清单',
                '备注': '【阶段一停招核查】经与官网 27/28 A-Z 清单核对，该项目已停招或合并，建议系统及时下线',
                '审批意见': '待审批'
            })

    print(f"[*] 阶段一完成：登记【漏录】{len(off_only)} 条，【停招】{len(db_only)} 条。")

    # 写入阶段一记录
    for item in stage1_records:
        ws_sheet2.cell(cur_row, 1, cur_row - 1).alignment = ALIGN_CENTER
        ws_sheet2.cell(cur_row, 2, item['院校中英文']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 3, item['专业中英文']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 4, item['官网URL']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 5, item['数据库ID']).alignment = ALIGN_CENTER
        ws_sheet2.cell(cur_row, 6, item['报错类型']).alignment = ALIGN_CENTER
        ws_sheet2.cell(cur_row, 7, item['数据库现有值']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 8, item['官网真实值']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 9, item['官网原文依据']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 10, item['备注']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 11, item['审批意见']).alignment = ALIGN_CENTER

        for c in range(1, 12):
            cell = ws_sheet2.cell(cur_row, c)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            cell.fill = STAGE1_FILL
        cur_row += 1

    # ---------------- 阶段二：任务 2 —— 高价值核心字段核查 ----------------
    print(f"[*] 正在执行【阶段二：任务 2 —— 高价值核心字段核查】...")
    base_headers = [ws_base.cell(1, c).value for c in range(1, ws_base.max_column + 1)]
    base_items = []
    for r in range(2, ws_base.max_row + 1):
        row_dict = {base_headers[c - 1]: ws_base.cell(r, c).value for c in range(1, ws_base.max_column + 1)}
        base_items.append(row_dict)

    existing_keys = set()
    for r in range(2, cur_row):
        db_id = ws_sheet2.cell(r, 5).value
        err_type = ws_sheet2.cell(r, 6).value
        if db_id and err_type: existing_keys.add((str(db_id), str(err_type)))

    stage2_records = []
    for row in base_items:
        db_id = row.get('数据库ID')
        if not db_id: continue
        db_id = int(db_id)
        db_info = slice_map.get(db_id, {})
        # 严格遵守导师规则一：is_valid == False 的停办专业不进入字段核验（已在阶段一处理停招）
        if str(db_info.get('is_valid', True)).lower() in ['false', '0']:
            continue

        url = row.get('官网URL')
        fpath = get_html_path(url, html_dir)
        html_data = extract_from_html(fpath) if fpath else None
        if not html_data: continue

        prog_name = row['专业中英文 (专业名/学位)']
        school_name = '香港大学 (The University of Hong Kong)'
        db_tuition = str(row['数据库现有学费']) if pd.notna(row['数据库现有学费']) else ''
        db_app = str(row['数据库现有申请时间']) if pd.notna(row['数据库现有申请时间']) else ''
        db_lang = str(row['数据库现有语言要求']) if pd.notna(row['数据库现有语言要求']) else ''
        db_dur = str(row['学制']) if pd.notna(row['学制']) else ''

        off_fee = html_data['fees_raw']
        off_deadlines = html_data['deadlines_raw']
        off_dur = html_data['duration']
        off_ielts = html_data['ielts_raw']
        off_toefl = html_data['toefl_raw']

        # 1. 申请时间核对
        if off_deadlines and off_deadlines != 'To be updated':
            # 提前批错位标为 Round 1
            if 'early admission' in off_deadlines.lower() and ('early admission' not in db_app.lower() and '提前批' not in db_app):
                stage2_records.append({
                    '院校中英文': school_name,
                    '专业中英文': prog_name,
                    '官网URL': url,
                    '数据库ID': db_id,
                    '报错类型': '其他异常',
                    '数据库现有值': db_app,
                    '官网真实值': off_deadlines,
                    '官网原文依据': f"Application Deadline: {off_deadlines}",
                    '备注': '【轮次命名错位纠偏】官网明确设置提前批(Early Admission Round)，库中误将提前批标为Round 1，导致后续轮次命名错位',
                    '审批意见': '待审批'
                })
            # 门槛未更新：停留在 2025/2026 旧轮次
            elif '2027' in off_deadlines and ('2025' in db_app or '26年秋季' in db_app):
                stage2_records.append({
                    '院校中英文': school_name,
                    '专业中英文': prog_name,
                    '官网URL': url,
                    '数据库ID': db_id,
                    '报错类型': '门槛未更新',
                    '数据库现有值': db_app,
                    '官网真实值': off_deadlines,
                    '官网原文依据': f"Application Deadline: {off_deadlines}",
                    '备注': '【申请时间停留在旧年】库中记录停留在 2026 招生季旧轮次，官网已公布 2027/28 学年最新截止日',
                    '审批意见': '待审批'
                })

        # 2. 国际生学费核对
        if off_fee and off_fee != 'To be updated':
            off_fee_num = parse_fee_number(off_fee)
            db_fee_num = parse_fee_number(db_tuition)
            dur_years = parse_duration_years(off_dur if off_dur else db_dur)

            if off_fee_num and db_fee_num:
                is_local_fee_in_db = (db_fee_num in [48250, 42100])
                annual_off_fee = off_fee_num / dur_years if dur_years > 1 and off_fee_num > db_fee_num * 1.3 else off_fee_num
                diff_amount = abs(annual_off_fee - db_fee_num)

                # 本地生/国际生混淆
                if is_local_fee_in_db and off_fee_num > 100000:
                    stage2_records.append({
                        '院校中英文': school_name,
                        '专业中英文': prog_name,
                        '官网URL': url,
                        '数据库ID': db_id,
                        '报错类型': '其他异常',
                        '数据库现有值': db_tuition,
                        '官网真实值': f"HKD {off_fee_num:,} ({off_fee})",
                        '官网原文依据': f"Fees: {off_fee}",
                        '备注': f'【本地生资助标准混淆】原库误录为香港本地生资助学费({db_fee_num:,})，国际生/非本地生实际学费为 {off_fee_num:,} 港币',
                        '审批意见': '待审批'
                    })
                # 门槛未更新：学费上涨或调整
                elif diff_amount > 10000 and abs(off_fee_num - db_fee_num) > 10000:
                    if not (dur_years > 1 and abs(off_fee_num / dur_years - db_fee_num) < 2000):
                        stage2_records.append({
                            '院校中英文': school_name,
                            '专业中英文': prog_name,
                            '官网URL': url,
                            '数据库ID': db_id,
                            '报错类型': '门槛未更新',
                            '数据库现有值': db_tuition,
                            '官网真实值': f"HKD {off_fee_num:,} ({off_fee})",
                            '官网原文依据': f"Fees: {off_fee} (Expected Duration: {off_dur})",
                            '备注': f'【2027学年学费调整】官网最新公布学费为 {off_fee}，与数据库现有值存在实质差额(差额约 HK${int(diff_amount):,})',
                            '审批意见': '待审批'
                        })

        # 3. 语言小分核对
        if off_ielts:
            m_overall = re.search(r'overall\s*(?:band\s*)?(?:of\s*)?([0-9](?:\.[0-9])?)', off_ielts, re.I)
            m_sub = re.search(r'no\s*(?:sub-?test|sub-?score|band)\s*(?:lower\s*than|below)\s*([0-9](?:\.[0-9])?)', off_ielts, re.I)
            off_overall = float(m_overall.group(1)) if m_overall else None
            off_sub = float(m_sub.group(1)) if m_sub else None

            db_overall_m = re.search(r'总分要求:\s*([0-9](?:\.[0-9])?)', db_lang)
            db_overall = float(db_overall_m.group(1)) if db_overall_m else None
            
            if off_overall and db_overall and abs(off_overall - db_overall) >= 0.5:
                stage2_records.append({
                    '院校中英文': school_name,
                    '专业中英文': prog_name,
                    '官网URL': url,
                    '数据库ID': db_id,
                    '报错类型': '小分有误',
                    '数据库现有值': db_lang,
                    '官网真实值': off_ielts,
                    '官网原文依据': off_ielts,
                    '备注': f'【语言门槛小分纠偏】官网明确总分要求为 {off_overall} 分，数据库现有总分要求为 {db_overall} 分',
                    '审批意见': '待审批'
                })

    # 写入阶段二记录
    added_stage2 = 0
    for item in stage2_records:
        key = (str(item['数据库ID']), str(item['报错类型']))
        if key in existing_keys:
            continue
        existing_keys.add(key)
        
        ws_sheet2.cell(cur_row, 1, cur_row - 1).alignment = ALIGN_CENTER
        ws_sheet2.cell(cur_row, 2, item['院校中英文']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 3, item['专业中英文']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 4, item['官网URL']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 5, item['数据库ID']).alignment = ALIGN_CENTER
        ws_sheet2.cell(cur_row, 6, item['报错类型']).alignment = ALIGN_CENTER
        ws_sheet2.cell(cur_row, 7, item['数据库现有值']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 8, item['官网真实值']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 9, item['官网原文依据']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 10, item['备注']).alignment = ALIGN_LEFT
        ws_sheet2.cell(cur_row, 11, item['审批意见']).alignment = ALIGN_CENTER

        for c in range(1, 12):
            cell = ws_sheet2.cell(cur_row, c)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            cell.fill = STAGE2_FILL
        cur_row += 1
        added_stage2 += 1

    print(f"[*] 阶段二完成：登记高价值字段差异 {added_stage2} 条。")
    print(f"[*] 全流程核验完毕！累计输出总差异条数: {cur_row - 2} 条。")

    # 调整列宽
    col_widths = {
        'A': 8,   # 序号
        'B': 25,  # 院校中英文
        'C': 35,  # 专业中英文
        'D': 35,  # 官网URL
        'E': 12,  # 数据库ID
        'F': 14,  # 报错类型
        'G': 35,  # 数据库现有值
        'H': 35,  # 官网真实值
        'I': 40,  # 官网原文依据
        'J': 35,  # 备注
        'K': 12   # 审批意见
    }
    for col_letter, width in col_widths.items():
        ws_sheet2.column_dimensions[col_letter].width = width

    # 保存
    wb.save(data_out_file)
    wb.save(desktop_out_file)
    print(f"[√] 最终标准交付表已同步更新至：\n    -> {data_out_file}\n    -> {desktop_out_file}")
    return desktop_out_file

if __name__ == '__main__':
    build_sop_feedback_table("hku")
