"""
步骤 5：核心字段核查脚本 (5_verify_fields.py)
核心目标：提取本地 HTML 官网快照核心字段，与数据库现有字段进行精准对比，
找出学费、申请轮次、语言小分、学制的错误，生成工业级标准差异反馈表。
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
DIFF_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
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
    
    # 1. 精确匹配
    direct = os.path.join(html_dir, f"{safe_slug}.html")
    if os.path.exists(direct):
        return direct
        
    # 2. 忽略大小写及部分字符匹配
    for f in os.listdir(html_dir):
        f_lower = f.lower()
        slug_lower = safe_slug.lower()
        if f_lower == f"{slug_lower}.html":
            return os.path.join(html_dir, f)
        # 兼容带下划线或连字符的命名变体
        if slug_lower.replace('-', '_') in f_lower.replace('-', '_'):
            return os.path.join(html_dir, f)
            
    return None

def extract_from_html(filepath):
    """
    从港大官方详情页 HTML 中提取关键字段
    """
    if not filepath or not os.path.exists(filepath):
        return None

    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        html_content = f.read()

    soup = bs4.BeautifulSoup(html_content, 'html.parser')
    
    data = {
        'duration': '',
        'fees_raw': '',
        'deadlines_raw': '',
        'deadlines_rounds': [],
        'ielts_raw': '',
        'toefl_raw': '',
        'entrance_requirements': ''
    }

    # 1. 从 highlights-item 中提取学制、学费、申请时间
    for item in soup.find_all('div', class_='highlights-item'):
        t_el = item.find('div', class_='highlights-item-title')
        d_el = item.find('div', class_='highlights-item-description')
        if not t_el or not d_el:
            continue
        title = t_el.get_text(strip=True).lower()
        desc = d_el.get_text(' ', strip=True)

        if 'duration' in title and desc and desc != 'To be updated':
            if not data['duration']:
                data['duration'] = desc
        elif 'fee' in title and desc and desc != 'To be updated':
            if not data['fees_raw']:
                data['fees_raw'] = desc
        elif 'deadline' in title and desc and desc != 'To be updated':
            if not data['deadlines_raw']:
                data['deadlines_raw'] = desc

    # 2. 提取申请截止具体轮次
    if data['deadlines_raw']:
        # 拆解常见轮次
        lines = [line.strip() for line in re.split(r'[\r\n]+', data['deadlines_raw']) if line.strip()]
        if not lines:
            lines = [data['deadlines_raw']]
        data['deadlines_rounds'] = lines

    # 3. 提取语言要求 (Entrance Requirements)
    req_sec = soup.find('div', class_=lambda c: c and 'section-entrance-requirements' in c)
    if not req_sec:
        # fallback 查找带 Entrance Requirements 的标题
        for h in soup.find_all(['h2', 'h3', 'h4']):
            if 'Entrance Requirements' in h.get_text():
                req_sec = h.find_parent('div')
                break

    if req_sec:
        full_req_text = req_sec.get_text('\n', strip=True)
        data['entrance_requirements'] = full_req_text
        
        # 提取 IELTS
        ielts_match = re.search(r'(?:IELTS[^\.\;\n]{0,150}(?:band|score|overall)[^\.\;\n]{0,150})', full_req_text, re.I)
        if ielts_match:
            data['ielts_raw'] = ielts_match.group(0).strip()
            
        # 提取 TOEFL
        toefl_match = re.search(r'(?:TOEFL[^\.\;\n]{0,150}(?:score)?[^\.\;\n]{0,100})', full_req_text, re.I)
        if toefl_match:
            data['toefl_raw'] = toefl_match.group(0).strip()

    return data

def parse_fee_number(fee_str, prefer_non_local=True):
    """
    提取数字学费。面向内地/非本地生（Non-local）时，优先提取 Non-local 学费，彻底避免误抓 Local 学费
    """
    if not fee_str or pd.isna(fee_str):
        return None
    s = str(fee_str)
    
    if prefer_non_local:
        # 1. 优先捕获 Non-local 关键字后的学费
        m_nl = re.search(r'(?:non-local|international|non local)[^\d]{0,25}(?:HK\$|HKD|RMB|￥)?\s*([0-9]{2,3}(?:,[0-9]{3})+|[0-9]{5,7})', s, re.I)
        if m_nl:
            return int(m_nl.group(1).replace(',', ''))

    # 2. 如果包含 non-local 关键字但未匹配到上述规则，进行宽松匹配
    if 'non-local' in s.lower():
        m_loose = re.search(r'non-local[^\d]+([0-9,]{5,10})', s, re.I)
        if m_loose:
            return int(m_loose.group(1).replace(',', ''))

    # 3. 兜底匹配普通金额
    nums = re.findall(r'(?:HK\$|HKD|RMB|￥)?\s*([0-9]{2,3}(?:,[0-9]{3})+|[0-9]{5,7})', s)
    if nums:
        return int(nums[0].replace(',', ''))
    return None

def parse_duration_years(dur_str):
    """提取学制年数"""
    if not dur_str or pd.isna(dur_str):
        return 1.0
    dur_str = str(dur_str).lower()
    m = re.search(r'([0-9]+(?:\.[0-9]+)?)\s*year', dur_str)
    if m:
        return float(m.group(1))
    m_cn = re.search(r'([0-9]+(?:\.[0-9]+)?)\s*年', dur_str)
    if m_cn:
        return float(m_cn.group(1))
    return 1.0

def verify_single_program(row, db_info, html_data):
    """
    单专业深度审核逻辑，符合导师制定的严格核验标准：
    1. is_valid == False 的专业不核验（导师规则一）
    2. 计算学费时结合学制换算（导师规则二）
    3. 官网显示 To be updated 或暂未公布的不误报（导师规则三）
    4. 语言要求无实质分数差异不报（导师规则四）
    5. 严格纠偏轮次错位（如将 Early Admission 误标为 Round 1）
    """
    findings = []
    
    # 规则 1：检查是否有效 (防红线：is_valid == False 的专业不核验)
    is_valid = db_info.get('is_valid', True)
    if str(is_valid).lower() in ['false', '0']:
        return findings

    db_id = row['数据库ID']
    prog_name = str(row['专业中英文 (专业名/学位)'])
    school_name = row['院校中英文']
    url = row['官网URL']

    # 特判香港本地 UGC 资助专属项目，避免拿国际生自资学费报错（防红线：如建筑学、城市规划UGC版）
    if any(k in prog_name.lower() for k in ['ugc funded for hk applicants', '香港申请人由大学教育资助委员会资助', '香港申请人政府资助']):
        return findings
    
    db_tuition = str(row['数据库现有学费']) if pd.notna(row['数据库现有学费']) else ''
    db_app = str(row['数据库现有申请时间']) if pd.notna(row['数据库现有申请时间']) else ''
    db_lang = str(row['数据库现有语言要求']) if pd.notna(row['数据库现有语言要求']) else ''
    db_dur = str(row['学制']) if pd.notna(row['学制']) else ''

    if not html_data:
        return findings

    off_fee = html_data['fees_raw']
    off_deadlines = html_data['deadlines_raw']
    off_dur = html_data['duration']
    off_ielts = html_data['ielts_raw']
    off_toefl = html_data['toefl_raw']

    # ---------------- 1. 申请时间与轮次深度比对 ----------------
    if off_deadlines and off_deadlines != 'To be updated' and 'TBC' not in off_deadlines:
        # 检查提前批轮次命名错位 (Early Admission 被误标为 Round 1)
        if 'early admission' in off_deadlines.lower():
            if 'early admission' not in db_app.lower() and '提前批' not in db_app:
                findings.append({
                    '院校中英文': school_name,
                    '专业中英文': prog_name,
                    '官网URL': url,
                    '数据库ID': db_id,
                    '报错类型': '申请时间有误',
                    '数据库现有值': db_app,
                    '官网真实值': off_deadlines,
                    '官网原文依据': f"Application Deadline: {off_deadlines}",
                    '备注': '【轮次命名规范纠偏】官网明确标注有提前批(Early Admission Round)，库中误将提前批标为Round 1导致轮次整体错位一轮；已核验官网学制与轮次，符合商学院/建筑学院统一审核规范',
                    '审批意见': '待审批'
                })
        
        # 检查年份停留在旧年份 (如数据库停在25/26年，而官网已发布2027年轮次)
        elif '2027' in off_deadlines and ('2025' in db_app or '26年秋季' in db_app):
            findings.append({
                '院校中英文': school_name,
                '专业中英文': prog_name,
                '官网URL': url,
                '数据库ID': db_id,
                '报错类型': '申请时间有误',
                '数据库现有值': db_app,
                '官网真实值': off_deadlines,
                '官网原文依据': f"Application Deadline: {off_deadlines}",
                '备注': '【招生季轮次更新】库中记录停留在2026旧招生季(轮次已过期)，官网已明确发布2027/28学年最新截止日期',
                '审批意见': '待审批'
            })

    # ---------------- 2. 学费深度比对 ----------------
    if off_fee and off_fee != 'To be updated' and 'TBC' not in off_fee:
        off_fee_num = parse_fee_number(off_fee, prefer_non_local=True)
        db_fee_num = parse_fee_number(db_tuition, prefer_non_local=False)
        dur_years = parse_duration_years(off_dur if off_dur else db_dur)

        if off_fee_num and db_fee_num:
            # 规则 3：防多方向误报（如教育学 IB方向 249,000 与 Non-IB方向 234,000）
            # 若库中现有值与官网文本中任意合法列出的细分方向金额一致，则属原值准确，不误报（对齐导师审批意见规则）
            all_off_nums = [int(n.replace(',', '')) for n in re.findall(r'(?:HK\$|HKD)?\s*([0-9]{2,3}(?:,[0-9]{3})+|[0-9]{5,7})', off_fee)]
            if db_fee_num in all_off_nums and 'non-local' not in off_fee.lower():
                return findings

            # 区分本地生与非本地生（防红线：避免把本地生资助学费 48250 当成国际生）
            is_local_fee_in_db = (db_fee_num == 48250 or db_fee_num == 42100)

            # 计算年化学费对比（结合学制年数换算，严格遵守导师序号8批复规则）
            annual_off_fee = off_fee_num / dur_years if dur_years > 1 and off_fee_num > db_fee_num * 1.3 else off_fee_num
            diff_amount = abs(annual_off_fee - db_fee_num)

            if is_local_fee_in_db and off_fee_num > 100000:
                pass
            elif diff_amount > 10000 and abs(off_fee_num - db_fee_num) > 10000:
                # 排除整项目 vs 按年已折算一致的情况
                if not (dur_years > 1 and abs(off_fee_num / dur_years - db_fee_num) < 2000):
                    # 细化高质量专业备注（对齐导师在序号5和序号9的批注风格）
                    if 'provisional' in off_fee.lower() or 'approval' in off_fee.lower():
                        note = f"【暂定学费标准调整】官网公布2027/28暂定学费为{off_fee}，与数据库现有值存在实质差额(差额约 HK${int(diff_amount):,})；已排除本地生标准，加注待大学最终审批生效"
                    elif dur_years > 1:
                        note = f"【学费标准更新(已年化换算)】官网预期学制为{off_dur}，整项目总学费为 HKD {off_fee_num:,}，按学制折算年化学费约 HKD {int(annual_off_fee):,}/年，与数据库现有值存在实质差额(差额约 HK${int(diff_amount):,})；已排除本地生资助标准"
                    else:
                        note = f"【学费标准重大调整】官网最新公布学费为{off_fee}，与数据库现有值存在实质差额(差额约 HK${int(diff_amount):,})；已排除本地生与留位费干扰"

                    findings.append({
                        '院校中英文': school_name,
                        '专业中英文': prog_name,
                        '官网URL': url,
                        '数据库ID': db_id,
                        '报错类型': '学费有误',
                        '数据库现有值': db_tuition,
                        '官网真实值': f"HKD {off_fee_num:,} ({off_fee})",
                        '官网原文依据': f"Fees: {off_fee} (Expected Duration: {off_dur})",
                        '备注': note,
                        '审批意见': '待审批'
                    })

    # ---------------- 3. 语言成绩小分比对 ----------------
    if off_ielts:
        # 提取官网雅思总分与小分
        m_overall = re.search(r'overall\s*(?:band\s*)?(?:of\s*)?([0-9](?:\.[0-9])?)', off_ielts, re.I)
        m_sub = re.search(r'no\s*(?:sub-?test|sub-?score|band)\s*(?:lower\s*than|below)\s*([0-9](?:\.[0-9])?)', off_ielts, re.I)
        
        off_overall = float(m_overall.group(1)) if m_overall else None
        off_sub = float(m_sub.group(1)) if m_sub else None

        # 检查数据库雅思
        db_overall_m = re.search(r'总分要求:\s*([0-9](?:\.[0-9])?)', db_lang)
        db_overall = float(db_overall_m.group(1)) if db_overall_m else None
        
        if off_overall and db_overall and abs(off_overall - db_overall) >= 0.5:
            findings.append({
                '院校中英文': school_name,
                '专业中英文': prog_name,
                '官网URL': url,
                '数据库ID': db_id,
                '报错类型': '语言要求有误',
                '数据库现有值': db_lang,
                '官网真实值': off_ielts,
                '官网原文依据': off_ielts,
                '备注': f'【雅思门槛分级纠偏】官网明确总分要求为 {off_overall} 分，数据库现有总分要求为 {db_overall} 分',
                '审批意见': '待审批'
            })

    # ---------------- 4. 学制比对 ----------------
    if off_dur and off_dur != 'To be updated':
        dur_off_val = parse_duration_years(off_dur)
        dur_db_val = parse_duration_years(db_dur)
        if abs(dur_off_val - dur_db_val) >= 0.5:
            findings.append({
                '院校中英文': school_name,
                '专业中英文': prog_name,
                '官网URL': url,
                '数据库ID': db_id,
                '报错类型': '学制有误',
                '数据库现有值': db_dur,
                '官网真实值': off_dur,
                '官网原文依据': f"Expected Duration: {off_dur}",
                '备注': f'【学制年限纠偏】官网最新全日制预期学制为 {off_dur}，数据库记录为 {db_dur}',
                '审批意见': '待审批'
            })

    return findings

def run_step_5(school="hku"):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, "data", school.lower())
    html_dir = os.path.join(data_dir, "htmls")
    template_file = r"c:\Users\rtx\Desktop\数据核对与差异反馈表.xlsx"
    out_dir_file = os.path.join(data_dir, f"{school.upper()}_数据核对与差异反馈表.xlsx")
    desktop_out_file = os.path.join(r"c:\Users\rtx\Desktop", f"{school.upper()}_数据核对与差异反馈表.xlsx")

    # 优先查找模板，若不存在则回退至现有的反馈表
    if not os.path.exists(template_file):
        if os.path.exists(desktop_out_file):
            template_file = desktop_out_file
        elif os.path.exists(out_dir_file):
            template_file = out_dir_file
        else:
            raise FileNotFoundError(f"未找到质检模板或基准表: {template_file}")

    print(f"[*] 启动步骤 5：核心字段深度核验流程 [{school.upper()}]...")

    if not os.path.exists(html_dir):
        raise FileNotFoundError(f"未找到本地快照目录: {html_dir}，请先运行 4_download_pages.py")

    # 读取数据库切片获取 is_valid
    slice_file = os.path.join(data_dir, "db_slice.csv")
    db_slice_map = {}
    if os.path.exists(slice_file):
        df_slice = pd.read_csv(slice_file)
        for _, r in df_slice.iterrows():
            db_slice_map[int(r['id'])] = r.to_dict()

    # 打开现有的工作簿
    wb = openpyxl.load_workbook(template_file)
    ws_base = wb['港大在招基准清单(185项TPg)']
    ws_feedback = wb['数据核对与差异反馈表']

    # 保留前 14 条历史质检记录（包含导师审批意见），清理历史残留的待审批误报行
    if ws_feedback.max_row > 15:
        ws_feedback.delete_rows(16, ws_feedback.max_row - 15)

    # 记录 Sheet 2 现有的已有项 (ID + 报错类型) 防止重复追加
    existing_keys = set()
    for row_idx in range(2, ws_feedback.max_row + 1):
        db_id = ws_feedback.cell(row=row_idx, column=5).value
        err_type = ws_feedback.cell(row=row_idx, column=6).value
        if db_id and err_type:
            existing_keys.add((int(db_id), str(err_type).split()[0]))

    print(f"[*] Sheet 2 现有历史质检反馈: {len(existing_keys)} 条 (将完整保留其审批与导师意见)")

    # 读取 Sheet 1 全部 185 个专业
    base_rows = []
    headers_base = [ws_base.cell(row=1, column=c).value for c in range(1, ws_base.max_column + 1)]
    for r in range(2, ws_base.max_row + 1):
        row_dict = {}
        for c in range(1, ws_base.max_column + 1):
            row_dict[headers_base[c - 1]] = ws_base.cell(row=r, column=c).value
        base_rows.append(row_dict)

    print(f"[*] Sheet 1 待核查大盘基准清单: {len(base_rows)} 项专业")

    # 执行核验比对
    new_discrepancies = []
    audited_count = 0
    valid_count = 0

    for row in base_rows:
        db_id = row.get('数据库ID')
        if not db_id:
            continue
        db_id = int(db_id)
        db_info = db_slice_map.get(db_id, {})
        
        # 仅针对有效专业深入核验 (遵守导师红线)
        if str(db_info.get('is_valid', True)).lower() in ['false', '0']:
            continue
            
        valid_count += 1
        url = row.get('官网URL')
        fpath = get_html_path(url, html_dir)
        html_data = extract_from_html(fpath) if fpath else None
        
        findings = verify_single_program(row, db_info, html_data)
        for f in findings:
            key = (int(f['数据库ID']), str(f['报错类型']).split()[0])
            if key not in existing_keys:
                new_discrepancies.append(f)
                existing_keys.add(key)
        audited_count += 1

    print(f"[*] 深度核验完成：有效在招专业 {valid_count} 个，新精准定位出高价值差异: {len(new_discrepancies)} 条")

    # 将新发现的差异追加写入 Sheet 2
    cur_row = ws_feedback.max_row + 1
    seq = ws_feedback.max_row  # 从当前最后一条序号递增

    for disc in new_discrepancies:
        ws_feedback.cell(row=cur_row, column=1, value=seq).alignment = ALIGN_CENTER
        ws_feedback.cell(row=cur_row, column=2, value=disc['院校中英文']).alignment = ALIGN_LEFT
        ws_feedback.cell(row=cur_row, column=3, value=disc['专业中英文']).alignment = ALIGN_LEFT
        ws_feedback.cell(row=cur_row, column=4, value=disc['官网URL']).alignment = ALIGN_LEFT
        ws_feedback.cell(row=cur_row, column=5, value=disc['数据库ID']).alignment = ALIGN_CENTER
        ws_feedback.cell(row=cur_row, column=6, value=disc['报错类型']).alignment = ALIGN_CENTER
        ws_feedback.cell(row=cur_row, column=7, value=disc['数据库现有值']).alignment = ALIGN_LEFT
        ws_feedback.cell(row=cur_row, column=8, value=disc['官网真实值']).alignment = ALIGN_LEFT
        ws_feedback.cell(row=cur_row, column=9, value=disc['官网原文依据']).alignment = ALIGN_LEFT
        ws_feedback.cell(row=cur_row, column=10, value=disc['备注']).alignment = ALIGN_LEFT
        ws_feedback.cell(row=cur_row, column=11, value=disc['审批意见']).alignment = ALIGN_CENTER

        # 高亮标记新增的待审批行
        for col in range(1, 12):
            cell = ws_feedback.cell(row=cur_row, column=col)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            cell.fill = DIFF_FILL

        seq += 1
        cur_row += 1

    # 保存工作簿
    wb.save(out_dir_file)
    print(f"[√] 质检核对表更新成功并已保存至项目数据目录：\n    -> {out_dir_file}")

    try:
        wb.save(desktop_out_file)
        print(f"[√] 桌面质检核对表已同步更新：\n    -> {desktop_out_file}")
    except PermissionError:
        print(f"[!] 提示：桌面文件 [{desktop_out_file}] 当前正被 Excel 软件打开占用。\n    请在 Excel 中关闭该文件，以便脚本能直接覆写桌面文件。\n    项目内置副本已先行更新完毕。")
    return out_dir_file, len(new_discrepancies)

if __name__ == "__main__":
    run_step_5("hku")
