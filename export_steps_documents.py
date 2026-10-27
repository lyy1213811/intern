"""
批量导出步骤一至步骤四的产出文件为格式精美的 Excel 文档
包含：
1. 步骤1_数据库切片_db_slice.xlsx
2. 步骤2_官网在招列表_official_list.xlsx
3. 步骤3_大盘名单差集比对_diff_programs.xlsx
4. 步骤4_网页快照下载清单_htmls_index.xlsx
5. HKU_流水线_步骤一至四全量汇总表.xlsx (包含上述4个完整工作表)
"""

import os
import sys
import re
import urllib.parse
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# 目录定义
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data", "hku")
OUTPUT_DIRS = [
    os.path.join(DATA_DIR, "export_documents"),
    r"c:\Users\rtx\Desktop\HKU_流水线步骤一至四产出文档"
]

for d in OUTPUT_DIRS:
    os.makedirs(d, exist_ok=True)

# 样式定义
FONT_FAMILY = "微软雅黑"
HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
HEADER_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
REGULAR_FONT = Font(name=FONT_FAMILY, size=9)
BOLD_FONT = Font(name=FONT_FAMILY, size=9, bold=True)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

BORDER_THIN = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

# 差集比对分类着色
DIFF_FILLS = {
    "完全匹配": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
    "模糊匹配(疑似更名)": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
    "官网独有(疑似新增/漏录)": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
    "数据库独有(疑似停办/下线)": PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")
}

INACTIVE_FILL = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()

def style_sheet(ws, title_headers, is_diff=False, is_db=False):
    ws.views.sheetView[0].showGridLines = True
    ws.row_dimensions[1].height = 28

    # 表头样式
    for col_idx, h in enumerate(title_headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN

    # 启用自动筛选
    last_col_letter = get_column_letter(len(title_headers))
    ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"

    # 数据行样式
    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 20
        row_category = ws.cell(row=row, column=1).value if is_diff else None
        is_valid_val = None
        if is_db:
            # 找到 is_valid 所在列
            is_valid_val = ws.cell(row=row, column=len(title_headers)).value

        for col in range(1, len(title_headers) + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = REGULAR_FONT
            cell.border = BORDER_THIN
            val = cell.value

            # 默认对齐逻辑
            if isinstance(val, (int, float)):
                cell.alignment = ALIGN_CENTER
            elif isinstance(val, bool):
                cell.alignment = ALIGN_CENTER
            elif val and str(val).startswith("http"):
                cell.alignment = ALIGN_LEFT
            else:
                cell.alignment = ALIGN_LEFT

            # 差集分类高亮
            if is_diff and row_category in DIFF_FILLS:
                cell.fill = DIFF_FILLS[row_category]
            elif is_db and is_valid_val is False:
                cell.fill = INACTIVE_FILL

    # 自动列宽调整
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            # 中文字符长度计算加权
            line_len = sum(2 if ord(c) > 127 else 1 for c in val_str.split('\n')[0])
            max_len = max(max_len, line_len)
        adjusted_width = min(max(max_len + 3, 10), 45)
        ws.column_dimensions[col_letter].width = adjusted_width

def generate_step1_excel(df_db):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "步骤1_数据库切片"
    headers = list(df_db.columns)
    ws.append(headers)
    for row in df_db.itertuples(index=False):
        ws.append(list(row))
    style_sheet(ws, headers, is_db=True)
    return wb

def generate_step2_excel(df_off):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "步骤2_官网在招列表"
    headers = ["序号"] + list(df_off.columns)
    ws.append(headers)
    for idx, row in enumerate(df_off.itertuples(index=False), 1):
        ws.append([idx] + list(row))
    style_sheet(ws, headers)
    return wb

def generate_step3_excel(df_diff):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "步骤3_大盘名单比对"
    headers = list(df_diff.columns)
    ws.append(headers)
    for row in df_diff.itertuples(index=False):
        ws.append(list(row))
    style_sheet(ws, headers, is_diff=True)
    return wb

def generate_step4_excel(df_htmls):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "步骤4_网页快照清单"
    headers = list(df_htmls.columns)
    ws.append(headers)
    for row in df_htmls.itertuples(index=False):
        ws.append(list(row))
    style_sheet(ws, headers)
    return wb

def generate_combined_excel(df_db, df_off, df_diff, df_htmls):
    wb = openpyxl.Workbook()
    # 1
    ws1 = wb.active
    ws1.title = "步骤1_数据库切片"
    ws1.append(list(df_db.columns))
    for row in df_db.itertuples(index=False): ws1.append(list(row))
    style_sheet(ws1, list(df_db.columns), is_db=True)

    # 2
    ws2 = wb.create_sheet(title="步骤2_官网在招列表")
    headers2 = ["序号"] + list(df_off.columns)
    ws2.append(headers2)
    for idx, row in enumerate(df_off.itertuples(index=False), 1): ws2.append([idx] + list(row))
    style_sheet(ws2, headers2)

    # 3
    ws3 = wb.create_sheet(title="步骤3_大盘名单比对")
    ws3.append(list(df_diff.columns))
    for row in df_diff.itertuples(index=False): ws3.append(list(row))
    style_sheet(ws3, list(df_diff.columns), is_diff=True)

    # 4
    ws4 = wb.create_sheet(title="步骤4_网页快照清单")
    ws4.append(list(df_htmls.columns))
    for row in df_htmls.itertuples(index=False): ws4.append(list(row))
    style_sheet(ws4, list(df_htmls.columns))

    return wb

def main():
    print("[*] 开始读取步骤一至四数据源...")
    db_file = os.path.join(DATA_DIR, "db_slice.csv")
    off_file = os.path.join(DATA_DIR, "official_list.csv")
    diff_file = os.path.join(DATA_DIR, "diff_programs.csv")
    html_dir = os.path.join(DATA_DIR, "htmls")

    df_db = pd.read_csv(db_file)
    df_off = pd.read_csv(off_file)
    df_diff = pd.read_csv(diff_file)

    # 准备步骤4快照清单数据
    slug_map = {}
    for _, row in df_off.iterrows():
        url = str(row.get('official_url', '')).strip()
        parsed = urllib.parse.urlparse(url)
        slug = urllib.parse.parse_qs(parsed.query).get('programme', [''])[0]
        if not slug:
            slug = os.path.basename(parsed.path.rstrip('/'))
        safe_slug = sanitize_filename(slug)
        slug_map[safe_slug.lower()] = {
            "name": row['official_program_name'],
            "faculty": row['faculty'],
            "url": url
        }

    html_records = []
    if os.path.exists(html_dir):
        for idx, fname in enumerate(sorted(os.listdir(html_dir)), 1):
            if not fname.endswith('.html'): continue
            fpath = os.path.join(html_dir, fname)
            fsize_bytes = os.path.getsize(fpath)
            fsize_kb = round(fsize_bytes / 1024, 1)
            fsize_mb = round(fsize_bytes / (1024 * 1024), 2)
            
            slug_key = fname[:-5].lower()
            # 兼容带有 db_id 前缀的文件名
            if '_' in slug_key and slug_key.split('_')[0].isdigit():
                slug_key = '_'.join(slug_key.split('_')[1:])

            meta = slug_map.get(slug_key, {})
            prog_name = meta.get("name", "详见快照内容")
            faculty = meta.get("faculty", "")
            url = meta.get("url", "")

            html_records.append({
                "序号": idx,
                "快照文件名": fname,
                "文件大小(KB)": fsize_kb,
                "文件大小(MB)": fsize_mb,
                "对应官网专业名称": prog_name,
                "所属学院": faculty,
                "官网详情页URL": url,
                "抓取状态": "已完成本地缓存(200 OK)"
            })
    df_htmls = pd.DataFrame(html_records)

    print(f"[*] 数据加载完成：")
    print(f"    - 步骤1 数据库切片: {len(df_db)} 条")
    print(f"    - 步骤2 官网在招列表: {len(df_off)} 条")
    print(f"    - 步骤3 大盘差集比对: {len(df_diff)} 条")
    print(f"    - 步骤4 网页快照清单: {len(df_htmls)} 条")

    print("[*] 正在生成各步骤独立 Excel 文档...")
    wb1 = generate_step1_excel(df_db)
    wb2 = generate_step2_excel(df_off)
    wb3 = generate_step3_excel(df_diff)
    wb4 = generate_step4_excel(df_htmls)
    wb_all = generate_combined_excel(df_db, df_off, df_diff, df_htmls)

    file_mapping = {
        "步骤1_数据库切片_db_slice.xlsx": wb1,
        "步骤2_官网在招列表_official_list.xlsx": wb2,
        "步骤3_大盘名单差集比对_diff_programs.xlsx": wb3,
        "步骤4_网页快照下载清单_htmls_index.xlsx": wb4,
        "HKU_流水线_步骤一至四全量中间产出汇总表.xlsx": wb_all
    }

    for target_dir in OUTPUT_DIRS:
        print(f"\n[+] 正在导出至目录: {target_dir}")
        for filename, wb in file_mapping.items():
            out_path = os.path.join(target_dir, filename)
            wb.save(out_path)
            print(f"    -> 成功保存: {filename} ({os.path.getsize(out_path) // 1024} KB)")

    print("\n[OK] 步骤一至步骤四产出文档导出全部完成！")

if __name__ == "__main__":
    main()
