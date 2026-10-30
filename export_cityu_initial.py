r"""
香港城市大学 (CityU) 步骤一：数据库切片全量与失效专业导出脚本
产出：
1. c:\Users\rtx\Desktop\香港城市大学_数据库切片全量清单.xlsx
2. c:\Users\rtx\Desktop\香港城市大学_数据库失效专业清单(is_valid=False).xlsx
3. c:\Users\rtx\Desktop\香港城市大学_数据库失效专业清单(is_valid=False).csv
4. c:\Users\rtx\Desktop\CityU_步骤一至四\步骤1_数据库切片_db_slice.xlsx
5. c:\Users\rtx\Desktop\CityU_步骤一至四\步骤1附表_数据库失效专业_is_valid_false.xlsx
"""

import os
import sys
import psycopg2
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"C:\Users\rtx\Desktop\intern"
DATA_DIR = os.path.join(BASE_DIR, "data", "cityu")
os.makedirs(DATA_DIR, exist_ok=True)

DESKTOP_FOLDER = r"c:\Users\rtx\Desktop\CityU_步骤一至四"
os.makedirs(DESKTOP_FOLDER, exist_ok=True)

slice_csv = os.path.join(DATA_DIR, "db_slice.csv")

DB_CONFIG = {
    "host": os.getenv("DB_HOST", "pgm-bp1k176i7sid0e775o.pg.rds.aliyuncs.com"),
    "port": int(os.getenv("DB_PORT", 5432)),
    "user": os.getenv("DB_USER", "data_maintain"),
    "password": os.getenv("DB_PASSWORD", "Str0ngP@5s!"),
    "dbname": os.getenv("DB_NAME", "tunshuedu_db")
}

REQUIRED_FIELDS = [
    "id", "school_name_cn", "school_name_en", "program_name_en",
    "program_name_cn", "program_code", "degree", "program_duration",
    "program_tuition", "application_time", "language_requirements",
    "program_website", "official_code", "is_valid"
]

# 样式
FONT_FAMILY = "微软雅黑"
HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
HEADER_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
INACTIVE_FILL = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
WARN_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

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

def fetch_and_save_db():
    print("[*] 正在连接数据库导出 [香港城市大学] 专业切片...")
    conn = psycopg2.connect(**DB_CONFIG)
    try:
        conn.set_session(readonly=True)
        fields_str = ", ".join(REQUIRED_FIELDS)
        sql = f"""
            SELECT {fields_str}
            FROM ai_selection_programs
            WHERE (school_name_cn = '香港城市大学' OR school_name_en = 'City University of Hong Kong')
              AND degree = '硕士'
            ORDER BY id ASC
        """
        with conn.cursor() as cur:
            cur.execute(sql)
            rows = cur.fetchall()
            columns = [desc[0] for desc in cur.description]
            df = pd.DataFrame(rows, columns=columns)
            
        df.to_csv(slice_csv, index=False, encoding="utf-8-sig")
        print(f"[OK] 成功导出 {len(df)} 条专业记录至: {slice_csv}")
        return df
    finally:
        conn.close()

def build_excel(df, sheet_title):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.views.sheetView[0].showGridLines = True
    ws.row_dimensions[1].height = 28

    headers = [
        "序号", "数据库ID", "校区归属", "院校中文名", "院校英文名", "专业英文名",
        "专业中文名", "学位类型", "学制", "现有学费", "现有申请时间", "现有语言要求",
        "官网网址", "是否有效(is_valid)"
    ]
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER

    for idx, r in df.iterrows():
        row_num = idx + 1
        db_id = r['id']
        prog_cn = str(r.get('program_name_cn', '') or '')
        prog_en = str(r.get('program_name_en', '') or '')
        url = str(r.get('program_website', '') or '')
        
        campus = "九龙塘本部 (Kowloon Tong)"
        if "东莞" in prog_cn or "dongguan" in prog_en.lower() or "cityu-dg" in url.lower():
            campus = "东莞校区 (CityU Dongguan)"
        
        is_valid = r.get('is_valid', True)
        is_valid_bool = (str(is_valid).lower() == 'true')

        vals = [
            row_num,
            db_id,
            campus,
            str(r.get('school_name_cn', '') or ''),
            str(r.get('school_name_en', '') or ''),
            prog_en,
            prog_cn,
            str(r.get('degree', '') or ''),
            str(r.get('program_duration', '') or ''),
            str(r.get('program_tuition', '') or ''),
            str(r.get('application_time', '') or '').replace('<br/>', '\n').strip(),
            str(r.get('language_requirements', '') or '').replace('<br/>', '\n').strip(),
            url,
            "有效 (True)" if is_valid_bool else "失效 (False)"
        ]
        ws.append(vals)

        cur_row = ws.max_row
        ws.row_dimensions[cur_row].height = 22
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=cur_row, column=col_idx)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            if col_idx in [1, 2, 3, 4, 8, 9, 14]:
                cell.alignment = ALIGN_CENTER
            else:
                cell.alignment = ALIGN_LEFT
            
            if not is_valid_bool:
                cell.fill = INACTIVE_FILL
            if "东莞" in campus:
                cell.fill = WARN_FILL

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            first_line = val_str.split('\n')[0]
            line_len = sum(2 if ord(c) > 127 else 1 for c in first_line)
            max_len = max(max_len, line_len)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 45)

    ws.auto_filter.ref = f"A1:N{ws.max_row}"
    return wb

def main():
    df = fetch_and_save_db()
    total_cnt = len(df)
    false_df = df[df['is_valid'].astype(str).str.lower() == 'false'].copy().reset_index(drop=True)
    false_cnt = len(false_df)

    print(f"[*] 全量专业数: {total_cnt}，其中失效专业数: {false_cnt}")

    # 1. 导出全量清单至桌面
    desktop_all_file = os.path.join(r"c:\Users\rtx\Desktop", f"香港城市大学_数据库切片全量清单(共{total_cnt}项).xlsx")
    wb_all = build_excel(df, f"CityU_切片全量清单({total_cnt}项)")
    wb_all.save(desktop_all_file)
    print(f"[OK] 导出桌面全量文件: {desktop_all_file}")

    # 2. 导出失效专业清单至桌面 (.xlsx & .csv)
    desktop_false_xlsx = os.path.join(r"c:\Users\rtx\Desktop", f"香港城市大学_数据库失效专业清单(is_valid=False_共{false_cnt}项).xlsx")
    wb_false = build_excel(false_df, f"CityU_失效专业清单({false_cnt}项)")
    wb_false.save(desktop_false_xlsx)
    print(f"[OK] 导出桌面失效Excel: {desktop_false_xlsx}")

    desktop_false_csv = os.path.join(r"c:\Users\rtx\Desktop", f"香港城市大学_数据库失效专业清单(is_valid=False_共{false_cnt}项).csv")
    false_df.to_csv(desktop_false_csv, index=False, encoding="utf-8-sig")
    print(f"[OK] 导出桌面失效CSV: {desktop_false_csv}")

    # 3. 导出至 CityU_步骤一至四 目录
    step1_all = os.path.join(DESKTOP_FOLDER, "步骤1_数据库切片_db_slice.xlsx")
    wb_all.save(step1_all)
    print(f"[OK] 导出步骤1全量文件: {step1_all}")

    step1_false = os.path.join(DESKTOP_FOLDER, "步骤1附表_数据库失效专业_is_valid_false.xlsx")
    wb_false.save(step1_false)
    print(f"[OK] 导出步骤1附表文件: {step1_false}")

if __name__ == "__main__":
    main()
