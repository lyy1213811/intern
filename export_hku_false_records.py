r"""
脚本：导出香港大学 is_valid = false 的失效/停办专业清单
数据来源：Aliyun RDS tunshuedu_db.public.ai_selection_programs
产出：
1. c:\Users\rtx\Desktop\香港大学_数据库失效专业清单(is_valid=False_共36项).xlsx
2. c:\Users\rtx\Desktop\HKU_流水线步骤一至四产出文档\步骤1附表_数据库失效专业_is_valid_false.xlsx
3. c:\Users\rtx\Desktop\intern\data\hku\hku_db_slice_false.csv
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

DB_CONFIG = {
    "host": "pgm-bp1k176i7sid0e775o.pg.rds.aliyuncs.com",
    "port": 5432,
    "user": "data_maintain",
    "password": "Str0ngP@5s!",
    "dbname": "tunshuedu_db"
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

def export_hku_false():
    print("[*] 正在连接数据库提取 [香港大学] is_valid = false 失效专业记录...")
    conn = psycopg2.connect(**DB_CONFIG)
    conn.set_session(readonly=True)

    fields_sql = ", ".join(REQUIRED_FIELDS)
    sql = f"""
        SELECT {fields_sql}
        FROM ai_selection_programs
        WHERE (school_name_cn = '香港大学' OR school_name_en = 'The University of Hong Kong')
          AND is_valid IS FALSE
        ORDER BY id ASC
    """
    
    df = pd.read_sql(sql, conn)
    conn.close()
    
    print(f"[*] 成功从数据库检索到 is_valid = false 专业记录: {len(df)} 项")

    # 1. 保存 CSV
    csv_out = r"c:\Users\rtx\Desktop\intern\data\hku\hku_db_slice_false.csv"
    os.makedirs(os.path.dirname(csv_out), exist_ok=True)
    df.to_csv(csv_out, index=False, encoding='utf-8-sig')

    # 2. 导出精美格式 Excel
    desktop_excel = r"c:\Users\rtx\Desktop\香港大学_数据库失效专业清单(is_valid=False_共36项).xlsx"
    pipeline_excel = r"c:\Users\rtx\Desktop\HKU_流水线步骤一至四产出文档\步骤1附表_数据库失效专业_is_valid_false.xlsx"
    os.makedirs(os.path.dirname(pipeline_excel), exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "HKU_失效专业(is_valid=False)"

    # 样式配置
    HEADER_FILL = PatternFill(start_color="800000", end_color="800000", fill_type="solid") # 深红/酒红色表头标识失效
    HEADER_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    REGULAR_FONT = Font(name="微软雅黑", size=9)
    BOLD_FONT = Font(name="微软雅黑", size=9, bold=True)
    ROW_FILL = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")
    FALSE_BADGE_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # 浅红高亮False

    ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    BORDER_THIN = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    ws.views.sheetView[0].showGridLines = True
    ws.row_dimensions[1].height = 28

    # 表头
    headers = list(df.columns)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN

    # 数据行
    for r_idx, row in df.iterrows():
        row_num = r_idx + 2
        ws.row_dimensions[row_num].height = 24
        for c_idx, col_name in enumerate(headers, 1):
            val = row[col_name]
            cell = ws.cell(row_num, c_idx, val)
            cell.font = REGULAR_FONT
            cell.border = BORDER_THIN

            # 对齐
            if col_name in ['id', 'degree', 'program_duration', 'is_valid', 'official_code', 'program_code']:
                cell.alignment = ALIGN_CENTER
            else:
                cell.alignment = ALIGN_LEFT

            # is_valid 单元格重点高亮
            if col_name == 'is_valid':
                cell.fill = FALSE_BADGE_FILL
                cell.font = BOLD_FONT

    # 启用自动筛选
    last_col_letter = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"

    # 调整列宽
    col_widths = {
        'A': 10,  # id
        'B': 14,  # school_name_cn
        'C': 26,  # school_name_en
        'D': 38,  # program_name_en
        'E': 32,  # program_name_cn
        'F': 14,  # program_code
        'G': 10,  # degree
        'H': 12,  # program_duration
        'I': 20,  # program_tuition
        'J': 35,  # application_time
        'K': 35,  # language_requirements
        'L': 35,  # program_website
        'M': 14,  # official_code
        'N': 12   # is_valid
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # 保存
    wb.save(desktop_excel)
    wb.save(pipeline_excel)

    print(f"[√] 导出完成！已保存至：")
    print(f"    1. {desktop_excel}")
    print(f"    2. {pipeline_excel}")
    print(f"    3. {csv_out}")

    print("\n=== 36 项失效专业主要类型归纳 ===")
    print("1. 计算机科学（CS）旧拆分方向（ID: 1065, 1066, 1067, 1068）：")
    print("   -> 网络安全、金融计算、多媒体计算、常规方向已全部合并为多方向通用大类 (ID: 49700)")
    print("2. 教育学（MEd）旧专业方向（ID: 1083, 1084, 1085, 1086, 1087, 1088, 1090 等）：")
    print("   -> 往年部分细分方向停招或并入大类")
    print("3. 建筑与城市设计旧项目（ID: 1106 城市设计旧条目等）：")
    print("   -> 历史条目已被新开设的普通申请者项目 (ID: 44056) 替代")

    return df

if __name__ == '__main__':
    export_hku_false()
