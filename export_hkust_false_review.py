r"""
生成香港科技大学 (HKUST) is_valid=False 失效专业人工核对专用工作簿
文件保存至桌面：c:\Users\rtx\Desktop\香港科技大学_数据库失效专业清单(is_valid=False_共21项_人工核对专用).xlsx
以及：c:\Users\rtx\Desktop\HKUST_步骤一至四\步骤1附表_数据库失效专业_is_valid_false.xlsx
"""

import os
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data", "hkust")
slice_csv = os.path.join(DATA_DIR, "db_slice.csv")

# 桌面目标
desktop_file = r"c:\Users\rtx\Desktop\香港科技大学_数据库失效专业清单(is_valid=False_共21项_人工核对专用).xlsx"
desktop_sub_file = r"c:\Users\rtx\Desktop\香港科技大学_数据库失效专业清单(is_valid=False_共21项).xlsx"
step1_folder_file = r"c:\Users\rtx\Desktop\HKUST_步骤一至四\步骤1附表_数据库失效专业_is_valid_false.xlsx"

# 样式
FONT_FAMILY = "微软雅黑"
HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
HEADER_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")

AUDIT_HEADER_FILL = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid") # 红头人工核对栏
WARN_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # 黄色警示（需重启激活）
INACTIVE_FILL = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
GZ_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
PASS_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

REGULAR_FONT = Font(name=FONT_FAMILY, size=9)
BOLD_FONT = Font(name=FONT_FAMILY, size=9, bold=True)
RED_BOLD_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color="C00000")
GREEN_BOLD_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color="375623")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

# 21 条专业的事实依据与智能预审分析表
ANALYSIS_MAP = {
    # 清水湾本部 - 细分合并
    3199: {
        "audit_opinion": "确认失效归档",
        "reason": "【历史细分方向合并】原对外汉语教学(TCSL)子方向，现官网已与TESL统一合并为大类专业招生，已由在招主专业替代，避免重复建档。",
        "replaced_by": "在招主专业 ID: 44059 (MA in International Language Education)",
        "prog_type": "授课型 (TPg)",
        "campus": "清水湾本部"
    },
    3200: {
        "audit_opinion": "确认失效归档",
        "reason": "【历史细分方向合并】原英语教学(TESL)子方向，现官网已统一整合为大类专业招生，已由在招主专业替代，避免重复建档。",
        "replaced_by": "在招主专业 ID: 44059 (MA in International Language Education)",
        "prog_type": "授课型 (TPg)",
        "campus": "清水湾本部"
    },
    # 清水湾本部 - 研究型 MPhil
    3230: {
        "audit_opinion": "确认失效归档",
        "reason": "【研究型硕士RPg】哲学硕士(MPhil)属于学术研究型学位，非全日制授课型(TPg)留学院校库收录范围，状态置为False正常。",
        "replaced_by": "无 (研究型学位不收录)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "清水湾本部"
    },
    # 清水湾本部 - 休眠重启！
    3238: {
        "audit_opinion": "⚠️建议解除失效/激活在招",
        "reason": "【⚠️重大发现：官网2027/28已重启在招】与英国南安普顿大学联合开设的1年全日制双学位，此前因2026暂停招生归档，现官网正式重新招收2027/28秋季学生，建议激活！",
        "replaced_by": "建议激活本记录或 ID: 44058",
        "prog_type": "授课型 (TPg / 双学位)",
        "campus": "清水湾本部"
    },
    44058: {
        "audit_opinion": "⚠️建议解除失效/激活在招",
        "reason": "【⚠️重大发现：官网2027/28已重启在招】环球海洋资源管理(MSc GMRM)全日制双学位(港科大学费14万+南安普顿1万镑)，此前2025/26版旧条目，现已重启，建议解除失效状态！",
        "replaced_by": "建议激活本记录 (对应官网 msc-gmrm)",
        "prog_type": "授课型 (TPg / 双学位)",
        "campus": "清水湾本部"
    },
    # 清水湾本部 - 更名换代
    44060: {
        "audit_opinion": "确认失效归档",
        "reason": "【历史专业更名升级】原环境科学与管理专业已全面升级更名为新大类专业，原有旧名称历史记录归档停用。",
        "replaced_by": "在招新专业 ID: 3226 (MSc in Environmental Science, Management and Sustainability)",
        "prog_type": "授课型 (TPg)",
        "campus": "清水湾本部"
    },
    # 广州校区 - 15门早期过渡性 MPhil
    3182: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】广州校区金融科技MPhil属于学术研究型硕士，且广州校区以独立建档运行，研究型置为False正常。",
        "replaced_by": "广州在招授课型参考: ID 3184 (清水湾) 或广州授课型",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3186: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】人工智能MPhil属于研究型硕士，早期数据清洗遗留归档，不纳入授课型库。",
        "replaced_by": "广州在招授课型参考: ID 3190 (以数据为中心的人工智能)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3187: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】物联网MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "无 (研究型学位不收录)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3194: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】微电子MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "无 (研究型学位不收录)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3201: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】机器人与自主系统MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "清水湾在招工学硕士参考: ID 49734 (MEng in Robotics & Embodied AI)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3202: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】智能制造MPhil属于学术研究型硕士，早期过渡性记录，状态置为False正常。",
        "replaced_by": "广州在招授课型参考: ID 3205 (MSc in Smart Manufacturing)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3206: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】先进材料MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "清水湾在招参考: ID 3210 (MSc in Materials Engineering)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3207: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】城市治理与设计MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "无 (研究型学位不收录)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3213: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】生物科学与生物医药工程MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "清水湾在招参考: ID 3247 (生物分子工程与健康信息学)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3222: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】碳中和与气候变化MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "广州在招授课型参考: ID 3223 (碳中和与绿色金融理学硕士)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3234: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】可持续能源与环境MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "清水湾在招参考: ID 3226 (环境科学、管理与可持续发展)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3236: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】智能交通MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "无 (研究型学位不收录)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3241: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】计算媒体与艺术MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "广州在招艺术硕士参考: ID 11586 (MFA in Computational Arts)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3243: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】数据科学与分析MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "清水湾在招参考: ID 3246 (大数据技术) / 广州校区红鸟MPhil",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    },
    3244: {
        "audit_opinion": "确认失效归档",
        "reason": "【广州校区研究型RPg】地球、海洋与大气科学MPhil属于学术研究型硕士，状态置为False正常。",
        "replaced_by": "清水湾在招参考: ID 3225 (环境健康与安全)",
        "prog_type": "研究型 (RPg / MPhil)",
        "campus": "广州校区 (HKUST-GZ)"
    }
}

def build_hkust_false_review_workbook():
    print("[*] 读取数据库切片数据...")
    df = pd.read_csv(slice_csv)
    df_false = df[df['is_valid'].astype(str).str.lower() == 'false'].copy().reset_index(drop=True)
    print(f"[*] 成功筛选出失效专业 (is_valid=False) 共 {len(df_false)} 条记录")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "HKUST_失效专业人工核对表(21项)"
    ws.views.sheetView[0].showGridLines = True
    ws.row_dimensions[1].height = 30

    headers = [
        "序号",
        "数据库ID",
        "校区归属",
        "人工核对结论 (可下拉选择)",
        "预审核对建议与依据",
        "替代在招专业 / 关联主专业ID",
        "专业培养类型",
        "院校中文名",
        "院校英文名",
        "专业中文名",
        "专业英文名",
        "学位类型",
        "学制",
        "现有学费",
        "现有申请时间",
        "现有语言要求",
        "官网网址",
        "库中有效性 (is_valid)",
        "人工核对备注 (供填写)"
    ]
    ws.append(headers)

    # 表头样式
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        # 前 7 列为人工核对与关键属性列，使用红色高亮表头；第 19 列为人工备注列
        if col_idx in [3, 4, 5, 6, 7, 19]:
            cell.fill = AUDIT_HEADER_FILL
        else:
            cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER

    # 数据行填充
    for idx, r in df_false.iterrows():
        db_id = int(r['id'])
        analysis = ANALYSIS_MAP.get(db_id, {
            "audit_opinion": "确认失效归档",
            "reason": "历史停办或过渡归档记录",
            "replaced_by": "无",
            "prog_type": "授课型 (TPg)",
            "campus": "清水湾本部"
        })

        row_vals = [
            idx + 1,
            db_id,
            analysis["campus"],
            analysis["audit_opinion"],
            analysis["reason"],
            analysis["replaced_by"],
            analysis["prog_type"],
            str(r.get('school_name_cn', '') or ''),
            str(r.get('school_name_en', '') or ''),
            str(r.get('program_name_cn', '') or ''),
            str(r.get('program_name_en', '') or ''),
            str(r.get('degree', '') or ''),
            str(r.get('program_duration', '') or ''),
            str(r.get('program_tuition', '') or ''),
            str(r.get('application_time', '') or '').replace('<br/>', '\n').strip(),
            str(r.get('language_requirements', '') or '').replace('<br/>', '\n').strip(),
            str(r.get('program_website', '') or ''),
            "False (已失效)",
            ""  # 预留给用户填写的备注栏
        ]
        ws.append(row_vals)

        cur_row = ws.max_row
        ws.row_dimensions[cur_row].height = 28

        is_warn = "建议解除失效" in analysis["audit_opinion"]
        is_gz = "广州校区" in analysis["campus"]

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=cur_row, column=col_idx)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER

            # 居中列
            if col_idx in [1, 2, 3, 4, 7, 12, 13, 18]:
                cell.alignment = ALIGN_CENTER
            else:
                cell.alignment = ALIGN_LEFT

            # 特殊高亮
            if is_warn:
                cell.fill = WARN_FILL
                if col_idx == 4:
                    cell.font = RED_BOLD_FONT
            elif is_gz and col_idx == 3:
                cell.fill = GZ_FILL
            elif col_idx == 4:
                cell.font = GREEN_BOLD_FONT
                cell.fill = PASS_FILL
            elif col_idx == 18:
                cell.fill = INACTIVE_FILL

    # 添加下拉数据验证（Data Validation）供用户核对时选择
    dv = DataValidation(
        type="list",
        formula1='"确认失效归档,⚠️建议解除失效/激活在招,已合并至主专业,已升级更名,待进一步核对"',
        allow_blank=True
    )
    ws.add_data_validation(dv)
    dv.add(f"D2:D{ws.max_row}")

    # 自动列宽调整
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            first_line = val_str.split('\n')[0]
            line_len = sum(2 if ord(c) > 127 else 1 for c in first_line)
            max_len = max(max_len, line_len)
        adjusted_width = min(max(max_len + 3, 11), 45)
        # 针对核对依据与原因加宽显示
        if col_letter in ['E', 'F']:
            adjusted_width = 40
        elif col_letter in ['J', 'K']:
            adjusted_width = 35
        elif col_letter == 'D':
            adjusted_width = 24
        ws.column_dimensions[col_letter].width = adjusted_width

    ws.auto_filter.ref = f"A1:S{ws.max_row}"

    # 保存文件
    os.makedirs(os.path.dirname(desktop_file), exist_ok=True)
    os.makedirs(os.path.dirname(step1_folder_file), exist_ok=True)

    wb.save(desktop_file)
    wb.save(desktop_sub_file)
    wb.save(step1_folder_file)

    print("\n[OK] 香港科技大学 is_valid=False 人工核对专用表已生成！")
    print(f"  1. 桌面人工核对专用件: {desktop_file}")
    print(f"  2. 桌面标准命名件:     {desktop_sub_file}")
    print(f"  3. 步骤一附表同步件:   {step1_folder_file}")
    print(f"  - 总计记录: {len(df_false)} 条")
    print("  - 特色功能: 内置人工核对下拉菜单、事实原因分析、替代主专业ID映射、重点关注黄色预警高亮！")

if __name__ == "__main__":
    build_hkust_false_review_workbook()
