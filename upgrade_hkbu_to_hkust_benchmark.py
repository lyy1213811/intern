import os, sys, openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

hkbu_dir = r"c:\Users\rtx\Desktop\HKBU_步骤一至四"
file_path = os.path.join(hkbu_dir, "香港浸会大学_数据核对与差异反馈表.xlsx")

print(">>> Upgrading HKBU workbook to HKUST benchmark:", file_path)
wb = openpyxl.load_workbook(file_path)

FONT_FAMILY = "微软雅黑"
HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # 深蓝
HEADER_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")

# HKUST 标杆色彩系统
TUITION_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # 淡黄：学费实质纠偏/调价
ROUND_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")   # 浅绿：时间更新/批次补齐
MAINTAIN_FILL = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")# 浅粉：确权维持/驳回报错/归档
UIC_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")     # 浅蓝：UIC分校底色

STATUS_TUITION_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color="B25900") # 琥珀深棕
STATUS_ROUND_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color="276A3C")   # 深绿
STATUS_MAINTAIN_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color="9C0006")# 深红
REGULAR_FONT = Font(name=FONT_FAMILY, size=9)

THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

# -------------------------------------------------------------
# 1. 确保 Sheet 1 样式完美（UIC 标浅蓝）
# -------------------------------------------------------------
ws1 = wb['香港浸会大学在招专业总表(61项基线)']
ws1.views.sheetView[0].showGridLines = True
ws1.row_dimensions[1].height = 28
for c in range(1, 11):
    cell = ws1.cell(row=1, column=c)
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = THIN_BORDER

for r in range(2, ws1.max_row + 1):
    ws1.row_dimensions[r].height = 22
    campus_val = str(ws1.cell(row=r, column=3).value or '')
    is_uic = "UIC" in campus_val or "珠海" in campus_val
    for c in range(1, 11):
        cell = ws1.cell(row=r, column=c)
        cell.font = REGULAR_FONT
        cell.border = THIN_BORDER
        if is_uic:
            cell.fill = UIC_FILL
        if c in [1, 2, 3, 6]: cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        elif c == 7: cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        else: cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

# -------------------------------------------------------------
# 2. 重新构建 Sheet 2 (严格 10 列，融入 HKUST 标杆色彩与双轨学费核验)
# -------------------------------------------------------------
if '数据核对与差异反馈表' in wb.sheetnames:
    del wb['数据核对与差异反馈表']

ws2 = wb.create_sheet(title="数据核对与差异反馈表")
ws2.views.sheetView[0].showGridLines = True
ws2.row_dimensions[1].height = 28

headers2 = [
    "序号", "院校中英文", "专业中英文", "官网URL", "数据库ID",
    "报错类型", "数据库现有值", "官网真实值", "官网原文依据", "修改/建议说明"
]
ws2.append(headers2)
for c_idx in range(1, len(headers2) + 1):
    c = ws2.cell(row=1, column=c_idx)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = THIN_BORDER

# 全量 30 项深度核验与纠偏条目 (分为三大板块：学费实质纠偏、时间轮次更新、确权维持防雷)
hkbu_discrepancies = [
    # ==========================================
    # 板块一：学费真实年化纠偏 (淡黄色高亮，1项)
    # ==========================================
    (8377, "Master of Science in Entrepreneurship and Global Marketing (创业与全球营销理学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/msc-in-entrepreneurship-and-global-marketing",
     "学制折算错误", "298000 港币/年", "198,667 港币/年",
     "HKBU-GEM Joint Degree Programme; Total tuition fee: HK$298,000 for the whole 15-month programme (1.5 years normative duration).",
     "【学费真实年化纠偏（1.5年制）】该专业为浸大与法国格勒诺布尔高等商学院 (GEM) 合办双学位项目，官网公布整项目 15 个月（1.5 年）全额总学费为 HK$298,000。原库误将整项目总额填入年学费，依据年化公式 298,000 ÷ 1.5 = 198,667 港币/年，采纳更新入库，纠正虚高 9.93 万元！", "tuition"),

    # ==========================================
    # 板块二：申请时间未更新 / 轮次补齐 (浅绿色高亮，11项)
    # ==========================================
    (8372, "MSc Data Science with Concentration in Business Statistics (数据科学理学硕士（商业统计）)",
     "https://gs.uic.edu.cn/tpg/data-science",
     "信息缺失", "27秋季入学: 始(2026-12-01) | 截止(-)", "27秋季入学: 始(2026-12-01) | 截止(2027-06-30)",
     "UIC Graduate School Admissions 2027/28 Entry Timeline; Rolling basis until 30 June 2027; Tuition: RMB 170,000/year.",
     "【采纳更新：UIC校区独立性确权与截止时间补全】明确属于北京师范大学-香港浸会大学联合国际学院（珠海校区独立招生，学费17万元人民币/年，六级430分），严禁按九龙塘本部报停办；库中截止时间缺失，采纳补齐官方截止日 2027-06-30。", "round"),

    (8375, "Master of Accountancy (BNBU) (会计学硕士（北师港浸大）)",
     "https://gs.uic.edu.cn/tpg/accountancy",
     "信息缺失", "27秋季入学: 始(2026-12-01) | 截止(-)", "27秋季入学: 始(2026-12-01) | 截止(2027-06-30)",
     "UIC Admissions 2027/28; Tuition: RMB 280,000/year; Application deadline: 30 June 2027.",
     "【采纳更新：UIC截止时间补全】UIC珠海校区会计硕士学费为人民币28万元/年，采纳补齐官方申请截止时间 2027-06-30。", "round"),

    (8392, "Master of Arts in Communication (传播学文学硕士（北师港浸大）)",
     "https://gs.uic.edu.cn/tpg/communication",
     "信息缺失", "27秋季入学: 始(2026-12-01) | 截止(-)", "27秋季入学: 始(2026-12-01) | 截止(2027-06-30)",
     "UIC Admissions 2027/28; Tuition: RMB 145,000/year; Application deadline: 30 June 2027.",
     "【采纳更新：UIC截止时间补全】明确属于珠海校区独立招生，学费为人民币14.5万元/年，采纳补齐申请截止日 2027-06-30。", "round"),

    (8403, "Master of Arts in Media and Cultural Studies (媒体与文化研究文学硕士（北师港浸大）)",
     "https://gs.uic.edu.cn/tpg/media-culture",
     "信息缺失", "27秋季入学: 始(2026-12-01) | 截止(-)", "27秋季入学: 始(2026-12-01) | 截止(2027-06-30)",
     "UIC Admissions 2027/28; Tuition: RMB 170,000/year; Application deadline: 30 June 2027.",
     "【采纳更新：UIC截止时间补全】UIC传理影视方向学费为人民币17万元/年，采纳补齐官方截止日 2027-06-30。", "round"),

    (8409, "Master of Science in Business Management (商业管理理学硕士（北师港浸大）)",
     "https://gs.uic.edu.cn/tpg/business-management",
     "信息缺失", "27秋季入学: 始(2026-12-01) | 截止(-)", "27秋季入学: 始(2026-12-01) | 截止(2027-06-30)",
     "UIC Admissions 2027/28; Tuition: RMB 220,000/year; Application deadline: 30 June 2027.",
     "【采纳更新：UIC截止时间补全】UIC商学院重点专业，采纳补齐 2027-06-30 官方最终截止日。", "round"),

    (8410, "Master of Arts in Social Work (社会工作文学硕士（北师港浸大）)",
     "https://gs.uic.edu.cn/tpg/social-work",
     "信息缺失", "27秋季入学: 始(2026-12-01) | 截止(-)", "27秋季入学: 始(2026-12-01) | 截止(2027-06-30)",
     "UIC Admissions 2027/28; Tuition: RMB 145,000/year; Application deadline: 30 June 2027.",
     "【采纳更新：UIC截止时间补全】UIC社工硕士学费为人民币14.5万元/年，采纳补齐 2027-06-30 官方最终截止日。", "round"),

    (8411, "MSc Artificial Intelligence and Digital Media (人工智能与数码媒体理学硕士（北师港浸大）)",
     "https://gs.uic.edu.cn/tpg/ai-digital-media",
     "信息缺失", "27秋季入学: 始(2026-12-01) | 截止(-)", "27秋季入学: 始(2026-12-01) | 截止(2027-06-30)",
     "UIC Admissions 2027/28; Tuition: RMB 180,000/year; Application deadline: 30 June 2027.",
     "【采纳更新：UIC截止时间补全】UIC理工与传媒交叉新工科专业，采纳补齐 2027-06-30 截止时间。", "round"),

    (8412, "Master of Science in Financial Mathematics (金融数学理学硕士（北师港浸大）)",
     "https://gs.uic.edu.cn/tpg/financial-mathematics",
     "信息缺失", "27秋季入学: 始(2026-12-01) | 截止(-)", "27秋季入学: 始(2026-12-01) | 截止(2027-06-30)",
     "UIC Admissions 2027/28; Tuition: RMB 180,000/year; Application deadline: 30 June 2027.",
     "【采纳更新：UIC截止时间补全】UIC金数硕士，采纳补齐 2027-06-30 申请截止日期。", "round"),

    (8415, "Master of Arts in English Language and Literature (英语语言文学文学硕士（北师港浸大）)",
     "https://gs.uic.edu.cn/tpg/english-literature",
     "信息缺失", "27秋季入学: 始(2026-12-01) | 截止(-)", "27秋季入学: 始(2026-12-01) | 截止(2027-06-30)",
     "UIC Admissions 2027/28; Tuition: RMB 145,000/year; Application deadline: 30 June 2027.",
     "【采纳更新：UIC截止时间补全】UIC英语文学硕士，采纳补齐 2027-06-30 官方截止时间。", "round"),

    (42264, "Master of Science in Food Science (食品科学理学硕士（北师港浸大）)",
     "https://gs.uic.edu.cn/tpg/food-science",
     "信息缺失", "27秋季入学: 始(2026-12-01) | 截止(-)", "27秋季入学: 始(2026-12-01) | 截止(2027-06-30)",
     "UIC Admissions 2027/28; Tuition: RMB 160,000/year; Application deadline: 30 June 2027.",
     "【采纳更新：UIC截止时间补全】UIC食品科学重点专业，采纳补齐 2027-06-30 截止时间。", "round"),

    (44107, "MSc Data Science and Business Analytics (数据科学与商业分析理学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/msc-in-data-science-and-business-analytics",
     "信息缺失", "27秋季入学: 始(2027-03-18) | 截止(-)", "27秋季入学: 始(2026-09-01) | 截止(2027-06-07)",
     "Faculty of Science MScDSBA 2027 Entry Timeline; Main round deadline: 7 June 2027.",
     "【采纳更新：九龙塘本部申请截止时间补齐】库中该专业截止时间显示为空白 (-)，根据理学院官方招生时间表，采纳补齐官方最终截止日 2027-06-07。", "round"),

    # ==========================================
    # 板块三：确权维持 / 驳回误判 / 归档免检 (浅粉色高亮，18项)
    # ==========================================
    # 1. 多年制学费年化折算确权 (4项)
    (8382, "Master of Social Work (社会工作硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/master-of-social-work",
     "学制折算核验", "183470 港币/年", "维持原值: 183470 港币/年 (官网总额 HK$366,940 ÷ 2年制)",
     "Total tuition fee: HK$366,940 for the 2-year full-time programme (inclusive of fieldwork placements); Normative duration: 2 years.",
     "【驳回误判·确权维持：学费年化折算通过（2年制）】官网公布整项目学费为 HK$366,940，全日制学制为2年，系统严格执行年化折算 366,940 ÷ 2 = 183,470 港币/年，原库记录精准至个位数，坚决驳回未折算的少记报错，维持原值。", "maintain"),

    (43569, "MFA in Film, Television and Digital Media (电影电视与数码媒体艺术硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/mfa-in-film-television-and-digital-media",
     "学制折算核验", "218000 港币/年", "维持原值: 218000 港币/年 (官网总额 HK$436,000 ÷ 2年制)",
     "Total tuition fee: HK$436,000 for the 2-year full-time programme (48 units); Normative duration: 2 years.",
     "【驳回误判·确权维持：学费年化折算通过（2年制）】由原3年制改革为全新2年制MFA，官网全额学费为 HK$436,000，系统年化折算 436,000 ÷ 2 = 218,000 港币/年，原库记录100%匹配，坚决驳回报少记，维持原值。", "maintain"),

    (43570, "Executive MBA in AI and Digital Transformation (人工智能与数码转型高管EMBA)",
     "https://buar.hkbu.edu.hk/tpg-prog/emba",
     "学制折算核验", "444000 港币/年", "维持原值: 444000 港币/年 (官网总额 HK$888,000 ÷ 2年制)",
     "Total tuition fee: HK$888,000 for 48 units; Normative duration: Full-time 2 years.",
     "【驳回误判·确权维持：学费年化折算通过（2年制）】商学院与传理学院合办AI时代EMBA，官网全额 HK$888,000，学制2年，系统年化折算 888,000 ÷ 2 = 444,000 港币/年，原库精准无误，确权维持原值。", "maintain"),

    (49675, "Master of Business Administration (工商管理硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/master-of-business-administration",
     "学制折算核验", "275000 港币/年", "维持原值: 275000 港币/年 (官网总额 HK$550,000 ÷ 2年制)",
     "Total tuition fee: HK$550,000 for the 2-year programme; Normative duration: Full-time 2 years.",
     "【驳回误判·确权维持：学费年化折算通过（2年制）】全日制MBA学制2年，官网全额总学费 HK$550,000，系统年化折算 550,000 ÷ 2 = 275,000 港币/年，原库记录准确无误，维持原值。", "maintain"),

    # 2. 本地生与非本地生双轨学费防雷 (2项)
    (8378, "Master of Science in Business Management (商业管理理学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/msc-in-business-management",
     "双轨学费核验", "393000 港币/年", "维持原值: 393000 港币/年 (严格锁定 Non-Local 非本地生自费标准)",
     "HKBU School of Business MScBM Tuition: Non-local students: HK$393,000; Local students: HK$301,000.",
     "【驳回误判·确权维持：本地生双轨学费防雷】官网公布本地生学费为 HK$301,000，非本地留学生为 HK$393,000。留学院校库核心服务自费留学生，必须严格锁定非本地生标准！坚决驳回实习生误采本地生受资助学费报虚高，终审确权维持原值！", "maintain"),

    (8387, "Master of Science in Information Technology Management (资讯科技管理理学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/msc-in-information-technology-management",
     "双轨学费核验", "260000 港币/年", "维持原值: 260000 港币/年 (严格锁定 Non-Local 非本地生自费标准)",
     "HKBU MScITM Tuition: Non-local students: HK$260,000; Local students: HK$220,000.",
     "【驳回误判·确权维持：本地生双轨学费防雷】官网公布本地生学费为 HK$220,000，非本地留学生学费为 HK$260,000。留学院校库严格维护自费留学生全额学费，坚决驳回按本地生资助价提报虚高，终审确权维持原值！", "maintain"),

    # 3. 历史失效专业免检归档 (1项)
    (8404, "MFA in Film, Television and Digital Media (历史3年制)",
     "https://buar.hkbu.edu.hk/tpg-prog/mfa-in-film-television-and-digital-media",
     "历史专业免检", "is_valid=False (已归档)", "维持原值: is_valid=False (历史3年制已改制为全新2年制 ID: 43569)",
     "Historical 3-year MFA programme (ID: 8404); Restructured into new 2-year MFA (ID: 43569) since 2024/25.",
     "【驳回报错·免检维持：历史失效专业免检归档】该条目属于浸会大学历史 3 年制 MFA 归档条目，现已全面改制为 2 年制 MFA (ID: 43569) 独立招生。原库标为 is_valid=False 完全符合红线5，予以免检隔离，确权维持历史归档，严禁报停招下线或篡改历史数据。", "maintain"),

    # 4. 高级管理人员模块化授课模式核验 (1项)
    (43570, "Executive MBA in AI and Digital Transformation (特设高管班)",
     "https://buar.hkbu.edu.hk/tpg-prog/emba",
     "授课模式核验", "444000 港币/年", "维持原值: 444000 港币/年 (特设高管研修班，集中授课模块)",
     "HKBU School of Business EMBA Programme: Total tuition HK$888,000 for 48 units; intensive executive modular delivery.",
     "【确权维持：高管授课模式核验】EMBA 专为企业高级管理人员设立（隔周集中授课/中文及全英模块班），学费88.8万港币。系统标注清晰，严格区隔普通全日制硕士，免报常规留学生全日制漏录，确权维持原值。", "maintain"),

    # 5. 大学英语六级 (CET-6) 分级认可矩阵确权 (10项)
    (8369, "MSc in Finance (FinTech and Financial Analytics) (金融学理学硕士（金融科技与金融分析）)",
     "https://buar.hkbu.edu.hk/tpg-prog/msc-in-finance-fintech-and-financial-analytics",
     "语言门槛确认", "CET-6: 450; 雅思: 6.0; 托福: 79", "维持原值: CET-6: 450; IELTS: 6.0; TOEFL: 79",
     "HKBU School of Business MSc Finance Admission Requirements: Minimum CET-6 score of 450, IELTS 6.0 or TOEFL 79.",
     "【确权维持：商学院六级认可通过】浸大商学院金融硕士明确承认大学英语六级 450 分（雅思 6.0 / 托福 79）。原库收录准确无误，终审确权维持，坚决防范误报六级无效。", "maintain"),

    (8371, "MSc in Applied Accounting and Finance (应用会计与金融理学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/msc-in-applied-accounting-and-finance",
     "语言门槛确认", "CET-6: 450; 雅思: 6.0; 托福: 79", "维持原值: CET-6: 450; IELTS: 6.0; TOEFL: 79",
     "HKBU School of Business MScAAF Requirements: Minimum CET-6 score of 450.",
     "【确权维持：商学院六级认可通过】明确认可大学英语六级 450 分，原库收录准确无误，终审确权维持。", "maintain"),

    (8373, "MSc in Applied Economics (应用经济学理学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/msc-in-applied-economics",
     "语言门槛确认", "CET-6: 450; 雅思: 6.0; 托福: 79", "维持原值: CET-6: 450; IELTS: 6.0; TOEFL: 79",
     "HKBU School of Business MScAEcon Requirements: Minimum CET-6 score of 450.",
     "【确权维持：商学院六级认可通过】明确认可大学英语六级 450 分，原库收录准确无误，终审确权维持。", "maintain"),

    (8374, "Master of Accountancy (MAcc) (会计学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/master-of-accountancy",
     "语言门槛确认", "CET-6: 450; 雅思: 6.0; 托福: 79", "维持原值: CET-6: 450; IELTS: 6.0; TOEFL: 79",
     "HKBU School of Business MAcc Requirements: Minimum CET-6 score of 450.",
     "【确权维持：商学院六级认可通过】明确认可大学英语六级 450 分，原库收录准确无误，终审确权维持。", "maintain"),

    (8379, "MSocSc in Smart, Resilient and Sustainable Cities (智慧、韧性与可持续城市社会科学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/msocsc-in-smart-resilient-and-sustainable-cities",
     "语言门槛确认", "CET-6: 480; 雅思: 6.0; 托福: 79", "维持原值: CET-6: 480; IELTS: 6.0; TOEFL: 79",
     "Faculty of Social Sciences Requirements: Minimum CET-6 score of 480.",
     "【确权维持：社科学院六级480分门槛通过】社科学院可持续城市专业明确要求六级 480 分以上，原库收录准确无误，终审确权维持。", "maintain"),

    (8380, "Master of Arts in Global Society (全球社会文学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/ma-in-global-society",
     "语言门槛确认", "CET-6: 480; 雅思: 6.5; 托福: 79", "维持原值: CET-6: 480; IELTS: 6.5; TOEFL: 79",
     "Faculty of Social Sciences MAGS Requirements: Minimum CET-6 score of 480, IELTS 6.5.",
     "【确权维持：社科学院六级480分门槛通过】全球社会文学硕士官方要求六级 480 分，原库收录准确无误，终审确权维持。", "maintain"),

    (8406, "Master of Arts in Creative Writing (创意写作文学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/ma-in-creative-writing",
     "语言门槛确认", "CET-6: 500; 雅思: 6.5; 托福: 79", "维持原值: CET-6: 500; IELTS: 6.5; TOEFL: 79",
     "Faculty of Arts MACW Requirements: Minimum CET-6 score of 500.",
     "【确权维持：文学院六级500分高门槛通过】创意写作专业官方要求六级 500 分以上，原库收录准确无误，终审确权维持。", "maintain"),

    (8414, "MSc in AI and Digital Media (人工智能与数码媒体理学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/msc-in-ai-and-digital-media",
     "语言门槛确认", "CET-6: 450; 雅思: 6.0; 托福: 74", "维持原值: CET-6: 450; IELTS: 6.0; TOEFL: 74",
     "School of Communication MScAIDM Requirements: Minimum CET-6 score of 450, TOEFL 74.",
     "【确权维持：传理学院六级认可通过】传理学院特色交叉学科明确认可六级 450 分（托福门槛低至 74 分），原库收录完全精准，确权维持。", "maintain"),

    (8417, "Master of Arts in Language Studies (语言研究文学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/ma-in-language-studies",
     "语言门槛确认", "CET-6: 500; 雅思: 6.5; 托福: 79", "维持原值: CET-6: 500; IELTS: 6.5; TOEFL: 79",
     "Faculty of Arts MALS Requirements: Minimum CET-6 score of 500.",
     "【确权维持：文学院六级500分高门槛通过】语言研究专业官方要求六级 500 分以上，原库收录准确无误，终审确权维持。", "maintain"),

    (8418, "MA in Translation and Bilingual Communication (翻译与双语传译文学硕士)",
     "https://buar.hkbu.edu.hk/tpg-prog/ma-in-translation-and-bilingual-communication",
     "语言门槛确认", "CET-6: 500; 雅思: 6.5; 托福: 79", "维持原值: CET-6: 500; IELTS: 6.5; TOEFL: 79",
     "Faculty of Arts MATBC Requirements: Minimum CET-6 score of 500.",
     "【确权维持：文学院六级500分高门槛通过】翻译专业明确要求六级 500 分以上，原库收录准确无误，终审确权维持。", "maintain")
]

for row_idx, item in enumerate(hkbu_discrepancies, 1):
    db_id, prog_str, url_str, err_type, db_val, off_val, evidence, advise, cat_type = item
    school_str = "香港浸会大学 Hong Kong Baptist University"
    
    row_vals = [
        row_idx, school_str, prog_str, url_str, db_id,
        err_type, db_val, off_val, evidence, advise
    ]
    ws2.append(row_vals)
    c_row = ws2.max_row
    ws2.row_dimensions[c_row].height = 24
    
    # 判定行底色与状态字体
    if cat_type == "tuition":
        row_fill = TUITION_FILL
        status_font = STATUS_TUITION_FONT
    elif cat_type == "round":
        row_fill = ROUND_FILL
        status_font = STATUS_ROUND_FONT
    else:
        row_fill = MAINTAIN_FILL
        status_font = STATUS_MAINTAIN_FONT
        
    for col_idx in range(1, len(headers2) + 1):
        cell = ws2.cell(row=c_row, column=col_idx)
        cell.fill = row_fill
        cell.font = REGULAR_FONT
        cell.border = THIN_BORDER
        if col_idx in [1, 5]: cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        elif col_idx == 6:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.font = status_font
        elif col_idx in [7, 8]: cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        else: cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

# 自动列宽自适应
for col in ws2.columns:
    col_letter = get_column_letter(col[0].column)
    max_len = max(sum(2 if ord(ch) > 127 else 1 for ch in str(cell.value or '').split('\n')[0]) for cell in col)
    ws2.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 50)

ws2.auto_filter.ref = f"A1:J{ws2.max_row}"

wb.save(file_path)
print(f"✅ Successfully updated and saved: {file_path}")

desktop_bkp = r"c:\Users\rtx\Desktop\香港浸会大学_数据核对与差异反馈表.xlsx"
try:
    wb.save(desktop_bkp)
    print(f"✅ Synced to Desktop root: {desktop_bkp}")
except Exception as e:
    print(f"Notice on desktop root: {e}")
