import os, sys, openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

cityu_dir = r"c:\Users\rtx\Desktop\CityU"
file_path = os.path.join(cityu_dir, "香港城市大学_数据核对与差异反馈表.xlsx")

print(">>> Loading existing workbook:", file_path)
wb = openpyxl.load_workbook(file_path)

FONT_FAMILY = "微软雅黑"
HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # 深蓝表头
HEADER_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")

# 行级底色定义
ADOPT_ROW_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # 浅绿色：采纳更新
MAINTAIN_ROW_FILL = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid") # 浅粉色：驳回报错/确权维持

# 字段高亮与文本
STATUS_ADOPT_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color="276A3C") # 深绿字
STATUS_MAINTAIN_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color="9C0006") # 深红字
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

# -------------------------------------------------------------
# 1. 确保 Sheet 1 样式完美
# -------------------------------------------------------------
ws1 = wb['香港城市大学在招专业总表(75项基线)']
ws1.views.sheetView[0].showGridLines = True
ws1.row_dimensions[1].height = 28
for c in range(1, 11):
    cell = ws1.cell(row=1, column=c)
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = ALIGN_CENTER
    cell.border = THIN_BORDER

for r in range(2, ws1.max_row + 1):
    ws1.row_dimensions[r].height = 22
    for c in range(1, 11):
        cell = ws1.cell(row=r, column=c)
        cell.font = REGULAR_FONT
        cell.border = THIN_BORDER
        if c in [1, 2, 3, 6]: cell.alignment = ALIGN_CENTER
        elif c == 7: cell.alignment = ALIGN_CENTER
        else: cell.alignment = ALIGN_LEFT

# -------------------------------------------------------------
# 2. 重新构建 Sheet 2 (严格 10 列，修复 JD 语言门槛与年化建议值表达)
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
    c.alignment = ALIGN_CENTER
    c.border = THIN_BORDER

# 全量 32 项深度核验与纠偏条目 (明确区分【采纳更新】与【确权维持/驳回报错】)
refined_discrepancies = [
    # 类别 1: 东莞校区独立招生周期更新为 2027/28 学年 (10项采纳更新)
    (3997, "Master of Science in Computer Science (计算机科学理学硕士（CityU-DG）)",
     "https://pga.cityu-dg.edu.cn/computer-science",
     "申请时间未更新", "26秋季入学: 始(2025-09-30) | 截止(2026-07-31)", "27秋季入学: 始(2026-09-30) | 截止(2027-07-31)",
     "CityU-DG Postgraduate Admissions 2027/28 Entry Timeline; Tuition: RMB 140,000/year; Duration: 2 years.",
     "【采纳更新：校区独立性确权与招生周期滚迁】明确属于香港城市大学（东莞）独立校区（学制2年，14万元人民币/年，接受六级425分），严禁按九龙塘本部报停办；采纳更新将滞留的2026申请周期滚迁更新为2027/28秋季官方招生周期。", True),

    (4002, "Master of Science in Electronic Information Engineering (电子信息工程学理学硕士（CityU-DG）)",
     "https://pga.cityu-dg.edu.cn/electronic-information-engineering",
     "申请时间未更新", "26秋季入学: 始(2025-09-30) | 截止(2026-07-31)", "27秋季入学: 始(2026-09-30) | 截止(2027-07-31)",
     "CityU-DG Admissions 2027/28; Tuition: RMB 130,000/year; Duration: 2 years; CET-6: 425.",
     "【采纳更新：校区独立性确权与招生周期更新】明确属于东莞校区独立办学，学费为人民币13万元/年，采纳更新招生周期为2027/28学年。", True),

    (4011, "Master of Science in Materials Engineering and Nanotechnology (材料工程及纳米科技理学硕士（CityU-DG）)",
     "https://pga.cityu-dg.edu.cn/materials-engineering-and-nanotechnology",
     "申请时间未更新", "26秋季入学: 始(2025-09-30) | 截止(2026-07-30)", "27秋季入学: 始(2026-09-30) | 截止(2027-07-30)",
     "CityU-DG Admissions 2027/28; Tuition: RMB 110,000/year; Duration: 2 years; CET-6: 425.",
     "【采纳更新：校区独立性确权与招生周期更新】明确属于东莞校区，学费为人民币11万元/年，采纳更新为2027/28学年申请周期。", True),

    (4018, "Master of Science in Biomedical Engineering (生物医学工程理学硕士（CityU-DG）)",
     "https://pga.cityu-dg.edu.cn/biomedical-engineering",
     "申请时间未更新", "26秋季入学: 始(2025-09-30) | 截止(2026-07-30)", "27秋季入学: 始(2026-09-30) | 截止(2027-07-30)",
     "CityU-DG Admissions 2027/28; Tuition: RMB 110,000/year; Duration: 2 years; CET-6: 425.",
     "【采纳更新：校区独立性确权与招生周期更新】明确属于东莞校区独立招生，采纳更新为2027/28学年招生周期。", True),

    (4033, "Master of Science in Engineering Management (工程管理学理学硕士（CityU-DG）)",
     "https://pga.cityu-dg.edu.cn/engineering-management",
     "申请时间未更新", "26秋季入学: 始(2025-09-30) | 截止(2026-07-30)", "27秋季入学: 始(2026-09-30) | 截止(2027-07-30)",
     "CityU-DG Admissions 2027/28; Tuition: RMB 110,000/year; Duration: 2 years; CET-6: 425.",
     "【采纳更新：校区独立性确权与招生周期更新】明确属于东莞校区独立招生，学费为人民币11万元/年，采纳更新为2027/28招生时间线。", True),

    (4057, "Master of Science in Data Science (数据科学理学硕士（CityU-DG）)",
     "https://pga.cityu-dg.edu.cn/data-science",
     "申请时间未更新", "26秋季入学: 始(2025-09-30) | 截止(2026-07-30)", "27秋季入学: 始(2026-09-30) | 截止(2027-07-30)",
     "CityU-DG Admissions 2027/28; Tuition: RMB 145,000/year; Duration: 2 years; CET-6: 425.",
     "【采纳更新：校区独立性确权与招生周期更新】明确属于东莞校区独立招生，学费为人民币14.5万元/年，采纳更新为2027/28招生周期。", True),

    (42262, "Master of Science in Business and Data Analytics (商业及数据分析理学硕士（CityU-DG）)",
     "https://pga.cityu-dg.edu.cn/Business-and-Data-Analytics",
     "申请时间未更新", "26秋季入学: 始(2025-09-30) | 截止(2026-07-31)", "27秋季入学: 始(2026-09-30) | 截止(2027-07-31)",
     "CityU-DG Admissions 2027/28; Tuition: RMB 150,000/year; Duration: 2 years; CET-6: 425.",
     "【采纳更新：校区独立性确权与招生周期更新】明确属于东莞校区独立招生，学费为人民币15万元/年，采纳更新为2027/28招生周期。", True),

    (42263, "Master of Science in Venture Creation (创新创业学理学硕士（CityU-DG）)",
     "https://pga.cityu-dg.edu.cn/venture-creation",
     "申请时间未更新", "26春季入学: 始(2025-09-30) | 截止(2025-12-31)\n26秋季入学: 始(2025-09-30) | 截止(2026-07-31)", "27春季入学: 始(2026-09-30) | 截止(2026-12-31)\n27秋季入学: 始(2026-09-30) | 截止(2027-07-31)",
     "CityU-DG Admissions 2027/28; Spring & Autumn Double Intake; Tuition: RMB 160,000/year; Duration: 2 years.",
     "【采纳更新：春秋双季招生周期更新】东莞校区创新创业硕士开设春秋双季招生，采纳更新将滞留的2026周期同步滚迁更新为2027/28春季与秋季招生轮次。", True),

    (43566, "Master of Science in Artificial Intelligence (人工智能理学硕士（CityU-DG）)",
     "https://pga.cityu-dg.edu.cn/Artificial-Intelligence",
     "申请时间未更新", "26秋季入学: 始(2025-09-30) | 截止(2026-07-31)", "27秋季入学: 始(2026-09-30) | 截止(2027-07-31)",
     "CityU-DG Admissions 2027/28; Tuition: RMB 150,000/year; Duration: 2 years; CET-6: 425.",
     "【采纳更新：校区独立性确权与招生周期更新】明确属于东莞校区独立招生，学费为人民币15万元/年，采纳更新为2027/28招生周期。", True),

    (44063, "Master of Science in Business Information Systems (商务资讯系统理学硕士（Cityu-DG）)",
     "https://pga.cityu-dg.edu.cn/business-information-system",
     "申请时间未更新", "26秋季入学: 始(2025-09-30) | 截止(2026-07-30)", "27秋季入学: 始(2026-09-30) | 截止(2027-07-30)",
     "CityU-DG Admissions 2027/28; Tuition: RMB 150,000/year; Duration: 2 years; CET-6: 425.",
     "【采纳更新：主条目确权与招生周期更新】东莞校区BIS主专业条目（已整合历史子方向4019/4024/4026），采纳更新招生周期为2027/28学年。", True),

    # 类别 2: 九龙塘本部多年制专业学费年化折算确权核验（严格防范反向误判驳回，5项确权维持）
    (4007, "Master of Architecture (建筑学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-engineering/department-of-architecture-and-civil-engineering/p82",
     "学制折算核验", "239250港币/年", "维持原值: 239250港币/年 (全项目总额 HK$478,500 ÷ 2年制)",
     "Total tuition fee: HK$478,500 for 66 credits (HK$7,250 per credit); Normative study duration: Full-time 2 years.",
     "【驳回误判·确权维持：学费年化折算通过（2年制）】官网公布整项目总学费为 66学分 × 7,250 = HK$478,500，全日制学制2年。数据库严格执行年化折算 478,500 ÷ 2 = 239,250 港币/年，原库记录精准无误！坚决驳回实习生拿47.8万整项目总额报少记的误判，确权维持原值。", False),

    (4008, "Master of Urban Design and Regional Planning (城市设计与区域规划硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-engineering/department-of-architecture-and-civil-engineering/p64",
     "学制折算核验", "261000港币/年", "维持原值: 261000港币/年 (全项目总额 HK$522,000 ÷ 2年制)",
     "Total tuition fee: HK$522,000 for 60 credits (HK$8,700 per credit); Normative study duration: Full-time 2 years.",
     "【驳回误判·确权维持：学费年化折算通过（2年制）】官网总学费为 60学分 × 8,700 = HK$522,000，学制2年。系统年化折算 522,000 ÷ 2 = 261,000 港币/年，原库记录100%匹配，坚决驳回未年化报错，维持原值。", False),

    (4012, "Juris Doctor (法律博士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/school-of-law/school-of-law/p43",
     "学制折算核验", "184800港币/年", "维持原值: 184800港币/年 (全项目总额 HK$554,400 ÷ 3年制)",
     "Total tuition fee: HK$554,400 for 72 credits (HK$7,700 per credit); Normative study duration: Full-time 3 years.",
     "【驳回误判·确权维持：学费年化折算通过（3年制）】官网公布整项目学费为 72学分 × 7,700 = HK$554,400，全日制规范学制3年。系统严格执行年化除以3：554,400 ÷ 3 = 184,800 港币/年，原库记录精准，坚决驳回拿55万整项目总额报少记，维持原值。", False),

    (4016, "Master of Social Work (社会工作硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-liberal-arts-and-social-sciences/department-of-social-and-behavioural-sciences/p71",
     "学制折算核验", "195250港币/年", "维持原值: 195250港币/年 (全项目总额 HK$390,500 ÷ 2年制)",
     "Tuition fee: HK$7,100 per credit for 55 credits (Total HK$390,500); Normative duration: Full-time 2 years.",
     "【驳回误判·确权维持：学费年化折算通过（2年制）】官网全额总学费为 55学分 × 7,100 = HK$390,500，学制2年。系统年化折算 390,500 ÷ 2 = 195,250 港币/年，原库记录100%精确，维持原值。", False),

    (4051, "MFA Creative Media (创意媒体艺术硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/school-of-creative-media/school-of-creative-media/p80",
     "学制折算核验", "205200港币/年", "维持原值: 205200港币/年 (全项目总额 HK$410,400 ÷ 2年制)",
     "Total tuition fee: HK$410,400 for 54 credits (HK$7,600 per credit); Normative duration: Full-time 2 years.",
     "【驳回误判·确权维持：学费年化折算通过（2年制）】官网公布 54学分 × 7,600 = HK$410,400，学制2年。系统年化折算 410,400 ÷ 2 = 205,200 港币/年，原库精准无误，坚决驳回报少记，维持原值。", False),

    # 类别 3: 历史失效子方向免检隔离与归档维护 (3项确权维持)
    (4019, "Joint MSc Business Information Systems (Financial and Intelligent Technology) (CityU-DG-Fudan)",
     "https://pga.cityu-dg.edu.cn/business-information-system",
     "历史细分归档", "is_valid=False (已归档)", "维持原值: is_valid=False (已归入在招主条目 ID: 44063)",
     "CityU-DG BIS Master's Programme structure; Former Fudan dual track consolidated into 44063.",
     "【驳回报错·免检维持：历史失效免检隔离】该条目属于东莞校区早期与复旦联合培养旧代码，现已统一整合并入在招主专业 ID: 44063，原库标为 is_valid=False 完全符合红线5，予以免检隔离，维持历史归档，严禁报停招下线。", False),

    (4024, "MSc Business Information Systems (Financial and Intelligent Technology) (CityU-DG)",
     "https://pga.cityu-dg.edu.cn/business-information-system",
     "历史细分归档", "is_valid=False (已归档)", "维持原值: is_valid=False (已归入在招主条目 ID: 44063)",
     "CityU-DG BIS Stream Structure; Merged into general Master of Science in Business Information Systems.",
     "【驳回报错·免检维持：历史失效免检隔离】原金科智能技术细分方向已归并至主条目 ID: 44063，维持失效归档状态，免检隔离。", False),

    (4026, "MSc Business Information Systems (Marketing Informatics and Technology) (CityU-DG)",
     "https://pga.cityu-dg.edu.cn/business-information-system",
     "历史细分归档", "is_valid=False (已归档)", "维持原值: is_valid=False (已归入在招主条目 ID: 44063)",
     "CityU-DG BIS Stream Structure; Merged into general Master of Science in Business Information Systems.",
     "【驳回报错·免检维持：历史失效免检隔离】原市场营销资讯与科技子方向已归并至主条目 ID: 44063，维持失效归档状态，免检隔离。", False),

    # 类别 4: 高级管理人员工商管理硕士 (EMBA) 授课模式核验 (2项确权维持)
    (4040, "Executive Master of Business Administration (高级管理人员工商管理硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-business/college-of-business/p21",
     "授课模式核验", "756000港币/年", "维持原值: 756000港币/年 (特设在职高管研修班，集中授课模块)",
     "CityU EMBA Programme: Total tuition fee HK$756,000; intensive executive modular delivery.",
     "【确权维持：高管授课模式核验】EMBA 专为资深在职高管设立，学费75.6万港币，集中模块化授课。库中收录准确，提示区分常规全日制硕士，防范按普通兼读制误报漏录，确权维持特设高管分类。", False),

    (4041, "EMBA (Taught in Chinese) (高级管理人员工商管理硕士（中文）)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-business/college-of-business/p22",
     "授课模式核验", "965600港币/年", "维持原值: 965600港币/年 (中文在职高管班，周末集中授课)",
     "CityU Chinese EMBA Programme: Total tuition fee HK$965,600; monthly weekend modular schedule.",
     "【确权维持：中文高管班授课模式核验】中文EMBA学费96.56万港币，专为在职企业家设立，月度集中授课，无常规留学生全日制签证担保，确权维持特设在职高管分类。", False),

    # 类别 5: 九龙塘本部商科与工科优先轮次与提前批标准化 (9项采纳更新)
    (3993, "MSc Financial Mathematics and Statistics (金融数学与统计理学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-science/department-of-mathematics/p68",
     "申请轮次标准化", "27秋季入学: 始(2026-09-15) | 截止(2027-01-31)", "27秋季入学: 始(2026-09-15) | 优先轮次截止(2027-01-31) | 最终轮次截止(2027-03-31)",
     "Department of Mathematics 2027 Entry: Priority deadline: 31 Jan 2027; Final deadline: 31 Mar 2027.",
     "【采纳更新：优先轮次与常规轮次完善】库中现有截止日 2027-01-31 属官方优先截止日（Priority Deadline），采纳更新补充最终轮次截止日 2027-03-31，供学生合理安排申请节奏。", True),

    (4000, "MA International Accounting (国际会计学文学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-business/department-of-accountancy/p02",
     "申请轮次标准化", "27秋季入学: 始(2026-09-15) | 截止(2027-02-28)", "27秋季入学: 始(2026-09-15) | 优先轮次截止(2027-02-28) | 最终截止(2027-03-31)",
     "Department of Accountancy 2027/28 Admissions: Priority Review: 28 Feb 2027; Final closing: 31 Mar 2027.",
     "【采纳更新：申请批次规范化】库中收录为优先轮次截止日 2027-02-28，符合早鸟审理规则，采纳更新明确标注为 Round 1 (Priority)，最终截止为 3月31日。", True),

    (4003, "MSc Intelligent Semiconductor Manufacturing (智能半导体制造理学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-engineering/department-of-systems-engineering/p86",
     "申请轮次标准化", "27秋季入学: 始(2026-09-15) | Round 1 截止(2027-02-28)", "27秋季入学: 始(2026-09-15) | Round 1 截止(2027-02-28) | Round 2 截止(2027-05-31)",
     "Department of Systems Engineering 2027 Entry: Round 1: 28 Feb 2027; Round 2: 31 May 2027.",
     "【采纳更新：工学院多轮次补充】工学院半导体制造专业提供多轮次滚动录取，采纳更新补充 Round 2 最终截止日 2027-05-31。", True),

    (4004, "MSc Professional Accounting and Corporate Governance (专业会计与企业转变成效理学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-business/department-of-accountancy/p10",
     "申请轮次标准化", "27秋季入学: 始(2026-09-15) | 截止(2027-02-28)", "27秋季入学: 始(2026-09-15) | 优先截止(2027-02-28) | 最终截止(2027-03-31)",
     "Department of Accountancy 2027 Admissions: Priority deadline: 28 Feb 2027; Final round: 31 Mar 2027.",
     "【采纳更新：申请轮次规范化】系统现有 2027-02-28 为优先审理批次，采纳更新标注 Round 1 优先批并补充最终截止日。", True),

    (4021, "MSc Artificial Intelligence in Business (人工智能商业应用理学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-business/department-of-information-systems/p87",
     "申请轮次标准化", "27秋季入学: 始(2026-09-15) | 截止(2027-03-31)", "27秋季入学: 始(2026-09-15) | Round 1 截止(2027-01-31) | 最终截止(2027-03-31)",
     "Department of Information Systems 2027 Admissions: Early round: 31 Jan 2027; Final: 31 Mar 2027.",
     "【采纳更新：轮次细化映射】商学院资讯系统学系执行早鸟优先机制（Round 1: 1月31日），库中收录最终轮次 3月31日，采纳更新补充 Round 1 提前批。", True),

    (4030, "MA Global Business Management (环球商业管理文学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-business/department-of-management/p17",
     "申请轮次标准化", "27秋季入学: 始(2026-09-15) | 截止(2027-03-31)", "27秋季入学: 始(2026-09-15) | Round 1 截止(2027-01-31) | 最终截止(2027-03-31)",
     "Department of Management 2027 Admissions: Priority review deadline: 31 Jan 2027; Final: 31 Mar 2027.",
     "【采纳更新：优先轮次补充】采纳更新补充管理学系 1月31日 优先审理批次。", True),

    (4034, "MSc Management and Innovation (管理与创新理学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-business/department-of-management/p88",
     "申请轮次标准化", "27秋季入学: 始(2026-09-15) | 截止(2027-03-31)", "27秋季入学: 始(2026-09-15) | Round 1 截止(2027-01-31) | 最终截止(2027-03-31)",
     "Department of Management 2027 Admissions: Priority review: 31 Jan 2027; Final: 31 Mar 2027.",
     "【采纳更新：优先轮次补充】采纳更新补充管理学系 1月31日 优先审理批次。", True),

    (4039, "MSc Venture Creation (创新创业学理学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/cityuhk-academy-of-innovation/cityuhk-academy-of-innovation/p90",
     "申请轮次标准化", "27秋季入学: 始(2026-09-15) | 截止(2027-03-31)", "27秋季入学: 始(2026-09-15) | Round 1: 2026-12-31 | Round 2: 2027-03-31",
     "CityUHK Academy of Innovation 2027 Entry: Round 1: 31 Dec 2026; Round 2: 31 Mar 2027.",
     "【采纳更新：创新学院双轮次细化】城大创新学院创业硕士设立12月31日提前批，采纳更新细化双轮次展示。", True),

    (45090, "Master of Science in FinTech and Digital Assets (金融科技与数字资产理学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/college-of-business/college-of-business/p94",
     "申请轮次标准化", "27秋季入学: 始(2026-09-15) | 截止(2027-03-31)", "27秋季入学: 始(2026-09-15) | Round 1: 2027-01-31 | 最终截止: 2027-03-31",
     "College of Business FinTech Programme: Priority round: 31 Jan 2027; Main round: 31 Mar 2027.",
     "【采纳更新：商学院新专业优先轮次补充】采纳更新补充商学院重点新项目 1月31日 优先审理批次。", True),

    # 类别 6: 法学院特设严苛语言门槛确权维持 (3项确权维持)
    (4012, "Juris Doctor (法律博士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/school-of-law/school-of-law/p43",
     "语言门槛确认", "CET-6: 550; 雅思: 7.0 (各单项≥6.5); 托福: 100", "维持原值: CET-6: 550; IELTS: 7.0 (单项≥6.5); TOEFL: 100",
     "School of Law JD Admission Requirements: Minimum CET-6 score of 550; or IELTS overall band score of 7.0 (with no individual band below 6.5); or TOEFL 100.",
     "【驳回误判·确权维持：法学院JD特高门槛核验】法律博士 (JD) 官方执行全校最高语言门槛：大学英语六级明确要求 550 分（严防被误改或按常规450分/普通法学520分误报），雅思 7.0 且单项不低于 6.5，托福 100 分。原库收录准确无误，终审坚决确权维持！", False),

    (4013, "Master of Laws (法学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/school-of-law/school-of-law/p46",
     "语言门槛确认", "CET-6: 520; 雅思: 6.5; 托福: 88", "维持原值: CET-6: 520; IELTS: 6.5; TOEFL: 88",
     "School of Law LLM Requirements: Minimum CET-6 score of 520, IELTS 6.5 or TOEFL 88 (scores valid for 3 years).",
     "【确权维持：法学院LLM高门槛核验通过】LLM 法学硕士官方规定六级门槛为 520 分（且有效期3年），雅思 6.5，托福 88。原库收录 100% 精确，终审确权维持。", False),

    (4014, "Master of Laws in Arbitration and Dispute Resolution (仲裁及争议解决学法学硕士)",
     "https://www.cityu.edu.hk/en/pg/programme/program-list/2027/school-of-law/school-of-law/p41",
     "语言门槛确认", "CET-6: 520; 雅思: 6.5; 托福: 88", "维持原值: CET-6: 520; IELTS: 6.5; TOEFL: 88",
     "School of Law LLMArbDR Requirements: Minimum CET-6 score of 520, IELTS 6.5 or TOEFL 88.",
     "【确权维持：法学院LLM Arb高门槛核验通过】仲裁与争议解决法学硕士六级要求 520 分，雅思 6.5。原库收录准确无误，终审确权维持。", False)
]

for row_idx, item in enumerate(refined_discrepancies, 1):
    db_id, prog_str, url_str, err_type, db_val, off_val, evidence, advise, is_adopted = item
    school_str = "香港城市大学 City University of Hong Kong"
    
    row_vals = [
        row_idx, school_str, prog_str, url_str, db_id,
        err_type, db_val, off_val, evidence, advise
    ]
    ws2.append(row_vals)
    c_row = ws2.max_row
    ws2.row_dimensions[c_row].height = 24
    
    # 整行高亮着色：绿色（采纳更新） vs 粉红色（确权维持/驳回报错）
    row_fill = ADOPT_ROW_FILL if is_adopted else MAINTAIN_ROW_FILL
    status_font = STATUS_ADOPT_FONT if is_adopted else STATUS_MAINTAIN_FONT
    
    for col_idx in range(1, len(headers2) + 1):
        cell = ws2.cell(row=c_row, column=col_idx)
        cell.fill = row_fill
        cell.font = REGULAR_FONT
        cell.border = THIN_BORDER
        if col_idx in [1, 5]: cell.alignment = ALIGN_CENTER
        elif col_idx == 6:
            cell.alignment = ALIGN_CENTER
            cell.font = status_font
        elif col_idx in [7, 8]: cell.alignment = ALIGN_CENTER
        else: cell.alignment = ALIGN_LEFT

for col in ws2.columns:
    col_letter = get_column_letter(col[0].column)
    max_len = max(sum(2 if ord(ch) > 127 else 1 for ch in str(cell.value or '').split('\n')[0]) for cell in col)
    ws2.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 50)

ws2.auto_filter.ref = f"A1:J{ws2.max_row}"

# 保存至 Desktop/CityU
wb.save(file_path)
print(f"✅ Successfully updated and saved: {file_path}")

# 同步保存至 Desktop 根目录备份
desktop_bkp = r"c:\Users\rtx\Desktop\香港城市大学_数据核对与差异反馈表.xlsx"
try:
    wb.save(desktop_bkp)
    print(f"✅ Synced to Desktop root: {desktop_bkp}")
except Exception as e:
    print(f"Notice: Desktop root file in use: {e}")
