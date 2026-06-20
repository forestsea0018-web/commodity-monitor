#!/usr/bin/env python3
"""
生成半导体制造价值链标的Excel
按半导体工序顺序排列所有上市标的
"""

from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter

# ── 数据定义 ─────────────────────────────────────────────────────────────────
# 结构: (阶段, 步骤编号, 步骤名称, 类型, 说明, 标的名称, 标的代码, 市场, 细分行业, 币种)
SEMI_CHAIN = [
    # ── Stage 0: EDA设计 ──────────────────────────────────────────────────────
    ("0. Design EDA设计", "0", "EDA软件", "SEQ",
     "逻辑综合→布局布线→时序签核→GDSII",
     "Cadence Design Systems", "CDNS", "NASDAQ", "EDA软件", "USD"),
    ("0. Design EDA设计", "0", "EDA软件", "SEQ",
     "逻辑综合→布局布线→时序签核→GDSII",
     "Synopsys", "SNPS", "NASDAQ", "EDA软件", "USD"),

    # ── Stage 0.5: 掩膜制备 ───────────────────────────────────────────────────
    ("0.5 Mask 掩膜制备", "0.5", "掩膜写入+检测", "SEQ",
     "OPC修正→E-beam写入→掩膜检测→空白板供应",
     "Lasertec", "6920.T", "东京证交所", "掩膜检测设备", "JPY"),
    ("0.5 Mask 掩膜制备", "0.5", "掩膜写入+检测", "SEQ",
     "OPC修正→E-beam写入→掩膜检测→空白板供应",
     "JEOL", "6951.T", "东京证交所", "电子束写入设备", "JPY"),
    ("0.5 Mask 掩膜制备", "0.5", "掩膜写入+检测", "SEQ",
     "OPC修正→E-beam写入→掩膜检测→空白板供应",
     "Hoya Corporation", "7741.T", "东京证交所", "空白掩膜板材料", "JPY"),
    ("0.5 Mask 掩膜制备", "0.5", "掩膜写入+检测", "SEQ",
     "OPC修正→E-beam写入→掩膜检测→空白板供应",
     "信越化学 Shin-Etsu", "4063.T", "东京证交所", "空白掩膜板材料", "JPY"),

    # ── Stage 1: FEOL 前端晶体管形成 ×20-30层循环 ────────────────────────────
    # ①Coat
    ("1. FEOL 前端×20-30层", "①", "Coat 涂布", "LOOP×N",
     "光阻涂布 Resist Coat（TEL 93%垄断）",
     "Tokyo Electron / TEL", "8035.T", "东京证交所", "涂布/显影设备", "JPY"),

    # ②Expose
    ("1. FEOL 前端×20-30层", "②", "Expose 曝光", "LOOP×N",
     "光刻曝光 Lithography (EUV/DUV/i-line)",
     "ASML Holding", "ASML", "NASDAQ", "光刻机 EUV/DUV", "USD"),
    ("1. FEOL 前端×20-30层", "②", "Expose 曝光", "LOOP×N",
     "光刻曝光 Lithography (EUV/DUV/i-line)",
     "Canon", "7751.T", "东京证交所", "光刻机 i-line/DUV", "JPY"),
    ("1. FEOL 前端×20-30层", "②", "Expose 曝光", "LOOP×N",
     "光刻曝光 Lithography (EUV/DUV/i-line)",
     "Nikon", "7731.T", "东京证交所", "光刻机 DUV", "JPY"),

    # ③Develop
    ("1. FEOL 前端×20-30层", "③", "Develop 显影", "LOOP×N",
     "曝光后显影 Post-Expose Develop",
     "Tokyo Electron / TEL", "8035.T", "东京证交所", "涂布/显影设备", "JPY"),

    # ④a Etch
    ("1. FEOL 前端×20-30层", "④a", "Etch 蚀刻", "PAR",
     "干法蚀刻 Dry Etch（Lam ~70%存储 / TEL ~30%）",
     "Lam Research", "LRCX", "NASDAQ", "蚀刻设备", "USD"),
    ("1. FEOL 前端×20-30层", "④a", "Etch 蚀刻", "PAR",
     "干法蚀刻 Dry Etch（Lam ~70%存储 / TEL ~30%）",
     "Tokyo Electron / TEL", "8035.T", "东京证交所", "蚀刻设备", "JPY"),
    ("1. FEOL 前端×20-30层", "④a", "Etch 蚀刻", "PAR",
     "干法蚀刻 Dry Etch（中国本土替代）",
     "中微公司 AMEC", "688012.SS", "上交所科创板", "蚀刻设备", "CNY"),
    ("1. FEOL 前端×20-30层", "④a", "Etch 蚀刻", "PAR",
     "干法蚀刻 Dry Etch（中国本土替代）",
     "北方华创 Naura", "002371.SZ", "深交所", "蚀刻/CVD设备", "CNY"),

    # ④b Ion Implant
    ("1. FEOL 前端×20-30层", "④b", "Ion Impl. 离子注入", "PAR",
     "离子注入 Ion Implant（掺杂），AMAT主导",
     "Applied Materials / AMAT", "AMAT", "NASDAQ", "离子注入/CVD/CMP设备", "USD"),

    # ⑤Thermal
    ("1. FEOL 前端×20-30层", "⑤", "Thermal 热处理", "LOOP×N",
     "热处理/退火 Furnace / RTP（TEL 53% / KOKUSAI 30%）",
     "Tokyo Electron / TEL", "8035.T", "东京证交所", "热处理炉设备", "JPY"),
    ("1. FEOL 前端×20-30层", "⑤", "Thermal 热处理", "LOOP×N",
     "热处理/退火 Furnace / RTP（TEL 53% / KOKUSAI 30%）",
     "Kokusai Electric", "6525.T", "东京证交所", "热处理炉设备", "JPY"),

    # ⑥Clean
    ("1. FEOL 前端×20-30层", "⑥", "Clean 清洗", "LOOP×N",
     "清洗（占全部工艺步骤>30%！），SCREEN >50%",
     "SCREEN Holdings", "7735.T", "东京证交所", "清洗设备", "JPY"),
    ("1. FEOL 前端×20-30层", "⑥", "Clean 清洗", "LOOP×N",
     "清洗，中国本土替代",
     "ACM Research / ACMR", "ACMR", "NASDAQ", "清洗设备", "USD"),

    # ⑦Inspect
    ("1. FEOL 前端×20-30层", "⑦", "Inspect 检测", "LOOP×N",
     "在线缺陷检测 In-line Inspection（KLA >50%）",
     "KLA Corporation", "KLAC", "NASDAQ", "量检测设备", "USD"),
    ("1. FEOL 前端×20-30层", "⑦", "Inspect 检测", "LOOP×N",
     "在线缺陷检测 In-line Inspection",
     "Lasertec", "6920.T", "东京证交所", "量检测/掩膜检测", "JPY"),

    # ── Stage 2: BEOL 前端互连线 ×7-15铜层循环 ───────────────────────────────
    # ⑧CVD
    ("2. BEOL 互连×7-15铜层", "⑧", "CVD 成膜", "LOOP×N",
     "介电膜/阻挡层 CVD/ALD Deposition",
     "Applied Materials / AMAT", "AMAT", "NASDAQ", "CVD/ALD设备", "USD"),
    ("2. BEOL 互连×7-15铜层", "⑧", "CVD 成膜", "LOOP×N",
     "介电膜/阻挡层 CVD/ALD Deposition",
     "Tokyo Electron / TEL", "8035.T", "东京证交所", "CVD/ALD设备", "JPY"),
    ("2. BEOL 互连×7-15铜层", "⑧", "CVD 成膜", "LOOP×N",
     "介电膜/阻挡层 CVD/ALD Deposition",
     "Lam Research", "LRCX", "NASDAQ", "CVD/ALD设备", "USD"),
    ("2. BEOL 互连×7-15铜层", "⑧", "CVD 成膜", "LOOP×N",
     "介电膜/阻挡层 CVD/ALD Deposition",
     "Kokusai Electric", "6525.T", "东京证交所", "CVD/ALD设备", "JPY"),
    ("2. BEOL 互连×7-15铜层", "⑧", "CVD 成膜", "LOOP×N",
     "介电膜/阻挡层 CVD/ALD Deposition",
     "ASM International / ASMI", "ASM.AS", "阿姆斯特丹", "ALD设备", "EUR"),

    # ⑨PVD
    ("2. BEOL 互连×7-15铜层", "⑨", "PVD 溅射", "LOOP×N",
     "种子层/金属溅射 Metal Sputter（AMAT主导）",
     "Applied Materials / AMAT", "AMAT", "NASDAQ", "PVD溅射设备", "USD"),
    ("2. BEOL 互连×7-15铜层", "⑨", "PVD 溅射", "LOOP×N",
     "种子层/金属溅射 Metal Sputter",
     "Ulvac", "6728.T", "东京证交所", "PVD溅射设备", "JPY"),

    # ⑩Litho+Etch (same as FEOL ①-④)
    ("2. BEOL 互连×7-15铜层", "⑩", "Litho+Etch", "LOOP×N",
     "重复阶段1的①-④步（沟槽/通孔图案），同FEOL步骤①-④",
     "ASML Holding", "ASML", "NASDAQ", "光刻机 EUV/DUV", "USD"),
    ("2. BEOL 互连×7-15铜层", "⑩", "Litho+Etch", "LOOP×N",
     "重复阶段1的①-④步（沟槽/通孔图案），同FEOL步骤①-④",
     "Lam Research", "LRCX", "NASDAQ", "蚀刻设备", "USD"),

    # ⑪Plate
    ("2. BEOL 互连×7-15铜层", "⑪", "Plate 电镀", "LOOP×N",
     "铜填充 Cu Electroplating（Lam ~80%）",
     "Lam Research", "LRCX", "NASDAQ", "电镀设备", "USD"),
    ("2. BEOL 互连×7-15铜层", "⑪", "Plate 电镀", "LOOP×N",
     "铜填充 Cu Electroplating",
     "Applied Materials / AMAT", "AMAT", "NASDAQ", "电镀设备", "USD"),
    ("2. BEOL 互连×7-15铜层", "⑪", "Plate 电镀", "LOOP×N",
     "铜填充 Cu Electroplating",
     "ACM Research / ACMR", "ACMR", "NASDAQ", "电镀/清洗设备", "USD"),
    ("2. BEOL 互连×7-15铜层", "⑪", "Plate 电镀", "LOOP×N",
     "铜填充 Cu Electroplating",
     "荏原制作所 Ebara", "6361.T", "东京证交所", "电镀/CMP设备", "JPY"),

    # ⑫CMP
    ("2. BEOL 互连×7-15铜层", "⑫", "CMP 抛光", "LOOP×N",
     "化学机械抛光 Cu Polish（AMAT >50% / EBARA #2↑）",
     "Applied Materials / AMAT", "AMAT", "NASDAQ", "CMP设备", "USD"),
    ("2. BEOL 互连×7-15铜层", "⑫", "CMP 抛光", "LOOP×N",
     "化学机械抛光 Cu Polish（AMAT >50% / EBARA #2↑）",
     "荏原制作所 Ebara", "6361.T", "东京证交所", "CMP/电镀设备", "JPY"),

    # ⑬Clean+Insp
    ("2. BEOL 互连×7-15铜层", "⑬", "Clean+Insp", "LOOP×N",
     "同阶段1的⑥⑦步",
     "SCREEN Holdings", "7735.T", "东京证交所", "清洗设备", "JPY"),
    ("2. BEOL 互连×7-15铜层", "⑬", "Clean+Insp", "LOOP×N",
     "同阶段1的⑥⑦步",
     "KLA Corporation", "KLAC", "NASDAQ", "量检测设备", "USD"),

    # ── Stage 3: Mid-end 中端先进封装前道 2023年起新兴 ─────────────────────────
    # ⑭Grind
    ("3. Mid-end 先进封装前道", "⑭", "Grind 减薄", "SEQ",
     "晶圆背面减薄 Backside Thinning (<30μm)（Disco #1）",
     "Disco Corporation", "6146.T", "东京证交所", "晶圆减薄/切割", "JPY"),
    ("3. Mid-end 先进封装前道", "⑭", "Grind 减薄", "SEQ",
     "晶圆背面减薄 Backside Thinning (<30μm)",
     "Tokyo Seimitsu / Accretech", "7729.T", "东京证交所", "晶圆减薄/量测", "JPY"),

    # ⑮a Hybrid Bond
    ("3. Mid-end 先进封装前道", "⑮a", "Hybrid Bond", "NEW",
     "混合键合 W2W/C2W (CoWoS/SoIC) 探索中",
     "Tokyo Electron / TEL", "8035.T", "东京证交所", "键合设备", "JPY"),

    # ⑮b TSV
    ("3. Mid-end 先进封装前道", "⑮b", "TSV 硅通孔", "SEQ",
     "Through-Silicon Via（Lam蚀刻/EBARA电镀/AMAT CMP）",
     "Lam Research", "LRCX", "NASDAQ", "TSV蚀刻设备", "USD"),
    ("3. Mid-end 先进封装前道", "⑮b", "TSV 硅通孔", "SEQ",
     "Through-Silicon Via（Lam蚀刻/EBARA电镀/AMAT CMP）",
     "荏原制作所 Ebara", "6361.T", "东京证交所", "TSV电镀设备", "JPY"),
    ("3. Mid-end 先进封装前道", "⑮b", "TSV 硅通孔", "SEQ",
     "Through-Silicon Via（Lam蚀刻/EBARA电镀/AMAT CMP）",
     "Applied Materials / AMAT", "AMAT", "NASDAQ", "TSV CMP设备", "USD"),

    # ⑯RDL
    ("3. Mid-end 先进封装前道", "⑯", "RDL 再布线", "SEQ",
     "重分布层 Redistribution Layer",
     "Tokyo Electron / TEL", "8035.T", "东京证交所", "RDL涂布/曝光设备", "JPY"),
    ("3. Mid-end 先进封装前道", "⑯", "RDL 再布线", "SEQ",
     "重分布层 Redistribution Layer",
     "SCREEN Holdings", "7735.T", "东京证交所", "RDL清洗设备", "JPY"),
    ("3. Mid-end 先进封装前道", "⑯", "RDL 再布线", "SEQ",
     "重分布层 Redistribution Layer",
     "荏原制作所 Ebara", "6361.T", "东京证交所", "RDL电镀设备", "JPY"),

    # ⑰Probe Test
    ("3. Mid-end 先进封装前道", "⑰", "Probe Test", "SEQ",
     "晶圆级探针测试 Wafer Probe",
     "Tokyo Seimitsu / Accretech", "7729.T", "东京证交所", "探针台/量测", "JPY"),

    # ── Stage 4: Back-end 后端封装测试 ──────────────────────────────────────
    # ⑱Dice
    ("4. Back-end 后端封装测试", "⑱", "Dice 切割", "SEQ",
     "晶圆切割 Wafer Dicing（Disco >70%）",
     "Disco Corporation", "6146.T", "东京证交所", "晶圆切割设备", "JPY"),
    ("4. Back-end 后端封装测试", "⑱", "Dice 切割", "SEQ",
     "晶圆切割 Wafer Dicing",
     "Tokyo Seimitsu / Accretech", "7729.T", "东京证交所", "晶圆切割/量测", "JPY"),
    ("4. Back-end 后端封装测试", "⑱", "Dice 切割", "SEQ",
     "晶圆切割检测 Wafer Dicing Inspection",
     "KLA Corporation", "KLAC", "NASDAQ", "切割后检测", "USD"),

    # ⑲a FC Bond
    ("4. Back-end 后端封装测试", "⑲a", "FC Bond 倒装键合", "PAR",
     "倒装芯片键合 Flip Chip TCB/Ultrasonic（ASMPT / Besi #1）",
     "ASMPT", "0522.HK", "港交所", "倒装键合/贴片设备", "HKD"),
    ("4. Back-end 后端封装测试", "⑲a", "FC Bond 倒装键合", "PAR",
     "倒装芯片键合 Flip Chip TCB/Ultrasonic",
     "BE Semiconductor / Besi", "BESI.AS", "阿姆斯特丹", "倒装键合设备", "EUR"),
    ("4. Back-end 后端封装测试", "⑲a", "FC Bond 倒装键合", "PAR",
     "倒装芯片键合 Flip Chip TCB/Ultrasonic",
     "Hanmi Semiconductor", "042700.KQ", "韩国KOSDAQ", "TC键合设备", "KRW"),
    ("4. Back-end 后端封装测试", "⑲a", "FC Bond 倒装键合", "PAR",
     "倒装芯片键合 Flip Chip TCB/Ultrasonic",
     "芝浦机械 Shibaura Mechatronics", "6590.T", "东京证交所", "键合/封装设备", "JPY"),

    # ⑲b Wire Bond
    ("4. Back-end 后端封装测试", "⑲b", "Wire Bond 引线键合", "PAR",
     "引线焊接 Wire Bonding（金线/铜线），K&S垄断",
     "Kulicke & Soffa / K&S", "KLIC", "NASDAQ", "引线键合设备", "USD"),
    ("4. Back-end 后端封装测试", "⑲b", "Wire Bond 引线键合", "PAR",
     "引线焊接 Wire Bonding（金线/铜线）",
     "Shinkawa", "6274.T", "东京证交所", "引线键合设备", "JPY"),

    # ⑳Mold
    ("4. Back-end 后端封装测试", "⑳", "Mold 模塑", "SEQ",
     "压缩模塑封装 Compression Molding（TOWA垄断）",
     "TOWA Corporation", "6315.T", "东京证交所", "模塑封装设备", "JPY"),

    # ㉑a ATE Test
    ("4. Back-end 后端封装测试", "㉑a", "ATE Test 芯片测试", "SEQ",
     "芯片测试 ATE Test（存储/SoC），Advantest #1 / Teradyne #2",
     "Advantest", "6857.T", "东京证交所", "ATE测试设备", "JPY"),
    ("4. Back-end 后端封装测试", "㉑a", "ATE Test 芯片测试", "SEQ",
     "芯片测试 ATE Test（存储/SoC），Advantest #1 / Teradyne #2",
     "Teradyne", "TER", "NASDAQ", "ATE测试设备", "USD"),

    # ㉑b SLT
    ("4. Back-end 后端封装测试", "㉑b", "SLT 系统测试", "NEW",
     "系统级测试 System-Level Test（新兴增长）",
     "Advantest", "6857.T", "东京证交所", "ATE/SLT测试设备", "JPY"),
    ("4. Back-end 后端封装测试", "㉑b", "SLT 系统测试", "NEW",
     "系统级测试 System-Level Test（新兴增长）",
     "Keysight Technologies", "KEYS", "NYSE", "测试测量设备", "USD"),
    ("4. Back-end 后端封装测试", "㉑b", "SLT 系统测试", "NEW",
     "系统级测试 System-Level Test（新兴增长）",
     "Teradyne", "TER", "NASDAQ", "ATE/SLT测试设备", "USD"),

    # ㉒Burn-in
    ("4. Back-end 后端封装测试", "㉒", "Burn-in 老化筛选", "SEQ",
     "老化筛选 Reliability Burn-in",
     "Advantest", "6857.T", "东京证交所", "测试设备", "JPY"),
    ("4. Back-end 后端封装测试", "㉒", "Burn-in 老化筛选", "SEQ",
     "老化筛选 Reliability Burn-in",
     "Teradyne", "TER", "NASDAQ", "测试设备", "USD"),
]

# ── 颜色配置 ──────────────────────────────────────────────────────────────────
STAGE_COLORS = {
    "0. Design EDA设计":        "4472C4",  # 蓝色
    "0.5 Mask 掩膜制备":        "7030A0",  # 紫色
    "1. FEOL 前端×20-30层":     "C00000",  # 深红
    "2. BEOL 互连×7-15铜层":    "833C00",  # 橙棕 (仿图中颜色)
    "3. Mid-end 先进封装前道":   "ED7D31",  # 橙色
    "4. Back-end 后端封装测试":  "375623",  # 深绿
}

TYPE_COLORS = {
    "SEQ":    "00B0F0",  # 蓝
    "LOOP×N": "92D050",  # 绿
    "PAR":    "FF0000",  # 红
    "NEW":    "FFC000",  # 黄
}

HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="微软雅黑", size=10)
TITLE_FONT  = Font(bold=True, color="FFFFFF", name="微软雅黑", size=11)

thin = Side(style="thin", color="BFBFBF")
THIN_BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

thick_bottom = Side(style="medium", color="808080")
STAGE_BORDER = Border(left=thin, right=thin, top=thick_bottom, bottom=thick_bottom)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)


def hex_fill(hex_color, alpha=None):
    return PatternFill(fill_type="solid", fgColor=hex_color)


def build_excel(filename="半导体制造价值链标的.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "半导体价值链标的"

    # ── 列宽 ──────────────────────────────────────────────────────────────────
    col_widths = {
        "A": 22,  # 阶段 Stage
        "B": 6,   # 步骤编号
        "C": 14,  # 步骤名称
        "D": 8,   # 类型
        "E": 32,  # 说明
        "F": 30,  # 标的名称
        "G": 14,  # 标的代码
        "H": 14,  # 市场
        "I": 20,  # 主要/细分行业
        "J": 6,   # 币种
        "K": 12,  # 昨日收盘价
        "L": 12,  # 实时现价
        "M": 12,  # 2025年末价格
        "N": 12,  # Q1最终价格
        "O": 14,  # 过去三个月均价
        "P": 12,  # 过去半年趋势
        "Q": 12,  # MS目标价
        "R": 12,  # 高盛目标价
        "S": 12,  # 摩根大通
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # ── 标题行 ────────────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A1:S1")
    ws["A1"] = "半导体制造价值链 完整标的总览 Semiconductor Manufacturing Value Chain — Listed Stocks"
    ws["A1"].font = Font(bold=True, color="FFFFFF", name="微软雅黑", size=14)
    ws["A1"].fill = hex_fill("1F3864")
    ws["A1"].alignment = CENTER

    # ── 副标题 ────────────────────────────────────────────────────────────────
    ws.row_dimensions[2].height = 20
    ws.merge_cells("A2:S2")
    ws["A2"] = (
        "类型标注：SEQ=顺序  LOOP×N=多层循环  PAR=并行路径  NEW=新兴/探索中"
        "    数据来源：公开市场信息，价格栏待填入实时数据"
    )
    ws["A2"].font = Font(italic=True, color="595959", name="微软雅黑", size=9)
    ws["A2"].fill = hex_fill("D6E4F7")
    ws["A2"].alignment = LEFT

    # ── 表头行 ────────────────────────────────────────────────────────────────
    headers = [
        "阶段 Stage", "步骤#", "步骤名称 Step", "类型",
        "说明 Description",
        "标的名称", "标的代码", "市场",
        "主要/细分行业", "币种",
        "昨日收盘价", "实时现价",
        "2025年末价格", "Q1最终价格",
        "过去三个月均价", "过去半年趋势",
        "MS目标价", "高盛目标价", "摩根大通",
    ]
    ws.row_dimensions[3].height = 36
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    # ── 数据行 ────────────────────────────────────────────────────────────────
    row_num = 4
    prev_stage = None
    stage_start_row = 4

    # 先收集各阶段行范围，用于后续合并
    stage_rows = {}  # stage -> [start_row, end_row]
    step_rows  = {}  # (stage, step_num) -> [start_row, end_row]

    data_rows = []
    for rec in SEMI_CHAIN:
        (stage, step_num, step_name, step_type,
         desc, name, ticker, market, sub_industry, currency) = rec
        data_rows.append(rec)

    # 写数据并记录合并区间
    stage_map = {}   # stage -> list of row numbers
    step_map  = {}   # (stage, step_num, step_name) -> list of row numbers

    for rec in data_rows:
        (stage, step_num, step_name, step_type,
         desc, name, ticker, market, sub_industry, currency) = rec

        stage_color = STAGE_COLORS.get(stage, "808080")
        type_color  = TYPE_COLORS.get(step_type, "FFFFFF")

        ws.row_dimensions[row_num].height = 18

        values = [
            stage, step_num, step_name, step_type,
            desc, name, ticker, market, sub_industry, currency,
            "", "", "", "", "", "", "", "", ""
        ]
        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            cell.border = THIN_BORDER
            # 阶段列背景
            if col_idx == 1:
                cell.fill = hex_fill(stage_color)
                cell.font = Font(bold=True, color="FFFFFF", name="微软雅黑", size=9)
                cell.alignment = CENTER
            elif col_idx == 4:  # 类型
                cell.fill = hex_fill(type_color)
                cell.font = Font(bold=True, color="000000" if step_type in ("NEW",) else "FFFFFF",
                                  name="微软雅黑", size=9)
                cell.alignment = CENTER
            elif col_idx in (2, 3):
                cell.fill = hex_fill("F2F2F2")
                cell.font = Font(name="微软雅黑", size=9, bold=(col_idx==3))
                cell.alignment = CENTER
            elif col_idx == 5:
                cell.fill = hex_fill("FAFAFA")
                cell.font = Font(name="微软雅黑", size=8, color="404040")
                cell.alignment = LEFT
            elif col_idx == 6:
                cell.font = Font(bold=True, name="微软雅黑", size=9)
                cell.alignment = LEFT
            elif col_idx == 7:
                cell.font = Font(bold=True, color="1F3864", name="Courier New", size=9)
                cell.alignment = CENTER
            elif col_idx == 10:
                cell.font = Font(name="微软雅黑", size=9, color="595959")
                cell.alignment = CENTER
            else:
                cell.font = Font(name="微软雅黑", size=9)
                cell.alignment = CENTER if col_idx >= 10 else LEFT

        # 记录合并区间
        stage_map.setdefault(stage, []).append(row_num)
        step_key = (stage, step_num, step_name)
        step_map.setdefault(step_key, []).append(row_num)

        row_num += 1

    # ── 合并阶段列（A列） ────────────────────────────────────────────────────
    for stage, rows in stage_map.items():
        if len(rows) > 1:
            ws.merge_cells(f"A{rows[0]}:A{rows[-1]}")

    # ── 合并步骤列（B、C、D、E列） ──────────────────────────────────────────
    for step_key, rows in step_map.items():
        if len(rows) > 1:
            for col in (2, 3, 4, 5):  # B,C,D,E
                ws.merge_cells(
                    f"{get_column_letter(col)}{rows[0]}"
                    f":{get_column_letter(col)}{rows[-1]}"
                )

    # ── 冻结首行+表头 ────────────────────────────────────────────────────────
    ws.freeze_panes = "A4"

    # ── 筛选 ─────────────────────────────────────────────────────────────────
    ws.auto_filter.ref = f"A3:S{row_num - 1}"

    wb.save(filename)
    print(f"[OK] 已生成: {filename}  （共 {row_num - 4} 行数据）")
    return filename


if __name__ == "__main__":
    build_excel()
