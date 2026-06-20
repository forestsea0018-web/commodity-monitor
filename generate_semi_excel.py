#!/usr/bin/env python3
"""
半导体全产业链 — 单张工作表，按工序顺序整合设备/材料/AI算力标的
"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── 样式工具 ──────────────────────────────────────────────────────────────────
def hf(hex6):
    return PatternFill(fill_type="solid", fgColor=hex6)

def tf(bold=False, color="000000", size=9, name="微软雅黑", italic=False):
    return Font(bold=bold, color=color, size=size, name=name, italic=italic)

_thin = Side(style="thin",   color="BFBFBF")
_med  = Side(style="medium", color="595959")
BD   = Border(left=_thin, right=_thin, top=_thin,  bottom=_thin)
BD_T = Border(left=_med,  right=_med,  top=_med,   bottom=_med)

C = Alignment(horizontal="center", vertical="center", wrap_text=True)
L = Alignment(horizontal="left",   vertical="center", wrap_text=True)

# ─── 价格列 ───────────────────────────────────────────────────────────────────
PRICE_HDRS = [
    "昨日收盘价", "实时现价", "2025年末价格", "Q1最终价格",
    "过去三月均价", "过去半年趋势", "MS目标价", "高盛目标价", "摩根大通",
]

# ─── 数据定义 ─────────────────────────────────────────────────────────────────
# 每行: (阶段, 步骤/子类, 类别, 说明, 公司名, 代码, 市场, 细分行业, 币种, 参与度/排名)
# 类别: "设备" | "材料" | "AI算力"
# 参与度: ★垄断>50%  ●核心参与  ○参与   排名:Top1/2/3/4

ROWS = [

    # ══════════════════════════════════════════════════════════════════════════
    # 0. Design EDA 设计
    # ══════════════════════════════════════════════════════════════════════════
    ("0. Design\nEDA设计",  "EDA软件",  "设备",
     "逻辑综合→P&R→时序签核→GDSII",
     "Cadence Design Systems",    "CDNS",      "NASDAQ",      "EDA软件",              "USD", "★"),
    ("0. Design\nEDA设计",  "EDA软件",  "设备",
     "逻辑综合→P&R→时序签核→GDSII",
     "Synopsys",                  "SNPS",      "NASDAQ",      "EDA软件",              "USD", "★"),

    # ── 衬底材料（进入晶圆厂的原材料）────────────────────────────────────────
    ("0. Design\nEDA设计",  "化合物半导体衬底",  "材料",
     "高速光芯片/毫米波雷达核心底材，全球供应极紧，缺口持续扩大",
     "云南锗业",           "002428.SZ", "深交所",      "锗/化合物半导体材料",  "CNY", "—"),
    ("0. Design\nEDA设计",  "化合物半导体衬底",  "材料",
     "高速光芯片/毫米波雷达核心底材，全球供应极紧",
     "有研新材",           "600206.SH", "上交所",      "半导体新材料/靶材",    "CNY", "—"),
    ("0. Design\nEDA设计",  "化合物半导体衬底",  "材料",
     "化合物半导体光学材料",
     "光智科技",           "300161.SZ", "深交所创业板","化合物半导体光学材料", "CNY", "—"),
    ("0. Design\nEDA设计",  "SiC碳化硅衬底",  "材料",
     "800V高压快充/新能源车驱动核心材料，今年涨幅>50%，缺货至2028",
     "天岳先进",           "688234.SH", "上交所科创板","SiC衬底",             "CNY", "—"),
    ("0. Design\nEDA设计",  "SiC碳化硅衬底",  "材料",
     "SiC外延/器件，今年涨幅>50%",
     "露笑科技",           "002617.SZ", "深交所",      "SiC外延/器件",        "CNY", "—"),
    ("0. Design\nEDA设计",  "SiC碳化硅衬底",  "材料",
     "SiC/化合物半导体器件",
     "三安光电",           "600703.SH", "上交所",      "SiC/化合物器件",      "CNY", "—"),

    # ══════════════════════════════════════════════════════════════════════════
    # 0.5 Mask 掩膜制备
    # ══════════════════════════════════════════════════════════════════════════
    ("0.5 Mask\n掩膜制备",  "掩膜写入",  "设备",
     "E-beam写入掩膜（非上市，东芝子公司）",
     "NuFlare Technology",        "非上市",    "—",           "掩膜写入设备",        "JPY", "●"),
    ("0.5 Mask\n掩膜制备",  "掩膜写入",  "设备",
     "E-beam写入/电子束设备",
     "JEOL",                      "6951.T",    "东京证交所",  "电子束设备/检测",     "JPY", "○"),
    ("0.5 Mask\n掩膜制备",  "掩膜检测",  "设备",
     "掩膜缺陷检测，EUV掩膜绝对垄断",
     "Lasertec",                  "6920.T",    "东京证交所",  "掩膜检测设备",        "JPY", "★"),
    ("0.5 Mask\n掩膜制备",  "空白掩膜板", "材料",
     "空白板材料供应，全球垄断",
     "Hoya Corporation",          "7741.T",    "东京证交所",  "空白掩膜板/光学元件", "JPY", "★"),
    ("0.5 Mask\n掩膜制备",  "空白掩膜板", "材料",
     "空白板材料供应",
     "Shin-Etsu Chemical 信越化学","4063.T",   "东京证交所",  "空白掩膜板/硅片",     "JPY", "●"),

    # ══════════════════════════════════════════════════════════════════════════
    # 1. FEOL 前端晶体管 ×20-30层循环
    # ══════════════════════════════════════════════════════════════════════════

    # ① Coat 涂布 + 光刻胶材料
    ("1. FEOL\n前端×20-30层", "① Coat 涂布", "设备",
     "光阻涂布 Resist Coat，TEL 93%垄断",
     "Tokyo Electron / TEL 东京电子","8035.T",  "东京证交所",  "涂布/显影/热处理设备","JPY", "★"),
    ("1. FEOL\n前端×20-30层", "① Coat 涂布", "设备",
     "光阻涂布，三星设备子公司（非上市）",
     "SEMES",                     "非上市",    "—",           "涂布/清洗设备",       "KRW", "○"),
    ("1. FEOL\n前端×20-30层", "光刻胶材料",  "材料",
     "JSR/信越垄断，国内自给率<10%，高端光刻胶极度短缺",
     "彤程新材",           "603650.SH", "上交所",      "光刻胶/半导体材料",    "CNY", "—"),
    ("1. FEOL\n前端×20-30层", "光刻胶材料",  "材料",
     "ArF光刻胶/前驱体，国产替代",
     "南大光电",           "300346.SZ", "深交所创业板","ArF光刻胶/前驱体",    "CNY", "—"),
    ("1. FEOL\n前端×20-30层", "光刻胶材料",  "材料",
     "光刻胶/电镀液，国产替代",
     "上海新阳",           "300236.SZ", "深交所创业板","光刻胶/电镀液",       "CNY", "—"),

    # ② Expose 曝光
    ("1. FEOL\n前端×20-30层", "② Expose 曝光", "设备",
     "EUV全球垄断，DUV主力",
     "ASML Holding",              "ASML",      "NASDAQ",      "光刻机 EUV/DUV",      "USD", "★"),
    ("1. FEOL\n前端×20-30层", "② Expose 曝光", "设备",
     "i-line/DUV准分子激光光刻",
     "Canon 佳能",                "7751.T",    "东京证交所",  "光刻机 i-line/DUV",   "JPY", "●"),
    ("1. FEOL\n前端×20-30层", "② Expose 曝光", "设备",
     "DUV准分子激光光刻",
     "Nikon 尼康",                "7731.T",    "东京证交所",  "光刻机 DUV",          "JPY", "○"),

    # ③ Develop 显影
    ("1. FEOL\n前端×20-30层", "③ Develop 显影", "设备",
     "曝光后显影 Post-Expose Develop",
     "Tokyo Electron / TEL 东京电子","8035.T",  "东京证交所",  "涂布/显影设备",       "JPY", "★"),

    # ④ Etch 蚀刻 + 气体/化学品材料
    ("1. FEOL\n前端×20-30层", "④a Etch 蚀刻",  "设备",
     "干法蚀刻，存储市场~70%份额",
     "Lam Research",              "LRCX",      "NASDAQ",      "蚀刻设备",            "USD", "★"),
    ("1. FEOL\n前端×20-30层", "④a Etch 蚀刻",  "设备",
     "干法蚀刻，逻辑市场~30%份额",
     "Tokyo Electron / TEL 东京电子","8035.T",  "东京证交所",  "蚀刻/CVD设备",        "JPY", "●"),
    ("1. FEOL\n前端×20-30层", "④a Etch 蚀刻",  "设备",
     "干法蚀刻，中国本土替代（科创板）",
     "中微公司 AMEC",     "688012.SS", "上交所科创板","蚀刻设备",            "CNY", "○"),
    ("1. FEOL\n前端×20-30层", "④a Etch 蚀刻",  "设备",
     "干法蚀刻，中国本土替代",
     "北方华创 Naura",    "002371.SZ", "深交所",      "蚀刻/CVD/PVD设备",    "CNY", "○"),
    ("1. FEOL\n前端×20-30层", "④b Ion Impl. 离子注入", "设备",
     "离子注入掺杂，AMAT全球主导",
     "Applied Materials / AMAT 应用材料","AMAT","NASDAQ",     "离子注入/CVD/PVD/CMP","USD", "★"),
    ("1. FEOL\n前端×20-30层", "蚀刻气体/湿化学品", "材料",
     "半导体蚀刻/清洗/气相沉积核心气体，今年价格涨一倍",
     "凯美特气",           "002549.SZ", "深交所",      "电子特种气体",        "CNY", "—"),
    ("1. FEOL\n前端×20-30层", "蚀刻气体/湿化学品", "材料",
     "高纯氢气/特种气体",
     "华特气体",           "688268.SH", "上交所科创板","电子特种气体",        "CNY", "—"),
    ("1. FEOL\n前端×20-30层", "蚀刻气体/湿化学品", "材料",
     "半导体特种气体/前驱体",
     "雅克科技",           "002409.SZ", "深交所",      "半导体特种气体/前驱体","CNY","—"),
    ("1. FEOL\n前端×20-30层", "蚀刻气体/湿化学品", "材料",
     "芯片制造湿法清洗/刻蚀核心化学品，Stella垄断，涨幅>50%",
     "江化微",             "603078.SH", "上交所",      "电子级硫酸/湿化学品", "CNY", "—"),
    ("1. FEOL\n前端×20-30层", "蚀刻气体/湿化学品", "材料",
     "电子级氟化物/湿化学品",
     "多氟多",             "002407.SZ", "深交所",      "电子级氟化物/锂电材料","CNY","—"),
    ("1. FEOL\n前端×20-30层", "蚀刻气体/湿化学品", "材料",
     "电子级化学品/光刻胶",
     "晶瑞电材",           "300655.SZ", "深交所创业板","电子级化学品",        "CNY", "—"),

    # ⑤ Thermal 热处理
    ("1. FEOL\n前端×20-30层", "⑤ Thermal 热处理", "设备",
     "热处理/退火 Furnace/RTP，TEL 53%",
     "Tokyo Electron / TEL 东京电子","8035.T",  "东京证交所",  "热处理炉设备",        "JPY", "★"),
    ("1. FEOL\n前端×20-30层", "⑤ Thermal 热处理", "设备",
     "热处理/退火，KOKUSAI 30%",
     "Kokusai Electric",          "6525.T",    "东京证交所",  "热处理炉设备",        "JPY", "●"),
    ("1. FEOL\n前端×20-30层", "⑤ Thermal 热处理", "设备",
     "ALD/热处理，ASMI全球领先",
     "ASM International / ASMI",  "ASM.AS",    "阿姆斯特丹",  "ALD/热处理设备",      "EUR", "○"),
    ("1. FEOL\n前端×20-30层", "⑤ Thermal 热处理", "设备",
     "热处理/CVD，中国本土替代",
     "北方华创 Naura",    "002371.SZ", "深交所",      "热处理/CVD设备",      "CNY", "○"),

    # ⑥ Clean 清洗
    ("1. FEOL\n前端×20-30层", "⑥ Clean 清洗", "设备",
     "单片清洗>50%，占全工艺步骤>30%",
     "SCREEN Holdings 迪恩士",    "7735.T",    "东京证交所",  "清洗设备",            "JPY", "★"),
    ("1. FEOL\n前端×20-30层", "⑥ Clean 清洗", "设备",
     "单片清洗，中国本土替代",
     "ACM Research / ACMR 盛美",  "ACMR",      "NASDAQ",      "清洗/电镀设备",       "USD", "○"),
    ("1. FEOL\n前端×20-30层", "⑥ Clean 清洗", "设备",
     "清洗设备，三星设备子公司（非上市）",
     "SEMES",                     "非上市",    "—",           "清洗设备",            "KRW", "○"),
    ("1. FEOL\n前端×20-30层", "⑥ Clean 清洗", "设备",
     "清洗设备，TEL参与",
     "Tokyo Electron / TEL 东京电子","8035.T",  "东京证交所",  "清洗设备",            "JPY", "●"),

    # ⑦ Inspect 检测
    ("1. FEOL\n前端×20-30层", "⑦ Inspect 检测", "设备",
     "在线缺陷检测，KLA >50%",
     "KLA Corporation",           "KLAC",      "NASDAQ",      "量检测/计量设备",     "USD", "★"),
    ("1. FEOL\n前端×20-30层", "⑦ Inspect 检测", "设备",
     "电子束检测/量测",
     "Hitachi High-Tech 日立高新", "6501.T*",  "东京证交所(日立母公司)","电子束检测设备","JPY","●"),
    ("1. FEOL\n前端×20-30层", "⑦ Inspect 检测", "设备",
     "电子显微镜/材料分析",
     "Thermo Fisher Scientific",  "TMO",       "NYSE",        "电子显微镜/分析仪器", "USD", "○"),
    ("1. FEOL\n前端×20-30层", "⑦ Inspect 检测", "设备",
     "电子显微镜/束流设备",
     "JEOL",                      "6951.T",    "东京证交所",  "电子束设备/检测",     "JPY", "○"),

    # ══════════════════════════════════════════════════════════════════════════
    # 2. BEOL 前端互连线 ×7-15铜层
    # ══════════════════════════════════════════════════════════════════════════

    # ⑧ CVD/ALD
    ("2. BEOL\n互连×7-15铜层", "⑧ CVD/ALD 成膜", "设备",
     "介电膜/ALD Deposition",
     "Applied Materials / AMAT 应用材料","AMAT","NASDAQ",     "CVD/ALD设备",         "USD", "●"),
    ("2. BEOL\n互连×7-15铜层", "⑧ CVD/ALD 成膜", "设备",
     "介电膜/ALD，TEL强势品类",
     "Tokyo Electron / TEL 东京电子","8035.T",  "东京证交所",  "CVD/ALD设备",         "JPY", "●"),
    ("2. BEOL\n互连×7-15铜层", "⑧ CVD/ALD 成膜", "设备",
     "CVD/ALD，Lam参与",
     "Lam Research",              "LRCX",      "NASDAQ",      "CVD/ALD设备",         "USD", "●"),
    ("2. BEOL\n互连×7-15铜层", "⑧ CVD/ALD 成膜", "设备",
     "批次CVD炉管，KOKUSAI强势",
     "Kokusai Electric",          "6525.T",    "东京证交所",  "批次CVD炉管设备",     "JPY", "●"),
    ("2. BEOL\n互连×7-15铜层", "⑧ CVD/ALD 成膜", "设备",
     "ALD设备，ASMI全球领先",
     "ASM International / ASMI",  "ASM.AS",    "阿姆斯特丹",  "ALD设备",             "EUR", "●"),

    # ⑨ PVD 溅射 + 靶材材料
    ("2. BEOL\n互连×7-15铜层", "⑨ PVD 溅射", "设备",
     "金属溅射 Metal Sputter，AMAT主导",
     "Applied Materials / AMAT 应用材料","AMAT","NASDAQ",     "PVD溅射设备",         "USD", "★"),
    ("2. BEOL\n互连×7-15铜层", "⑨ PVD 溅射", "设备",
     "PVD溅射，中国本土替代",
     "北方华创 Naura",    "002371.SZ", "深交所",      "PVD溅射设备",         "CNY", "○"),
    ("2. BEOL\n互连×7-15铜层", "⑨ PVD 溅射", "设备",
     "PVD溅射",
     "Ulvac",                     "6728.T",    "东京证交所",  "PVD溅射/真空设备",    "JPY", "●"),
    ("2. BEOL\n互连×7-15铜层", "PVD靶材",   "材料",
     "半导体PVD溅射工艺不可或缺，今年涨幅80%，缺货至2027底",
     "金钼股份",           "601958.SH", "上交所",      "钼靶材/贵金属",       "CNY", "—"),
    ("2. BEOL\n互连×7-15铜层", "PVD靶材",   "材料",
     "钼铜靶材",
     "洛阳股份",           "000537.SZ", "深交所",      "钼铜靶材",            "CNY", "—"),
    ("2. BEOL\n互连×7-15铜层", "PVD靶材",   "材料",
     "半导体靶材/新材料",
     "有研新材",           "600206.SH", "上交所",      "半导体靶材/新材料",   "CNY", "—"),

    # ⑩ Litho+Etch (同FEOL①-④)
    ("2. BEOL\n互连×7-15铜层", "⑩ Litho+Etch\n(同FEOL①-④)", "设备",
     "重复FEOL①-④步（沟槽/通孔图案）",
     "ASML Holding",              "ASML",      "NASDAQ",      "光刻机 EUV/DUV",      "USD", "★"),
    ("2. BEOL\n互连×7-15铜层", "⑩ Litho+Etch\n(同FEOL①-④)", "设备",
     "重复FEOL蚀刻步骤",
     "Lam Research",              "LRCX",      "NASDAQ",      "蚀刻设备",            "USD", "★"),

    # ⑪ 铜电镀
    ("2. BEOL\n互连×7-15铜层", "⑪ Plate 铜电镀", "设备",
     "Cu Electroplating，Lam ~80%",
     "Lam Research",              "LRCX",      "NASDAQ",      "铜电镀设备",          "USD", "★"),
    ("2. BEOL\n互连×7-15铜层", "⑪ Plate 铜电镀", "设备",
     "铜电镀，AMAT参与",
     "Applied Materials / AMAT 应用材料","AMAT","NASDAQ",     "铜电镀/ECD设备",      "USD", "●"),
    ("2. BEOL\n互连×7-15铜层", "⑪ Plate 铜电镀", "设备",
     "铜电镀，中国本土替代",
     "ACM Research / ACMR 盛美",  "ACMR",      "NASDAQ",      "铜电镀/清洗设备",     "USD", "○"),
    ("2. BEOL\n互连×7-15铜层", "⑪ Plate 铜电镀", "设备",
     "铜电镀，Ebara全球参与",
     "荏原制作所 Ebara",  "6361.T",    "东京证交所",  "铜电镀/CMP设备",      "JPY", "●"),

    # ⑫ CMP
    ("2. BEOL\n互连×7-15铜层", "⑫ CMP 化学机械抛光", "设备",
     "Cu CMP，AMAT >50%",
     "Applied Materials / AMAT 应用材料","AMAT","NASDAQ",     "CMP设备",             "USD", "★"),
    ("2. BEOL\n互连×7-15铜层", "⑫ CMP 化学机械抛光", "设备",
     "CMP，Ebara全球第2",
     "荏原制作所 Ebara",  "6361.T",    "东京证交所",  "CMP/电镀设备",        "JPY", "●"),

    # ⑬ Clean+Insp
    ("2. BEOL\n互连×7-15铜层", "⑬ Clean+Insp", "设备",
     "BEOL清洗，同FEOL⑥",
     "SCREEN Holdings 迪恩士",    "7735.T",    "东京证交所",  "清洗设备",            "JPY", "★"),
    ("2. BEOL\n互连×7-15铜层", "⑬ Clean+Insp", "设备",
     "BEOL检测，同FEOL⑦",
     "KLA Corporation",           "KLAC",      "NASDAQ",      "量检测设备",          "USD", "★"),

    # ══════════════════════════════════════════════════════════════════════════
    # 3. Mid-end 中端先进封装前道
    # ══════════════════════════════════════════════════════════════════════════

    # ⑭ Grind
    ("3. Mid-end\n先进封装前道", "⑭ Grind 减薄", "设备",
     "晶圆背面减薄 <30μm，Disco #1",
     "Disco Corporation 迪斯科",  "6146.T",    "东京证交所",  "晶圆减薄/切割设备",   "JPY", "★"),
    ("3. Mid-end\n先进封装前道", "⑭ Grind 减薄", "设备",
     "晶圆减薄，Tokyo Seimitsu #2",
     "Tokyo Seimitsu / Accretech 东京精密","7729.T","东京证交所","晶圆减薄/量测/探针","JPY","●"),
    ("3. Mid-end\n先进封装前道", "⑭ Grind 减薄", "设备",
     "减薄后检测",
     "KLA Corporation",           "KLAC",      "NASDAQ",      "量检测设备",          "USD", "○"),

    # ⑮a Hybrid Bond
    ("3. Mid-end\n先进封装前道", "⑮a Hybrid Bond\n混合键合(NEW)", "设备",
     "W2W/C2W混合键合(CoWoS/SoIC)探索中，TEL参与",
     "Tokyo Electron / TEL 东京电子","8035.T",  "东京证交所",  "混合键合设备",        "JPY", "●"),
    ("3. Mid-end\n先进封装前道", "⑮a Hybrid Bond\n混合键合(NEW)", "设备",
     "W2W键合设备，EVG私有奥地利企业",
     "EV Group / EVG",            "非上市",    "—",           "W2W键合设备",         "EUR", "●"),

    # ⑮b TSV
    ("3. Mid-end\n先进封装前道", "⑮b TSV 硅通孔", "设备",
     "TSV Deep Etch，Lam主导",
     "Lam Research",              "LRCX",      "NASDAQ",      "TSV蚀刻/CVD设备",     "USD", "●"),
    ("3. Mid-end\n先进封装前道", "⑮b TSV 硅通孔", "设备",
     "TSV铜填充电镀",
     "荏原制作所 Ebara",  "6361.T",    "东京证交所",  "TSV电镀设备",         "JPY", "●"),
    ("3. Mid-end\n先进封装前道", "⑮b TSV 硅通孔", "设备",
     "TSV CMP",
     "Applied Materials / AMAT 应用材料","AMAT","NASDAQ",     "TSV CMP设备",         "USD", "●"),

    # ⑯ RDL
    ("3. Mid-end\n先进封装前道", "⑯ RDL 再布线", "设备",
     "重分布层 Redistribution Layer，涂布/曝光",
     "Tokyo Electron / TEL 东京电子","8035.T",  "东京证交所",  "RDL涂布/曝光设备",    "JPY", "●"),
    ("3. Mid-end\n先进封装前道", "⑯ RDL 再布线", "设备",
     "RDL清洗",
     "SCREEN Holdings 迪恩士",    "7735.T",    "东京证交所",  "RDL清洗设备",         "JPY", "●"),
    ("3. Mid-end\n先进封装前道", "⑯ RDL 再布线", "设备",
     "RDL铜电镀",
     "荏原制作所 Ebara",  "6361.T",    "东京证交所",  "RDL电镀设备",         "JPY", "●"),
    ("3. Mid-end\n先进封装前道", "⑯ RDL 再布线", "设备",
     "RDL铜电镀，中国本土替代",
     "ACM Research / ACMR 盛美",  "ACMR",      "NASDAQ",      "RDL电镀/清洗设备",    "USD", "○"),

    # ⑰ Probe Test
    ("3. Mid-end\n先进封装前道", "⑰ Probe Test\n探针测试", "设备",
     "晶圆级探针测试 Wafer Probe，探针台",
     "Tokyo Seimitsu / Accretech 东京精密","7729.T","东京证交所","探针台",            "JPY","●"),
    ("3. Mid-end\n先进封装前道", "⑰ Probe Test\n探针测试", "设备",
     "探针卡 Probe Card 制造商 #1",
     "Micronics Japan",           "6871.T",    "东京证交所",  "探针卡",              "JPY", "●"),
    ("3. Mid-end\n先进封装前道", "⑰ Probe Test\n探针测试", "设备",
     "探针卡 Probe Card 制造商 #2",
     "FormFactor",                "FORM",      "NASDAQ",      "探针卡",              "USD", "●"),

    # ══════════════════════════════════════════════════════════════════════════
    # 4. Back-end 后端封装测试
    # ══════════════════════════════════════════════════════════════════════════

    # ⑱ Dice
    ("4. Back-end\n后端封装测试", "⑱ Dice 切割", "设备",
     "晶圆切割 Wafer Dicing，Disco >70%",
     "Disco Corporation 迪斯科",  "6146.T",    "东京证交所",  "晶圆切割设备",        "JPY", "★"),
    ("4. Back-end\n后端封装测试", "⑱ Dice 切割", "设备",
     "晶圆切割，Tokyo Seimitsu #2",
     "Tokyo Seimitsu / Accretech 东京精密","7729.T","东京证交所","晶圆切割/量测",     "JPY","○"),
    ("4. Back-end\n后端封装测试", "⑱ Dice 切割", "设备",
     "切割后检测",
     "KLA Corporation",           "KLAC",      "NASDAQ",      "切割后检测",          "USD", "○"),

    # ⑲a FC Bond
    ("4. Back-end\n后端封装测试", "⑲a FC Bond\n倒装键合", "设备",
     "倒装键合/先进封装，ASMPT全球领先",
     "ASMPT",                     "0522.HK",   "港交所",      "倒装键合/贴片设备",   "HKD", "●"),
    ("4. Back-end\n后端封装测试", "⑲a FC Bond\n倒装键合", "设备",
     "TCB热压键合，Besi全球 #1",
     "BE Semiconductor / Besi",   "BESI.AS",   "阿姆斯特丹",  "倒装键合设备",        "EUR", "★"),
    ("4. Back-end\n后端封装测试", "⑲a FC Bond\n倒装键合", "设备",
     "TC键合(HBM关键)，Hanmi核心供应商",
     "Hanmi Semiconductor 韩美半导体","042700.KQ","韩国KOSDAQ","TC键合/视觉检测",    "KRW","●"),
    ("4. Back-end\n后端封装测试", "⑲a FC Bond\n倒装键合", "设备",
     "键合/封装设备",
     "Shibaura Mechatronics 芝浦机械","6590.T", "东京证交所",  "键合/封装设备",       "JPY", "●"),

    # ⑲b Wire Bond
    ("4. Back-end\n后端封装测试", "⑲b Wire Bond\n引线键合", "设备",
     "金/铜引线焊接，K&S全球垄断",
     "Kulicke & Soffa / K&S",     "KLIC",      "NASDAQ",      "引线键合设备",        "USD", "★"),
    ("4. Back-end\n后端封装测试", "⑲b Wire Bond\n引线键合", "设备",
     "引线焊接，日本第2",
     "Shinkawa 新川",             "6274.T",    "东京证交所",  "引线键合设备",        "JPY", "●"),

    # ⑳ Mold
    ("4. Back-end\n后端封装测试", "⑳ Mold 模塑", "设备",
     "压缩模塑封装，TOWA全球垄断",
     "TOWA Corporation",          "6315.T",    "东京证交所",  "压缩模塑设备",        "JPY", "★"),
    ("4. Back-end\n后端封装测试", "⑳ Mold 模塑", "设备",
     "压缩模塑，日本第2",
     "Apic Yamada",               "6300.T",    "东京证交所",  "压缩模塑/引线框架设备","JPY","●"),

    # ㉑ ATE Test
    ("4. Back-end\n后端封装测试", "㉑a ATE Test\n芯片测试", "设备",
     "存储/SoC ATE测试，Advantest #1",
     "Advantest",                 "6857.T",    "东京证交所",  "ATE测试设备",         "JPY", "★"),
    ("4. Back-end\n后端封装测试", "㉑a ATE Test\n芯片测试", "设备",
     "ATE测试，Teradyne #2",
     "Teradyne 泰瑞达",           "TER",       "NASDAQ",      "ATE测试设备",         "USD", "●"),

    # ㉑b SLT
    ("4. Back-end\n后端封装测试", "㉑b SLT 系统测试(NEW)", "设备",
     "系统级测试 System-Level Test 新兴赛道",
     "Advantest",                 "6857.T",    "东京证交所",  "ATE/SLT测试设备",     "JPY", "●"),
    ("4. Back-end\n后端封装测试", "㉑b SLT 系统测试(NEW)", "设备",
     "SLT测试",
     "Teradyne 泰瑞达",           "TER",       "NASDAQ",      "ATE/SLT测试设备",     "USD", "●"),
    ("4. Back-end\n后端封装测试", "㉑b SLT 系统测试(NEW)", "设备",
     "测试测量",
     "Keysight Technologies",     "KEYS",      "NYSE",        "测试测量仪器",         "USD", "●"),

    # ㉒ Burn-in
    ("4. Back-end\n后端封装测试", "㉒ Burn-in 老化筛选", "设备",
     "可靠性老化筛选",
     "Advantest",                 "6857.T",    "东京证交所",  "老化/测试设备",        "JPY", "●"),
    ("4. Back-end\n后端封装测试", "㉒ Burn-in 老化筛选", "设备",
     "可靠性老化筛选",
     "Teradyne 泰瑞达",           "TER",       "NASDAQ",      "老化/测试设备",        "USD", "●"),

    # ── 封装材料 ──────────────────────────────────────────────────────────────
    ("4. Back-end\n后端封装测试", "ABF载板", "材料",
     "CPU/GPU封装基板，ABF膜日本垄断，今年涨幅>70%，供应极紧",
     "深南电路",           "002916.SZ", "深交所",      "PCB/封装基板",        "CNY", "—"),
    ("4. Back-end\n后端封装测试", "ABF载板", "材料",
     "HDI/封装基板，国产替代",
     "兴森科技",           "002436.SZ", "深交所",      "HDI/封装基板",        "CNY", "—"),
    ("4. Back-end\n后端封装测试", "ABF载板", "材料",
     "PCB/高密度板，国产替代",
     "崇达技术",           "002815.SZ", "深交所",      "PCB/高密度板",        "CNY", "—"),
    ("4. Back-end\n后端封装测试", "高端PCB载板", "材料",
     "AI服务器需求爆发，今年大涨60%，算力硬件关键基材",
     "生益科技",           "600183.SH", "上交所",      "覆铜板/PCB材料",      "CNY", "—"),
    ("4. Back-end\n后端封装测试", "高端PCB载板", "材料",
     "高频PCB材料/覆铜板",
     "华正新材",           "603186.SH", "上交所",      "高频PCB材料/覆铜板",  "CNY", "—"),
    ("4. Back-end\n后端封装测试", "高端PCB载板", "材料",
     "高速高频PCB",
     "沪电股份",           "002463.SZ", "深交所",      "高端PCB/通信板",      "CNY", "—"),
    ("4. Back-end\n后端封装测试", "钽电容", "材料",
     "积小/容大/耐高温，军工/航天/AI服务器刚需，今年涨幅>80%",
     "宏达电子",           "300726.SZ", "深交所创业板","钽电容/被动元件",     "CNY", "—"),
    ("4. Back-end\n后端封装测试", "钽电容", "材料",
     "钽电容/军用电子元件",
     "火炬电子",           "600650.SH", "上交所",      "钽电容/军用元件",     "CNY", "—"),
    ("4. Back-end\n后端封装测试", "钽电容", "材料",
     "军用电子元件",
     "振华科技",           "000733.SZ", "深交所",      "军用电子元件",        "CNY", "—"),
    ("4. Back-end\n后端封装测试", "MLCC电容", "材料",
     "电子工业大米，今年涨幅>50%，全球供应缺口明显",
     "风华高科",           "000636.SZ", "深交所",      "MLCC/片式电容",       "CNY", "—"),
    ("4. Back-end\n后端封装测试", "MLCC电容", "材料",
     "MLCC/电子陶瓷",
     "三环集团",           "300408.SZ", "深交所创业板","MLCC/电子陶瓷",       "CNY", "—"),
    ("4. Back-end\n后端封装测试", "MLCC电容", "材料",
     "MLCC载带/隔离纸",
     "洁美科技",           "002859.SZ", "深交所",      "MLCC载带/隔离纸",     "CNY", "—"),
    ("4. Back-end\n后端封装测试", "铜箔/电子布", "材料",
     "锂电负极/PCB覆铜板材料，今年价格翻倍",
     "铜冠铜箔",           "301217.SZ", "深交所",      "电解铜箔",            "CNY", "—"),
    ("4. Back-end\n后端封装测试", "铜箔/电子布", "材料",
     "高端电解铜箔",
     "嘉元科技",           "688388.SH", "上交所科创板","高端电解铜箔",        "CNY", "—"),
    ("4. Back-end\n后端封装测试", "铜箔/电子布", "材料",
     "覆铜板关键增强材料，高端电子布价格翻倍，缺货至2028",
     "中国巨石",           "600176.SH", "上交所",      "玻纤/电子布",         "CNY", "—"),
    ("4. Back-end\n后端封装测试", "铜箔/电子布", "材料",
     "高端电子布",
     "宏和科技",           "603256.SH", "上交所",      "高端电子布",          "CNY", "—"),
    ("4. Back-end\n后端封装测试", "铜箔/电子布", "材料",
     "玻纤/复合材料",
     "中材科技",           "002080.SZ", "深交所",      "玻纤/复合材料",       "CNY", "—"),

    # ══════════════════════════════════════════════════════════════════════════
    # 5. AI算力应用层
    # ══════════════════════════════════════════════════════════════════════════
    ("5. AI算力\n应用层", "CPO 共封装光学", "AI算力",
     "数据中心光电共封装，AI算力核心互连技术",
     "中际旭创",       "300308.SZ","深交所",      "CPO/光模块",          "CNY","Top1"),
    ("5. AI算力\n应用层", "CPO 共封装光学", "AI算力",
     "数据中心光电共封装",
     "新易盛",         "300502.SZ","深交所创业板","CPO/光模块",          "CNY","Top2"),
    ("5. AI算力\n应用层", "CPO 共封装光学", "AI算力",
     "CPO光互连组件",
     "天孚通信",       "300394.SZ","深交所创业板","CPO光互连组件",       "CNY","Top3"),
    ("5. AI算力\n应用层", "CPO 共封装光学", "AI算力",
     "激光/光通信器件",
     "华工科技",       "000988.SZ","深交所",      "激光/光通信器件",     "CNY","Top4"),

    ("5. AI算力\n应用层", "OCS 光交换", "AI算力",
     "光学交换技术，替代电交换降低数据中心能耗",
     "腾景科技",       "688195.SH","上交所科创板","光开关/OCS器件",      "CNY","Top1"),
    ("5. AI算力\n应用层", "OCS 光交换", "AI算力",
     "晶体光学/激光器件",
     "福晶科技",       "002222.SZ","深交所",      "晶体光学/激光器件",   "CNY","Top2"),
    ("5. AI算力\n应用层", "OCS 光交换", "AI算力",
     "OCS光学组件",
     "光库科技",       "300620.SZ","深交所创业板","OCS光学组件",         "CNY","Top3"),
    ("5. AI算力\n应用层", "OCS 光交换", "AI算力",
     "光模块/相干光器件",
     "德科立",         "688205.SH","上交所科创板","光模块/相干光器件",   "CNY","Top4"),

    ("5. AI算力\n应用层", "光芯片 Photonic Chip", "AI算力",
     "AI算力光互连核心芯片",
     "源杰科技",       "688498.SH","上交所科创板","激光/光芯片",         "CNY","Top1"),
    ("5. AI算力\n应用层", "光芯片 Photonic Chip", "AI算力",
     "光芯片/PLC分路器",
     "仕佳光子",       "688313.SH","上交所科创板","光芯片/PLC",          "CNY","Top2"),
    ("5. AI算力\n应用层", "光芯片 Photonic Chip", "AI算力",
     "光芯片/光模块",
     "光迅科技",       "002281.SZ","深交所",      "光芯片/光模块",       "CNY","Top3"),
    ("5. AI算力\n应用层", "光芯片 Photonic Chip", "AI算力",
     "高功率激光芯片",
     "长光华芯",       "688048.SH","上交所科创板","高功率激光芯片",      "CNY","Top4"),

    ("5. AI算力\n应用层", "AI服务器 PCB", "AI算力",
     "AI服务器PCB，算力硬件关键基材，今年大涨60%",
     "胜宏科技",       "300476.SZ","深交所创业板","AI服务器PCB",         "CNY","Top1"),
    ("5. AI算力\n应用层", "AI服务器 PCB", "AI算力",
     "高端PCB/FPC",
     "东山精密",       "002384.SZ","深交所",      "高端PCB/FPC",         "CNY","Top2"),
    ("5. AI算力\n应用层", "AI服务器 PCB", "AI算力",
     "AI服务器/封装基板",
     "深南电路",       "002916.SZ","深交所",      "AI服务器/封装基板",   "CNY","Top3"),
    ("5. AI算力\n应用层", "AI服务器 PCB", "AI算力",
     "高速高频PCB",
     "沪电股份",       "002463.SZ","深交所",      "高速高频PCB",         "CNY","Top4"),

    ("5. AI算力\n应用层", "AI服务器整机", "AI算力",
     "AI服务器代工/组装",
     "工业富联",       "601138.SH","上交所",      "AI服务器代工/组装",   "CNY","Top1"),
    ("5. AI算力\n应用层", "AI服务器整机", "AI算力",
     "AI服务器整机",
     "浪潮信息",       "000977.SZ","深交所",      "AI服务器整机",        "CNY","Top2"),
    ("5. AI算力\n应用层", "AI服务器整机", "AI算力",
     "AI服务器/网络设备",
     "紫光股份",       "000938.SZ","深交所",      "AI服务器/网络设备",   "CNY","Top3"),
    ("5. AI算力\n应用层", "AI服务器整机", "AI算力",
     "AI服务器/超算",
     "中科曙光",       "603019.SH","上交所",      "AI服务器/超算",       "CNY","Top4"),

    ("5. AI算力\n应用层", "AI芯片", "AI算力",
     "国产GPU/DCU芯片，AI训练推理",
     "海光信息",       "688041.SH","上交所科创板","GPU/DCU芯片",         "CNY","Top1"),
    ("5. AI算力\n应用层", "AI芯片", "AI算力",
     "AI推理/训练芯片",
     "寒武纪",         "688256.SH","上交所科创板","AI推理/训练芯片",     "CNY","Top2"),
    ("5. AI算力\n应用层", "AI芯片", "AI算力",
     "GPU架构芯片（非上市）",
     "沐曦股份",       "非上市",  "—",           "GPU架构芯片",         "CNY","Top3"),
    ("5. AI算力\n应用层", "AI芯片", "AI算力",
     "GPU芯片（非上市）",
     "摩尔线程",       "非上市",  "—",           "GPU芯片",             "CNY","Top4"),

    ("5. AI算力\n应用层", "光纤光缆", "AI算力",
     "AI数据中心高速互连光纤",
     "长飞光纤",       "601869.SH","上交所",      "光纤预制棒/光缆",     "CNY","Top1"),
    ("5. AI算力\n应用层", "光纤光缆", "AI算力",
     "光纤光缆/海缆",
     "亨通光电",       "600487.SH","上交所",      "光纤光缆/海缆",       "CNY","Top2"),
    ("5. AI算力\n应用层", "光纤光缆", "AI算力",
     "光纤光缆/海缆",
     "中天科技",       "600522.SH","上交所",      "光纤光缆/海缆",       "CNY","Top3"),
    ("5. AI算力\n应用层", "光纤光缆", "AI算力",
     "光纤光缆/通信设备",
     "烽火通信",       "600498.SH","上交所",      "光纤光缆/通信设备",   "CNY","Top4"),

    ("5. AI算力\n应用层", "存储芯片", "AI算力",
     "AI训练推理大规模存储需求",
     "德明利",         "001323.SZ","深交所",      "NAND Flash存储模组",  "CNY","Top1"),
    ("5. AI算力\n应用层", "存储芯片", "AI算力",
     "NOR Flash/MCU",
     "兆易创新",       "603986.SH","上交所",      "NOR Flash/MCU",       "CNY","Top2"),
    ("5. AI算力\n应用层", "存储芯片", "AI算力",
     "NAND存储模组",
     "佰维存储",       "688525.SH","上交所科创板","NAND存储模组",        "CNY","Top3"),
    ("5. AI算力\n应用层", "存储芯片", "AI算力",
     "NAND/eMMC存储",
     "江波龙",         "301308.SZ","深交所",      "NAND/eMMC存储",       "CNY","Top4"),

    ("5. AI算力\n应用层", "高速链接 Interconnect", "AI算力",
     "AI服务器高速铜缆/连接器",
     "立讯精密",       "002475.SZ","深交所",      "高速连接器/线缆",     "CNY","Top1"),
    ("5. AI算力\n应用层", "高速链接 Interconnect", "AI算力",
     "高速铜缆/连接器",
     "兆龙互连",       "603557.SH","上交所",      "高速铜缆/连接器",     "CNY","Top2"),
    ("5. AI算力\n应用层", "高速链接 Interconnect", "AI算力",
     "高速线缆/热缩材料",
     "沃尔核材",       "002130.SZ","深交所",      "高速线缆/热缩材料",   "CNY","Top3"),
    ("5. AI算力\n应用层", "高速链接 Interconnect", "AI算力",
     "高速连接器",
     "鼎通科技",       "301291.SZ","深交所",      "高速连接器",          "CNY","Top4"),

    ("5. AI算力\n应用层", "液冷 Liquid Cooling", "AI算力",
     "AI数据中心液冷散热，H100/H200热耗密度极高",
     "英维克",         "002837.SZ","深交所",      "液冷/机房热管理",     "CNY","Top1"),
    ("5. AI算力\n应用层", "液冷 Liquid Cooling", "AI算力",
     "液冷系统",
     "高澜股份",       "300499.SZ","深交所创业板","液冷系统",            "CNY","Top2"),
    ("5. AI算力\n应用层", "液冷 Liquid Cooling", "AI算力",
     "液冷AI服务器/超算",
     "中科曙光",       "603019.SH","上交所",      "液冷AI服务器/超算",   "CNY","Top3"),

    ("5. AI算力\n应用层", "电源 Power Supply", "AI算力",
     "数据中心电源/UPS",
     "中恒电气",       "002364.SZ","深交所",      "数据中心电源",        "CNY","Top1"),
    ("5. AI算力\n应用层", "电源 Power Supply", "AI算力",
     "UPS/储能电源",
     "圣阳股份",       "002580.SZ","深交所",      "UPS/储能电源",        "CNY","Top2"),
    ("5. AI算力\n应用层", "电源 Power Supply", "AI算力",
     "服务器电源",
     "欧陆通",         "300870.SZ","深交所创业板","服务器电源",          "CNY","Top3"),
    ("5. AI算力\n应用层", "电源 Power Supply", "AI算力",
     "工业电源/服务器电源",
     "麦格米特",       "002851.SZ","深交所",      "工业电源/服务器电源", "CNY","Top4"),

    ("5. AI算力\n应用层", "AIDC 数据中心", "AI算力",
     "AI数据中心建设运营",
     "润泽科技",       "300442.SZ","深交所",      "AIDC建设运营",        "CNY","Top1"),
    ("5. AI算力\n应用层", "AIDC 数据中心", "AI算力",
     "CDN/AIDC",
     "网宿科技",       "300017.SZ","深交所创业板","CDN/AIDC",            "CNY","Top2"),
    ("5. AI算力\n应用层", "AIDC 数据中心", "AI算力",
     "IDC/AIDC",
     "光环新网",       "300383.SZ","深交所创业板","IDC/AIDC",            "CNY","Top3"),
    ("5. AI算力\n应用层", "AIDC 数据中心", "AI算力",
     "IDC/AIDC建设运营",
     "数据港",         "603881.SH","上交所",      "IDC/AIDC建设运营",    "CNY","Top4"),

    ("5. AI算力\n应用层", "算电协同", "AI算力",
     "算力+电力协同，绿电支撑AI数据中心",
     "豫能控股",       "001896.SZ","深交所",      "电力/算力协同",       "CNY","Top1"),
    ("5. AI算力\n应用层", "算电协同", "AI算力",
     "绿电/算力协同",
     "协鑫能科",       "002015.SZ","深交所",      "绿电/算力协同",       "CNY","Top2"),
    ("5. AI算力\n应用层", "算电协同", "AI算力",
     "绿电/数字能源",
     "韶能股份",       "000601.SZ","深交所",      "绿电/数字能源",       "CNY","Top3"),

    ("5. AI算力\n应用层", "算力租赁", "AI算力",
     "GPU云算力租赁",
     "利通电子",       "603629.SH","上交所",      "算力租赁/GPU云",      "CNY","Top1"),
    ("5. AI算力\n应用层", "算力租赁", "AI算力",
     "算力/数据中心租赁",
     "协创数据",       "300456.SZ","深交所创业板","算力/数据中心租赁",   "CNY","Top2"),
    ("5. AI算力\n应用层", "算力租赁", "AI算力",
     "算力云/GPU租赁",
     "优刻得",         "688158.SH","上交所科创板","算力云/GPU租赁",      "CNY","Top3"),
]

# ─── 颜色映射 ─────────────────────────────────────────────────────────────────
STAGE_BG = {
    "0. Design\nEDA设计":       "4472C4",
    "0.5 Mask\n掩膜制备":       "7030A0",
    "1. FEOL\n前端×20-30层":    "C00000",
    "2. BEOL\n互连×7-15铜层":   "833C00",
    "3. Mid-end\n先进封装前道":  "ED7D31",
    "4. Back-end\n后端封装测试": "375623",
    "5. AI算力\n应用层":        "1F3864",
}

CAT_BG = {
    "设备":  None,       # 继承阶段色（淡化）
    "材料":  "FFF2CC",   # 淡黄
    "AI算力":"EEF2FF",   # 淡蓝紫
}
CAT_FONT = {
    "设备":  ("1F3864","微软雅黑"),
    "材料":  ("7F6000","微软雅黑"),
    "AI算力":("203864","微软雅黑"),
}

RANK_BG = {
    "Top1":"C00000","Top2":"ED7D31","Top3":"4472C4","Top4":"808080",
    "★":"C00000","●":"4472C4","○":"A6A6A6","—":"F2F2F2",
}
RANK_FC = {
    "Top1":"FFFFFF","Top2":"FFFFFF","Top3":"FFFFFF","Top4":"FFFFFF",
    "★":"FFFFFF","●":"FFFFFF","○":"FFFFFF","—":"808080",
}


# ─── 构建工作表 ───────────────────────────────────────────────────────────────
def build(filename="半导体全产业链标的.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "半导体全产业链标的"

    TOTAL_COLS = 10 + len(PRICE_HDRS)   # A-S = 19列

    # 列宽
    widths = {
        "A": 14,  # 阶段
        "B": 18,  # 步骤/子类
        "C": 7,   # 类别
        "D": 30,  # 说明
        "E": 26,  # 标的名称
        "F": 13,  # 代码
        "G": 14,  # 市场
        "H": 20,  # 细分行业
        "I": 6,   # 币种
        "J": 7,   # 参与度/排名
    }
    for ci, ph in enumerate(PRICE_HDRS, start=11):
        widths[get_column_letter(ci)] = 12
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ── 标题行 ────────────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A1:{get_column_letter(TOTAL_COLS)}1")
    c = ws["A1"]
    c.value = "半导体全产业链投资标的总览  Semiconductor Full Value Chain — 按工序顺序整合设备 / 材料 / AI算力"
    c.font, c.fill, c.alignment = (
        tf(bold=True, color="FFFFFF", size=13),
        hf("1F3864"), C
    )

    ws.row_dimensions[2].height = 18
    ws.merge_cells(f"A2:{get_column_letter(TOTAL_COLS)}2")
    c2 = ws["A2"]
    c2.value = (
        "类别：设备=半导体制造设备    材料=关键原材料    AI算力=算力产业链终端应用"
        "    参与度：★垄断/主导(>50%)  ●核心参与  ○参与    排名：Top1~4"
        "    注：非上市=暂未公开上市；6501.T*=日立高新2020年已私有化至日立集团"
    )
    c2.font    = tf(italic=True, color="595959", size=8)
    c2.fill    = hf("D6E4F7")
    c2.alignment = L

    # ── 表头行 ────────────────────────────────────────────────────────────────
    HDRS = ["阶段 Stage","步骤/子类","类别","说明 Description",
            "标的名称","标的代码","市场","主要/细分行业","币种","参与度/排名"] + PRICE_HDRS
    ws.row_dimensions[3].height = 34
    for ci, h in enumerate(HDRS, 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.font, c.fill, c.alignment, c.border = (
            tf(bold=True, color="FFFFFF", size=9),
            hf("1F3864"), C, BD
        )

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(TOTAL_COLS)}3"

    # ── 数据行 ────────────────────────────────────────────────────────────────
    stage_rows   = {}   # stage -> [row_numbers]
    substep_rows = {}   # (stage,substep) -> [row_numbers]

    for ri, rec in enumerate(ROWS, start=4):
        stage, substep, cat, desc, name, ticker, market, sub, cur, note = rec
        ws.row_dimensions[ri].height = 16

        sg_color = STAGE_BG.get(stage, "808080")
        cat_bg   = CAT_BG.get(cat)
        is_unlisted = ticker in ("非上市", "—")

        # 行底色（材料/AI算力有独立底色）
        row_fill = hf(cat_bg) if cat_bg else None

        vals = [stage, substep, cat, desc, name, ticker, market, sub, cur, note] + [""]*len(PRICE_HDRS)
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.border = BD

            if ci == 1:  # 阶段
                cell.fill = hf(sg_color)
                cell.font = tf(bold=True, color="FFFFFF", size=9)
                cell.alignment = C

            elif ci == 2:  # 步骤/子类
                bg = "F2F2F2" if not cat_bg else cat_bg
                cell.fill = hf(bg)
                cell.font = tf(bold=True, size=8,
                               color=sg_color if cat == "设备" else CAT_FONT[cat][0])
                cell.alignment = C

            elif ci == 3:  # 类别标签
                tag_colors = {"设备":"4472C4","材料":"7F6000","AI算力":"7030A0"}
                tag_bgs    = {"设备":"D6E4F7","材料":"FFF2CC","AI算力":"E8EAF6"}
                cell.fill = hf(tag_bgs.get(cat, "F2F2F2"))
                cell.font = tf(bold=True, size=8, color=tag_colors.get(cat,"000000"))
                cell.alignment = C

            elif ci == 4:  # 说明
                if row_fill:
                    cell.fill = row_fill
                cell.font = tf(size=8, color="404040",
                               italic=(cat == "材料"))
                cell.alignment = L

            elif ci == 5:  # 标的名称
                if row_fill:
                    cell.fill = row_fill
                cell.font = tf(bold=True, size=9,
                               color="808080" if is_unlisted else "000000",
                               italic=is_unlisted)
                cell.alignment = L

            elif ci == 6:  # 代码
                if row_fill:
                    cell.fill = row_fill
                cell.font = Font(bold=not is_unlisted, color="808080" if is_unlisted else "1F3864",
                                 size=9, name="Courier New", italic=is_unlisted)
                cell.alignment = C

            elif ci == 10:  # 参与度/排名
                nb = RANK_BG.get(note, "F2F2F2")
                nf = RANK_FC.get(note, "000000")
                cell.fill = hf(nb)
                cell.font = tf(bold=True, color=nf, size=9)
                cell.alignment = C

            else:  # 市场/行业/币种/价格
                if row_fill and ci <= 9:
                    cell.fill = row_fill
                cell.font = tf(size=9, color="595959" if ci == 9 else "000000")
                cell.alignment = C if ci >= 7 else L

        stage_rows.setdefault(stage, []).append(ri)
        substep_rows.setdefault((stage, substep), []).append(ri)

    # ── 合并阶段列 A ──────────────────────────────────────────────────────────
    for stage, rows in stage_rows.items():
        if len(rows) > 1:
            ws.merge_cells(f"A{rows[0]}:A{rows[-1]}")
            ws[f"A{rows[0]}"].alignment = C

    # ── 合并步骤列 B ──────────────────────────────────────────────────────────
    for (st, ss), rows in substep_rows.items():
        if len(rows) > 1:
            ws.merge_cells(f"B{rows[0]}:B{rows[-1]}")
            ws[f"B{rows[0]}"].alignment = C

    wb.save(filename)
    print(f"[OK] {filename}  共 {len(ROWS)} 行")


if __name__ == "__main__":
    build()
