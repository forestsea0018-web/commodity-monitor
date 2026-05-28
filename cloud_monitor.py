#!/usr/bin/env python3
"""
商品期货新闻云端监控 - 微信推送版 (PushPlus)
设计为 GitHub Actions 每 10 分钟运行一次。
无状态：只推送"最近 15 分钟内发布"且命中关键词的新闻。
"""

import os
import sys
import json
import time
import hashlib
import socket
import requests
import feedparser
import concurrent.futures
from pathlib import Path
from datetime import datetime, timezone, timedelta

socket.setdefaulttimeout(6)

# ── 环境变量 ──────────────────────────────────────────────
PUSHPLUS_TOKEN = os.environ.get("PUSHPLUS_TOKEN", "").strip()
# 群组推送：在 pushplus.plus 建群后填入群 topic（留空则只推给自己）
PUSHPLUS_TOPIC = os.environ.get("PUSHPLUS_TOPIC", "").strip()
MAX_PUSH_PER_RUN = int(os.environ.get("MAX_PUSH_PER_RUN", "30"))
DIGEST_MODE = os.environ.get("DIGEST_MODE", "1") == "1"
MAX_AGE_HOURS = int(os.environ.get("MAX_AGE_HOURS", "24"))

SEEN_FILE = Path("seen_articles.json")

# ── 关键词（与本地版保持一致） ────────────────────────────
COMMODITIES = {
    "🟤 铜": [
        "copper", "铜", "comex copper", "lme copper", "shfe copper",
        "cu futures", "cathode", "阴极铜", "精炼铜", "铜矿",
        "codelco", "freeport", "escondida", "grasberg", "bhp copper",
        "rio tinto copper", "antofagasta", "first quantum", "glencore copper",
        "southern copper", "scco", "teck copper", "zijin", "紫金矿业",
        "江西铜业", "jiangxi copper", "tongling", "铜陵有色",
        "zambia copper", "chile copper", "peru copper", "congo copper",
    ],
    "⚪ 铝": [
        "aluminum", "aluminium", "铝", "lme aluminum", "shfe aluminum",
        "alumina", "氧化铝", "bauxite", "铝土矿", "电解铝",
        "rusal", "chalco", "中铝", "hongqiao", "宏桥", "alcoa",
        "rio tinto aluminum", "norsk hydro", "emirates global aluminium",
        "south32", "vedanta aluminium", "weiqiao", "魏桥",
        "guinea bauxite", "indonesia bauxite", "russia aluminium",
    ],
    "⬜ 铂金": [
        "platinum", "铂金", "铂", "pgm", "palladium", "钯金",
        "amplats", "anglo platinum", "sibanye", "sibanye-stillwater",
        "impala", "implats", "northam platinum", "stillwater",
        "north american palladium", "nornickel", "诺镍",
        "south africa platinum", "zimbabwe platinum", "russia palladium",
    ],
    "🌿 美棉": [
        "cotton", "棉花", "美棉", "ice cotton", "ct futures",
        "usda cotton", "cotton yield", "cotton export",
        "texas cotton", "xinjiang cotton", "新疆棉", "pakistan cotton",
        "india cotton", "brazil cotton", "australia cotton",
        "cotlook", "棉纱", "upland cotton", "pima cotton",
        "national cotton council", "cotton incorporated",
        "olam cotton", "louis dreyfus cotton",
    ],
    "🟡 棕榈油": [
        "palm oil", "棕榈油", "cpo", "mpob", "palm",
        "crude palm oil", "palm kernel", "palm olein", "rbd palm",
        "indonesia palm", "malaysia palm", "bmd palm", "大马棕", "印尼棕",
        "wilmar", "丰益国际", "sime darby plantation", "ioi corporation",
        "kuala lumpur kepong", "genting plantations", "golden agri",
        "musim mas", "astra agro", "first resources", "felda",
        "indofood agri", "gapki", "mpoc", "rspo", "oil world",
    ],
    "🌽 玉米": [
        "corn", "玉米", "maize", "cbot corn", "c futures",
        "usda corn", "corn belt", "corn yield", "corn export",
        "brazil corn", "argentina corn", "ukraine corn", "us corn",
        "safrinha", "china corn", "中国玉米",
        "ethanol", "玉米淀粉", "feed corn", "玉米饲料",
        "adm", "archer daniels midland", "bunge", "cargill", "cofco",
        "中粮", "louis dreyfus", "syngenta corn", "bayer corn",
        # EU/JRC 作物监测（MARS Bulletin 月报）
        "MARS bulletin", "JRC MARS", "eu crop forecast", "european crop forecast",
        "eu maize", "european maize", "eu corn", "crop monitoring europe",
        # 美国玉米深度关键词
        "crop tour", "planting pace", "crop condition", "trendline yield",
        "corn rootworm", "derecho", "stonex grain", "allendale",
        "us grains council", "growth energy", "farmdoc",
        "grain stocks", "corn stocks", "corn acreage", "prospective plantings",
        "crop progress", "iowa corn", "illinois corn", "nebraska corn",
        "e15", "ethanol blend", "rfs mandate", "renewable fuel standard",
    ],
    "🌳 橡胶": [
        "rubber", "橡胶", "natural rubber", "tsr", "rss sheet",
        "tocom rubber", "shfe rubber", "sicom rubber", "anrpc",
        "天然橡胶", "合成橡胶",
        "thailand rubber", "indonesia rubber", "vietnam rubber",
        "malaysia rubber", "ivory coast rubber", "india rubber",
        "轮胎", "tire demand", "tire production", "bridgestone",
        "michelin", "goodyear", "continental tire", "pirelli",
        "hankook", "sumitomo rubber", "中策橡胶", "zc rubber",
        "halcyon agri", "sri trang", "von bundit",
    ],
}

PRIORITY_KEYWORDS = [
    "weather", "drought", "flood", "frost", "rain", "hurricane", "typhoon",
    "heatwave", "el nino", "la nina", "monsoon", "cyclone",
    "天气", "干旱", "洪水", "霜冻", "降雨", "台风", "热浪", "厄尔尼诺",
    "harvest", "crop", "production", "output", "yield", "planting",
    "acreage", "inventory", "stockpile", "warehouse",
    "supply", "demand", "surplus", "shortage", "deficit",
    "减产", "增产", "产量", "收成", "库存", "供应", "需求", "过剩", "短缺",
    "tariff", "ban", "sanction", "export", "import", "quota", "subsidy",
    "embargo", "trade war", "anti-dumping", "section 232", "section 301",
    "ustr", "wto",
    "关税", "禁令", "制裁", "出口", "进口", "配额", "反倾销", "贸易战", "加征",
    "lme stocks", "lme inventory", "shfe stocks", "comex stocks",
    "warehouse receipts", "position limit", "delivery notice",
    "open interest", "cot report",
    "上期所库存", "保税区库存", "交割", "持仓", "仓单",
    "mine", "smelter", "refinery", "strike", "shutdown", "closure",
    "accident", "force majeure", "disruption",
    "矿山", "冶炼厂", "炼厂", "罢工", "停产", "事故", "不可抗力",
    "usda", "anrpc", "lme", "shfe", "cftc", "opec", "wasde", "ico",
    "mpob", "cofco", "report", "forecast", "outlook",
    "上调", "下调", "预测", "报告",
    # EU JRC 报告
    "MARS bulletin", "JRC MARS", "JRC crop", "EU crop",
]

# 命中则跳过（财报/股东会噪音）
EXCLUDE_KEYWORDS = [
    "earnings", "quarterly results", "q1 result", "q2 result", "q3 result",
    "q4 result", "first quarter", "second quarter", "third quarter",
    "fourth quarter", "half-year", "interim result", "annual result",
    "annual report", "trading update", "agm", "annual general meeting",
    "investor day", "capital markets day", "dividend", "share buyback",
    "stock split", "rights issue", "ipo", "secondary offering",
    "财报", "年报", "季报", "半年报", "中报", "业绩公告", "业绩报告",
    "营业收入", "净利润", "扣非", "派息", "分红", "股东大会", "股票回购",
]

RSS_FEEDS = [
    # 金属/矿业
    ("Kitco Metals",            "https://www.kitco.com/rss/rss.xml"),
    ("Mining.com",              "https://www.mining.com/feed/"),
    ("Northern Miner",          "https://www.northernminer.com/feed/"),
    ("Mining Weekly",           "https://www.miningweekly.com/topic/mining/feed"),
    ("Metal Miner",             "https://agmetalminer.com/feed/"),
    ("Aluminum Insider",        "https://aluminiuminsider.com/feed/"),
    ("Steel Guru",              "https://www.steelguru.com/rss/latest/all"),
    ("SMM 有色网",                "https://news.smm.cn/rss/1"),
    ("我的钢铁网",                "https://news.mysteel.com/rss/all.xml"),
    # 农产品
    ("Farm Progress",           "https://www.farmprogress.com/rss.xml"),
    ("AgriMoney",               "https://www.agrimoney.com/rss"),
    ("DTN Ag",                  "https://www.dtnpf.com/agriculture/web/ag/rss/news"),
    ("Brownfield Ag",           "https://brownfieldagnews.com/feed/"),
    ("Feedstuffs",              "https://www.feedstuffs.com/rss.xml"),
    ("Farm Futures",            "https://www.farmfutures.com/rss.xml"),
    ("Pro Farmer",              "https://www.profarmer.com/rss.xml"),
    ("Agri Pulse",              "https://www.agri-pulse.com/rss"),
    ("USDA FAS",                "https://www.fas.usda.gov/rss-feeds/all"),
    ("FAO News",                "https://www.fao.org/news/rss-feed/en/"),
    # 棕榈油/橡胶（东南亚）
    ("The Edge Markets",        "https://www.theedgemarkets.com/rss.xml"),
    ("Bernama Business",        "https://www.bernama.com/en/business/rss.php"),
    ("Malaysia Reserve",        "https://themalaysianreserve.com/feed/"),
    ("Rubber News",             "https://www.rubbernews.com/rss/all"),
    ("Tire Business",           "https://www.tirebusiness.com/rss/all"),
    # 棉花（印度/巴基斯坦消费大国）
    ("Economic Times Commodities","https://economictimes.indiatimes.com/markets/commodities/rssfeeds/1808151502.cms"),
    ("LiveMint Markets",        "https://www.livemint.com/rss/markets"),
    ("Moneycontrol Commodities","https://www.moneycontrol.com/rss/commodityfutures.xml"),
    ("Business Recorder PK",    "https://www.brecorder.com/feed"),
    # 综合财经
    ("Yahoo Metals Futures",    "https://feeds.finance.yahoo.com/rss/2.0/headline?s=HG%3DF%2CALI%3DF%2CPL%3DF%2CPA%3DF&region=US&lang=en-US"),
    ("Yahoo Agri Futures",      "https://feeds.finance.yahoo.com/rss/2.0/headline?s=CT%3DF%2CC%3DF%2CS%3DF%2CW%3DF&region=US&lang=en-US"),
    ("Investing Commodities CN","https://cn.investing.com/rss/news_285.rss"),
    ("Investing Commodities EN","https://www.investing.com/rss/commodities.rss"),
    ("CNBC Commodities",        "https://www.cnbc.com/id/20910258/device/rss/rss.html"),
    ("OilPrice Commodities",    "https://oilprice.com/rss/commodity-news"),
    ("FXStreet",                "https://www.fxstreet.com/rss/news"),
    ("Trading Economics",       "https://tradingeconomics.com/rss/news.aspx"),
    ("Nikkei Asia Business",    "https://asia.nikkei.com/rss/feed/nar"),
    ("Caixin Global",           "https://www.caixinglobal.com/rss.xml"),
    ("SCMP Business",           "https://www.scmp.com/rss/92/feed"),
    ("第一财经",                  "https://www.yicai.com/news/rss.xml"),
    ("期货日报",                  "http://www.qhrb.com.cn/rss/news.xml"),
    ("百川资讯",                  "https://www.100ppi.com/rss/news.xml"),
    # 交易所 / 贸易政策
    ("CME News",                "https://www.cmegroup.com/CmeWS/mvc/news/rss/news.xml"),
    ("LME News",                "https://www.lme.com/en/news-and-events/news/rss"),
    ("USTR",                    "https://ustr.gov/about-us/policy-offices/press-office/press-releases/feed"),
    ("White House",             "https://www.whitehouse.gov/briefing-room/feed/"),
    ("WTO News",                "https://www.wto.org/english/news_e/news_e.xml"),
    ("EU Trade",                "https://policy.trade.ec.europa.eu/news/rss_en"),
    ("JRC Publications",        "https://publications.jrc.ec.europa.eu/repository/feed/rss_2.0/site"),
    ("China MOFCOM",            "http://english.mofcom.gov.cn/rss/news.xml"),
    # USDA 玉米专属高频
    ("USDA WASDE",              "https://www.usda.gov/oce/commodity/wasde/rss/wasde.xml"),
    ("USDA NASS 作物进度",       "https://www.nass.usda.gov/rss/cropprogress.xml"),
    ("USDA NASS 官方报告",       "https://www.nass.usda.gov/Newsroom/Syndication/News/index.php"),
    ("USDA ERS",                "https://www.ers.usda.gov/rss/feed.xml"),
    ("AgWeb News",              "https://www.agweb.com/rss.xml"),
    ("Farm Journal",            "https://www.farmjournal.com/rss.xml"),
    # Farm Journal YouTube（已确认 Channel ID）
    ("Farm Journal YouTube",    "https://www.youtube.com/feeds/videos.xml?channel_id=UCqTizkFSbhBojVtIkjMEoZQ"),
    # 美国玉米专属深度源
    ("US Grains Council",       "https://grains.org/feed/"),
    ("Growth Energy 乙醇",       "https://growthenergy.org/feed/"),
    ("farmdoc Daily",           "https://farmdocdaily.illinois.edu/feed"),
    ("Iowa State 作物生产",       "https://crops.extension.iastate.edu/rss/category/crop-production"),
    ("AgWired",                 "https://agwired.com/feed/"),
    ("EIA Today in Energy",     "https://www.eia.gov/rss/todayinenergy.xml"),
    ("High Plains Journal",     "https://hpj.com/feed/"),
    ("DTN Ag News",             "https://www.dtnpf.com/agriculture/web/ag/rss/news"),
    ("Brownfield Ag News",      "https://brownfieldagnews.com/feed/"),
    ("Pro Farmer",              "https://www.profarmer.com/rss.xml"),
    # 钢联 Mysteel 各板块
    ("Mysteel 有色",             "https://list.mysteel.com/rss/c-1024.xml"),
    ("Mysteel 铜",               "https://list.mysteel.com/rss/c-1024-1045.xml"),
    ("Mysteel 铝",               "https://list.mysteel.com/rss/c-1024-1046.xml"),
    ("Mysteel 农产品",           "https://list.mysteel.com/rss/c-1165.xml"),
    ("Mysteel 玉米",             "https://list.mysteel.com/rss/c-1165-1166.xml"),
    ("Mysteel 棉花",             "https://list.mysteel.com/rss/c-1165-1170.xml"),
    ("Mysteel 橡胶",             "https://list.mysteel.com/rss/c-1024-1055.xml"),
    # 生意社各板块
    ("生意社-有色",              "https://www.100ppi.com/rss/news-metal.xml"),
    ("生意社-农副",              "https://www.100ppi.com/rss/news-agri.xml"),
    ("生意社-橡塑",              "https://www.100ppi.com/rss/news-rubber.xml"),
    # 权威财经媒体（Bloomberg 无公开 RSS，通过 Google News 聚合抓取）
    ("Reuters Business",          "https://feeds.reuters.com/reuters/businessNews"),
    ("Reuters Commodities",       "https://feeds.reuters.com/reuters/companyNews"),
    ("FT Commodities",            "https://www.ft.com/commodities?format=rss"),
    ("MarketWatch Markets",       "https://feeds.marketwatch.com/marketwatch/marketpulse/"),
    # Google News RSS：聚合 Bloomberg / Reuters / FT 等权威来源
    ("GNews Bloomberg 铝",        "https://news.google.com/rss/search?q=bloomberg+aluminum+market&hl=en-US&gl=US&ceid=US:en"),
    ("GNews Bloomberg 铜",        "https://news.google.com/rss/search?q=bloomberg+copper+commodity&hl=en-US&gl=US&ceid=US:en"),
    ("GNews Bloomberg 棕榈油",    "https://news.google.com/rss/search?q=bloomberg+palm+oil&hl=en-US&gl=US&ceid=US:en"),
    ("GNews Bloomberg 橡胶",      "https://news.google.com/rss/search?q=bloomberg+rubber+commodity&hl=en-US&gl=US&ceid=US:en"),
    ("GNews Bloomberg 棉花",      "https://news.google.com/rss/search?q=bloomberg+cotton+commodity&hl=en-US&gl=US&ceid=US:en"),
    ("GNews Bloomberg 玉米",      "https://news.google.com/rss/search?q=bloomberg+corn+grain+market&hl=en-US&gl=US&ceid=US:en"),
    ("GNews Bloomberg 铂钯",      "https://news.google.com/rss/search?q=bloomberg+platinum+palladium&hl=en-US&gl=US&ceid=US:en"),
    ("GNews 商品期货综合",         "https://news.google.com/rss/search?q=commodity+futures+bloomberg+reuters&hl=en-US&gl=US&ceid=US:en"),
    # 公司（保留但财报已被 EXCLUDE 过滤）
    ("Freeport-McMoRan",        "https://investors.fcx.com/rss/news-releases.xml"),
    ("Glencore",                "https://www.glencore.com/rss/news"),
    ("Anglo American",          "https://www.angloamerican.com/rss/news"),
    ("Wilmar",                  "https://ir.wilmar-international.com/rss/announcements.xml"),
    ("Bridgestone",             "https://www.bridgestone.com/corporate/news/rss.xml"),
    ("Michelin",                "https://www.michelin.com/en/feed/"),
    ("Alcoa",                   "https://news.alcoa.com/press-releases?format=rss"),
]

# ── 匹配 ──────────────────────────────────────────────────
import re
import urllib.parse

_CACHE = {}
_TRANSLATE_CACHE = {}

def translate_to_zh(text: str) -> str:
    """Google 免费翻译接口，失败返回原文"""
    if not text or not text.strip():
        return text
    zh_chars = sum(1 for c in text if "一" <= c <= "鿿")
    if zh_chars / max(len(text), 1) > 0.3:
        return text
    text_cut = text[:300]
    if text_cut in _TRANSLATE_CACHE:
        return _TRANSLATE_CACHE[text_cut]
    try:
        url = (
            "https://translate.googleapis.com/translate_a/single"
            "?client=gtx&sl=auto&tl=zh-CN&dt=t&q="
            + urllib.parse.quote(text_cut)
        )
        r = requests.get(url, timeout=5,
                         headers={"User-Agent": "Mozilla/5.0"})
        data = r.json()
        result = "".join(seg[0] for seg in data[0] if seg[0])
        _TRANSLATE_CACHE[text_cut] = result
        return result
    except Exception:
        return text

def _kw_match(text_lower, kw):
    kwl = kw.lower()
    if any("\u4e00" <= c <= "\u9fff" for c in kwl):
        return kwl in text_lower
    if len(kwl) <= 4:
        p = _CACHE.get(kwl)
        if p is None:
            p = re.compile(r"\b" + re.escape(kwl) + r"\b")
            _CACHE[kwl] = p
        return bool(p.search(text_lower))
    return kwl in text_lower

def match_commodity(text):
    t = text.lower()
    if any(_kw_match(t, kw) for kw in EXCLUDE_KEYWORDS):
        return []
    return [name for name, kws in COMMODITIES.items()
            if any(_kw_match(t, kw) for kw in kws)]

def is_priority(text):
    t = text.lower()
    return any(_kw_match(t, kw) for kw in PRIORITY_KEYWORDS)

# ── 抓取 ──────────────────────────────────────────────────
def parse_pub_time(entry):
    """尝试解析文章发布时间，失败返回 None"""
    for field in ("published_parsed", "updated_parsed"):
        t = getattr(entry, field, None)
        if t:
            try:
                return datetime(*t[:6], tzinfo=timezone.utc)
            except Exception:
                pass
    return None

def article_id(entry_link, entry_title):
    raw = (entry_link or entry_title or "").strip()
    return hashlib.md5(raw.encode("utf-8")).hexdigest()

def fetch_feed(nu):
    name, url = nu
    try:
        feed = feedparser.parse(
            url, request_headers={"User-Agent": "Mozilla/5.0"}
        )
        out = []
        for e in feed.entries[:50]:
            link = getattr(e, "link", "")
            title = getattr(e, "title", "")
            out.append({
                "id":      article_id(link, title),
                "title":   title,
                "summary": getattr(e, "summary", "")[:400],
                "link":    link,
                "source":  name,
                "pub":     parse_pub_time(e),
            })
        return out
    except Exception:
        return []

def fetch_all():
    articles = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as ex:
        futures = {ex.submit(fetch_feed, nu): nu for nu in RSS_FEEDS}
        done, not_done = concurrent.futures.wait(futures, timeout=45)
        for f in done:
            try:
                articles.extend(f.result(timeout=1))
            except Exception:
                pass
        for f in not_done:
            f.cancel()
        if not_done:
            print(f"[超时] {len(not_done)} 个 RSS 源未返回")
    return articles

# ── 推送 ──────────────────────────────────────────────────
def push_wechat(title, content_html):
    """PushPlus 发微信，支持群组推送。"""
    if not PUSHPLUS_TOKEN:
        print("[错误] 未设置 PUSHPLUS_TOKEN 环境变量")
        return False
    try:
        payload = {
            "token":    PUSHPLUS_TOKEN,
            "title":    title[:80],
            "content":  content_html,
            "template": "html",
        }
        # 群组推送：有 topic 时同时推给所有群成员
        if PUSHPLUS_TOPIC:
            payload["topic"] = PUSHPLUS_TOPIC
        r = requests.post(
            "http://www.pushplus.plus/send",
            json=payload,
            timeout=15,
        )
        resp_json = r.json()
        ok = r.status_code == 200 and resp_json.get("code") == 200
        if ok:
            mode = f"群组 topic={PUSHPLUS_TOPIC}" if PUSHPLUS_TOPIC else "个人"
            print(f"[推送成功] {mode} | 返回: {resp_json.get('msg', '')}")
        else:
            print(f"[推送失败] HTTP={r.status_code} code={resp_json.get('code')} msg={resp_json.get('msg','')}")
        return ok
    except Exception as e:
        print(f"[推送异常] {e}")
        return False

# ── 主逻辑 ────────────────────────────────────────────────
def load_seen():
    if SEEN_FILE.exists():
        try:
            return set(json.loads(SEEN_FILE.read_text(encoding="utf-8")))
        except Exception:
            pass
    return set()

def save_seen(seen):
    items = list(seen)[-5000:]
    SEEN_FILE.write_text(json.dumps(items), encoding="utf-8")

def main():
    if not PUSHPLUS_TOKEN:
        print("错误：请设置 PUSHPLUS_TOKEN 环境变量（GitHub Secrets）")
        sys.exit(1)

    # 启动时打印推送模式，方便在 Actions 日志里确认
    if PUSHPLUS_TOPIC:
        print(f"[推送模式] 群组推送 topic={PUSHPLUS_TOPIC!r}（群组成员均可收到）")
    else:
        print("[推送模式] 个人推送（仅 token 持有人收到）— 如需群推请设置 PUSHPLUS_TOPIC Secret")

    seen = load_seen()
    first_run = len(seen) == 0
    print(f"[{datetime.now()}] 开始抓取... (已读库: {len(seen)} 篇)")

    articles = fetch_all()
    print(f"共抓取 {len(articles)} 篇")

    max_age_cutoff = datetime.now(timezone.utc) - timedelta(hours=MAX_AGE_HOURS)
    hits = []
    for a in articles:
        if a["id"] in seen:
            continue  # 已推过，跳过
        combined = f"{a['title']} {a['summary']}"
        commodities = match_commodity(combined)
        if not commodities:
            seen.add(a["id"])
            continue
        # 兜底：有发布时间且太旧（超过 MAX_AGE_HOURS）的不推
        if a["pub"] and a["pub"] < max_age_cutoff:
            seen.add(a["id"])
            continue
        a["commodities"] = commodities
        a["priority"] = is_priority(combined)
        hits.append(a)
        seen.add(a["id"])

    # 去重（同标题）
    seen_titles = set()
    unique = []
    for a in hits:
        k = a["title"].strip().lower()
        if k and k not in seen_titles:
            seen_titles.add(k)
            unique.append(a)

    # 首次运行：只记录，不刷屏
    if first_run:
        print(f"首次运行：记录 {len(unique)} 条命中但不推送（避免刷屏）")
        save_seen(seen)
        return

    unique = unique[:MAX_PUSH_PER_RUN]
    print(f"本轮新命中：{len(unique)} 条")
    save_seen(seen)

    if not unique:
        print("无新命中，退出")
        return

    if DIGEST_MODE:
        lines = [f"<p>📊 <b>商品期货新命中 {len(unique)} 条</b></p><hr/>"]
        for a in unique:
            tag = " ".join(a["commodities"])
            prio = "⚡ " if a["priority"] else ""
            zh_title   = translate_to_zh(a["title"])
            zh_summary = translate_to_zh(a["summary"][:200] or a["title"])
            lines.append(
                f'<p>{prio}<b>{tag}</b><br/>'
                f'<a href="{a["link"]}">{zh_title}</a><br/>'
                f'<small>[{a["source"]}] {zh_summary}</small></p>'
            )
        title = f"🔔 商品期货 {len(unique)} 条 ({datetime.now().strftime('%H:%M')})"
        push_wechat(title, "".join(lines))
    else:
        for a in unique:
            tag = " ".join(a["commodities"])
            prio = "⚡ " if a["priority"] else ""
            zh_title   = translate_to_zh(a["title"])
            zh_summary = translate_to_zh(a["summary"] or a["title"])
            title = f"{prio}{tag}｜{zh_title[:40]}"
            link = a["link"]
            source = a["source"]
            content = (
                f'<p><b>{zh_title}</b></p>'
                f'<p>{zh_summary}</p>'
                f'<p><a href="{link}">查看原文</a></p>'
                f'<p><small>来源: {source}</small></p>'
            )
            push_wechat(title, content)
            time.sleep(0.5)

    print("推送完成")

if __name__ == "__main__":
    main()
