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

socket.setdefaulttimeout(10)

# ── 环境变量 ──────────────────────────────────────────────
PUSHPLUS_TOKEN = os.environ.get("PUSHPLUS_TOKEN", "").strip()
# 最大推送条数（防刷屏 & 控制 PushPlus 日额度）
MAX_PUSH_PER_RUN = int(os.environ.get("MAX_PUSH_PER_RUN", "30"))
# 是否合并成一条汇总消息（推荐 True，节省 PushPlus 额度）
DIGEST_MODE = os.environ.get("DIGEST_MODE", "1") == "1"
# 兜底时间窗（小时）：即使是没见过的文章，超过这个时间也不推（避免首次跑刷屏）
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
]

RSS_FEEDS = [
    ("Mining.com",            "https://www.mining.com/feed/"),
    ("Northern Miner",        "https://www.northernminer.com/feed/"),
    ("Kitco Metals",          "https://www.kitco.com/rss/rss.xml"),
    ("Farm Progress",         "https://www.farmprogress.com/rss.xml"),
    ("AgriMoney",             "https://www.agrimoney.com/rss"),
    ("Investing CN",          "https://cn.investing.com/rss/news_285.rss"),
    ("Investing EN",          "https://www.investing.com/rss/commodities.rss"),
    ("Yahoo Metals",          "https://feeds.finance.yahoo.com/rss/2.0/headline?s=HG%3DF%2CALI%3DF%2CPL%3DF%2CPA%3DF&region=US&lang=en-US"),
    ("Yahoo Agri",            "https://feeds.finance.yahoo.com/rss/2.0/headline?s=CT%3DF%2CC%3DF%2CS%3DF%2CW%3DF&region=US&lang=en-US"),
    ("CNBC Commodities",      "https://www.cnbc.com/id/20910258/device/rss/rss.html"),
    ("OilPrice Commodities",  "https://oilprice.com/rss/commodity-news"),
    ("Rubber News",           "https://www.rubbernews.com/rss/all"),
    ("Tire Business",         "https://www.tirebusiness.com/rss/all"),
    ("The Edge Markets",      "https://www.theedgemarkets.com/rss.xml"),
    ("Bernama Business",      "https://www.bernama.com/en/business/rss.php"),
    ("USTR",                  "https://ustr.gov/about-us/policy-offices/press-office/press-releases/feed"),
    ("White House",           "https://www.whitehouse.gov/briefing-room/feed/"),
    ("WTO",                   "https://www.wto.org/english/news_e/news_e.xml"),
    ("Caixin Global",         "https://www.caixinglobal.com/rss.xml"),
    ("SCMP Business",         "https://www.scmp.com/rss/92/feed"),
]

# ── 匹配 ──────────────────────────────────────────────────
import re
_CACHE = {}

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
        for result in ex.map(fetch_feed, RSS_FEEDS, timeout=60):
            articles.extend(result)
    return articles

# ── 推送 ──────────────────────────────────────────────────
def push_wechat(title, content_html):
    """PushPlus 发微信。content 支持 HTML。"""
    if not PUSHPLUS_TOKEN:
        print("[错误] 未设置 PUSHPLUS_TOKEN 环境变量")
        return False
    try:
        r = requests.post(
            "http://www.pushplus.plus/send",
            json={
                "token":    PUSHPLUS_TOKEN,
                "title":    title[:80],
                "content":  content_html,
                "template": "html",
            },
            timeout=15,
        )
        ok = r.status_code == 200 and r.json().get("code") == 200
        if not ok:
            print(f"[推送失败] {r.status_code} {r.text[:200]}")
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
        # 汇总成一条 HTML 消息
        lines = [
            f"<p>📊 <b>商品期货新命中 {len(unique)} 条</b></p><hr/>"
        ]
        for a in unique:
            tag = " ".join(a["commodities"])
            prio = "⚡ " if a["priority"] else ""
            lines.append(
                f'<p>{prio}<b>{tag}</b><br/>'
                f'<a href="{a["link"]}">{a["title"]}</a><br/>'
                f'<small>[{a["source"]}] {a["summary"][:150]}</small></p>'
            )
        title = f"🔔 商品期货 {len(unique)} 条 ({datetime.now().strftime('%H:%M')})"
        push_wechat(title, "".join(lines))
    else:
        for a in unique:
            tag = " ".join(a["commodities"])
            prio = "⚡ " if a["priority"] else ""
            title = f"{prio}{tag}｜{a['title'][:40]}"
            content = (
                f'<p><b>{a["title"]}</b></p>'
                f'<p>{a["summary"]}</p>'
                f'<p><a href="{a["link"]}">查看原文</a></p>'
                f'<p><small>来源: {a["source"]}</small></p>'
            )
            push_wechat(title, content)
            time.sleep(0.5)

    print("推送完成")

if __name__ == "__main__":
    main()
