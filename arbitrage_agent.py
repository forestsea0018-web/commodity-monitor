#!/usr/bin/env python3
"""
MSTR-BMNR 套利交易 AI Agent
功能：
  1. 接收 TradingView 策略信号（via GitHub repository_dispatch）
  2. 实时抓取行情数据 & 计算技术指标
  3. Claude AI (claude-opus-4-7) 深度分析 + 自适应思考
  4. 微信实时推送 (PushPlus)
  5. Alpaca 券商 API 自动化执行交易（可选）

GitHub Secrets 需要配置：
  PUSHPLUS_TOKEN      - PushPlus 微信推送 Token
  ANTHROPIC_API_KEY   - Anthropic Claude API Key
  ALPACA_API_KEY      - Alpaca 交易账户 Key（可选）
  ALPACA_SECRET_KEY   - Alpaca 交易账户 Secret（可选）

TradingView Webhook 接入：
  Webhook URL: https://api.github.com/repos/{owner}/{repo}/dispatches
  Headers:    Authorization: token {GITHUB_PAT}
              Content-Type: application/json
  Body:       {"event_type":"tradingview-signal","client_payload":{
                "symbol":"MSTR","action":"BUY","price":500.0,
                "strategy":"arbitrage","timeframe":"1H"}}
"""

import os
import sys
import json
import requests
import numpy as np
from datetime import datetime, timezone, timedelta

# ── 环境变量 ─────────────────────────────────────────────────────────────────
PUSHPLUS_TOKEN     = os.environ.get("PUSHPLUS_TOKEN", "").strip()
ANTHROPIC_API_KEY  = os.environ.get("ANTHROPIC_API_KEY", "").strip()
DEEPSEEK_API_KEY   = os.environ.get("DEEPSEEK_API_KEY", "").strip()
# "anthropic" 或 "deepseek"，未设置时根据哪个 key 存在自动选择
AI_PROVIDER        = os.environ.get("AI_PROVIDER", "").strip().lower()
ALPACA_API_KEY     = os.environ.get("ALPACA_API_KEY", "").strip()
ALPACA_SECRET_KEY  = os.environ.get("ALPACA_SECRET_KEY", "").strip()

# TradingView 信号（由 GitHub Actions 通过 repository_dispatch 传入）
TV_PAYLOAD = os.environ.get("TV_PAYLOAD", "").strip()

# 交易配置
FORCE_ANALYSIS   = os.environ.get("FORCE_ANALYSIS", "false") == "true"
AUTO_TRADE       = os.environ.get("AUTO_TRADE", "0") == "1"
PAPER_TRADING    = os.environ.get("PAPER_TRADING", "1") == "1"
SYMBOL_A         = os.environ.get("SYMBOL_A", "MSTR")
SYMBOL_B         = os.environ.get("SYMBOL_B", "BMNR")
ENTRY_ZSCORE     = float(os.environ.get("ENTRY_ZSCORE", "2.0"))
EXIT_ZSCORE      = float(os.environ.get("EXIT_ZSCORE", "0.5"))
LOOKBACK_DAYS    = int(os.environ.get("LOOKBACK_DAYS", "30"))
POSITION_PCT     = float(os.environ.get("POSITION_PCT", "0.05"))   # 每次交易占账户比例
STOP_LOSS_PCT    = float(os.environ.get("STOP_LOSS_PCT", "0.03"))   # 止损 3%
TAKE_PROFIT_PCT  = float(os.environ.get("TAKE_PROFIT_PCT", "0.06")) # 止盈 6%

ALPACA_BASE = (
    "https://paper-api.alpaca.markets"
    if PAPER_TRADING
    else "https://api.alpaca.markets"
)

# ── 系统提示（缓存到 Claude API，节省 token 费用）────────────────────────────
SYSTEM_PROMPT = """你是一位专业的量化交易 AI 分析师，专职负责 MSTR（MicroStrategy）和 BMNR 之间的统计套利分析。

## 核心背景
- MSTR: MicroStrategy，大量持有比特币的上市公司，其股价与 BTC NAV（净资产值）存在溢价或折价
- BMNR: 比特币矿工相关股票，与 BTC 价格高度相关但 Beta 系数不同
- 两者均受比特币价格驱动，但因公司基本面和市场情绪不同，价差会周期性偏离和回归

## 交易策略框架
本系统采用统计套利（均值回归）策略：
- 计算 log(MSTR/BMNR) 价差的 Z-Score
- Z-Score > +{entry}: MSTR 相对高估 → 做空 MSTR / 做多 BMNR（负向配对交易）
- Z-Score < -{entry}: MSTR 相对低估 → 做多 MSTR / 做空 BMNR（正向配对交易）
- |Z-Score| < {exit}: 价差收敛 → 平仓获利

## 信号质量判断标准
优质信号（综合评分 ≥ 70）需满足多数条件：
1. Z-Score 在多个时间窗口（10/20/30日）方向一致
2. RSI 显示同向动量（MSTR RSI > 70 或 < 30 时更有说服力）
3. 成交量放大（Volume Ratio > 1.5x 确认趋势）
4. BTC 市场方向支持该方向
5. 无重大财报/宏观事件风险

## 严格输出格式
你必须且只能以下列 JSON 格式响应（无额外文字）：
{{
  "decision": "PAIRS_LONG_A_SHORT_B" | "PAIRS_LONG_B_SHORT_A" | "CLOSE_POSITION" | "HOLD",
  "confidence": <整数 0-100>,
  "action_mstr": "BUY" | "SELL" | "HOLD",
  "action_bmnr": "BUY" | "SELL" | "HOLD",
  "position_size_pct": <建议仓位占账户比例, 0.01-0.10>,
  "stop_loss_pct": <止损比例, 如 0.03>,
  "take_profit_pct": <止盈比例, 如 0.06>,
  "reasoning": "<详细中文分析，至少 100 字>",
  "risk_level": "LOW" | "MEDIUM" | "HIGH",
  "market_regime": "TRENDING" | "MEAN_REVERTING" | "VOLATILE" | "UNCLEAR",
  "key_factors": ["关键因素1", "关键因素2", "关键因素3"],
  "tv_signal_consistent": true | false | null
}}
""".format(entry=ENTRY_ZSCORE, exit=EXIT_ZSCORE)


# ── 技术指标计算 ──────────────────────────────────────────────────────────────
def ema(prices: np.ndarray, period: int) -> np.ndarray:
    k = 2.0 / (period + 1)
    out = np.empty_like(prices)
    out[0] = prices[0]
    for i in range(1, len(prices)):
        out[i] = prices[i] * k + out[i - 1] * (1 - k)
    return out


def calc_rsi(prices: np.ndarray, period: int = 14) -> float:
    if len(prices) < period + 1:
        return 50.0
    deltas = np.diff(prices)
    gains = np.where(deltas > 0, deltas, 0.0)
    losses = np.where(deltas < 0, -deltas, 0.0)
    avg_g = np.mean(gains[:period])
    avg_l = np.mean(losses[:period])
    for i in range(period, len(deltas)):
        avg_g = (avg_g * (period - 1) + gains[i]) / period
        avg_l = (avg_l * (period - 1) + losses[i]) / period
    if avg_l == 0:
        return 100.0
    return float(100 - 100 / (1 + avg_g / avg_l))


def calc_macd(prices: np.ndarray) -> tuple[float, float, float]:
    if len(prices) < 35:
        return 0.0, 0.0, 0.0
    fast = ema(prices, 12)
    slow = ema(prices, 26)
    macd_line = fast - slow
    signal = ema(macd_line, 9)
    hist = macd_line - signal
    return float(macd_line[-1]), float(signal[-1]), float(hist[-1])


def calc_bollinger_pct(prices: np.ndarray, period: int = 20) -> float:
    """Bollinger %B: 0=下轨, 0.5=中轨, 1=上轨"""
    if len(prices) < period:
        return 0.5
    window = prices[-period:]
    mean = np.mean(window)
    std = np.std(window)
    if std == 0:
        return 0.5
    upper = mean + 2 * std
    lower = mean - 2 * std
    return float((prices[-1] - lower) / (upper - lower))


def calc_zscore(series: np.ndarray, window: int) -> float:
    if len(series) < window + 1:
        return 0.0
    recent = series[-(window + 1):-1]  # exclude current point
    mean = np.mean(recent)
    std = np.std(recent)
    if std == 0:
        return 0.0
    return float((series[-1] - mean) / std)


def calc_volume_ratio(volumes: np.ndarray, period: int = 20) -> float:
    if len(volumes) < period + 1:
        return 1.0
    avg = np.mean(volumes[-period - 1:-1])
    return float(volumes[-1] / avg) if avg > 0 else 1.0


# ── 行情数据抓取 ──────────────────────────────────────────────────────────────
def fetch_market_data() -> dict | None:
    """通过 yfinance 抓取 MSTR、BMNR、BTC 行情及技术指标"""
    try:
        import yfinance as yf
    except ImportError:
        print("ERROR: yfinance not installed. Run: pip install yfinance")
        sys.exit(1)

    symbols = {
        "mstr": SYMBOL_A,
        "bmnr": SYMBOL_B,
        "btc": "BTC-USD",
    }
    period = f"{LOOKBACK_DAYS + 15}d"
    result = {}

    for key, sym in symbols.items():
        try:
            ticker = yf.Ticker(sym)
            hist = ticker.history(period=period)
            if hist.empty or len(hist) < 10:
                print(f"WARNING: Insufficient data for {sym}")
                if key in ("mstr", "bmnr"):
                    return None
                continue

            closes = hist["Close"].values.astype(float)
            volumes = hist["Volume"].values.astype(float)

            rsi = calc_rsi(closes)
            macd_v, macd_sig, macd_hist = calc_macd(closes)
            bb_pct = calc_bollinger_pct(closes)
            vol_ratio = calc_volume_ratio(volumes)
            sma20 = float(np.mean(closes[-20:])) if len(closes) >= 20 else float(closes[-1])
            change_24h = float((closes[-1] / closes[-2] - 1) * 100) if len(closes) >= 2 else 0.0

            result[key] = {
                "symbol": sym,
                "price": float(closes[-1]),
                "change_24h": change_24h,
                "rsi": rsi,
                "macd": macd_v,
                "macd_signal": macd_sig,
                "macd_hist": macd_hist,
                "bb_pct": bb_pct,
                "sma20": sma20,
                "volume_ratio": vol_ratio,
                "closes": closes,
            }
        except Exception as e:
            print(f"ERROR fetching {sym}: {e}")
            if key in ("mstr", "bmnr"):
                return None

    if "mstr" not in result or "bmnr" not in result:
        return None

    # 计算价差 Z-Score（多个时间窗口）
    closes_a = result["mstr"]["closes"]
    closes_b = result["bmnr"]["closes"]
    min_len = min(len(closes_a), len(closes_b))
    closes_a = closes_a[-min_len:]
    closes_b = closes_b[-min_len:]

    log_ratio = np.log(closes_a / closes_b)
    current_log_ratio = float(log_ratio[-1])

    z30 = calc_zscore(log_ratio, 30)
    z20 = calc_zscore(log_ratio, 20)
    z10 = calc_zscore(log_ratio, 10)

    # 价差趋势方向
    if z30 > 0.5 and z20 > 0.5:
        spread_trend = "上升（MSTR相对强势）"
    elif z30 < -0.5 and z20 < -0.5:
        spread_trend = "下降（BMNR相对强势）"
    else:
        spread_trend = "震荡"

    result["spread"] = {
        "log_ratio": current_log_ratio,
        "zscore_30d": z30,
        "zscore_20d": z20,
        "zscore_10d": z10,
        "mean_30d": float(np.mean(log_ratio[-31:-1])) if len(log_ratio) > 31 else 0.0,
        "std_30d": float(np.std(log_ratio[-31:-1])) if len(log_ratio) > 31 else 1.0,
        "trend": spread_trend,
    }

    # 移除原始 closes 数组（不需要传给 Claude）
    for key in ("mstr", "bmnr", "btc"):
        if key in result:
            result[key].pop("closes", None)

    return result


# ── Claude AI 分析 ────────────────────────────────────────────────────────────
def analyze_with_claude(market_data: dict, tv_signal: dict | None = None) -> dict:
    """调用 Claude claude-opus-4-7 进行深度套利分析，使用自适应思考 + 提示缓存"""
    import anthropic

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    md = market_data
    sp = md["spread"]
    btc = md.get("btc", {})

    tv_section = ""
    if tv_signal:
        tv_section = f"""
### TradingView 策略信号
```json
{json.dumps(tv_signal, ensure_ascii=False, indent=2)}
```
请判断 TradingView 信号与统计套利信号是否一致（tv_signal_consistent 字段）。
"""

    user_prompt = f"""## 当前市场快照 — {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}

### 价格与波动
| 指标 | {SYMBOL_A} | {SYMBOL_B} |
|------|-----------|-----------|
| 当前价格 | ${md['mstr']['price']:.2f} | ${md['bmnr']['price']:.2f} |
| 24h 变化 | {md['mstr']['change_24h']:+.2f}% | {md['bmnr']['change_24h']:+.2f}% |
| 20日均线 | ${md['mstr']['sma20']:.2f} | ${md['bmnr']['sma20']:.2f} |
| 成交量比率 | {md['mstr']['volume_ratio']:.2f}x | {md['bmnr']['volume_ratio']:.2f}x |

### 技术指标
| 指标 | {SYMBOL_A} | {SYMBOL_B} |
|------|-----------|-----------|
| RSI(14) | {md['mstr']['rsi']:.1f} | {md['bmnr']['rsi']:.1f} |
| MACD | {md['mstr']['macd']:.4f} | {md['bmnr']['macd']:.4f} |
| MACD Signal | {md['mstr']['macd_signal']:.4f} | {md['bmnr']['macd_signal']:.4f} |
| MACD Hist | {md['mstr']['macd_hist']:.4f} | {md['bmnr']['macd_hist']:.4f} |
| Bollinger %B | {md['mstr']['bb_pct']:.3f} | {md['bmnr']['bb_pct']:.3f} |

### 套利价差（核心信号）
- **Z-Score (30日)**: {sp['zscore_30d']:+.3f} {'⚠️ 入场区域' if abs(sp['zscore_30d']) >= ENTRY_ZSCORE else ('✅ 平仓区域' if abs(sp['zscore_30d']) <= EXIT_ZSCORE else '⏳ 观望区域')}
- **Z-Score (20日)**: {sp['zscore_20d']:+.3f}
- **Z-Score (10日)**: {sp['zscore_10d']:+.3f}
- **当前 log 价差**: {sp['log_ratio']:.5f}
- **30日均值**: {sp['mean_30d']:.5f} | **标准差**: {sp['std_30d']:.5f}
- **价差趋势**: {sp['trend']}

### 比特币背景
- BTC 价格: ${btc.get('price', 0):.0f}
- BTC 24h: {btc.get('change_24h', 0):+.2f}%
{tv_section}
## 请综合以上数据，给出交易决策（严格 JSON 格式）。"""

    try:
        with client.messages.stream(
            model="claude-opus-4-7",
            max_tokens=2048,
            thinking={"type": "adaptive"},
            system=[{
                "type": "text",
                "text": SYSTEM_PROMPT,
                "cache_control": {"type": "ephemeral"},  # 缓存稳定的系统提示
            }],
            messages=[{"role": "user", "content": user_prompt}],
        ) as stream:
            response = stream.get_final_message()

        # 记录缓存使用情况
        usage = response.usage
        if hasattr(usage, "cache_read_input_tokens") and usage.cache_read_input_tokens:
            print(f"[Claude] 缓存命中: {usage.cache_read_input_tokens} tokens 已缓存读取")

        # 从响应中提取 JSON
        for block in response.content:
            if block.type == "text":
                text = block.text.strip()
                # 提取 JSON 块
                if "```json" in text:
                    json_str = text.split("```json")[1].split("```")[0].strip()
                elif text.startswith("{"):
                    json_str = text
                else:
                    start = text.find("{")
                    end = text.rfind("}") + 1
                    if start >= 0 and end > start:
                        json_str = text[start:end]
                    else:
                        continue
                try:
                    return json.loads(json_str)
                except json.JSONDecodeError:
                    continue

        return {"decision": "HOLD", "confidence": 0, "reasoning": "JSON 解析失败，保持观望", "risk_level": "HIGH"}

    except Exception as e:
        print(f"ERROR calling Claude API: {e}")
        return {"decision": "HOLD", "confidence": 0, "reasoning": f"Claude 调用失败: {e}", "risk_level": "HIGH"}


def analyze_with_deepseek(market_data: dict, tv_signal: dict | None = None) -> dict:
    """调用 DeepSeek API（OpenAI 兼容格式）进行套利分析"""
    sp = market_data["spread"]
    md = market_data
    btc = md.get("btc", {})

    tv_section = ""
    if tv_signal:
        tv_section = f"\n### TradingView 信号\n```json\n{json.dumps(tv_signal, ensure_ascii=False, indent=2)}\n```"

    user_prompt = f"""## 当前市场快照 — {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}

| 指标 | {SYMBOL_A} | {SYMBOL_B} |
|------|-----------|-----------|
| 价格 | ${md['mstr']['price']:.2f} | ${md['bmnr']['price']:.2f} |
| 24h% | {md['mstr']['change_24h']:+.2f}% | {md['bmnr']['change_24h']:+.2f}% |
| RSI | {md['mstr']['rsi']:.1f} | {md['bmnr']['rsi']:.1f} |
| MACD Hist | {md['mstr']['macd_hist']:.4f} | {md['bmnr']['macd_hist']:.4f} |
| Bollinger %B | {md['mstr']['bb_pct']:.3f} | {md['bmnr']['bb_pct']:.3f} |
| 成交量比率 | {md['mstr']['volume_ratio']:.2f}x | {md['bmnr']['volume_ratio']:.2f}x |

### 套利价差
- Z-Score (30日): {sp['zscore_30d']:+.3f}
- Z-Score (20日): {sp['zscore_20d']:+.3f}
- Z-Score (10日): {sp['zscore_10d']:+.3f}
- 趋势: {sp['trend']}

### BTC 背景
- 价格: ${btc.get('price', 0):.0f} | 24h: {btc.get('change_24h', 0):+.2f}%
{tv_section}

请给出交易决策，严格 JSON 格式：
{{
  "decision": "PAIRS_LONG_A_SHORT_B" | "PAIRS_LONG_B_SHORT_A" | "CLOSE_POSITION" | "HOLD",
  "confidence": 整数0-100,
  "action_mstr": "BUY"|"SELL"|"HOLD",
  "action_bmnr": "BUY"|"SELL"|"HOLD",
  "position_size_pct": 0.01-0.10,
  "stop_loss_pct": 止损比例,
  "take_profit_pct": 止盈比例,
  "reasoning": "详细中文分析",
  "risk_level": "LOW"|"MEDIUM"|"HIGH",
  "market_regime": "TRENDING"|"MEAN_REVERTING"|"VOLATILE"|"UNCLEAR",
  "key_factors": ["因素1","因素2","因素3"],
  "tv_signal_consistent": true|false|null
}}"""

    payload = {
        "model": "deepseek-reasoner",   # 带推理链；如想更快可用 "deepseek-chat"
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": user_prompt},
        ],
        "temperature": 0.1,
        "max_tokens": 2048,
        "stream": False,
    }

    try:
        r = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={
                "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
                "Content-Type": "application/json",
            },
            json=payload,
            timeout=120,
        )
        r.raise_for_status()
        text = r.json()["choices"][0]["message"]["content"].strip()

        # 提取 JSON
        if "```json" in text:
            json_str = text.split("```json")[1].split("```")[0].strip()
        elif text.startswith("{"):
            json_str = text
        else:
            start, end = text.find("{"), text.rfind("}") + 1
            json_str = text[start:end] if start >= 0 and end > start else "{}"

        return json.loads(json_str)

    except Exception as e:
        print(f"ERROR calling DeepSeek API: {e}")
        return {"decision": "HOLD", "confidence": 0, "reasoning": f"DeepSeek 调用失败: {e}", "risk_level": "HIGH"}


def _pick_ai_provider() -> str:
    """自动选择 AI 提供商：优先使用 AI_PROVIDER 环境变量，其次看哪个 key 存在"""
    if AI_PROVIDER in ("anthropic", "deepseek"):
        return AI_PROVIDER
    if DEEPSEEK_API_KEY:
        return "deepseek"
    if ANTHROPIC_API_KEY:
        return "anthropic"
    return "anthropic"  # 默认，后续会因 key 缺失报错


def analyze(market_data: dict, tv_signal: dict | None = None) -> dict:
    """统一入口：根据配置调用 Claude 或 DeepSeek"""
    provider = _pick_ai_provider()
    print(f"[AI] 使用提供商: {provider.upper()}")
    if provider == "deepseek":
        return analyze_with_deepseek(market_data, tv_signal)
    return analyze_with_claude(market_data, tv_signal)


# ── Alpaca 交易执行 ───────────────────────────────────────────────────────────
def alpaca_headers() -> dict:
    return {
        "APCA-API-KEY-ID": ALPACA_API_KEY,
        "APCA-API-SECRET-KEY": ALPACA_SECRET_KEY,
        "Content-Type": "application/json",
    }


def get_account() -> dict:
    try:
        r = requests.get(f"{ALPACA_BASE}/v2/account", headers=alpaca_headers(), timeout=10)
        return r.json()
    except Exception as e:
        return {"error": str(e)}


def get_positions() -> list:
    try:
        r = requests.get(f"{ALPACA_BASE}/v2/positions", headers=alpaca_headers(), timeout=10)
        data = r.json()
        return data if isinstance(data, list) else []
    except Exception:
        return []


def place_order(symbol: str, notional: float, side: str) -> dict:
    """按美元金额下市价单"""
    payload = {
        "symbol": symbol,
        "notional": f"{notional:.2f}",
        "side": side,
        "type": "market",
        "time_in_force": "day",
    }
    try:
        r = requests.post(
            f"{ALPACA_BASE}/v2/orders",
            headers=alpaca_headers(),
            json=payload,
            timeout=15,
        )
        return r.json()
    except Exception as e:
        return {"error": str(e)}


def close_position(symbol: str) -> dict:
    try:
        r = requests.delete(
            f"{ALPACA_BASE}/v2/positions/{symbol}",
            headers=alpaca_headers(),
            timeout=15,
        )
        return r.json() if r.text else {"status": "closed"}
    except Exception as e:
        return {"error": str(e)}


def execute_trade(decision: dict, account: dict) -> list[dict]:
    """根据 AI 决策执行交易"""
    results = []
    if not AUTO_TRADE:
        return results

    d = decision.get("decision", "HOLD")
    confidence = decision.get("confidence", 0)
    if confidence < 65:
        print(f"[交易] 置信度 {confidence}% 低于阈值 65%，跳过执行")
        return results

    try:
        equity = float(account.get("equity", 0))
    except (ValueError, TypeError):
        print("[交易] 无法获取账户净值，跳过")
        return results

    pos_pct = min(decision.get("position_size_pct", POSITION_PCT), 0.10)
    trade_amount = equity * pos_pct

    if d == "PAIRS_LONG_A_SHORT_B":
        results.append(place_order(SYMBOL_A, trade_amount, "buy"))
        results.append(place_order(SYMBOL_B, trade_amount, "sell"))
    elif d == "PAIRS_LONG_B_SHORT_A":
        results.append(place_order(SYMBOL_A, trade_amount, "sell"))
        results.append(place_order(SYMBOL_B, trade_amount, "buy"))
    elif d == "CLOSE_POSITION":
        for sym in (SYMBOL_A, SYMBOL_B):
            results.append(close_position(sym))

    return results


# ── PushPlus 微信推送 ─────────────────────────────────────────────────────────
def push_wechat(title: str, content: str) -> bool:
    if not PUSHPLUS_TOKEN:
        print("[推送] PUSHPLUS_TOKEN 未配置，跳过微信推送")
        return False
    try:
        r = requests.post(
            "http://www.pushplus.plus/send",
            json={
                "token": PUSHPLUS_TOKEN,
                "title": title,
                "content": content,
                "template": "html",
            },
            timeout=15,
        )
        data = r.json()
        ok = data.get("code") == 200
        if not ok:
            print(f"[推送] PushPlus 返回: {data}")
        return ok
    except Exception as e:
        print(f"[推送] 推送失败: {e}")
        return False


def build_alert_html(market_data: dict, decision: dict, trade_results: list,
                     tv_signal: dict | None) -> str:
    """生成结构化的微信推送 HTML"""
    sp = market_data["spread"]
    md = market_data
    d = decision

    # 决策颜色编码
    COLOR_MAP = {
        "PAIRS_LONG_A_SHORT_B": "#e74c3c",
        "PAIRS_LONG_B_SHORT_A": "#27ae60",
        "CLOSE_POSITION": "#f39c12",
        "HOLD": "#95a5a6",
    }
    decision_text_map = {
        "PAIRS_LONG_A_SHORT_B": f"🔴 做多 {SYMBOL_A} / 做空 {SYMBOL_B}",
        "PAIRS_LONG_B_SHORT_A": f"🟢 做多 {SYMBOL_B} / 做空 {SYMBOL_A}",
        "CLOSE_POSITION": "🟡 平仓",
        "HOLD": "⚪ 观望",
    }

    dec = d.get("decision", "HOLD")
    color = COLOR_MAP.get(dec, "#95a5a6")
    dec_text = decision_text_map.get(dec, dec)
    confidence = d.get("confidence", 0)
    risk = d.get("risk_level", "N/A")
    regime = d.get("market_regime", "N/A")

    trade_html = ""
    if trade_results:
        trade_rows = "".join(
            f"<tr><td>{r.get('symbol','?')}</td><td>{r.get('side','?')}</td>"
            f"<td>{r.get('status','?')}</td><td>{r.get('id','?')[:8] if r.get('id') else 'ERR'}</td></tr>"
            for r in trade_results
        )
        trade_html = f"""
<h3>📋 成交记录</h3>
<table border="1" cellpadding="4" style="border-collapse:collapse;width:100%">
  <tr style="background:#34495e;color:white"><th>标的</th><th>方向</th><th>状态</th><th>订单ID</th></tr>
  {trade_rows}
</table>"""

    tv_html = ""
    if tv_signal:
        tv_html = f"""
<h3>📡 TradingView 信号</h3>
<p>标的: <b>{tv_signal.get('symbol','?')}</b> | 操作: <b>{tv_signal.get('action','?')}</b> |
价格: <b>${tv_signal.get('price',0):.2f}</b> | 策略: {tv_signal.get('strategy','?')}</p>
<p>AI 与 TV 信号一致: <b>{'✅ 是' if d.get('tv_signal_consistent') else ('❌ 否' if d.get('tv_signal_consistent') is False else '— N/A')}</b></p>"""

    key_factors_html = "".join(
        f"<li>{f}</li>" for f in d.get("key_factors", [])
    )

    return f"""
<div style="font-family:Arial,sans-serif;max-width:600px">

<h2 style="background:{color};color:white;padding:10px;border-radius:5px">
  {dec_text}
</h2>

<table border="1" cellpadding="6" style="border-collapse:collapse;width:100%">
  <tr><td><b>置信度</b></td><td><b style="color:{color}">{confidence}%</b></td>
      <td><b>风险等级</b></td><td>{risk}</td></tr>
  <tr><td><b>市场状态</b></td><td>{regime}</td>
      <td><b>时间</b></td><td>{datetime.now(timezone.utc).strftime('%m-%d %H:%M UTC')}</td></tr>
</table>

<h3>📊 套利价差</h3>
<table border="1" cellpadding="6" style="border-collapse:collapse;width:100%">
  <tr style="background:#2c3e50;color:white">
    <th>Z-Score 窗口</th><th>数值</th><th>信号</th></tr>
  <tr><td>30日</td><td><b>{sp['zscore_30d']:+.3f}</b></td>
      <td>{'🚨 入场' if abs(sp['zscore_30d'])>=ENTRY_ZSCORE else ('✅ 平仓' if abs(sp['zscore_30d'])<=EXIT_ZSCORE else '⏳ 等待')}</td></tr>
  <tr><td>20日</td><td>{sp['zscore_20d']:+.3f}</td><td>—</td></tr>
  <tr><td>10日</td><td>{sp['zscore_10d']:+.3f}</td><td>—</td></tr>
  <tr><td>趋势</td><td colspan="2">{sp['trend']}</td></tr>
</table>

<h3>📈 价格与技术指标</h3>
<table border="1" cellpadding="6" style="border-collapse:collapse;width:100%">
  <tr style="background:#2c3e50;color:white">
    <th>指标</th><th>{SYMBOL_A}</th><th>{SYMBOL_B}</th></tr>
  <tr><td>当前价格</td><td><b>${md['mstr']['price']:.2f}</b></td>
      <td><b>${md['bmnr']['price']:.2f}</b></td></tr>
  <tr><td>24h 变化</td><td>{md['mstr']['change_24h']:+.2f}%</td>
      <td>{md['bmnr']['change_24h']:+.2f}%</td></tr>
  <tr><td>RSI(14)</td><td>{md['mstr']['rsi']:.1f}</td>
      <td>{md['bmnr']['rsi']:.1f}</td></tr>
  <tr><td>MACD Hist</td><td>{md['mstr']['macd_hist']:.4f}</td>
      <td>{md['bmnr']['macd_hist']:.4f}</td></tr>
  <tr><td>Bollinger %B</td><td>{md['mstr']['bb_pct']:.3f}</td>
      <td>{md['bmnr']['bb_pct']:.3f}</td></tr>
  <tr><td>成交量比率</td><td>{md['mstr']['volume_ratio']:.2f}x</td>
      <td>{md['bmnr']['volume_ratio']:.2f}x</td></tr>
</table>

<h3>🧠 AI 分析</h3>
<p style="background:#f8f9fa;padding:10px;border-left:4px solid {color}">
  {d.get('reasoning','N/A')}
</p>

{'<h3>🎯 关键因素</h3><ul>' + key_factors_html + '</ul>' if key_factors_html else ''}

{tv_html}
{trade_html}

<p style="color:#95a5a6;font-size:12px">
  ⚠️ 本系统仅供参考，不构成投资建议。交易有风险，请结合自身判断。<br>
  模式: {"纸交易" if PAPER_TRADING else "实盘"} | 自动执行: {"开启" if AUTO_TRADE else "关闭"}
</p>
</div>"""


# ── 市场时间检查 ──────────────────────────────────────────────────────────────
def is_market_open() -> bool:
    """判断当前是否在美股交易时段（粗略检查，不含假日）"""
    now_et = datetime.now(timezone.utc) - timedelta(hours=4)  # UTC-4 (EDT)
    if now_et.weekday() >= 5:  # 周六日
        return False
    market_open = now_et.replace(hour=9, minute=30, second=0, microsecond=0)
    market_close = now_et.replace(hour=16, minute=0, second=0, microsecond=0)
    return market_open <= now_et <= market_close


# ── 主流程 ────────────────────────────────────────────────────────────────────
def main():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 套利 Agent 启动")
    print(f"  交易对: {SYMBOL_A} / {SYMBOL_B}")
    print(f"  模式: {'纸交易' if PAPER_TRADING else '实盘'} | 自动执行: {'开启' if AUTO_TRADE else '关闭'}")

    # ── 1. 解析 TradingView 信号 ──────────────────────────────────────────────
    tv_signal = None
    if TV_PAYLOAD:
        try:
            tv_signal = json.loads(TV_PAYLOAD)
            print(f"[TV信号] 接收到: {tv_signal}")
        except json.JSONDecodeError:
            print(f"[TV信号] 解析失败: {TV_PAYLOAD}")

    # ── 2. 检查市场是否开盘 ───────────────────────────────────────────────────
    if not is_market_open() and not tv_signal:
        print("[市场] 当前不在交易时段，跳过（TradingView 信号例外）")
        return

    # ── 3. 抓取行情数据 ───────────────────────────────────────────────────────
    print("[数据] 抓取行情数据...")
    market_data = fetch_market_data()
    if not market_data:
        push_wechat(
            f"⚠️ {SYMBOL_A}/{SYMBOL_B} 数据抓取失败",
            "<p>行情数据获取失败，请检查网络或 yfinance 服务。</p>",
        )
        sys.exit(1)

    sp = market_data["spread"]
    print(f"[数据] {SYMBOL_A}=${market_data['mstr']['price']:.2f} | "
          f"{SYMBOL_B}=${market_data['bmnr']['price']:.2f} | "
          f"Z-Score(30d)={sp['zscore_30d']:+.3f}")

    # ── 4. 快速判断：是否需要 AI 分析 ─────────────────────────────────────────
    needs_analysis = (
        FORCE_ANALYSIS                                  # 手动强制触发
        or tv_signal is not None                        # TradingView 有信号
        or abs(sp["zscore_30d"]) >= ENTRY_ZSCORE * 0.8 # Z-Score 接近入场阈值
    )

    if not needs_analysis:
        print(f"[跳过] Z-Score={sp['zscore_30d']:+.3f} 未达分析阈值，发送日常简报")
        _send_daily_summary(market_data)
        return

    # ── 5. AI 分析（Claude 或 DeepSeek）─────────────────────────────────────
    provider = _pick_ai_provider()
    if provider == "deepseek" and not DEEPSEEK_API_KEY:
        print("ERROR: AI_PROVIDER=deepseek 但 DEEPSEEK_API_KEY 未配置")
        sys.exit(1)
    if provider == "anthropic" and not ANTHROPIC_API_KEY:
        print("ERROR: ANTHROPIC_API_KEY 未配置")
        sys.exit(1)

    print(f"[AI] 调用 {provider.upper()} 分析中...")
    decision = analyze(market_data, tv_signal)
    print(f"[AI] 决策: {decision.get('decision','?')} | 置信度: {decision.get('confidence',0)}%")
    print(f"[AI] 分析: {decision.get('reasoning','')[:100]}...")

    # ── 6. 执行交易 ───────────────────────────────────────────────────────────
    trade_results = []
    account = {}
    if AUTO_TRADE and ALPACA_API_KEY:
        account = get_account()
        if "error" not in account:
            trade_results = execute_trade(decision, account)
            if trade_results:
                print(f"[交易] 已执行 {len(trade_results)} 笔订单")
    elif AUTO_TRADE and not ALPACA_API_KEY:
        print("[交易] AUTO_TRADE=1 但 ALPACA_API_KEY 未配置，跳过执行")

    # ── 7. 微信推送 ───────────────────────────────────────────────────────────
    dec = decision.get("decision", "HOLD")
    confidence = decision.get("confidence", 0)

    title_map = {
        "PAIRS_LONG_A_SHORT_B": f"🔴 套利信号: 做多{SYMBOL_A}/做空{SYMBOL_B} ({confidence}%)",
        "PAIRS_LONG_B_SHORT_A": f"🟢 套利信号: 做多{SYMBOL_B}/做空{SYMBOL_A} ({confidence}%)",
        "CLOSE_POSITION": f"🟡 套利平仓信号 ({confidence}%)",
        "HOLD": f"⚪ {SYMBOL_A}/{SYMBOL_B} 套利监控 — 观望",
    }
    title = title_map.get(dec, f"{SYMBOL_A}/{SYMBOL_B} 套利分析")
    content = build_alert_html(market_data, decision, trade_results, tv_signal)

    pushed = push_wechat(title, content)
    print(f"[推送] 微信推送{'成功' if pushed else '失败'}: {title}")

    # 非 HOLD 决策且置信度高时，在控制台也输出完整分析
    if dec != "HOLD" and confidence >= 65:
        print("\n" + "=" * 60)
        print(f"🚀 交易信号触发！")
        print(f"   决策: {dec}")
        print(f"   置信度: {confidence}%")
        print(f"   风险: {decision.get('risk_level','?')}")
        print(f"   分析: {decision.get('reasoning','')}")
        print("=" * 60 + "\n")


def _send_daily_summary(market_data: dict):
    """发送日常监控简报（无需 AI 分析）"""
    sp = market_data["spread"]
    md = market_data
    content = f"""
<h3>📊 {SYMBOL_A}/{SYMBOL_B} 套利监控简报</h3>
<p>{datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}</p>
<table border="1" cellpadding="6" style="border-collapse:collapse">
  <tr><th>{SYMBOL_A} 价格</th><td>${md['mstr']['price']:.2f} ({md['mstr']['change_24h']:+.2f}%)</td></tr>
  <tr><th>{SYMBOL_B} 价格</th><td>${md['bmnr']['price']:.2f} ({md['bmnr']['change_24h']:+.2f}%)</td></tr>
  <tr><th>Z-Score (30日)</th><td>{sp['zscore_30d']:+.3f}</td></tr>
  <tr><th>Z-Score (20日)</th><td>{sp['zscore_20d']:+.3f}</td></tr>
  <tr><th>价差趋势</th><td>{sp['trend']}</td></tr>
  <tr><th>{SYMBOL_A} RSI</th><td>{md['mstr']['rsi']:.1f}</td></tr>
  <tr><th>{SYMBOL_B} RSI</th><td>{md['bmnr']['rsi']:.1f}</td></tr>
</table>
<p style="color:#95a5a6">入场阈值: |Z| ≥ {ENTRY_ZSCORE} | 平仓阈值: |Z| ≤ {EXIT_ZSCORE}</p>
"""
    push_wechat(f"📊 {SYMBOL_A}/{SYMBOL_B} 套利监控 — Z:{sp['zscore_30d']:+.2f}", content)


if __name__ == "__main__":
    main()
