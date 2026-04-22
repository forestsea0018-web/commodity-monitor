# 商品期货云端监控 - 微信推送

每 10 分钟自动扫描全球 RSS 源，命中关键词 → 推送到微信。

## 部署步骤（5 分钟）

### 1. 注册 PushPlus 拿 token
- 打开 https://www.pushplus.plus/
- 用微信扫码登录
- 首页顶部就能看到你的 token（32位字符串）

### 2. 建 GitHub 仓库
- 打开 https://github.com/new
- 仓库名随意（比如 `commodity-monitor`），可设为 **Private**
- 创建后，把本目录（`commodity_cloud/`）整个上传：
  - `cloud_monitor.py`
  - `.github/workflows/monitor.yml`
  - `README.md`

本地命令（装了 git 的话）：
```bash
cd C:\Users\lenovo\Desktop\commodity_cloud
git init
git add .
git commit -m "init"
git branch -M main
git remote add origin https://github.com/你的用户名/commodity-monitor.git
git push -u origin main
```

### 3. 配置 Secret
- 仓库页面 → Settings → Secrets and variables → Actions → New repository secret
- Name: `PUSHPLUS_TOKEN`
- Value: 粘贴第 1 步拿到的 token
- Add secret

### 4. 启用 Actions
- 仓库页面 → Actions 标签
- 如果提示 "I understand my workflows"，点 Enable
- 首次可手动触发：Actions → Commodity News Monitor → Run workflow

## 验证
- 手动触发后 1-2 分钟，查看 Actions 运行日志
- 成功的话微信会收到 "PushPlus" 公众号推送
- 之后每 10 分钟自动运行（GitHub cron 可能有 5-15 分钟延迟，属正常）

## 注意事项
- **PushPlus 免费版每天 200 条**。启用了 `DIGEST_MODE=1`（多条合并成一条），通常一天 20-50 条汇总消息，远未达上限
- **GitHub Actions 免费额度**：公开仓库无限；Private 仓库每月 2000 分钟。本任务每次 ~30 秒，一个月约 130 分钟，富余
- **时间窗口**：脚本只推送"最近 15 分钟内发布"的文章。窗口略大于 cron 间隔（10 分钟）以防漏
- **部分源没有发布时间** → 会保守放行，可能会重复推送，忽略即可

## 调整参数
在 `.github/workflows/monitor.yml` 的 env 区：
```yaml
WINDOW_MIN: "15"      # 时间窗口分钟数
DIGEST_MODE: "1"      # 1=合并发送, 0=每条单发
MAX_PUSH_PER_RUN: "20"  # 单次最多推送条数
```

## 停止监控
仓库 → Settings → Actions → Disable Actions
