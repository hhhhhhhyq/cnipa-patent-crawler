# CNIPA 专利检索爬虫

复用你**已登录**的 Chrome 会话（CDP），按 Excel 里的公司名批量检索国家知识产权局
[常规检索](https://pss-system.cponline.cnipa.gov.cn/conventionalSearch)，按申请日时间范围抓全，汇总成一个 Excel。

> 已在华为等真实公司上端到端跑通：自动识别检索→服务端按申请日精确过滤→自动翻页抓全→字段清洗入表。

## 工作原理（为什么这么做）
- 系统**必须登录**、登录有滑块验证码 → 复用你手动登录的浏览器（登录令牌在 localStorage，同源跨标签共享）。
- 检索接口带 **WAF 动态令牌**且检索式 AES 加密 → 不能直接构造/重放请求，只能**驱动 UI** 让页面自身的 XHR 去请求；脚本用 CDP 抓取这些 XHR 的 JSON 响应。
- 结果页的**日期范围是明文参数**：脚本在结果页填“申请日”起止日期并点“确定”，即**服务端精确过滤**，再自动翻页（每页 40 条）抓完全部。

## 一、安装依赖
```bash
npm install -g agent-browser
python3 -m pip install pandas openpyxl   # 已装可跳过
```

## 二、启动 debug Chrome 并登录
```bash
node ~/.claude/skills/browser-cdp/scripts/setup-cdp-chrome.js 9222 --yes
```
在弹出的 Chrome 里打开并**登录** https://pss-system.cponline.cnipa.gov.cn/ （含滑块验证码）。
登录一次即可，之后脚本自动复用。

> 上面那行是本机 Claude Code 环境专用的便捷脚本。**换到其他电脑**（尤其没装 Claude Code 的）时，
> 改用下面的手动方式启动调试版 Chrome，效果完全一样。见「在其他电脑上运行」。

## 三、运行
输入 Excel：**第一列为公司名**（带表头）。
```bash
cd /Users/gaotu/cnipa_patent_crawler

# 按申请日时间范围抓全（推荐：范围越具体越快）
python3 cnipa_crawler.py run --input 公司名.xlsx --output 专利汇总.xlsx \
        --start 2020-01-01 --end 2023-12-31

# 不限时间（抓每家公司全部专利；大公司会很多很慢，慎用）
python3 cnipa_crawler.py run --input 公司名.xlsx --output 专利汇总.xlsx

# 额外抓“被引证次数”（逐条点击，慢很多，慎用）
python3 cnipa_crawler.py run --input 公司名.xlsx --output 专利汇总.xlsx \
        --start 2015-01-01 --end 2015-12-31 --citations
```
- `--start` / `--end`：申请日范围（服务端过滤）。写 `2020-01-01` 或 `20200101` 都行，可只给其一。
- 日期格式自动归一；时间范围由 CNIPA 服务端按**申请日**精确过滤，脚本再做一次客户端兜底过滤。
- `--citations`：额外抓每条专利的**被引证次数**。它不在列表接口里，需对**每条专利**单独点“被引证”再请求，
  **速度慢很多、请求量大**（每条 ~1.5 秒、易触发风控）。不加此参数时该列留空。

## 输出
- `专利汇总.xlsx` —— 每行一条专利，列：
  检索公司、申请号、专利名称、申请日、公开(公告)号、公开(公告)日、
  公开日期(公开)、公开日期(授权)、专利类型、
  申请人、发明人、代理机构、申请人所在国/地区、主IPC分类号、IPC分类号、被引证次数、摘要。
  （`公开日期(授权)` 来自著录项目 SQ_PD，仅已授权专利有值；未授权则为空。
  `被引证次数` 仅在加 `--citations` 时填充，否则留空。）
- `专利汇总.xlsx.raw.jsonl` —— 原始 JSON 备份（每行一条，含全部原始字段）。

- `专利汇总.xlsx.progress.txt` —— 断点续抓进度（已完成的公司名，每行一个）。

边爬边增量写盘，中途中断不丢已抓数据。

## 断点续抓
- 每抓完一家公司就记进 `*.progress.txt`。**中断后用同一命令重跑**会：自动**跳过已完成的公司**、从备份**恢复已抓数据**（不覆盖、不重复），只继续抓没抓过的。
- 适合配合限速“分多次慢慢跑”，或被风控/断网后接着跑。
- 输入 Excel 里**新增公司**后重跑，只会抓新增的那些。
- **想从头重抓**：删掉 `专利汇总.xlsx.progress.txt`（及 `专利汇总.xlsx` / `*.raw.jsonl`）即可。

## 在其他电脑上运行
流程、脚本、命令完全相同，只有「启动 debug Chrome」这一步改成手动方式。

**前提（每台机器一次性）**
1. 装 Node.js（含 npm）
2. 装 Google Chrome
3. `npm install -g agent-browser`
4. `pip install pandas openpyxl`（Python 3）
5. 拷贝 `cnipa_crawler.py` 过去

**手动启动调试版 Chrome（替代第二步那行脚本）**
先**彻底退出已开的 Chrome**，再按系统执行其一：

```bash
# macOS
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
  --remote-debugging-port=9222 --user-data-dir="$HOME/chrome-debug-profile"

# Linux
google-chrome --remote-debugging-port=9222 --user-data-dir="$HOME/chrome-debug-profile"
```
```powershell
# Windows (PowerShell)
& "C:\Program Files\Google\Chrome\Application\chrome.exe" `
  --remote-debugging-port=9222 --user-data-dir="$env:USERPROFILE\chrome-debug-profile"
```
然后在这个 Chrome 里**登录 CNIPA**（滑块验证码），再照「三、运行」执行即可。

**注意**
- 脚本会自动判断：找不到 Claude 的便捷脚本时**直接连 9222 端口**，所以手动起好 Chrome 后脚本无需改动。
- 端口必须是 **9222**（要改的话改脚本顶部 `CDP_PORT`，并相应改启动命令的端口）。
- 登录态按「机器 + 调试 profile」独立保存：**每台新机器都要重新登录一次** CNIPA。

## 防风控措施
- **复用真实登录浏览器**：所有请求走你真实 Chrome 的 XHR（带正确 WAF 令牌/指纹/会话），是绕过风控的根本。
- **随机等待**：所有停顿都是随机区间（见脚本顶部 `WAIT_*` / `DELAY_COMPANIES`），无固定节奏。想更保守可把区间整体调大。
- **串行不并发**：一次一家公司、一页一页翻。
- **风控/异常检测 + 自动退避重试**：检测到①登录态失效（token 丢失/跳登录页）②限流提示弹窗（频繁/繁忙/人机验证/稍后等）时，自动**指数退避**（30s→…→最多 10 分钟）并重新检查，恢复后继续；
  - 登录失效时会提示你**在 Chrome 重新登录**，并留足等待时间自动恢复；
  - 翻页中途异常会区分“真的没下一页”还是“被风控”，避免**静默少抓**；
  - 连续无法恢复则**安全停止**（已抓数据已增量落盘，恢复后重跑即可）。
- 调参：脚本顶部 `BASE_BACKOFF` / `MAX_BACKOFF` / `RECOVER_TRIES` 控制退避节奏与次数。

## 说明与注意
- **运行时不要手动操作那个 Chrome 标签**（脚本在驱动它）。脚本会把 CNIPA 标签自动置于前台。
- 公司名用“自动识别”检索（系统按申请人匹配）。同名/简称/合并申请人会一并出现在结果里，可按“申请人”列再筛。
- **法律状态**未抓取：它不在检索结果接口里，需逐条点开详情页单独请求（很慢、易触发风控）。如需要可再加。
- 仅抓取公开可检索数据，请遵守网站使用条款、控制频率、自用为主。
- 登录态过期：重跑第二步（可加 `--reset --yes` 清缓存重登）。
- 速度参考：受每页等待与翻页数影响，范围越窄越快；大公司宽范围会较慢（每 40 条约 4 秒一页）。
