# SUNSHINE SHOPPING CENTER 订货系统

阳光集团 7 家分店订货系统 / Ordering system for 7 branch stores.

## 安装 / Installation

```bash
pip install streamlit pandas openpyxl
```

## 准备商品文件 / Prepare products file

将商品主数据 `products.xlsx` 放在 `app.py` 同目录下，列名必须为：

Place `products.xlsx` next to `app.py` with these columns:

| ItemCode | Barcode | Name | Unit | Price |
|----------|---------|------|------|-------|

> 如果文件不存在，系统会自动创建一份 5 行的示例文件。
> If the file is missing, a 5-row demo file will be auto-generated on first launch.

## 启动 / Run

```bash
streamlit run app.py --server.port 8502 --server.address 0.0.0.0
```

可选：通过环境变量覆盖默认密码（推荐生产使用）  
Optional: override default passwords via environment variables (recommended for production)

```bash
# Windows PowerShell
$env:SUNSHINE_WAREHOUSE_PASSWORD="your-warehouse-password"
$env:SUNSHINE_ADMIN_PASSWORD="your-admin-password"
$env:GEMINI_API_KEY="your-gemini-api-key"
$env:GEMINI_MODEL="gemini-2.0-flash"
streamlit run app.py --server.port 8502 --server.address 0.0.0.0
```

通过 Tailscale 内网，分店访问 `http://<主机Tailscale IP>:8502` 即可。
Branches access via Tailscale at `http://<host-tailscale-ip>:8502`.

## 登录 / Login

| 角色 Role | 密码 Password |
|-----------|---------------|
| 🛒 分店员工 Branch staff | (无 / none — just pick the branch) |
| 📦 仓库员工 Warehouse | `sunshine888` |
| 🏭 管理员 Admin | `sunshine` |

> 若设置了环境变量 `SUNSHINE_WAREHOUSE_PASSWORD` / `SUNSHINE_ADMIN_PASSWORD`，将优先使用环境变量值。  
> If `SUNSHINE_WAREHOUSE_PASSWORD` / `SUNSHINE_ADMIN_PASSWORD` are set, they take precedence.
>
> 分店登录为“按钮直选”模式（非下拉框）：点击店铺按钮即可进入。最近使用店铺会有 ⭐ 标记。  
> Branch login uses direct store buttons (not a dropdown): click a store to enter. Most recently used store is marked with ⭐.

## 数据存储 / Data

- SQLite database file: `orders.db` (auto-created on first launch in working directory)
- 表 / Tables: `orders`, `shipments`, `shortages`
- Runtime log file: `app_runtime.log` (key non-blocking exceptions for diagnosis)

## 主要功能 / Key features

- **双语界面** Bilingual UI (English / 中文) — switch any time from the EN/中文 buttons in the top right.
- **扫码枪支持** Bluetooth scanner — search and receive pages accept scanner input directly (no camera needed, works over plain HTTP).
- **手机优化** Mobile-optimized layout.
- **箱/个分别输入** Cartons & Each pcs entered separately, never converted.
- **缺货闭环** Full shortage workflow: branch reports → warehouse resends or marks out-of-stock → branch confirms.
- **三种 Excel 报表** Three Excel exports with styled headers:
  - Picking list — one sheet per branch
  - Reconciliation — dispatched vs actual, SHORT rows red, OVER rows yellow
  - Shortage report
- **管理员仪表盘** Admin dashboard — 4 metric cards + branch status table + latest orders & shortages.
- **缺货红色角标** Red badge on title bar when there are open shortages.
- **防重复提交** Duplicate-submit protection on order submit / dispatch submit.
- **登录记忆** Login remembers last role and last branch for faster re-login.
- **待发货高效筛选** Pending-dispatch page supports:
  - keyword filter (order id / item name / barcode)
  - branch multi-select + one-click "All branches"
  - date range + quick actions ("Recent 7 days" / "All dates")
  - expand/collapse controls ("Expand all" / "Collapse all" / match-only expand)
  - filter summary strip (branches/date/keyword at a glance)
- **分店下单搜索分页** Branch product search shows 5 products per page.
- **下单完成页** After branch submits an order, app navigates to a dedicated
  "order sent" page with:
  - Continue creating another order
  - View my orders
- **发货效率增强** Dispatch form supports:
  - one-click "Fill with ordered qty"
  - one-click "Clear all dispatch qty"
  - row-level changed marker when dispatch qty differs from ordered qty
- **Gemini AI助手** AI assistant for all roles (Branch/Warehouse/Admin):
  - asks operational questions and troubleshooting guidance
  - optional context payload (current role/branch, pending orders, open shortages, recent orders)
  - rate-limited and read-only guidance (no direct write actions)

## AI 助手配置 / AI Assistant Setup

- Required: `GEMINI_API_KEY` (either set as environment variable **or** put it in `gemini.env` next to `app.py`)
- Optional: `GEMINI_MODEL` in env or `gemini.env` (default: `gemini-2.0-flash`). If you still have `gemini-1.5-*` in a **Windows system** environment variable, the app automatically maps it to `gemini-2.0-flash` (1.5 often returns 404 on current API).
- `gemini.env` format (one line per variable):

```text
GEMINI_API_KEY=your-key-here
GEMINI_MODEL=gemini-2.0-flash
# Optional: if primary hits 429 quota, try these in order (comma-separated)
# GEMINI_FALLBACK_MODELS=gemini-2.5-flash,gemini-2.0-flash-lite
```

- `gemini.env` is listed in `.gitignore` — do not commit API keys.
- AI入口: side navigation `🤖 AI Assistant` for all roles
- Safety notes:
  - assistant outputs are advisory only
  - validate critical operations against actual system data
  - API key is never written to the database or email logs
- If you see **HTTP 429 / RESOURCE_EXHAUSTED / quota limit: 0**, your Google project free tier may not allow that model, or daily quota is exhausted. Enable billing or wait, then retry. The app can try **fallback models** via `GEMINI_FALLBACK_MODELS`.

## Railway 部署 / Deploy on Railway

[Railway](https://railway.com/) 不会上传本机被 `.gitignore` 忽略的 `gemini.env`，也不会自动带上你只在本地保存的 `email_config.json`。必须在 **Project → Variables** 里配置，然后 **Redeploy**。

| Variable | 说明 |
|----------|------|
| `GEMINI_API_KEY` | **必填**（AI）。在 [Google AI Studio](https://aistudio.google.com/apikey) 新建密钥；若日志出现 `leaked` / 403，旧密钥已作废，必须换新。 |
| `GEMINI_MODEL` | 可选，默认 `gemini-2.0-flash` |
| `SUNSHINE_EMAIL_ENABLED` | `true` 开启邮件 |
| `SUNSHINE_SMTP_HOST` | 如 `smtp.gmail.com` |
| `SUNSHINE_SMTP_PORT` | 如 `587` |
| `SUNSHINE_SMTP_USER` | SMTP 登录名 |
| `SUNSHINE_SMTP_PASSWORD` | Gmail 请用[应用专用密码](https://myaccount.google.com/apppasswords) |
| `SUNSHINE_SMTP_FROM` | 发件人地址 |
| `SUNSHINE_SMTP_USE_TLS` | `true`（587 端口） |
| `SUNSHINE_EMAIL_NOTIFY_TO` | 逗号分隔收件人，会用于全部通知类型 |
| `SUNSHINE_WAREHOUSE_PASSWORD` | 建议生产修改仓库密码 |
| `SUNSHINE_ADMIN_PASSWORD` | 建议生产修改管理员密码 |

仓库根目录的 `railway.toml` 已配置 Streamlit 监听 `$PORT`。部署后：管理员 → **邮件通知** → **发送测试邮件** 验证 SMTP；侧边栏 **AI Assistant** 试一条问题。

可选：为 `orders.db` / `email_config.json` 挂载 [Railway Volume](https://docs.railway.com/guides/volumes)，避免每次重新部署丢失数据与在界面里保存的邮件配置。

### Railway 上邮件发不出去？（AI 正常、邮件失败）

[Railway 官方说明](https://docs.railway.com/reference/outbound-networking)：**Free / Trial / Hobby 套餐会封锁出站 SMTP**（Gmail 的 587/465 端口连不上），与你在 Variables 里填的账号密码无关。可选方案：

| 方案 | 做法 |
|------|------|
| **A. Resend（推荐，Hobby 可用）** | 注册 [Resend](https://resend.com) → 验证发信域名 → API Keys 创建密钥 → Railway 增加 `RESEND_API_KEY`、`RESEND_FROM`（如 `SUNSHINE <orders@你的域名.com>`）→ Redeploy |
| **B. 升级 Pro** | 升级到 Pro 后对服务 **Redeploy**，再继续用 Gmail SMTP 环境变量 |

设置 `RESEND_API_KEY` 后，系统优先走 HTTPS 发信，不再连接 `smtp.gmail.com`。

## 订单状态流程 / Order status flow

```
Pending (待发货) → Dispatched (已发货) → Received (已收货)
```

## 缺货状态流程 / Shortage status flow

```
Open (待处理) → Resending (补发中) or Out of Stock (缺货) → Resolved (已解决)
```

## 分店库存临期/过期预警 / Branch expiry alerts

为每个分店独立管理**批次库存 + 过期日期**，并每天定时扫描临期/过期商品，
按分店通知对应店长（消息中心 + 邮件）。

### 数据表 / Table

新增 `branch_inventory_batches`（随 `app.py` 启动自动建表，幂等安全）：

| 字段 | 说明 |
|------|------|
| `branch` | 分店 ID（属于 7 家分店之一），各分店库存独立 |
| `name` / `item_code` / `barcode` / `unit` | 商品信息 |
| `batch_no` | 批次号（可选，便于追溯） |
| `qty_cartons` / `qty_pcs` | 该批次剩余库存（箱 / 个） |
| `production_date` | 生产/入库日期（可选） |
| `shelf_life_days` | 保质期天数（可选） |
| `expire_date` | **过期日期（核心，必填）** |
| `status` | active / depleted / discarded |

> 店长联系方式复用 `email_config.json → branch_emails[分店]`。

### 入库 / Add a batch (business logic)

`app.py` 内 `add_branch_batch(...)`：入库/调拨时必须能确定过期日期
（直接给 `expire_date`，或给 `production_date` + `shelf_life_days` 自动推算）。

```python
import app
app.add_branch_batch(
    branch="SUNSHINE MARKET", name="鲜奶 1L", item_code="M001",
    qty_pcs=20, expire_date="2026-06-10",
)
```

### 预警天数 / Warning window

默认 **15 天**；可用环境变量覆盖：

```bash
SUNSHINE_EXPIRY_WARN_DAYS=15
```

判定逻辑：`当前日期 + 预警天数 >= 过期日期` 且 `库存 > 0` 且 `status='active'`。

### 定时任务 / Cron job（每天凌晨 1 点）

运行入口：

```bash
python expiry_alert_job.py            # 正式运行
python expiry_alert_job.py --no-email # 演练（只写消息中心，不发邮件）
python expiry_alert_job.py --days 7   # 临时指定预警天数
```

脚本自带「每日去重」，同一天重复触发只执行一次（`--force` 可强制重跑）。

#### A. Railway（线上，推荐）

在 Railway 项目内新增一个 **Cron 服务**（与主 Web 服务共用同一仓库 / Volume）：

- **Schedule**：`0 1 * * *`（每天 01:00）
- **Start Command**：`python expiry_alert_job.py`
- 若挂载了数据 Volume，设 `SUNSHINE_DATA_DIR=/data`，与主服务指向同一个 `orders.db`。

> 注意时区：Railway 默认 **UTC**。要按当地凌晨 1 点运行，请把 cron 换算为 UTC，
> 或设置容器时区环境变量 `TZ=Asia/Shanghai`（按你所在时区）。

#### B. Windows 计划任务（本地主机）

```bat
schtasks /Create /TN "SunshineExpiryAlert" /TR "python \"d:\sunshine 系统\sunshine-ordering-system-main\expiry_alert_job.py\"" /SC DAILY /ST 01:00
```

> 若 Python 不在全局 PATH，请把 `python` 换成完整路径（如 `C:\...\python.exe`）。

#### C. Linux/Mac crontab

```cron
0 1 * * * cd /path/to/app && /usr/bin/python expiry_alert_job.py >> expiry_job.log 2>&1
```
