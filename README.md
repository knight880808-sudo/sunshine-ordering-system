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

## 订单状态流程 / Order status flow

```
Pending (待发货) → Dispatched (已发货) → Received (已收货)
```

## 缺货状态流程 / Shortage status flow

```
Open (待处理) → Resending (补发中) or Out of Stock (缺货) → Resolved (已解决)
```
