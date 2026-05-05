# Supermarket Ordering System — Architecture & Design Reference

## 1. 系统定位

多分店（7 家固定）向中央仓库订货，仓库发货，分店收货，支持缺货闭环；管理员侧报表/库存/账号/备份；可选 POS 子系统共享同一商品库与 SQLite。

---

## 2. 分层架构

```
┌─────────────────────────────────────────────────────────────┐
│                         客户端层                             │
│  Streamlit app.py (主系统UI)   │   pos-web (React/Vite POS) │
└────────────────┬───────────────┴──────────────┬─────────────┘
                 │  HTTP / REST                  │ /api (proxy)
┌────────────────▼───────────────────────────────▼─────────────┐
│                         API 层                                │
│   api_v2.py  (FastAPI :5056)   │   pos_api.py (FastAPI :5055)│
│   所有订货/库存/账号业务        │   仅 POS 收银（兼容保留）   │
└────────────────────────────────┬─────────────────────────────┘
                                  │ SQLite / pandas
┌─────────────────────────────────▼─────────────────────────────┐
│                         数据层                                 │
│  orders.db (SQLite WAL)         products.xlsx                 │
│  db_schema.py (schema/migrate)  product_images/               │
│  backups/  email_config.json                                  │
└────────────────────────────────────────────────────────────────┘
```

> **重要约定**  
> - `app.py`（Streamlit）直连 `orders.db`，无需经过 API  
> - `api_v2.py` 与 `app.py` 共享同一个 `orders.db`，两者并行运行互不干扰  
> - `pos_api.py` 保留以兼容 `pos-web`；`api_v2.py` 同时提供 `/api/v2/pos/*` 端点，两者均写 `pos_orders` 表  

---

## 3. 角色与认证

| 角色 | 代号 | 认证方式 | Token TTL |
|------|------|----------|-----------|
| 分店员工 | `branch` | 按钮选店（无密码）或账号/密码 | 12 h |
| 仓库员工 | `warehouse` | 环境变量 `SUNSHINE_WAREHOUSE_PASSWORD` | 12 h |
| 管理员 | `admin` | 环境变量 `SUNSHINE_ADMIN_PASSWORD` | 12 h |

Token 格式：HMAC-SHA256 签名的 JSON，HEX 编码，无外部依赖。  
生产环境必须设置 `API_V2_SECRET_KEY` 环境变量。

---

## 4. 数据库 ER（实体关系）

### 表清单（对应需求第 4 节）

| 需求表名 | 实际表名 | 说明 |
|---------|---------|------|
| 订单行 | `orders` | 一行一订单明细；`order_group_id` 聚合同批提交 |
| 发货 | `orders.dispatch_*` 字段 | 发货信息内联于 orders 行，status=Dispatched |
| 收货 | `orders.receive_*` 字段 | 收货信息内联于 orders 行，status=Received |
| 缺货 | `shortages` | 关联 order_id，双向备注，四态状态机 |
| 消息 | `notifications` | 按 role/branch/account 路由，is_read 标记 |
| 进货通知 | `stock_arrivals` | 标题/正文/商品清单，is_active 控制展示 |
| 库存 | `inventory` | 按 (item_code, barcode) 唯一，权威数量 |
| 库存流水 | `inventory_txn` | IN/OUT/ADJUST/DISPATCH/RECEIVE/IMPORT |
| 价格覆盖 | `product_prices` | 运行时覆盖 Excel 档案价 |
| 分店账号 | `user_accounts` | pending/approved/rejected + permissions JSON |
| 审计日志 | `audit_log` | login/order_submit/receive_confirm 等 |
| 购物车草稿 | `branch_cart_draft` | 按 account_id+branch 唯一 |
| 供货商订单 | `supplier_orders` | 仓库→供货商，发消息+邮件 |
| POS 订单 | `pos_orders` | 独立收银订单头 |
| POS 订单行 | `pos_order_lines` | 独立收银行，FK→pos_orders |

### 核心关系图

```
user_accounts ──(account_id)──▶ orders
                                    │
                            ┌───────┴───────┐
                            ▼               ▼
                        shortages      inventory_txn
                            │
                        notifications (ref_shortage_id)

orders ──(ref_order_id)──▶ notifications
orders ──(ref_order_id)──▶ inventory_txn

product_prices ──(item_code/barcode)──▶ [overlay on products.xlsx]
inventory      ──(item_code/barcode)──▶ [overlay on products.xlsx]
```

---

## 5. 状态机

### 订单状态

```
Pending ──[仓库发货]──▶ Dispatched ──[分店收货]──▶ Received
```

- `Pending`：分店提交，等待仓库发货  
- `Dispatched`：仓库填写发货箱/个，写 `dispatch_*` 字段，扣减 inventory  
- `Received`：分店确认，填写实收箱/个；若有差量自动创建 shortage

### 缺货状态

```
Open ──[仓库处理]──▶ Resending ──[再次收货/关闭]──▶ Resolved
  │
  └──[确认无货]──▶ Out of Stock ──[知悉关闭]──▶ Resolved
```

---

## 6. API 端点清单

### 认证
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v2/branches` | 分店列表 |
| POST | `/api/v2/auth/login` | 登录，返回 token |

### 商品
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v2/products?q=&limit=` | 商品搜索 |
| GET | `/api/v2/product-image/{key}` | 商品图片 |
| POST | `/api/v2/product-image/{key}` | 上传商品图片 (admin/warehouse) |

### 订单
| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/v2/orders` | 提交订单（分店） |
| GET | `/api/v2/orders?branch=&status=&q=&date_from=&date_to=` | 订单列表 |
| GET | `/api/v2/orders/{id}` | 订单详情 |
| POST | `/api/v2/orders/dispatch` | 批量发货（仓库/管理员） |
| POST | `/api/v2/orders/receive` | 批量收货确认（分店） |

### 缺货
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v2/shortages?branch=&status=&q=` | 缺货列表 |
| PUT | `/api/v2/shortages/{id}` | 更新状态/备注（双方） |

### 库存
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v2/inventory?q=` | 当前库存列表 |
| POST | `/api/v2/inventory/adjust` | 手动调整（箱/个 delta） |
| POST | `/api/v2/inventory/import` | Excel 批量导入 (append/overwrite) |
| GET | `/api/v2/inventory/txn` | 库存流水 |
| GET | `/api/v2/templates/inventory` | 下载导入模板 |

### 价格
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v2/prices?q=` | 档案价+当前生效价对照 |
| POST | `/api/v2/prices` | 批量更新价格（admin） |
| POST | `/api/v2/prices/import` | Excel 批量导入价格（admin） |
| GET | `/api/v2/templates/price` | 下载价格导入模板 |

### 消息/通知
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v2/notifications?unread_only=` | 消息列表 |
| POST | `/api/v2/notifications/{id}/read` | 标记已读 |
| POST | `/api/v2/notifications/read-all` | 全部标为已读 |
| POST | `/api/v2/notifications` | 创建消息（admin/warehouse） |

### 进货通知
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v2/stock-arrivals?active_only=` | 到货公告 |
| POST | `/api/v2/stock-arrivals` | 发布到货通知 |
| PUT | `/api/v2/stock-arrivals/{id}` | 更新到货通知 |

### 供货商下单
| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/v2/supplier-orders` | 提交供货商订单 |
| GET | `/api/v2/supplier-orders` | 历史记录 |

### 账号管理
| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/v2/accounts/register` | 分店员工自助注册（→pending） |
| GET | `/api/v2/accounts?status=&branch=` | 账号列表（admin） |
| POST | `/api/v2/accounts` | 直接创建账号（admin） |
| PUT | `/api/v2/accounts/{id}` | 审批/改权限/改密码（admin） |
| DELETE | `/api/v2/accounts/{id}` | 删除账号（admin） |

### 报表导出
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v2/reports/picking-list` | 拣货单 Excel |
| GET | `/api/v2/reports/shortages` | 缺货报表 Excel |
| GET | `/api/v2/reports/reconciliation` | 对账报表 Excel |
| GET | `/api/v2/audit-log?export_csv=true` | 审计日志 CSV |

### 备份
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v2/backups` | 备份文件列表（admin） |
| POST | `/api/v2/backups` | 创建快照（admin） |
| GET | `/api/v2/backups/{filename}` | 下载备份文件（admin） |

### 仪表盘 & 系统
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v2/dashboard` | 关键指标 (admin/warehouse) |
| GET | `/api/v2/health` | 健康检查 |

### POS（共享商品/库存/价格）
| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/v2/pos/checkout` | 提交收银订单 |
| GET | `/api/v2/pos/stats/summary` | 收银汇总 |

---

## 7. 迁移步骤（从旧 orders.db + products.xlsx）

### 步骤一：运行 schema 迁移（幂等，安全）

```bash
cd C:\Users\Administrator\Desktop\Supermarket
python db_schema.py
# 输出: [db_schema] Migration complete: orders.db
```

所有旧数据保留，仅追加缺失列与新表。

### 步骤二：验证关键表存在

```bash
python -c "
import sqlite3
conn = sqlite3.connect('orders.db')
tables = [r[0] for r in conn.execute(\"SELECT name FROM sqlite_master WHERE type='table'\").fetchall()]
print(tables)
"
```

期望输出包含：`orders, shortages, inventory, inventory_txn, product_prices,
user_accounts, notifications, stock_arrivals, audit_log, branch_cart_draft,
supplier_orders, pos_orders, pos_order_lines`

### 步骤三：启动 api_v2.py

```bash
pip install fastapi uvicorn pandas openpyxl python-multipart
python -m uvicorn api_v2:app --host 127.0.0.1 --port 5056 --reload
```

### 步骤四：配置环境变量（生产）

```env
SUNSHINE_WAREHOUSE_PASSWORD=<strong-password>
SUNSHINE_ADMIN_PASSWORD=<strong-password>
API_V2_SECRET_KEY=<64-char-random>
```

可写入 `gemini.env`（与 `app.py` 同目录），格式与现有保持一致。

### 步骤五：现有 Streamlit app.py 无需修改

`app.py` 继续直连 `orders.db`，与 `api_v2.py` 并行运行，共享同一文件。  
未来如需将 `app.py` 切换为调用 API，只需将各 `db_conn()` 调用替换为 HTTP 请求。

### 步骤六（可选）：迁移 pos_api.py 流量

`pos-web` 的 Vite proxy 目前指向 `:5055`（`pos_api.py`）。  
可将 proxy 改为 `:5056` 使用 `api_v2.py` 的 `/api/v2/...` 端点，或保留两个服务并行。

---

## 8. 禁止事项（验收检查点）

| # | 禁止行为 | 如何规避 |
|---|---------|---------|
| 1 | 把箱/个合并为单一整数 | `inventory` 表永远保留 `stock_cartons` + `stock_pcs` 两列；`orders` 同理 |
| 2 | 省略缺货闭环 | `shortages` 表 + `PUT /api/v2/shortages/{id}` + 状态机 Open→Resolved |
| 3 | 省略订单三状态 | `status IN ('Pending','Dispatched','Received')` 硬编码检查在 dispatch/receive 接口 |
| 4 | SMTP 密钥进主数据库 | `email_config.json` 独立存储，备份 `orders.db` 不含邮件密钥 |
| 5 | API 密钥进数据库 | `GEMINI_API_KEY` 只读 env/gemini.env，从不写库 |

---

## 9. 风险与开放问题

| 风险 | 级别 | 建议 |
|------|------|------|
| SQLite 写并发（多人同时提交）| 中 | 已启用 WAL 模式；若并发>20 同时写，考虑迁移 PostgreSQL |
| token 无法主动吊销 | 低 | TTL=12h；如需立即失效，可在 DB 加 token 黑名单表 |
| products.xlsx 被覆盖导致价格消失 | 中 | 每次导入前备份旧文件；建议仅通过 `product_prices` 表管理生效价 |
| 图片目录无权限控制 | 低 | 图片按 ItemCode/Barcode 命名，无敏感信息；如需控制可加 auth 检查 |
| 备份文件与 orders.db 在同一磁盘 | 中 | 建议定期将 `backups/` 目录同步到外部存储（NAS/云盘） |
| `api_v2.py` 与 `app.py` 同时写同一 SQLite | 低 | WAL 模式支持多写者；长事务（如批量导入）应快速提交 |
| 旧 `shortages` 表列可能缺失 | 低 | `db_schema.py` 的 `MIGRATION_COLUMNS` 已处理 ADD COLUMN |
| 分店无密码登录无审计 | 中 | `/api/v2/auth/login` 写 audit_log；Streamlit `app.py` 侧也有审计 |

---

## 10. 运行命令汇总

```bash
# 迁移数据库
python db_schema.py

# 启动主 Streamlit 系统（现有）
streamlit run app.py --server.port 8502

# 启动新 REST API
python -m uvicorn api_v2:app --host 127.0.0.1 --port 5056 --reload

# 启动旧 POS API（保留兼容）
python -m uvicorn pos_api:app --host 127.0.0.1 --port 5055 --reload

# 启动 POS 前端（Vite）
cd pos-web && npm run dev
```

---

*生成时间: 2026-05-04 — 对应 api_v2.py v2.0.0 / db_schema.py v1.0*
