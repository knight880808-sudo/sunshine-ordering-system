# Git 上传与忽略参考（SUNSHINE / Supermarket_clean）

本文说明：**哪些文件建议提交到 Git**，**哪些不要提交**，以及**日常该怎么操作**。可随时打开本文件对照，不必先删除 GitHub 上的文件再上传——用 `.gitignore` + 正常 `commit` / `push` 即可维护干净仓库。

---

## 1. 建议务必上传（程序与可复现运行）

更换电脑后，靠这些内容应能重新安装依赖并跑起来：

| 类别 | 典型路径 / 文件 |
|------|------------------|
| Streamlit 主程序 | `app.py` |
| API / 其它后端入口 | `pos_api.py`、`api_v2.py`（若你仍在使用） |
| 启动脚本 | `run_supermarket.bat`、`run_pos_api.bat` |
| Python 依赖清单 | `requirements.txt`、`requirements-pos-api.txt` |
| 业务模块 | `modules/`、`views/`、`ui/` |
| 路由与配置相关 | `router.py`、`auth.py`、`config.py`、`database.py`、`db_schema.py`、`i18n.py`、`sunmarket_bundle.py`（若适用） |
| POS 前端源码 | `pos-web/`（含 `package.json`、`src/` 等；**不要**上传 `node_modules/`） |
| Streamlit 配置 | `.streamlit/` |
| 文档 | `README.md`、`ARCHITECTURE.md`、`Operation_Guide_EN.md`、`新手操作指南.md` 等 |
| 构建辅助 | `build_bundle.py`（若仍用切片管线） |

---

## 2. 可以上传，但请自行权衡

| 内容 | 说明 |
|------|------|
| `products.xlsx` | 若为**小样本/脱敏**，可提交；若为**全量真实商品档案**且体积大或涉密，建议不放仓库，另用网盘或内部备份。 |
| `price.xlsx` | 同上，看是否含敏感定价。 |
| `商品档案_*.xlsx` | 一般为业务主数据，**优先本地或私有存储**；确需共享时用私有仓库并控制权限。 |
| `images/` 商品图 | 可提交；体积很大时可考虑 Git LFS、对象存储或文档中说明「图片自行部署」。 |

---

## 3. 不要上传（数据、密钥、缓存、本地生成）

以下容易**泄密、冲突或撑爆仓库**，应通过 `.gitignore` 排除，且**不要**把真实生产库提交到公开仓库：

| 类型 | 典型路径 / 文件 |
|------|------------------|
| 订单与业务数据库 | `orders.db` |
| 自动数据库快照 | `backups/*.db` |
| API / 模型密钥 | `gemini.env` |
| 邮件 SMTP 等敏感配置 | `email_config.json`（若含账号密码） |
| 运行日志 | `app_runtime.log`、`email_log.json` |
| Excel 临时锁文件 | `~$*.xlsx` |
| Python 缓存 | `__pycache__/`、`*.pyc` |
| 前端依赖目录 | `pos-web/node_modules/` |

> **说明**：若 `orders.db` 等曾被误提交，仅从网页删文件不够，需在本地 `git rm --cached` 并提交，历史里可能仍保留记录；公开仓库建议轮换密钥并视情况清理历史。

---

## 4. 备份与历史副本（不必进仓库）

| 路径 | 建议 |
|------|------|
| `备份/`、`备份/历史版本_*` | 旧版 `app.py` 等快照，**留在本机或网盘**即可；主仓库只保留当前主线代码。 |

---

## 5. 推荐操作流程（不必先清空 GitHub）

1. 在本机维护好 `.gitignore`（根目录已有基础规则，可按上表补充）。  
2. `git status` 检查：不应出现 `orders.db`、日志、`.db` 备份等被加入暂存区。  
3. 只添加「第 1 节」及你确认要共享的「第 2 节」文件。  
4. `git commit` → `git push`。  

若要以本地为准覆盖远端分支，在**确认无他人协作或已沟通**的前提下再考虑 `git push --force-with-lease`（慎用）。

---

## 6. 当前仓库内 `.gitignore` 已包含（便于对照）

以下为编写本说明时，项目中已配置的忽略项（若之后有变动，以 `.gitignore` 文件为准）：

- `gemini.env`
- `__pycache__/`、`*.pyc`
- `backups/*.db`
- `app_runtime.log`
- `email_log.json`
- `~$products.xlsx`

**建议自行核对**：`orders.db` 是否仍被 Git 跟踪；若被跟踪，应加入 `.gitignore` 并对已跟踪文件执行 `git rm --cached orders.db` 后提交。

---

## 7. 一句话记忆

- **上传**：代码、依赖列表、必要配置模板、文档。  
- **不上传**：数据库、日志、密钥、大体积私有业务表、缓存与 `node_modules`。  
- **备份目录**：本机保留，一般不进 Git。

---

*文档随项目演进可继续补充；修改忽略规则时记得同步更新本节与 `.gitignore`。*
