"""
expiry_alert_job.py — 分店库存临期/过期预警 · 定时任务入口
==========================================================
每天凌晨运行：扫描所有分店批次库存，找出临期/已过期且仍有库存的商品，
按分店归类后写入消息中心并发邮件给对应店长。

用法 / Usage:
    python expiry_alert_job.py                 # 用默认预警天数（环境变量或 15 天）
    python expiry_alert_job.py --days 7        # 临时指定预警天数
    python expiry_alert_job.py --today 2026-06-01   # 指定"当前日期"（测试用）
    python expiry_alert_job.py --no-email      # 只写消息中心，不发邮件（演练）

部署见 README「临期预警定时任务」一节（Railway Cron / Windows 计划任务）。

设计要点：
  - 复用主系统 app.py 的数据库、邮件、通知逻辑，避免重复实现。
  - 自带"每日去重"保护：同一天重复触发只真正执行一次（防止多次唤醒重复发信）。
    去重标记写在数据 dir 下的 .expiry_job_lastrun 文件。
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path


def _log(msg: str) -> None:
    print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} [expiry_job] {msg}")


def _lastrun_path(app) -> Path:
    # 与 orders.db 同目录，跟随 Railway Volume / 本地工作目录。
    return Path(app.DB_PATH).parent / ".expiry_job_lastrun"


def _already_ran_today(app, today_label: str, force: bool) -> bool:
    if force:
        return False
    p = _lastrun_path(app)
    try:
        return p.exists() and p.read_text(encoding="utf-8").strip() == today_label
    except Exception:
        return False


def _mark_ran_today(app, today_label: str) -> None:
    try:
        _lastrun_path(app).write_text(today_label, encoding="utf-8")
    except Exception as e:
        _log(f"warn: 无法写入去重标记: {e}")


def main() -> int:
    parser = argparse.ArgumentParser(description="分店临期/过期预警定时任务")
    parser.add_argument("--days", type=int, default=None, help="预警天数（默认取环境变量或 15）")
    parser.add_argument("--today", default="", help="指定当前日期 YYYY-MM-DD（测试用）")
    parser.add_argument("--no-email", action="store_true", help="只写消息中心，不发邮件")
    parser.add_argument("--force", action="store_true", help="忽略每日去重，强制执行")
    args = parser.parse_args()

    # 导入主系统（不会启动 Streamlit 服务，只加载业务逻辑与配置）。
    try:
        import app
    except Exception as e:
        _log(f"FATAL: 无法导入 app.py: {type(e).__name__}: {e}")
        return 2

    # 确保数据库与新表已就绪（CREATE TABLE IF NOT EXISTS，幂等安全）。
    try:
        app.init_db()
    except Exception as e:
        _log(f"FATAL: init_db 失败: {type(e).__name__}: {e}")
        return 2

    today_label = (args.today or app.today_str()).strip()

    if _already_ran_today(app, today_label, args.force):
        _log(f"今日（{today_label}）已执行过，跳过。用 --force 可强制重跑。")
        return 0

    days = args.days if args.days is not None else app.expiry_warn_days()
    _log(f"开始扫描：预警天数={days}，基准日期={today_label}，发邮件={not args.no_email}")

    try:
        summary = app.run_expiry_scan_and_notify(
            warn_days=days,
            today=today_label,
            send_email=not args.no_email,
        )
    except Exception as e:
        _log(f"FATAL: 扫描/通知失败: {type(e).__name__}: {e}")
        return 2

    _mark_ran_today(app, today_label)
    _log(
        "完成："
        f"预警分店 {summary['branches_alerted']} 家，"
        f"商品 {summary['items_total']} 项，"
        f"发出邮件 {summary['emails_sent']} 封。"
    )
    if summary["details"]:
        for branch, n in summary["details"].items():
            _log(f"  - {branch}: {n} 项")
    # 邮件为后台线程异步发送，稍等片刻以便线程完成投递。
    if not args.no_email and summary["emails_sent"] > 0:
        import time
        time.sleep(8)
    return 0


if __name__ == "__main__":
    sys.exit(main())
