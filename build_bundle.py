"""
从 Desktop/Supermarket/app.py 按行切片生成模块化文件，并拼接为 sunmarket_bundle.py。
运行: python build_bundle.py
"""
from __future__ import annotations

import os
from pathlib import Path

# 源文件固定为当前项目根目录的 app.py，避免误指向旧目录副本
ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "app.py"

# 输出文件路径 -> (起始行, 结束行) 均为 1-based 闭区间
CHUNKS: list[tuple[str, tuple[int, int]]] = [
    ("config.py", (37, 202)),
    ("i18n.py", (204, 1049)),
    ("database.py", (1051, 2096)),
    (os.path.join("modules", "backup.py"), (2098, 2315)),
    (os.path.join("modules", "email_notify.py"), (2316, 2556)),
    (os.path.join("modules", "products.py"), (2557, 2958)),
    (os.path.join("modules", "core_utils.py"), (2960, 3867)),
    (os.path.join("ui", "style.py"), (3869, 4529)),
    (os.path.join("ui", "helpers.py"), (4531, 4628)),
    ("auth.py", (4630, 5477)),
    (os.path.join("views", "branch.py"), (5478, 6272)),
    (os.path.join("views", "warehouse.py"), (6274, 7186)),
    (os.path.join("views", "admin.py"), (7190, 7638)),
    (os.path.join("modules", "export.py"), (7640, 9186)),
    (os.path.join("views", "navigation.py"), (9188, 9407)),
    ("router.py", (9409, 9448)),
    ("app_main.py", (9453, 9488)),
]

IMPORT_HEADER_LINES = (11, 35)  # 1-based，与原 app 顶部 import 一致


def main() -> None:
    if not SOURCE.exists():
        raise SystemExit(f"源文件不存在: {SOURCE}")

    lines = SOURCE.read_text(encoding="utf-8").splitlines()
    hdr_lo, hdr_hi = IMPORT_HEADER_LINES
    header = "\n".join(lines[hdr_lo - 1 : hdr_hi])

    for rel, (a, b) in CHUNKS:
        out = ROOT / rel
        out.parent.mkdir(parents=True, exist_ok=True)
        body = "\n".join(lines[a - 1 : b])
        banner = (
            f'"""\n切片来源: app.py 第 {a}-{b} 行\n'
            f"运行入口请勿单独执行本文件；请使用根目录 app.py + sunmarket_bundle.py\n"
            f'"""\n\n'
        )
        out.write_text(banner + body + "\n", encoding="utf-8")
        print(f"Wrote {rel} ({b - a + 1} lines)")

    bundle_parts: list[str] = [
        '"""AUTO-GENERATED — do not edit by hand. Run build_bundle.py after changing slices."""\n',
        "from __future__ import annotations\n",
        header + "\n\n",
    ]
    for rel, (a, b) in CHUNKS:
        bundle_parts.append(f"\n# --- {rel} (app.py {a}-{b}) ---\n")
        bundle_parts.append("\n".join(lines[a - 1 : b]))
        bundle_parts.append("\n")

    bundle_path = ROOT / "sunmarket_bundle.py"
    bundle_path.write_text("".join(bundle_parts), encoding="utf-8")
    print(f"\nWrote {bundle_path.name} ({bundle_path.stat().st_size // 1024} KB approx)")


if __name__ == "__main__":
    main()
