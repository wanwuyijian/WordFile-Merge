# -*- coding: utf-8 -*-
"""PyInstaller 入口：命令行版。

同 entry_gui.py 的理由——入口必须绝对导入 `wmtool` 包，
否则包内的相对导入会因为"没有父包"而失败。

另外这里顺手做一件事：把控制台设为 UTF-8，
否则在中文 Windows 的 cmd 里打印中文进度可能出现乱码。
"""

from __future__ import annotations

import sys
from pathlib import Path

if not getattr(sys, "frozen", False):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# 控制台编码尽量拉到 UTF-8（失败就算了，不影响功能）
try:
    if sys.stdout is not None and hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if sys.stderr is not None and hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from wmtool.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
