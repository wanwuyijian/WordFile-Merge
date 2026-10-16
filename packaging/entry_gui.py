# -*- coding: utf-8 -*-
"""PyInstaller 入口：图形界面版。

为什么需要这个文件：打包时如果直接把 `src/wmtool/gui.py` 当作入口脚本，
PyInstaller 会以 `__main__` 的形式运行它——此时它**不是** `wmtool` 包的成员，
文件里的 `from . import ...`（相对导入）就会报
"attempted relative import with no known parent package"。

正确做法是让入口脚本只做一件事：把 `wmtool` 包**绝对导入**进来再调 main()。
这样包身份是完整的，包内所有相对导入都能正常工作。
"""

from __future__ import annotations

import sys
from pathlib import Path

# 源码运行时（未打包）把 src 加进搜索路径；打包后 PyInstaller 已经收好了包，这步是空操作。
if not getattr(sys, "frozen", False):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from wmtool.gui import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
