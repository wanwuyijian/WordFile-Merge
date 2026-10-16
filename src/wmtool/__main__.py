# -*- coding: utf-8 -*-
"""让 ``python -m wmtool`` 直接启动图形界面。"""

from __future__ import annotations

import sys


def main() -> int:
    """根据运行环境选择 GUI 或 CLI。"""
    # 打包成 exe 且不带参数时开 GUI；否则走命令行
    if getattr(sys, "frozen", False) and len(sys.argv) == 1:
        from .gui import main as gui_main

        return gui_main()
    if len(sys.argv) > 1:
        from .cli import main as cli_main

        return cli_main()
    try:
        from .gui import main as gui_main

        return gui_main()
    except Exception as exc:  # pragma: no cover
        print(f"图形界面启动失败（{exc}），改为显示命令行帮助。", file=sys.stderr)
        from .cli import main as cli_main

        return cli_main(["--help"])


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
