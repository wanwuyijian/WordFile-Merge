# -*- coding: utf-8 -*-
"""一键把所有项目内 .py 编译检查一遍（语法 + 字节码）。

比 `python -m compileall` 多做的事：**把每个模块都 import 一次**，
这样"语法没错但导入就炸"的问题（循环导入、写错模块名、缺依赖）也能当场暴露。

用法::

    python tools/compile_check.py
"""

from __future__ import annotations

import importlib
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))


def main() -> int:
    """编译并导入全部模块。"""
    failures: list[str] = []

    print("1) 语法检查（compile）")
    files = sorted(SRC.rglob("*.py")) + sorted((ROOT / "tools").rglob("*.py"))
    files += sorted((ROOT / "packaging").rglob("*.py"))
    for f in files:
        try:
            source = f.read_text(encoding="utf-8")
            compile(source, str(f), "exec")
        except SyntaxError as exc:
            failures.append(f"{f.relative_to(ROOT)}: 第 {exc.lineno} 行语法错误: {exc.msg}")
            print(f"   [FAIL] {f.relative_to(ROOT)}: {exc.msg}")
        except Exception as exc:
            failures.append(f"{f.relative_to(ROOT)}: {exc}")
            print(f"   [FAIL] {f.relative_to(ROOT)}: {exc}")
    print(f"   检查了 {len(files)} 个文件")

    print("\n2) 导入检查（import）")
    importable = [
        "wmtool",
        "wmtool.paths",
        "wmtool.logging_setup",
        "wmtool.settings",
        "wmtool.geometry",
        "wmtool.docx_util",
        "wmtool.mdparse",
        "wmtool.verify",
        "wmtool.engine",
        "wmtool.selftest",
        "wmtool.cli",
        "wmtool.gui",
        "wmtool.converters",
        "wmtool.converters.base",
        "wmtool.converters.image",
        "wmtool.converters.pdf",
        "wmtool.converters.docx",
        "wmtool.converters.xlsx",
        "wmtool.converters.pptx",
        "wmtool.converters.text",
        "wmtool.converters.legacy",
    ]
    for name in importable:
        try:
            importlib.import_module(name)
            print(f"   [OK]   {name}")
        except Exception as exc:
            failures.append(f"导入 {name} 失败: {exc}")
            print(f"   [FAIL] {name}: {exc}")
            traceback.print_exc()

    print("\n" + "=" * 62)
    if failures:
        print(f"编译检查【失败】{len(failures)} 项：")
        for f in failures:
            print(f"  ✗ {f}")
        return 1
    print(f"编译检查【全部通过】（{len(files)} 个文件 + {len(importable)} 个模块）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
