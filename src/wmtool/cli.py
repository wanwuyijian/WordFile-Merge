# -*- coding: utf-8 -*-
"""命令行入口（自测与批量脚本用）。

用 ``python -m wmtool.cli`` 或打包后的 ``WordMergeTool-CLI.exe`` 调用。

示例::

    # 合并整个目录
    WordMergeTool-CLI.exe --input "D:/资料" --output "D:/合并结果.docx"

    # 指定多个文件（按给定顺序）
    WordMergeTool-CLI.exe --files a.docx b.pdf c.png --output out.docx

    # 无界面自检：生成样例并跑一遍
    WordMergeTool-CLI.exe --selftest "D:/tmp/selftest"
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from . import __app_name_cn__, __version__
from .converters import supported_suffixes
from .engine import MergeEngine, MergeReport
from .logging_setup import get_logger, install_crash_handlers, setup_logging
from .paths import app_dir, logs_dir
from .settings import load_settings, save_settings


def build_parser() -> argparse.ArgumentParser:
    """构造命令行参数。"""
    p = argparse.ArgumentParser(
        prog="WordMergeTool",
        description=f"{__app_name_cn__} v{__version__} —— 把多种格式的文件合并成一个 Word 文档",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  WordMergeTool-CLI.exe --input D:\\资料 --output D:\\合并.docx\n"
            "  WordMergeTool-CLI.exe --files a.docx b.pdf c.png --output out.docx\n"
            "  WordMergeTool-CLI.exe --list-formats\n"
        ),
    )
    p.add_argument("--input", "-i", help="输入目录（递归扫描其中的支持格式）")
    p.add_argument("--files", "-f", nargs="+", help="显式指定文件列表（按给定顺序合并）")
    p.add_argument("--output", "-o", help="输出 .docx 路径")
    p.add_argument("--recursive", "-r", action="store_true", help="递归扫描子目录（默认只扫一层）")
    p.add_argument("--no-titles", action="store_true", help="不插入每个文件的标题页")
    p.add_argument("--no-toc", action="store_true", help="不生成目录")
    p.add_argument("--no-bookmarks", action="store_true", help="不加书签")
    p.add_argument("--timeout", type=float, default=180.0, help="单文件超时秒数（默认 180）")
    p.add_argument("--list-formats", action="store_true", help="列出支持的格式后退出")
    p.add_argument("--selftest", metavar="DIR", help="在指定目录生成各格式样例并跑一次端到端合并")
    p.add_argument("--version", "-V", action="version", version=f"{__app_name_cn__} {__version__}")
    return p


def collect_from_dir(root: Path, recursive: bool) -> list[Path]:
    """从目录收集支持的文件（自然序：a1, a2, a10 而不是 a1, a10, a2）。"""
    exts = set(supported_suffixes())
    it = root.rglob("*") if recursive else root.glob("*")
    files = [p for p in it if p.is_file() and p.suffix.lower() in exts and not p.name.startswith("~$")]
    return sorted(files, key=natural_key)


def natural_key(path: Path):  # type: ignore[no-untyped-def]
    """自然排序键：把文件名里的数字当数字比较。"""
    import re

    parts = re.split(r"(\d+)", path.name)
    return [int(t) if t.isdigit() else t.lower() for t in parts]


def print_report(report: MergeReport) -> None:
    """把结果打到控制台。"""
    print()
    print("=" * 66)
    print(f" 结果: {report.summary_line()}")
    if report.output_path:
        print(f" 输出: {report.output_path}")
    if report.error:
        print(f" 问题: {report.error}")
    print("=" * 66)
    for i, o in enumerate(report.outcomes, start=1):
        flag = {"ok": "[OK]  ", "skipped": "[跳过]", "failed": "[失败]"}.get(o.status, "[?]   ")
        line = f"{flag} {i:>3}. {o.name}"
        if o.status == "ok" and o.summary:
            line += f"  —— {o.summary}"
        if o.status != "ok" and o.error:
            line += f"  —— {o.error}"
        print(line)
        for w in o.warnings:
            print(f"        警告: {w}")
    print()


def run_selftest(target: Path) -> int:
    """端到端自检：造样例 -> 合并 -> 校验产物。"""
    from .selftest import run_selftest as _run

    return _run(target)


def main(argv: list[str] | None = None) -> int:
    """CLI 主入口。"""
    setup_logging()
    install_crash_handlers()
    log = get_logger("cli")

    parser = build_parser()
    args = parser.parse_args(argv)

    if args.list_formats:
        print(f"{__app_name_cn__} v{__version__} 支持的格式：")
        for e in supported_suffixes():
            print(f"  {e}")
        print(f"\n共 {len(supported_suffixes())} 种扩展名。")
        print(f"日志目录: {logs_dir()}")
        return 0

    if args.selftest:
        return run_selftest(Path(args.selftest))

    # 收集文件
    files: list[Path] = []
    if args.files:
        files = [Path(f).resolve() for f in args.files]
    elif args.input:
        root = Path(args.input).resolve()
        if not root.is_dir():
            print(f"错误：输入目录不存在 - {root}", file=sys.stderr)
            return 2
        files = collect_from_dir(root, args.recursive)

    if not files:
        parser.print_help()
        print("\n错误：没有可合并的文件（请用 --input 或 --files 指定）", file=sys.stderr)
        return 2

    missing = [f for f in files if not f.exists()]
    if missing:
        for m in missing:
            print(f"错误：文件不存在 - {m}", file=sys.stderr)
        return 2

    # 输出路径
    if args.output:
        output = Path(args.output).resolve()
    else:
        stamp = time.strftime("%Y%m%d-%H%M%S")
        out_dir = Path(args.input).resolve() if args.input else Path.cwd()
        output = out_dir / f"合并文档-{stamp}.docx"
    output.parent.mkdir(parents=True, exist_ok=True)

    settings = load_settings()
    settings.include_titles = not args.no_titles
    settings.include_toc = not args.no_toc
    settings.include_bookmarks = not args.no_bookmarks

    log.info("命令行参数: 文件 %d 个，输出 %s", len(files), output)

    def progress(index: int, total: int, message: str) -> None:
        print(f"\r[{index}/{total}] {message[:60]:<60}", end="", flush=True)

    engine = MergeEngine(settings)
    report = engine.merge(files, output, progress=progress, timeout_per_file=args.timeout)
    print("\r" + " " * 70 + "\r", end="")
    print_report(report)

    save_settings(settings)

    if report.cancelled:
        return 130
    if report.failed and report.ok_count == 0:
        return 3
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
