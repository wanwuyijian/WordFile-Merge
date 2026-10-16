# -*- coding: utf-8 -*-
"""统一测试入口：一条命令跑完全部验证。

包含四层，从快到慢、从细到粗：

1. **单元测试** —— 几何/方向决策/Markdown/docx 封装/日志与崩溃日志（秒级）
2. **端到端自检** —— 造各格式样例，真合并一次，校验产物（十几秒）
3. **压力测试** —— 极端尺寸输入 + 损坏文件容错（半分钟）
4. **GUI 检查** —— 界面布局断言 + 交互冒烟（十几秒；无图形环境时自动跳过）

用法::

    python tools/run_all_tests.py                # 全部
    python tools/run_all_tests.py --no-gui       # 跳过界面（无桌面环境）
    python tools/run_all_tests.py --exe dist\\WordMergeTool-CLI.exe   # 顺带验 exe
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
TOOLS = ROOT / "tools"


def run_step(title: str, cmd: list[str], *, cwd: Path | None = None) -> tuple[bool, str]:
    """跑一个步骤，返回 (是否成功, 输出)。"""
    print("\n" + "=" * 70)
    print(f" {title}")
    print("=" * 70)
    env = dict(os.environ)
    env["PYTHONPATH"] = str(SRC)
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"

    started = time.time()
    try:
        proc = subprocess.run(
            cmd,
            cwd=str(cwd or ROOT),
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=1800,
        )
    except subprocess.TimeoutExpired:
        print("  [超时] 该步骤超过 30 分钟")
        return False, "timeout"

    output = (proc.stdout or "") + (proc.stderr or "")
    # 过滤掉日志行，只留人类可读的断言输出
    for line in output.splitlines():
        s = line.rstrip()
        if not s:
            continue
        if len(s) > 19 and s[4] == "-" and s[7] == "-" and "[" in s:
            continue  # 形如 "2026-10-16 09:00:00.000 [INFO ] ..." 的日志行
        print("  " + s)

    elapsed = time.time() - started
    ok = proc.returncode == 0
    print(f"  --> {'通过' if ok else '失败'}（{elapsed:.1f}s，退出码 {proc.returncode}）")
    return ok, output


def has_display() -> bool:
    """有没有图形环境（Windows 上基本都有；这里主要防 CI）。"""
    if os.name != "nt":
        return bool(os.environ.get("DISPLAY"))
    return True


def main() -> int:
    ap = argparse.ArgumentParser(description="Word 合并工具 全量测试")
    ap.add_argument("--no-gui", action="store_true", help="跳过 GUI 相关测试")
    ap.add_argument("--no-stress", action="store_true", help="跳过压力测试")
    ap.add_argument("--exe", help="额外的 exe 路径，用它再跑一遍端到端自检")
    ap.add_argument("--keep", action="store_true", help="保留测试产生的临时文件")
    args = ap.parse_args()

    python = sys.executable
    workdir = Path(tempfile.mkdtemp(prefix="wmtool-tests-"))
    results: list[tuple[str, bool]] = []

    print("Word 合并工具 —— 全量验证")
    print(f"项目根目录: {ROOT}")
    print(f"临时目录:   {workdir}")

    try:
        # 1. 单元测试
        ok, _ = run_step("1/4 单元测试", [python, str(TOOLS / "run_unit_tests.py")])
        results.append(("单元测试", ok))

        # 2. 端到端自检
        selftest_dir = workdir / "selftest"
        ok, _ = run_step(
            "2/4 端到端自检（造样例 -> 合并 -> 校验）",
            [python, "-m", "wmtool.cli", "--selftest", str(selftest_dir)],
        )
        results.append(("端到端自检", ok))

        # 3. 压力测试
        if args.no_stress:
            print("\n(已跳过压力测试)")
        else:
            ok, _ = run_step(
                "3/4 压力测试（极端尺寸 + 损坏文件）",
                [python, str(TOOLS / "stress_test.py"), str(workdir / "stress")],
            )
            results.append(("压力测试", ok))

        # 4. GUI
        if args.no_gui or not has_display():
            print("\n(已跳过 GUI 测试)")
        else:
            ok, _ = run_step(
                "4/4a GUI 布局校验",
                [python, str(TOOLS / "gui_layout_check.py")],
            )
            results.append(("GUI 布局校验", ok))

            # 冒烟测试需要有样例，复用自检生成的目录
            sample_dir = selftest_dir
            if sample_dir.is_dir():
                ok, _ = run_step(
                    "4/4b GUI 交互冒烟",
                    [python, str(TOOLS / "gui_smoketest.py"), str(sample_dir)],
                )
                results.append(("GUI 交互冒烟", ok))
            else:
                print("\n(跳过 GUI 冒烟：没有找到样例目录)")

        # 5. 可选：验证打包好的 exe
        if args.exe:
            exe = Path(args.exe)
            if not exe.is_absolute():
                exe = ROOT / exe
            if exe.is_file():
                ok, _ = run_step(
                    f"附加：验证 exe（{exe.name}）",
                    [str(exe), "--selftest", str(workdir / "exe-selftest")],
                )
                results.append((f"exe 验证（{exe.name}）", ok))
            else:
                print(f"\n(找不到 exe：{exe})")
                results.append((f"exe 验证（{exe.name}）", False))

    finally:
        if not args.keep:
            shutil.rmtree(workdir, ignore_errors=True)
        else:
            print(f"\n临时文件保留在: {workdir}")

    print("\n" + "=" * 70)
    print(" 汇总")
    print("=" * 70)
    for name, ok in results:
        print(f"  {'✔ 通过' if ok else '✘ 失败'}   {name}")
    failed = [n for n, ok in results if not ok]
    print("-" * 70)
    if failed:
        print(f" 共 {len(results)} 项，失败 {len(failed)} 项：{'、'.join(failed)}")
        return 1
    print(f" 共 {len(results)} 项，全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
