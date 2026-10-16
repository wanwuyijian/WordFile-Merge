# -*- coding: utf-8 -*-
"""GUI 冒烟测试：真的把窗口建起来，驱动一遍关键交互，然后自己关掉。

跑在无人值守环境里也能执行——它不依赖人去点，
却覆盖了最容易出问题的几处：控件构建、卡片增删、拖拽排序、进度更新、日志渲染。
"""

from __future__ import annotations

import sys
import time
from pathlib import Path


def main() -> int:
    """构建窗口 -> 增删卡片 -> 排序 -> 模拟进度 -> 关闭。"""
    root_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path.cwd() / "selftest_out"
    src_dir = root_dir / "源文件"
    if not src_dir.is_dir():
        print(f"找不到样例目录：{src_dir}", file=sys.stderr)
        return 2

    from wmtool import gui

    app = gui.App()
    root = app.root
    failures: list[str] = []

    def check(cond: bool, label: str) -> None:
        print(("  [OK]   " if cond else "  [FAIL] ") + label)
        if not cond:
            failures.append(label)

    try:
        root.update()
        print("1) 窗口构建")
        check(root.winfo_exists() == 1, "主窗口已创建")
        check(app.grid is not None, "卡片网格已创建")

        print("2) 添加文件")
        files = sorted(src_dir.iterdir())
        n = app.grid.add_paths(files)
        root.update()
        check(n > 0, f"添加了 {n} 个文件")
        check(len(app.grid.cards) == n, f"卡片数 = {len(app.grid.cards)}")
        check(len(app.grid.ordered_paths()) == n, "顺序路径数一致")

        print("3) 缩略图")
        with_img = sum(1 for c in app.grid.cards if c._photo is not None)
        check(with_img == n, f"{with_img}/{n} 张卡片有缩略图")

        print("4) 调整顺序（把第 1 个往后移 2 位）")
        before = [p.name for p in app.grid.ordered_paths()]
        app.grid.move(app.grid.cards[0], 2)
        root.update()
        after = [p.name for p in app.grid.ordered_paths()]
        check(before != after, "顺序确实变了")
        check(sorted(before) == sorted(after), "文件集合没变（只是顺序变了）")
        print(f"        {before[0]} -> 位置 3；现在第 3 个是 {after[2]}")

        print("5) 选中与移除")
        app.grid.cards[0].set_selected(True)
        app.grid.remove_selected()
        root.update()
        check(len(app.grid.cards) == n - 1, f"移除后剩 {len(app.grid.cards)} 个")

        print("6) 进度与状态")
        app._on_progress_thread(3, 10, "正在处理测试文件")
        root.update()
        time.sleep(0.3)
        root.update()
        check(float(app.progress["value"]) > 0, f"进度条已推进（{app.progress['value']}）")

        print("7) 日志窗口")
        app._enqueue_log("2026-01-01 00:00:00.000 [ERROR   ] 测试错误行")
        root.update()
        time.sleep(0.3)
        root.update()
        content = app.log_text.get("1.0", "end")
        check("测试错误行" in content, "日志已渲染到窗口")

        print("8) 清空")
        app.grid.clear()
        root.update()
        check(len(app.grid.cards) == 0, "网格已清空")

        print("9) 关闭流程")
        app._closing = True
        root.update()
        root.destroy()
        print("  [OK]   窗口已关闭")
    except Exception as exc:
        import traceback

        traceback.print_exc()
        failures.append(f"异常: {exc}")
    finally:
        try:
            root.destroy()
        except Exception:
            pass

    print("-" * 60)
    if failures:
        print(f"GUI 冒烟测试【失败】{len(failures)} 项：")
        for f in failures:
            print(f"  ✗ {f}")
        return 1
    print("GUI 冒烟测试【全部通过】")
    return 0


if __name__ == "__main__":
    sys.exit(main())
