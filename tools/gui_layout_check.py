# -*- coding: utf-8 -*-
"""界面布局校验：在离屏状态下把窗口真正建起来，检查每个控件的位置与尺寸。

比截图更可靠的几点：
* 能断言"控件可见、尺寸非零、没互相压住"这类客观事实；
* 不依赖人眼看图，也不需要视觉模型；
* 能输出一份布局清单，出问题时直接看出是谁错了。

用法::

    python tools/gui_layout_check.py
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))


def main() -> int:
    from wmtool import gui

    app = gui.App()
    root = app.root
    failures: list[str] = []
    warnings: list[str] = []

    def check(cond: bool, label: str, detail: str = "") -> None:
        tag = "[OK]  " if cond else "[FAIL]"
        print(f"  {tag} {label}" + (f"  ({detail})" if detail else ""))
        if not cond:
            failures.append(label)

    try:
        # 让窗口完成一次完整布局
        root.update_idletasks()
        root.update()

        print("窗口")
        w, h = root.winfo_width(), root.winfo_height()
        check(w > 600 and h > 400, "窗口尺寸合理", f"{w}x{h}")
        check("Word" in root.title(), "标题正确", root.title())

        print("\n关键控件可见性与尺寸")
        widgets = {
            "卡片网格 canvas": app.grid.canvas,
            "卡片滚动条": app.grid.scroll,
            "日志文本框": app.log_text,
            "进度条": app.progress,
            "开始合并按钮": app.btn_merge,
            "取消按钮": app.btn_cancel,
            "状态栏": app.status,
        }
        for name, wdg in widgets.items():
            ww = wdg.winfo_width()
            hh = wdg.winfo_height()
            check(
                ww > 1 and hh > 1,
                f"{name} 已渲染且有尺寸",
                f"{ww}x{hh}",
            )

        print("\n布局合理性")
        # 进度条不能比窗口还宽
        check(app.progress.winfo_width() <= w, "进度条宽度未超出窗口")
        # 日志区起点应在卡片区下方（垂直堆叠）
        check(
            app.log_text.winfo_rooty() >= app.grid.canvas.winfo_rooty(),
            "日志区在文件区下方（垂直分布正确）",
        )
        # 按钮都在窗口内部，没有被挤出可见区域
        for name, btn in (("开始合并", app.btn_merge), ("取消", app.btn_cancel)):
            bx = btn.winfo_rootx() - root.winfo_rootx()
            by = btn.winfo_rooty() - root.winfo_rooty()
            check(
                0 <= bx < w and 0 <= by < h,
                f"{name} 按钮在窗口可视范围内",
                f"({bx},{by})",
            )

        print("\n选项勾选框")
        opts_found = 0
        for child in root.winfo_children():
            for sub in child.winfo_children():
                for wdg in sub.winfo_children():
                    if wdg.winfo_class() == "TCheckbutton":
                        txt = ""
                        try:
                            txt = wdg.cget("text")
                        except Exception:
                            pass
                        if txt:
                            opts_found += 1
        check(opts_found == 3, "三个选项勾选框都在", f"找到 {opts_found} 个")

        print("\n空状态提示")
        # 没有文件时应显示提示文字
        check(
            app.grid._placeholder.winfo_ismapped(),
            "空状态提示可见",
        )

        print("\n添加文件后的布局")
        sample = Path(__file__).resolve().parents[1] / "selftest_out" / "源文件"
        if sample.is_dir():
            app.grid.add_paths(sorted(sample.iterdir())[:6])
            root.update_idletasks()
            root.update()
            n_cards = len(app.grid.cards)
            check(n_cards == 6, f"卡片已创建（{n_cards} 个）")
            mapped = sum(1 for c in app.grid.cards if c.winfo_ismapped())
            check(mapped == n_cards, f"卡片全部可见（{mapped}/{n_cards}）")
            # 卡片不该互相压住
            boxes = [
                (c.winfo_rootx(), c.winfo_rooty(), c.winfo_width(), c.winfo_height())
                for c in app.grid.cards
            ]
            overlap = 0
            for i in range(len(boxes)):
                for j in range(i + 1, len(boxes)):
                    ax, ay, aw, ah = boxes[i]
                    bx, by, bw, bh = boxes[j]
                    if ax < bx + bw and bx < ax + aw and ay < by + bh and by < ay + ah:
                        overlap += 1
            check(overlap == 0, "卡片之间没有重叠", f"重叠对数 {overlap}")
            check(
                not app.grid._placeholder.winfo_ismapped(),
                "有文件后空状态提示已隐藏",
            )
            app.grid.clear()
            root.update()
        else:
            warnings.append("跳过卡片布局检查（未找到样例目录）")

    finally:
        try:
            app._closing = True
            root.destroy()
        except Exception:
            pass

    print("-" * 62)
    for wmsg in warnings:
        print(f"  [warn] {wmsg}")
    if failures:
        print(f"布局校验【失败】{len(failures)} 项：")
        for f in failures:
            print(f"  ✗ {f}")
        return 1
    print("布局校验【全部通过】")
    return 0


if __name__ == "__main__":
    sys.exit(main())
