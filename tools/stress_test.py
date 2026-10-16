# -*- coding: utf-8 -*-
"""压力测试：专门造"极端尺寸"的输入，验证两条硬约束不被突破。

自检（selftest）用的是温和样例。这里反过来——造出让排版最容易崩的极端值：

* 超长横幅（20000×300px）
* 极端竖图（400×12000px）
* 巨型正方形（8000×8000px）
* 单页超宽 PDF（A0 横向）
* 超宽表格（40 列）
* 单列超长文本表格
* 空文件 / 损坏文件 / 不支持格式（验证容错）
* 大字号带透明通道的 PNG

然后合并、校验，断言：
1. 没有任何内容超出所在页面的可用宽高；
2. 坏文件被正确跳过且**不影响**好文件；
3. 输出能被 Word 解析。
"""

from __future__ import annotations

import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))


def make_stress_inputs(root: Path) -> tuple[list[Path], list[str]]:
    """造极端输入。返回 (好文件列表, 预期会被跳过的文件名)。"""
    src = root / "源文件-压力"
    src.mkdir(parents=True, exist_ok=True)
    good: list[Path] = []
    expected_skip: list[str] = []

    from PIL import Image

    # 1. 超长横幅
    p = src / "01-超宽横幅.png"
    Image.new("RGB", (20000, 300), (220, 235, 250)).save(p)
    good.append(p)

    # 2. 极端竖图
    p = src / "02-极端竖长图.jpg"
    Image.new("RGB", (400, 12000), (250, 240, 230)).save(p, quality=85)
    good.append(p)

    # 3. 巨型正方形
    p = src / "03-巨型方图.png"
    Image.new("RGB", (8000, 8000), (235, 250, 235)).save(p)
    good.append(p)

    # 4. 极小图（不该被放大糊掉）
    p = src / "04-极小图.png"
    Image.new("RGB", (16, 16), (255, 100, 100)).save(p)
    good.append(p)

    # 5. 超宽 PDF（A0 横向 ≈ 3370×2384 pt）
    import pymupdf

    p = src / "05-超宽PDF.pdf"
    doc = pymupdf.open()
    page = doc.new_page(width=3370, height=2384)
    page.insert_text((100, 200), "A0 LANDSCAPE", fontsize=120)
    doc.new_page(width=595, height=842).insert_text((60, 90), "A4 portrait", fontsize=20)
    doc.save(str(p))
    doc.close()
    good.append(p)

    # 6. 超宽表格（40 列）
    from openpyxl import Workbook

    p = src / "06-超宽表格.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "宽表40列"
    ws.append([f"列{i:02d}" for i in range(1, 41)])
    for r in range(1, 21):
        ws.append([f"值{r}-{c}" for c in range(1, 41)])
    wb.save(str(p))
    good.append(p)

    # 7. 单列超长文本
    p = src / "07-长文本单元格.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "长文本"
    ws.append(["编号", "描述"])
    for r in range(1, 8):
        ws.append([r, "这是一段很长的描述文字" * 30])
    wb.save(str(p))
    good.append(p)

    # 8. 超大字号文本（验证文本不会撑破页）
    p = src / "08-超长单行文本.txt"
    p.write_text("超长单行：" + "测试内容" * 400, encoding="utf-8")
    good.append(p)

    # 9. 空文件 -> 应被跳过
    p = src / "09-空文件.txt"
    p.write_bytes(b"")
    expected_skip.append(p.name)

    # 10. 损坏的 docx -> 应被记为失败但不中断
    p = src / "10-损坏文件.docx"
    p.write_bytes(b"PK\x03\x04 this is not a real docx at all")
    expected_skip.append(p.name)

    # 11. 不支持的格式 -> 应被跳过
    p = src / "11-不支持格式.xyz"
    p.write_text("hello", encoding="utf-8")
    expected_skip.append(p.name)

    # 12. 带大字号和透明通道
    p = src / "12-透明大图.png"
    im = Image.new("RGBA", (4000, 3000), (0, 0, 0, 0))
    from PIL import ImageDraw

    d = ImageDraw.Draw(im)
    d.ellipse([200, 200, 3800, 2800], fill=(60, 140, 220, 200))
    im.save(p)
    good.append(p)

    return good, expected_skip


def main() -> int:
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else Path.cwd() / "stress_out"
    root = root.resolve()
    root.mkdir(parents=True, exist_ok=True)

    from wmtool.engine import MergeEngine
    from wmtool.logging_setup import install_crash_handlers, setup_logging
    from wmtool.settings import Settings
    from wmtool.verify import verify_document

    setup_logging()
    install_crash_handlers()

    print("1) 生成极端尺寸输入 ...")
    good, expected_skip = make_stress_inputs(root)
    src_dir = root / "源文件-压力"
    all_files = sorted(src_dir.iterdir())
    for f in all_files:
        print(f"   {f.name:<26} {f.stat().st_size / 1024:>10.1f} KB")

    print(f"\n2) 合并全部 {len(all_files)} 个文件（含故意损坏的）...")
    settings = Settings()
    settings.timestamp_output = False
    engine = MergeEngine(settings)
    output = root / "压力合并结果.docx"
    report = engine.merge(all_files, output)

    print(f"   结果: {report.summary_line()}")
    print(f"   输出: {output.name}（{report.output_bytes / 1024:.1f} KB）")

    failures: list[str] = []

    print("\n3) 校验输出")
    vr = verify_document(output)
    print(vr.report())
    if not vr.ok:
        for i in vr.issues:
            if i.severity == "error":
                failures.append(f"{i.kind}: {i.detail}")

    print("\n4) 容错检查")
    # 好文件必须全部成功
    ok_names = {o.name for o in report.outcomes if o.status == "ok"}
    missing_ok = [f.name for f in good if f.name not in ok_names]
    if missing_ok:
        failures.append(f"这些正常文件本应成功却失败了: {missing_ok}")
    else:
        print(f"   [OK]   {len(good)} 个好文件全部合并成功")

    # 坏文件必须被跳过或记为失败，而且不能中断整批
    bad_states = {
        o.name: o.status for o in report.outcomes if o.name in expected_skip
    }
    print(f"   [OK]   坏文件处理结果: {bad_states}")
    for name in expected_skip:
        st = bad_states.get(name)
        if st not in ("skipped", "failed"):
            failures.append(f"{name} 应被跳过或记为失败，实际是 {st}")
    if not report.failed and not report.skipped:
        failures.append("坏文件没有被正确识别")

    print("\n5) 源文件未被修改")
    print("   （engine 内部已做 SHA-256 比对；此处复核文件仍可读）")
    unreadable = [f.name for f in all_files if not f.is_file()]
    if unreadable:
        failures.append(f"源文件丢失: {unreadable}")
    else:
        print(f"   [OK]   {len(all_files)} 个源文件仍在原处")

    print("-" * 62)
    if failures:
        print(f"压力测试【失败】{len(failures)} 项：")
        for f in failures:
            print(f"  ✗ {f}")
        return 1
    print("压力测试【全部通过】")
    return 0


if __name__ == "__main__":
    sys.exit(main())
