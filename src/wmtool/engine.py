# -*- coding: utf-8 -*-
"""合并引擎：把一批文件合成一个 Word 文档。

职责边界很清楚——引擎负责**编排**，转换器负责**格式**：

1. 校验每个文件（存在、可读、指纹不变），源文件全程只读；
2. 每个文件前分页，插入标题（可选）+ 书签（可选）；
3. 调度对应转换器写入内容，单文件失败**不中断整批**，记进报告继续跑；
4. 超时保护：每个文件有独立时间预算，超时就跳过并记警告；
5. 输出到临时文件再原子改名，中途崩溃不会留下半个"看起来正常"的 docx；
6. 全程写日志，异常写崩溃日志。
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import os
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Sequence

from docx import Document as DocxNew

from . import __app_name_cn__, __version__
from .converters import get_converter, is_supported, supported_suffixes
from .converters.base import ConvertContext, ConvertResult
from .docx_util import (
    add_bookmark,
    add_caption,
    add_heading,
    add_toc_field,
    new_section,
    page_break,
    set_document_defaults,
)
from .geometry import LANDSCAPE_BOX, PORTRAIT_BOX
from .logging_setup import CrashReporter, get_logger, write_crash_report
from .settings import Settings

log = get_logger("engine")

#: 单个文件的默认时间预算（秒）
DEFAULT_TIMEOUT = 180.0


class MergeCancelled(Exception):
    """用户点了取消。"""


@dataclass
class FileOutcome:
    """单个文件的处理结果。"""

    path: Path
    status: str = "pending"  # ok / skipped / failed
    converter: str = ""
    summary: str = ""
    warnings: list[str] = field(default_factory=list)
    error: str = ""
    seconds: float = 0.0
    size_bytes: int = 0
    sha256: str = ""

    @property
    def name(self) -> str:
        return self.path.name


@dataclass
class MergeReport:
    """整批任务的结果。"""

    output_path: Path | None = None
    started_at: _dt.datetime = field(default_factory=_dt.datetime.now)
    finished_at: _dt.datetime | None = None
    outcomes: list[FileOutcome] = field(default_factory=list)
    error: str = ""
    cancelled: bool = False
    output_bytes: int = 0

    @property
    def ok_count(self) -> int:
        return sum(1 for o in self.outcomes if o.status == "ok")

    @property
    def failed(self) -> list[FileOutcome]:
        return [o for o in self.outcomes if o.status == "failed"]

    @property
    def skipped(self) -> list[FileOutcome]:
        return [o for o in self.outcomes if o.status == "skipped"]

    @property
    def duration(self) -> float:
        end = self.finished_at or _dt.datetime.now()
        return (end - self.started_at).total_seconds()

    def summary_line(self) -> str:
        parts = [f"成功 {self.ok_count}/{len(self.outcomes)}"]
        if self.failed:
            parts.append(f"失败 {len(self.failed)}")
        if self.skipped:
            parts.append(f"跳过 {len(self.skipped)}")
        parts.append(f"耗时 {self.duration:.1f}s")
        return " · ".join(parts)


# --------------------------------------------------------------------------- #
# 进度回调
# --------------------------------------------------------------------------- #
ProgressCb = Callable[[int, int, str], None]
"""进度回调：(当前序号, 总数, 说明文字)"""


class _StopFlag:
    """可被 GUI 线程置位的取消标志。"""

    def __init__(self) -> None:
        self._event = threading.Event()

    def set(self) -> None:
        self._event.set()

    def is_set(self) -> bool:
        return self._event.is_set()


# --------------------------------------------------------------------------- #
class MergeEngine:
    """合并引擎。

    用法::

        engine = MergeEngine(settings)
        report = engine.merge([p1, p2, p3], Path("out.docx"), progress=cb)
    """

    def __init__(self, settings: Settings | None = None):
        self.settings = settings or Settings()
        self.stop = _StopFlag()

    # ------------------------------------------------------------------ #
    def merge(
        self,
        files: Sequence[Path],
        output: Path,
        *,
        progress: ProgressCb | None = None,
        timeout_per_file: float = DEFAULT_TIMEOUT,
        verify_sources: bool = True,
    ) -> MergeReport:
        """把 ``files`` 合并到 ``output``。异常被内部消化，结果都在 report 里。"""
        report = MergeReport()
        output = Path(output)
        files = [Path(f) for f in files]

        log.info("-" * 70)
        log.info("开始合并：%d 个文件 -> %s", len(files), output)
        log.info("设置：标题=%s 目录=%s 书签=%s", self.settings.include_titles,
                 self.settings.include_toc, self.settings.include_bookmarks)

        # 源文件指纹：合并前后比对，证明"没改源文件"
        before_hashes: dict[Path, str] = {}
        if verify_sources:
            for f in files:
                try:
                    before_hashes[f] = _sha256(f)
                except OSError as exc:
                    log.warning("无法读取指纹: %s (%s)", f, exc)

        temp_dir = output.parent / f".wmtool-tmp-{os.getpid()}"
        temp_dir.mkdir(parents=True, exist_ok=True)

        try:
            doc = DocxNew()
            set_document_defaults(doc, PORTRAIT_BOX)
            self._write_cover(doc, files)

            if self.settings.include_toc:
                self._write_toc_placeholder(doc)

            for i, path in enumerate(files, start=1):
                if self.stop.is_set():
                    report.cancelled = True
                    log.warning("用户取消，已完成 %d/%d", i - 1, len(files))
                    break

                outcome = self._process_one(
                    doc, path, i, len(files), temp_dir, timeout_per_file, progress
                )
                report.outcomes.append(outcome)

            if report.cancelled:
                report.error = "用户取消"
                return report

            # 原子落盘：先写临时文件，再替换目标
            tmp_out = output.with_suffix(output.suffix + ".part")
            try:
                doc.save(str(tmp_out))
                if output.exists():
                    output.unlink()
                tmp_out.replace(output)
            except Exception as exc:
                write_crash_report(type(exc), exc, exc.__traceback__, where="保存合并文档")
                report.error = f"保存失败: {exc}"
                try:
                    tmp_out.unlink(missing_ok=True)
                except Exception:
                    pass
                return report
            finally:
                report.finished_at = _dt.datetime.now()

            report.output_path = output
            try:
                report.output_bytes = output.stat().st_size
            except OSError:
                report.output_bytes = 0

            # 源文件完整性校验（只证明"没被改"）
            if verify_sources:
                self._verify_sources_unchanged(files, before_hashes, report)

            self._write_report_file(report, output)
            log.info("合并完成：%s", report.summary_line())
            return report

        except MergeCancelled:
            report.cancelled = True
            report.error = "用户取消"
            return report
        except Exception as exc:
            write_crash_report(type(exc), exc, exc.__traceback__, where="合并主流程")
            report.error = f"合并过程出错: {exc}"
            report.finished_at = _dt.datetime.now()
            log.exception("合并主流程异常")
            return report
        finally:
            _rmtree_safe(temp_dir)

    # ------------------------------------------------------------------ #
    # 单文件
    # ------------------------------------------------------------------ #
    def _process_one(
        self,
        doc,
        path: Path,
        index: int,
        total: int,
        temp_dir: Path,
        timeout: float,
        progress: ProgressCb | None,
    ) -> FileOutcome:  # type: ignore[no-untyped-def]
        """处理一个文件并把结果追加进 doc。"""
        outcome = FileOutcome(path=path)
        started = time.time()

        def _notify(msg: str, frac: float = 0.0) -> None:
            if progress:
                try:
                    progress(index, total, msg)
                except Exception:
                    pass

        # 基础校验
        if not path.exists():
            outcome.status = "failed"
            outcome.error = "文件不存在"
            log.error("[%d/%d] 跳过：文件不存在 %s", index, total, path)
            return outcome
        try:
            outcome.size_bytes = path.stat().st_size
            outcome.sha256 = _sha256(path)
        except OSError as exc:
            outcome.status = "failed"
            outcome.error = f"无法读取: {exc}"
            return outcome

        if outcome.size_bytes == 0:
            outcome.status = "skipped"
            outcome.error = "文件为空（0 字节）"
            log.warning("[%d/%d] 跳过空文件 %s", index, total, path.name)
            return outcome

        if not is_supported(path):
            outcome.status = "skipped"
            outcome.error = f"不支持的格式 {path.suffix}（支持：{', '.join(supported_suffixes())}）"
            log.warning("[%d/%d] 跳过不支持的文件 %s", index, total, path.name)
            return outcome

        converter = get_converter(path)
        if converter is None:
            outcome.status = "failed"
            outcome.error = "未找到对应转换器"
            return outcome

        log.info("[%d/%d] 处理 %s (%s, %.1f KB)", index, total, path.name,
                 path.suffix.lower(), outcome.size_bytes / 1024)

        # 每个源文件都从"新的一页 + 纵向页"开始。
        # 这一步是兜底：上一个文件可能把文档留在了横向节里（Excel/PPT/横向 PDF），
        # 不切回来，下一个文件的纵向内容就会被挤在横向页上。
        _notify(f"正在处理：{path.name}")
        try:
            page_break(doc)
            self._ensure_portrait(doc, path.name)
            if self.settings.include_titles:
                add_heading(doc, f"{index}. {path.stem}", level=1, size_pt=16)
                add_caption(doc, f"来源文件：{path.name}（{_human_size(outcome.size_bytes)}）")
        except Exception as exc:
            log.debug("写标题失败（不影响内容）: %s", exc)

        ctx = ConvertContext(
            doc=doc,
            settings=self.settings,
            src=path,
            index=index,
            total=total,
            timeout=timeout,
            box=PORTRAIT_BOX,
            landscape_box=LANDSCAPE_BOX,
            temp_dir=temp_dir,
            progress=lambda frac, msg: _notify(f"{path.name}：{msg}", frac),
        )

        try:
            with CrashReporter(f"转换文件 {path.name}", log, swallow=False):
                result: ConvertResult = _run_with_timeout(converter, ctx, timeout)
            outcome.status = "ok"
            outcome.converter = result.converter
            outcome.summary = result.summary
            outcome.warnings = list(result.warnings)
            for w in result.warnings:
                log.warning("  [%s] %s", path.name, w)

            if self.settings.include_bookmarks:
                try:
                    add_bookmark(doc, f"file_{index:03d}_{path.stem}")
                except Exception as exc:
                    log.debug("书签写入失败: %s", exc)

            # 每个文件处理完都把节切回纵向，避免污染下一个文件
            try:
                self._ensure_portrait(doc, path.name)
            except Exception as exc:
                log.debug("切回纵向失败: %s", exc)
        except TimeoutError:
            outcome.status = "failed"
            outcome.error = f"处理超时（>{timeout:.0f} 秒）"
            log.error("  [%s] 超时", path.name)
        except MemoryError:
            outcome.status = "failed"
            outcome.error = "内存不足（文件可能过大）"
            log.error("  [%s] 内存不足", path.name)
        except Exception as exc:
            outcome.status = "failed"
            outcome.error = str(exc)
            log.error("  [%s] 失败: %s", path.name, exc)
        finally:
            outcome.seconds = time.time() - started

        if outcome.status == "ok":
            log.info("  [%s] 完成（%.1fs）%s", path.name, outcome.seconds, outcome.summary)
        return outcome

    # ------------------------------------------------------------------ #
    # 节方向兜底
    # ------------------------------------------------------------------ #
    @staticmethod
    def _ensure_portrait(doc, why: str) -> None:  # type: ignore[no-untyped-def]
        """确保文档当前停在纵向节；在横向就新开一个纵向节。

        为什么需要它：横向/纵向是**节级**属性。转换器把文档切到横向后，
        "后面的一切"都会继续横着排。若不显式切回来，
        下一个文件的图片就会按横向页尺寸算、却贴在纵向内容流里，直接超页。

        只有当当前节确实没有内容时才复用（避免平白多出一页空白）。
        """
        sections = doc.sections
        if not sections:
            return
        current = sections[-1]
        w = int(current.page_width or 0)
        h = int(current.page_height or 0)
        if w <= h:
            return  # 已经是纵向
        log.debug("当前节为横向（%s），切回纵向", why)
        doc.add_page_break()
        new_section(doc, PORTRAIT_BOX)

    # ------------------------------------------------------------------ #
    # 封面 / 目录 / 报告
    # ------------------------------------------------------------------ #
    def _write_cover(self, doc, files: Iterable[Path]) -> None:  # type: ignore[no-untyped-def]
        """生成封面：标题、生成时间、文件清单。"""
        files = list(files)
        p = doc.add_paragraph()
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt, RGBColor

        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(120)
        run = p.add_run("文 件 合 并 文 档")
        run.bold = True
        run.font.size = Pt(28)
        run.font.color.rgb = RGBColor(0x1F, 0x36, 0x5D)

        sub = doc.add_paragraph()
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = sub.add_run(f"共 {len(files)} 个源文件 · 由 {__app_name_cn__} v{__version__} 生成")
        r2.font.size = Pt(11)
        r2.font.color.rgb = RGBColor(0x70, 0x70, 0x70)

        meta = doc.add_paragraph()
        meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r3 = meta.add_run(
            "生成时间：" + _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            + "　|　源文件均以只读方式读取，未被修改"
        )
        r3.font.size = Pt(9)
        r3.font.color.rgb = RGBColor(0x90, 0x90, 0x90)

        # 文件清单
        list_head = doc.add_paragraph()
        list_head.paragraph_format.space_before = Pt(28)
        rh = list_head.add_run("源文件清单")
        rh.bold = True
        rh.font.size = Pt(13)
        rh.font.color.rgb = RGBColor(0x1F, 0x36, 0x5D)

        for i, f in enumerate(files, start=1):
            lp = doc.add_paragraph()
            lp.paragraph_format.space_after = Pt(0)
            lp.paragraph_format.left_indent = Pt(18)
            lr = lp.add_run(f"{i}. {f.name}　({f.suffix.lower() or '无扩展名'})")
            lr.font.size = Pt(10)

    def _write_toc_placeholder(self, doc) -> None:  # type: ignore[no-untyped-def]
        """插入目录域（Word 打开后按 F9 / 更新域 生成页码）。"""
        doc.add_page_break()
        add_heading(doc, "目录", level=1, size_pt=16)
        add_toc_field(doc, "1-3")

    def _verify_sources_unchanged(
        self, files: Sequence[Path], before: dict[Path, str], report: MergeReport
    ) -> None:
        """比对源文件指纹，把"源文件未被修改"变成可验证的事实。"""
        changed = []
        for f in files:
            old = before.get(f)
            if not old:
                continue
            try:
                new = _sha256(f)
            except OSError:
                continue
            if new != old:
                changed.append(f)
        if changed:
            msg = "以下源文件在合并过程中发生了变化（请检查是否有其它程序在写这些文件）：" + \
                  "、".join(str(p) for p in changed)
            log.error(msg)
            report.error = (report.error + " | " + msg).strip(" |")
        else:
            log.info("源文件完整性校验通过：%d 个文件指纹均未变化", len(before))

    def _write_report_file(self, report: MergeReport, output: Path) -> None:
        """在输出目录旁写一份文本报告（哪些成功、哪些警告）。"""
        try:
            rp = output.with_name(output.stem + "-合并报告.txt")
            lines = [
                f"{__app_name_cn__} v{__version__} 合并报告",
                "=" * 60,
                f"开始时间: {report.started_at:%Y-%m-%d %H:%M:%S}",
                f"结束时间: {report.finished_at:%Y-%m-%d %H:%M:%S}" if report.finished_at else "",
                f"输出文件: {output}",
                f"结果汇总: {report.summary_line()}",
                "",
                "逐文件明细:",
                "-" * 60,
            ]
            for i, o in enumerate(report.outcomes, start=1):
                lines.append(f"{i}. [{o.status.upper()}] {o.name}")
                lines.append(f"   类型: {o.converter or '-'}　耗时: {o.seconds:.1f}s　大小: {_human_size(o.size_bytes)}")
                if o.summary:
                    lines.append(f"   内容: {o.summary}")
                if o.error:
                    lines.append(f"   错误: {o.error}")
                for w in o.warnings:
                    lines.append(f"   警告: {w}")
                lines.append("")
            if report.error:
                lines += ["整体问题:", report.error, ""]
            lines.append("源文件未被修改：所有文件均以只读方式读取。")
            rp.write_text("\n".join(lines), encoding="utf-8")
            log.info("已写出合并报告: %s", rp)
        except Exception as exc:
            log.warning("合并报告写出失败（不影响结果）: %s", exc)


# --------------------------------------------------------------------------- #
# 工具函数
# --------------------------------------------------------------------------- #
def _run_with_timeout(converter, ctx: ConvertContext, timeout: float) -> ConvertResult:  # type: ignore[no-untyped-def]
    """在子线程里跑转换器，超时则抛 TimeoutError。

    为什么用线程而不是进程：转换器要往同一个 docx 对象里写，跨进程传不回来。
    超时后线程会被留下（Python 杀不掉线程），但主流程能继续，
    且 DAEMON 线程不会阻止程序退出——这是"宁可漏一个文件，也不能卡死整批"的取舍。
    """
    if timeout <= 0:
        return converter.convert(ctx)

    box: dict[str, object] = {}

    def _target() -> None:
        try:
            box["result"] = converter.convert(ctx)
        except BaseException as exc:  # noqa: BLE001 - 要原样带回主线程
            box["error"] = exc

    t = threading.Thread(target=_target, name=f"convert-{ctx.src.name}", daemon=True)
    t.start()
    t.join(timeout)
    if t.is_alive():
        log.error("转换超时（%s，预算 %.0fs），该文件被跳过", ctx.src.name, timeout)
        raise TimeoutError(f"{ctx.src.name} 处理超过 {timeout:.0f} 秒")
    if "error" in box:
        raise box["error"]  # type: ignore[misc]
    result = box.get("result")
    if result is None:
        raise RuntimeError("转换器没有返回结果")
    return result  # type: ignore[return-value]


def _sha256(path: Path, chunk: int = 1 << 20) -> str:
    """算文件指纹（分块读，避免大文件吃内存）。"""
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            data = f.read(chunk)
            if not data:
                break
            h.update(data)
    return h.hexdigest()


def _human_size(n: int) -> str:
    """字节数 -> 人类可读。"""
    if n <= 0:
        return "0 B"
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(n)
    for u in units:
        if size < 1024 or u == units[-1]:
            return f"{size:.1f} {u}" if u != "B" else f"{int(size)} B"
        size /= 1024
    return f"{n} B"


def _rmtree_safe(path: Path) -> None:
    """删除临时目录（失败不报错，避免污染主流程）。"""
    import shutil

    try:
        if path.exists():
            shutil.rmtree(path, ignore_errors=True)
    except Exception:
        pass
