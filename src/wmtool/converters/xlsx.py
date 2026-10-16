# -*- coding: utf-8 -*-
"""Excel (.xlsx/.xlsm/.csv) -> Word。

对应需求第 4 条：**不损失内容**、**不出页面**。策略是"分块 + 自适应方向"：

1. 每个工作表先算"需要多宽"（按各列最长内容的估算像素宽 × 列数）；
2. 窄表 -> 纵向页；中等 -> 纵向页 + 压缩列宽；宽表 -> **自动切横版页**；
3. 极宽表 -> 横版页 + 缩小字号到可读下限，仍然列数过多时按行拆表，
   保证没有一列被裁掉（宁可多占几页，也不丢内容）；
4. 工作表名作为小标题，数据表首行作为跨页重复的表头。

.xlsx 用 openpyxl 读（data_only=True 取公式的**缓存值**，也就是 Excel 里看到的
数字，而不是公式文本；取不到缓存值才退回公式串）。.xls 需要 Microsoft Excel
或 LibreOffice 参与（另见 xls_legacy.py）。
"""

from __future__ import annotations

import csv
import datetime as _dt
from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Emu, Pt

from ..docx_util import (
    add_caption,
    add_heading,
    new_section,
    repeat_header_row,
    set_cell_text,
    set_row_height_pt,
    set_table_layout_fixed,
    style_data_table,
)
from ..geometry import PageBox, twip_to_emu
from ..logging_setup import get_logger
from .base import ConvertContext, ConvertResult

log = get_logger("conv.xlsx")

SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm", ".csv", ".tsv"}

#: 一个字符的估算宽度（EMU），10.5pt 微软雅黑下大约 7.5px ≈ 71437 EMU
CHAR_WIDTH_EMU = 71437
#: 列宽下限（EMU）：太窄看不清
MIN_COL_EMU = int(0.8 * 360000)      # 0.8cm
#: 列宽上限（EMU）：单列最宽到 4cm，再宽说明内容确实长，但那不是"需要横向页"的理由
MAX_COL_EMU = int(4.0 * 360000)      # 4cm
#: 单元格最长按多少"视觉字符"估算（避免一个超长备注把整表判成超宽）
MAX_CELL_CHARS = 40
#: 表格字号范围
FONT_MAX_PT = 10.0
FONT_MIN_PT = 5.5
#: 单个工作表最多渲染的行数（超出部分另起一页继续，不丢）
ROWS_PER_TABLE = 400


class XlsxConverter:
    """Excel 工作簿 -> Word 表格。"""

    name = "xlsx"
    display = "Excel"

    def convert(self, ctx: ConvertContext) -> ConvertResult:
        result = ConvertResult(converter=self.name)
        src = ctx.src
        suffix = src.suffix.lower()

        if suffix in (".csv", ".tsv"):
            sheets = [(src.stem, self._read_csv(src, delimiter="\t" if suffix == ".tsv" else None))]
        else:
            sheets = self._read_xlsx(src)

        if not sheets:
            raise ValueError("工作簿里没有可读取的数据")

        wide_seen = False
        for s_idx, (sheet_name, rows) in enumerate(sheets):
            if not rows:
                continue

            # 决定这一张表用纵向还是横向页。
            # 判据是"哪个方向能让字号更大"——横向可用宽度是纵向的 1.5 倍左右，
            # 所以只要纵向需要把字号压到比横向更小，就该转横版（需求第 4 条）。
            widths = self._estimate_widths(rows)
            natural = sum(widths)
            box, font_pt = self._choose_orientation(ctx, natural)
            usable = box.usable_width

            if box.landscape and not wide_seen:
                wide_seen = True
            if box.landscape:
                ctx.doc.add_page_break()
                new_section(ctx.doc, box)

            actual_widths = self._fit_widths(widths, usable)

            add_heading(ctx.doc, f"工作表：{sheet_name}", level=2, size_pt=13)
            n_rows_written = 0
            for chunk_start in range(0, len(rows), ROWS_PER_TABLE):
                chunk = rows[chunk_start:chunk_start + ROWS_PER_TABLE]
                if chunk_start > 0:
                    ctx.doc.add_page_break()
                    add_caption(
                        ctx.doc,
                        f"（{sheet_name} 接上页，第 {chunk_start + 1} 行起）",
                    )
                self._write_table(ctx, chunk, actual_widths, font_pt=font_pt)
                n_rows_written += len(chunk)

            result.table_count += 1
            result.warn(
                f"工作表「{sheet_name}」：{len(rows)} 行 × {len(widths)} 列，"
                f"字号 {font_pt:.1f}pt，{'横向' if box.landscape else '纵向'}排版"
            )
            ctx.report((s_idx + 1) / max(1, len(sheets)), f"已写入工作表 {sheet_name}")

        # 横向页用完，切回纵向，不影响后续文件
        if wide_seen:
            result.landscape = True
            ctx.doc.add_page_break()
            new_section(ctx.doc, ctx.box)

        total_rows = sum(len(r) for _, r in sheets)
        result.summary = f"Excel {len(sheets)} 个工作表 · {total_rows} 行"
        ctx.report(1.0, "Excel 已写入")
        return result

    # ------------------------------------------------------------------ #
    # 读取
    # ------------------------------------------------------------------ #
    def _read_xlsx(self, src: Path) -> list[tuple[str, list[list[str]]]]:
        """读 xlsx/xlsm，返回 [(表名, 二维文本)]。"""
        try:
            from openpyxl import load_workbook
        except ImportError as exc:  # pragma: no cover
            raise ValueError("缺少 openpyxl，无法读取 Excel 文件") from exc

        # data_only=True：取公式缓存值（Excel 里显示的数字）；读不到再退回公式
        values_wb = None
        formula_wb = None
        try:
            values_wb = load_workbook(str(src), data_only=True, read_only=True)
        except Exception as exc:
            raise ValueError(f"Excel 打开失败: {exc}") from exc

        out: list[tuple[str, list[list[str]]]] = []
        try:
            for ws in values_wb.worksheets:
                if ws.sheet_state != "visible":
                    continue
                rows = self._dump_sheet(ws)
                # 整表都是 None（说明是纯公式表且没缓存值）-> 用公式文本兜底
                if rows and all(all(c == "" for c in r) for r in rows):
                    try:
                        formula_wb = formula_wb or load_workbook(
                            str(src), data_only=False, read_only=True
                        )
                        rows = self._dump_sheet(formula_wb[ws.title])
                    except Exception:
                        pass
                rows = self._trim(rows)
                if rows:
                    out.append((ws.title, rows))
        finally:
            for wb in (values_wb, formula_wb):
                try:
                    if wb is not None:
                        wb.close()
                except Exception:
                    pass
        return out

    @staticmethod
    def _dump_sheet(ws) -> list[list[str]]:  # type: ignore[no-untyped-def]
        """把工作表读成二维字符串。"""
        rows: list[list[str]] = []
        for row in ws.iter_rows(values_only=True):
            rows.append([XlsxConverter._cell_to_text(v) for v in row])
        return rows

    @staticmethod
    def _read_csv(src: Path, delimiter: str | None = None) -> list[list[str]]:
        """读 CSV/TSV（自动猜编码：utf-8-sig -> gbk -> latin-1）。"""
        for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030", "latin-1"):
            try:
                with src.open("r", encoding=enc, newline="") as f:
                    sample = f.read(8192)
                    f.seek(0)
                    if delimiter is None:
                        try:
                            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                        except Exception:
                            dialect = csv.excel
                        reader = csv.reader(f, dialect)
                    else:
                        reader = csv.reader(f, delimiter=delimiter)
                    return [[c.strip() for c in row] for row in reader]
            except UnicodeDecodeError:
                continue
            except Exception as exc:
                raise ValueError(f"CSV 读取失败: {exc}") from exc
        raise ValueError("CSV 编码无法识别（已尝试 utf-8 / gbk / latin-1）")

    @staticmethod
    def _cell_to_text(value) -> str:  # type: ignore[no-untyped-def]
        """单元格值 -> 文本（日期/时间格式化，浮点去掉多余小数）。"""
        if value is None:
            return ""
        if isinstance(value, _dt.datetime):
            if value.hour == value.minute == value.second == 0:
                return value.strftime("%Y-%m-%d")
            return value.strftime("%Y-%m-%d %H:%M:%S")
        if isinstance(value, _dt.date):
            return value.strftime("%Y-%m-%d")
        if isinstance(value, _dt.time):
            return value.strftime("%H:%M:%S")
        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return f"{value:g}"
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        text = str(value).strip()
        # 超长文本在单元格里截断显示？不——需求要求不丢内容，这里保留全量
        return text

    @staticmethod
    def _trim(rows: list[list[str]]) -> list[list[str]]:
        """去掉尾部全空行与全空列（前导空行保留，因为它可能是缩进布局）。"""
        while rows and all(c == "" for c in rows[-1]):
            rows.pop()
        if not rows:
            return rows
        max_cols = max((len(r) for r in rows), default=0)
        while max_cols > 0 and all(
            (len(r) < max_cols or r[max_cols - 1] == "") for r in rows
        ):
            max_cols -= 1
        return [r[:max_cols] for r in rows]

    # ------------------------------------------------------------------ #
    # 排版
    # ------------------------------------------------------------------ #
    @staticmethod
    def _choose_orientation(ctx: ConvertContext, natural: int) -> tuple[PageBox, float]:
        """选纵向还是横向页，并给出该方向下需要的字号。

        规则（按优先级）：
        1. 纵向放得下且不用缩字号 -> 纵向（最自然，大多数表都是这种）
        2. 比较两个方向各自能给出的字号，谁大用谁
        3. **两个方向都被压到最低字号时，选横向** —— 此时字号已经到顶了，
           横向能给出更多可用宽度，列宽更宽裕，可读性严格更好

        第 3 条是关键：如果只看字号，25 列和 40 列这种超宽表在横竖两边
        都是"最低字号"，比较不出差别，就会错误地留在纵向。
        """
        portrait_usable = ctx.box.usable_width
        landscape_usable = ctx.landscape_box.usable_width

        def font_for(usable: int) -> float:
            if natural <= usable:
                return FONT_MAX_PT
            return max(FONT_MIN_PT, FONT_MAX_PT * usable / natural)

        p_font = font_for(portrait_usable)
        l_font = font_for(landscape_usable)

        # 1. 纵向够用就不折腾
        if p_font >= FONT_MAX_PT:
            return ctx.box, FONT_MAX_PT

        # 3. 两边都触底：横向必然更好（可用宽度 1.5 倍），直接横版
        if p_font <= FONT_MIN_PT and l_font <= FONT_MIN_PT:
            return ctx.landscape_box, l_font

        # 2. 横向要明显更好才转，避免 1.15 倍以内的来回横竖跳动
        if l_font > p_font * 1.15:
            return ctx.landscape_box, l_font
        return ctx.box, p_font

    @staticmethod
    def _estimate_widths(rows: list[list[str]]) -> list[int]:
        """按每列最长内容估算自然列宽（EMU）。

        注意这里**下限是 MIN_COL_EMU、上限是 MAX_COL_EMU**，但估算值本身
        不受可用宽度约束——它代表"这张表最舒服需要多宽"。
        真正压缩到页面里是后面 :meth:`_fit_widths` 的事。
        如果在这里就把宽度卡死，方向判断（哪边字号更大）会永远得不出结论。
        """
        n_cols = max((len(r) for r in rows), default=0)
        widths = [MIN_COL_EMU] * n_cols
        for row in rows:
            for i, cell in enumerate(row):
                if i >= n_cols:
                    continue
                # 中文按 2 个字符宽算
                text = cell or ""
                visual = sum(2 if ord(ch) > 0x2E80 else 1 for ch in text)
                visual = min(visual, MAX_CELL_CHARS)
                need = int(visual * CHAR_WIDTH_EMU) + 6 * CHAR_WIDTH_EMU
                widths[i] = max(widths[i], min(need, MAX_COL_EMU))
        return widths

    @staticmethod
    def _fit_widths(widths: list[int], usable: int) -> list[int]:
        """把估算列宽等比压到可用宽度内（压不动就保持，交给分表兜底）。"""
        total = sum(widths)
        if total <= usable:
            return widths
        scale = usable / total
        fitted = [max(MIN_COL_EMU // 2, int(w * scale)) for w in widths]
        # 缩放后仍超宽（列太多，每列都触底）时按可用宽度均分
        if sum(fitted) > usable:
            each = max(1, usable // len(widths))
            fitted = [each] * len(widths)
        return fitted

    def _write_table(
        self,
        ctx: ConvertContext,
        rows: list[list[str]],
        widths: list[int],
        *,
        font_pt: float,
    ) -> None:
        """把二维数据写成一个 Word 表格。"""
        n_rows = len(rows)
        n_cols = max((len(r) for r in rows), default=0)
        if n_rows == 0 or n_cols == 0:
            return

        table = ctx.doc.add_table(rows=n_rows, cols=n_cols)
        style_data_table(table, font_pt=font_pt)
        set_table_layout_fixed(table)

        for r_idx, row in enumerate(rows):
            for c_idx in range(n_cols):
                text = row[c_idx] if c_idx < len(row) else ""
                cell = table.cell(r_idx, c_idx)
                align = None
                if c_idx > 0 and text:
                    # 数字右对齐，读起来更像表格
                    try:
                        float(text.replace(",", "").replace("%", ""))
                        align = WD_ALIGN_PARAGRAPH.RIGHT
                    except ValueError:
                        align = None
                set_cell_text(
                    cell, text, bold=(r_idx == 0), size_pt=font_pt, align=align
                )

        for c_idx, w in enumerate(widths[:n_cols]):
            for row in table.rows:
                try:
                    row.cells[c_idx].width = Emu(w)
                except Exception:
                    pass
            try:
                table.columns[c_idx].width = Emu(w)
            except Exception:
                pass

        repeat_header_row(table)
        ctx.doc.add_paragraph()
