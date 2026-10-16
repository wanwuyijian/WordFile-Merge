# -*- coding: utf-8 -*-
"""可读的 Markdown 解析器（不引第三方依赖）。

只支持**合并工具真正需要**的那一小撮语法，够用且行为可预期：

* ``#`` ~ ``######`` 标题
* 段落、``---`` 分隔线
* ``-`` / ``*`` / ``+`` / ``1.`` 列表（支持一层缩进嵌套，简化为两级）
* ```` ``` ```` 围栏代码块（等宽字体保留换行）
* ``>`` 引用
* ``| a | b |`` 表格（带 ``|---|`` 分隔行）
* 行内 ``**粗体**``、``*斜体*``、`` `代码` ``、``[文字](链接)``

解析结果是**中间表示**（``Block`` 列表），由 :mod:`wmtool.converters.md` 渲染进 docx。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Literal

BlockKind = Literal["heading", "paragraph", "list", "code", "quote", "table", "rule"]

_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_ULIST_RE = re.compile(r"^(\s*)([-*+])\s+(.*)$")
_OLIST_RE = re.compile(r"^(\s*)(\d+)[.)]\s+(.*)$")
_RULE_RE = re.compile(r"^\s*([-*_])(\s*\1){2,}\s*$")
_FENCE_RE = re.compile(r"^\s*(```|~~~)\s*(\S*)\s*$")
_TABLE_SEP_RE = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{1,}:?\s*)*\|?\s*$")


@dataclass
class Block:
    """一个块级元素。"""

    kind: BlockKind
    text: str = ""
    level: int = 0
    ordered: bool = False
    items: list[tuple[int, str]] = field(default_factory=list)  # (缩进级别, 文本)
    rows: list[list[str]] = field(default_factory=list)
    language: str = ""


@dataclass
class Inline:
    """行内片段。"""

    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False
    link: str = ""


_INLINE_RE = re.compile(
    r"(\*\*.+?\*\*)"          # 粗体
    r"|(`[^`]+`)"             # 行内代码
    r"|(\[.+?\]\(.+?\))"      # 链接
    r"|(\*[^*\n]+\*)"         # 斜体
)


def parse_inline(text: str) -> list[Inline]:
    """把一行文本切成带格式的片段。"""
    out: list[Inline] = []
    pos = 0
    for m in _INLINE_RE.finditer(text):
        if m.start() > pos:
            out.append(Inline(text[pos:m.start()]))
        token = m.group(0)
        if token.startswith("**"):
            out.append(Inline(token[2:-2], bold=True))
        elif token.startswith("`"):
            out.append(Inline(token[1:-1], code=True))
        elif token.startswith("["):
            label, _, url = token[1:-1].partition("](")
            out.append(Inline(label, link=url.rstrip(")")))
        else:
            out.append(Inline(token[1:-1], italic=True))
        pos = m.end()
    if pos < len(text):
        out.append(Inline(text[pos:]))
    return out or [Inline("")]


def parse_markdown(text: str) -> list[Block]:
    """解析 Markdown 文本为 Block 列表。"""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks: list[Block] = []
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # 空行
        if not stripped:
            i += 1
            continue

        # 围栏代码块
        m = _FENCE_RE.match(line)
        if m:
            fence = m.group(1)
            lang = m.group(2)
            i += 1
            buf: list[str] = []
            while i < n and not lines[i].strip().startswith(fence):
                buf.append(lines[i])
                i += 1
            i += 1  # 跳过结束围栏
            blocks.append(Block(kind="code", text="\n".join(buf), language=lang))
            continue

        # 分隔线
        if _RULE_RE.match(line):
            blocks.append(Block(kind="rule"))
            i += 1
            continue

        # 标题
        m = _HEADING_RE.match(line)
        if m:
            blocks.append(Block(kind="heading", level=len(m.group(1)), text=m.group(2).strip()))
            i += 1
            continue

        # 引用（连续多行合并）
        if stripped.startswith(">"):
            buf = []
            while i < n and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip().lstrip(">").strip())
                i += 1
            blocks.append(Block(kind="quote", text=" ".join(buf)))
            continue

        # 表格：当前行含 | 且下一行是分隔行
        if "|" in line and i + 1 < n and _TABLE_SEP_RE.match(lines[i + 1]):
            header = _split_table_row(line)
            rows = [header]
            i += 2
            while i < n and "|" in lines[i] and lines[i].strip():
                rows.append(_split_table_row(lines[i]))
                i += 1
            blocks.append(Block(kind="table", rows=rows))
            continue

        # 列表（连续的列表项算一个块）
        if _ULIST_RE.match(line) or _OLIST_RE.match(line):
            ordered = bool(_OLIST_RE.match(line))
            items: list[tuple[int, str]] = []
            while i < n:
                cur = lines[i]
                mu = _ULIST_RE.match(cur)
                mo = _OLIST_RE.match(cur)
                if mu:
                    indent = len(mu.group(1)) // 2
                    items.append((indent, mu.group(3).strip()))
                elif mo:
                    indent = len(mo.group(1)) // 2
                    items.append((indent, mo.group(3).strip()))
                    ordered = True
                else:
                    break
                i += 1
            blocks.append(Block(kind="list", ordered=ordered, items=items))
            continue

        # 普通段落（连续非空行合并，Markdown 的软换行按空格处理）
        buf = [stripped]
        i += 1
        while i < n and lines[i].strip() and not _starts_new_block(lines[i]):
            buf.append(lines[i].strip())
            i += 1
        blocks.append(Block(kind="paragraph", text=" ".join(buf)))

    return blocks


def _starts_new_block(line: str) -> bool:
    """判断该行是否是新的块起始（用于段落收集时提前收尾）。"""
    s = line.strip()
    return bool(
        _HEADING_RE.match(line)
        or _FENCE_RE.match(line)
        or _RULE_RE.match(line)
        or _ULIST_RE.match(line)
        or _OLIST_RE.match(line)
        or s.startswith(">")
    )


def _split_table_row(line: str) -> list[str]:
    """切分表格行，处理首尾竖线。"""
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def plain_text(md_text: str) -> str:
    """把 Markdown 压成纯文本（用于生成摘要 / 反向搜索）。"""
    out: list[str] = []
    for b in parse_markdown(md_text):
        if b.kind in ("heading", "paragraph", "quote"):
            out.append(b.text)
        elif b.kind == "list":
            out.extend(t for _, t in b.items)
        elif b.kind == "code":
            out.append(b.text)
        elif b.kind == "table":
            for row in b.rows:
                out.append(" | ".join(row))
    return "\n".join(out)
