#!/usr/bin/env python3
"""Static preflight for research-style Quarto PowerPoint source files.

Two layers of checks:
  1. Content contract — title length, list budget, body length, conclusion-led
     titles, source lines, image existence (original behaviour).
  2. Layout quality gate — the machine-checkable half of `references/slide-quality.md`
     (the slide-auditor checklist localised to QMD): top-level image/table not wrapped
     in `.columns`, content after a columns block, near-empty content pages, and
     manual font shrinking. These are the failure modes that make a Quarto deck look
     bare, and qmd-contract already forbids them — this gate stops them slipping through.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


FRONTMATTER = re.compile(r"\A---\s*\n(.*?)\n---\s*(?:\n|\Z)", re.DOTALL)
SLIDE_HEADING = re.compile(r"(?m)^##\s+(.+?)\s*$")
IMAGE = re.compile(r"!\[[^\]]*\]\(([^)\s]+)(?:\s+[^)]*)?\)")
SOURCE = re.compile(r"(?:资料)?来源\s*[：:]", re.IGNORECASE)
FENCE_START = re.compile(r"^[ \t]{0,3}(`{3,}|~{3,})(?:[^`]*)$")
COLUMNS_OPEN = re.compile(r"^(:{3,})\s*\{([^}]*)\}\s*$")
DIV_CLOSE = re.compile(r"^(:{3,})\s*$")
TABLE_ROW = re.compile(r"^\s*\|.*\|\s*$")
BULLET = re.compile(r"(?m)^\s*(?:[-*+] |\d+[.)]\s+)")
INLINE_SHRINK = re.compile(r"font-size|\.smaller\b|\\(?:small|footnotesize|scriptsize|tiny)\b")


def _strip_fenced_code(text: str) -> str:
    """Remove fenced-code contents while preserving line boundaries."""
    output: list[str] = []
    fence_char = ""
    fence_length = 0
    for line in text.splitlines(keepends=True):
        stripped = line.rstrip("\r\n")
        if not fence_char:
            match = FENCE_START.match(stripped)
            if match:
                fence = match.group(1)
                fence_char = fence[0]
                fence_length = len(fence)
                output.append("\n" if line.endswith("\n") else "")
                continue
            output.append(line)
            continue
        if re.match(rf"^[ \t]{{0,3}}{re.escape(fence_char)}{{{fence_length},}}[ \t]*$", stripped):
            fence_char = ""
            fence_length = 0
        output.append("\n" if line.endswith("\n") else "")
    return "".join(output)


def _format_is_pptx(frontmatter: str) -> bool:
    direct = re.search(r"(?m)^format\s*:\s*pptx\s*$", frontmatter)
    nested = re.search(r"(?ms)^format\s*:\s*\n(?:[ \t]+.*\n)*?[ \t]+pptx\s*:", frontmatter)
    return bool(direct or nested)


def _slides(body: str) -> list[tuple[str, str]]:
    headings = list(SLIDE_HEADING.finditer(body))
    return [
        (match.group(1).strip(), body[match.end() : headings[index + 1].start() if index + 1 < len(headings) else len(body)])
        for index, match in enumerate(headings)
    ]


def _div_spans(lines: list[str]) -> list[tuple[str, int, int]]:
    """Return (classes, start_line, end_line) for each fenced div, nesting-aware."""
    stack: list[tuple[str, int]] = []
    spans: list[tuple[str, int, int]] = []
    for index, line in enumerate(lines):
        stripped = line.strip()
        opener = COLUMNS_OPEN.match(stripped)
        if opener:
            stack.append((opener.group(2), index))
            continue
        if DIV_CLOSE.match(stripped) and stack:
            classes, start = stack.pop()
            spans.append((classes, start, index))
    return spans


def _layout_findings(number: int, title: str, content: str, has_card: bool = False) -> tuple[list[str], list[str]]:
    """Machine-checkable subset of references/slide-quality.md."""
    errors: list[str] = []
    warnings: list[str] = []
    lines = content.splitlines()
    spans = _div_spans(lines)

    def in_columns(idx: int) -> bool:
        return any("columns" in classes and start < idx <= end for classes, start, end in spans)

    has_columns = any("columns" in classes for classes, _, _ in spans)

    # 1) top-level image outside .columns — the #1 "bare slide" cause
    for match in IMAGE.finditer(content):
        if match.group(1).startswith("data:"):
            continue
        line_index = content[: match.start()].count("\n")
        if not in_columns(line_index):
            errors.append(
                f"第 {number} 页“{title}”有顶层图片未放进 .columns —— Pandoc 会让图片独占一页、"
                f"把来源挤到无标题延续页；请用 :::: {{.columns}} 组织（图一栏、要点一栏，来源写栏内最后一段）"
            )
            break

    # 2) top-level pipe table outside .columns
    for index, line in enumerate(lines):
        if TABLE_ROW.match(line) and not in_columns(index):
            warnings.append(
                f"第 {number} 页“{title}”有顶层表格未放进 .columns；宽表可能挤走来源行，"
                f"建议放进栏内或移入附录"
            )
            break

    # 3) top-level content after a columns block (qmd-contract hard rule)
    if has_columns:
        last_end = max(end for classes, _, end in spans if "columns" in classes)
        tail = [line for line in lines[last_end + 1 :] if line.strip()]
        if tail:
            errors.append(
                f"第 {number} 页“{title}”在 .columns 之后还有顶层内容（{tail[0].strip()[:20]}…）—— "
                f"columns 之后不得再跟顶层段落，来源行应写进栏内"
            )

    # 4) near-empty content page — the "section-y whitespace" look
    visible = re.sub(r"(?m)^:{3,}.*$", "", content)
    visible = SOURCE.sub("", visible)
    visible = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", visible)
    stripped = re.sub(r"\s+", "", visible)
    has_exhibit = bool(IMAGE.search(content) or any(TABLE_ROW.match(line) for line in lines))
    has_bullet = bool(BULLET.search(content))
    if not has_exhibit and not has_bullet and not has_card and len(stripped) < 12:
        warnings.append(
            f"第 {number} 页“{title}”内容近乎空白 —— 研报密度下建议补一个关键数字、一句核心判断"
            f"或一个展项，避免大片留白"
        )

    # 5) manual font shrinking instead of restructuring (title attrs like
    #    `## 标题 {.smaller}` are the most common page-level shrink)
    if INLINE_SHRINK.search(title) or INLINE_SHRINK.search(content):
        warnings.append(
            f"第 {number} 页“{title}”疑似手动缩小字号（font-size/.smaller）—— "
            f"溢出时优先按 删减 → 换版式 → 拆页 处理，不靠缩字塞入"
        )

    return errors, warnings


def validate(path: Path) -> dict[str, object]:
    path = Path(path)
    text = path.read_text(encoding="utf-8")
    errors: list[str] = []
    warnings: list[str] = []
    match = FRONTMATTER.search(text)
    if not match:
        errors.append("缺少 YAML front matter")
        frontmatter = ""
        body = text
    else:
        frontmatter = match.group(1)
        body = text[match.end() :]
    if not _format_is_pptx(frontmatter):
        errors.append("format 必须为 pptx")

    slides = _slides(_strip_fenced_code(body))
    raw_content = {title: content for title, content in _slides(body)}
    if not slides:
        errors.append("至少需要一个二级标题（##）作为内容页")

    for number, (title, content) in enumerate(slides, 1):
        # A `#` level-1 heading starts the next section — it is not part of this
        # `##` slide's body. Cut it off so section headers after a table/columns
        # page don't get mis-flagged as "content after columns".
        section_break = re.search(r"(?m)^#[ \t]", content)
        if section_break:
            content = content[: section_break.start()]
        if len(title) > 30:
            errors.append(f"第 {number} 页标题超过 30 字，页面标题字号下无法单行放下，应改写：{title}")
        bullets = re.findall(r"(?m)^\s*(?:[-*+] |\d+[.)]\s+)", content)
        if len(bullets) > 8:
            errors.append(f"第 {number} 页“{title}”有 {len(bullets)} 个列表项，超过 8 个（研报/咨询密度上限）")
        plain = re.sub(r"```.*?```", "", content, flags=re.DOTALL)
        if len(plain.strip()) > 1200:
            errors.append(f"第 {number} 页“{title}”正文超过 1200 个字符，应删减或拆页")
        if title in {"分析", "现状", "结论", "研究", "背景"}:
            warnings.append(f"第 {number} 页标题“{title}”不是结论式标题")
        has_exhibit = bool(IMAGE.search(content) or re.search(r"(?m)^\s*\|.*\|\s*$", content))
        has_data = bool(re.search(r"\d+(?:\.\d+)?\s*(?:%|亿元|亿|万|倍|bp|bps)\b", content, re.IGNORECASE))
        if (has_exhibit or has_data) and not SOURCE.search(content):
            errors.append(f"第 {number} 页“{title}”包含图表或数据但缺少来源")
        for resource in IMAGE.findall(content):
            if re.match(r"https?://", resource):
                continue
            if not (path.parent / resource).exists():
                errors.append(f"第 {number} 页图片不存在：{resource}")
        has_card = bool(re.search(r"```\{=ppt-(?:kpi|takeaway|cards|flow|compare)\}", raw_content.get(title, "")))
        layout_errors, layout_warnings = _layout_findings(number, title, content, has_card)
        errors.extend(layout_errors)
        warnings.extend(layout_warnings)

    return {"path": str(path), "slides": len(slides), "errors": errors, "warnings": warnings}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("qmd", type=Path)
    args = parser.parse_args()
    result = validate(args.qmd)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if result["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
