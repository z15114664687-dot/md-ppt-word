#!/usr/bin/env python3
"""Validate the semantic contract for a fixed-format research report."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


IMAGE = re.compile(r"!\[[^\]]*\]\(([^)\s]+)(?:\s+[^)]*)?\)")
TABLE = re.compile(r"(?m)(?:^\s*\|.*\|\s*\n){2,}")
SOURCE = re.compile(r"(?:资料)?来源\s*[：:]")
# 国金参考稿：图和表共用一套 "图表N：题" 编号，全角冒号。
EXHIBIT_TITLE = re.compile(r"(?m)^图表\s*(\d+)\s*：\s*.+$")
# 旧式分开编号（图 1 / 表 1）或半角冒号，一律要求改写。
EXHIBIT_TITLE_MALFORMED = re.compile(r"(?m)^(?:[图表]\s*\d+\s+\S.*|图表\s*\d+\s*:.*)$")


def _math_counts(text: str) -> tuple[int, bool]:
    display_tokens = re.findall(r"(?<!\\)\$\$", text)
    display_unbalanced = len(display_tokens) % 2 != 0
    displays = re.findall(r"(?s)(?<!\\)\$\$(.+?)(?<!\\)\$\$", text)
    without_displays = re.sub(r"(?s)(?<!\\)\$\$.*?(?<!\\)\$\$", "", text)
    inline_tokens = re.findall(r"(?<!\\)(?<!\$)\$(?!\$)", without_displays)
    inline_unbalanced = len(inline_tokens) % 2 != 0
    inlines = re.findall(r"(?s)(?<!\\)(?<!\$)\$(?!\$)(.+?)(?<!\\)\$(?!\$)", without_displays)
    return len(displays) + len(inlines), display_unbalanced or inline_unbalanced


def validate(path: Path) -> dict[str, object]:
    path = Path(path)
    text = path.read_text(encoding="utf-8")
    errors: list[str] = []
    warnings: list[str] = []

    headings = [match.group(1).strip() for match in re.finditer(r"(?m)^#{1,6}\s+(.+?)\s*$", text)]
    for required in ("摘要", "方法", "回测", "结论", "附录"):
        if not any(heading == required or heading.startswith(required + " ") for heading in headings):
            errors.append(f"缺少必需章节：{required}")

    math_expressions, unbalanced = _math_counts(text)
    if unbalanced:
        errors.append("公式分隔符不成对")
    if re.search(r"\\(?:newcommand|renewcommand|begin\{align\*?\}|tag)\b", text):
        warnings.append("检测到可能不受支持的 LaTeX 宏或环境，应先改写为标准公式")

    exhibit_numbers = [int(match.group(1)) for match in EXHIBIT_TITLE.finditer(text)]
    if exhibit_numbers != list(range(1, len(exhibit_numbers) + 1)):
        errors.append(f"图表编号必须从 1 起连续递增（图表1：、图表2：…），当前为 {exhibit_numbers}")
    for match in EXHIBIT_TITLE_MALFORMED.finditer(text):
        errors.append(f"题注格式应为“图表N：题”（图表统一编号、全角冒号）：{match.group(0).strip()}")

    for match in IMAGE.finditer(text):
        resource = match.group(1)
        if Path(resource).suffix.lower() in {".gif", ".webp"}:
            errors.append(f"图片需先转为 PNG/JPEG：{resource}")
        if not re.match(r"https?://", resource) and not (path.parent / resource).exists():
            errors.append(f"图片不存在：{resource}")
        before = text[max(0, match.start() - 240) : match.start()]
        after = text[match.end() : min(len(text), match.end() + 360)]
        if not EXHIBIT_TITLE.search(before):
            errors.append(f"图片缺少上方“图表N：”题注：{resource}")
        if not SOURCE.search(after):
            errors.append(f"图片缺少来源：{resource}")

    for index, match in enumerate(TABLE.finditer(text), 1):
        before = text[max(0, match.start() - 240) : match.start()]
        after = text[match.end() : min(len(text), match.end() + 360)]
        if not EXHIBIT_TITLE.search(before):
            errors.append(f"第 {index} 个表格缺少上方“图表N：”题注")
        if not SOURCE.search(after):
            errors.append(f"第 {index} 个表格缺少来源")

    images = len(IMAGE.findall(text))
    tables = len(TABLE.findall(text))
    if exhibit_numbers and len(exhibit_numbers) != images + tables:
        warnings.append(f"题注数（{len(exhibit_numbers)}）与图表对象数（{images + tables}）不一致，请检查是否有孤立题注")

    return {
        "path": str(path),
        "headings": len(headings),
        "images": images,
        "tables": tables,
        "exhibits": len(exhibit_numbers),
        "math_expressions": math_expressions,
        "errors": errors,
        "warnings": warnings,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("markdown", type=Path)
    args = parser.parse_args()
    result = validate(args.markdown)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if result["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
