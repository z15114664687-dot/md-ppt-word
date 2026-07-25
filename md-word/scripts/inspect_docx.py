#!/usr/bin/env python3
"""Inspect DOCX equations, fields, fonts, visible branding, and media names."""

from __future__ import annotations

import argparse
import html
import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[2]
TOKENS_PATH = ROOT / "shared-assets" / "design-tokens.json"
INSTR_TEXT = re.compile(r"<w:instrText[^>]*>([^<]*)</w:instrText>")
BANNED_VISIBLE_TEXT = ("国金证券", "LOGO")


def _local_name(value: str) -> str:
    return value.rsplit("}", 1)[-1]


def _attribute(element: ET.Element, name: str) -> str | None:
    for key, value in element.attrib.items():
        if _local_name(key) == name:
            return value
    return None


def _field_instructions(xml: str) -> list[str]:
    return [html.unescape(match.group(1)) for match in INSTR_TEXT.finditer(xml)]


def _field_records(xml: str) -> list[dict[str, str]]:
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []
    records: list[dict[str, str]] = []
    current: dict[str, object] | None = None
    for element in root.iter():
        name = _local_name(element.tag)
        if name == "fldChar":
            kind = _attribute(element, "fldCharType")
            if kind == "begin":
                current = {"instruction": "", "result": [], "separated": False}
            elif kind == "separate" and current is not None:
                current["separated"] = True
            elif kind == "end" and current is not None:
                records.append(
                    {
                        "instruction": str(current["instruction"]),
                        "result": "\n".join(str(item) for item in current["result"]).strip(),
                    }
                )
                current = None
        elif name == "instrText" and current is not None and not current["separated"]:
            current["instruction"] = str(current["instruction"]) + (element.text or "")
        elif name == "t" and current is not None and current["separated"]:
            current["result"].append(element.text or "")
    return records


def _visible_text(xml: str) -> list[str]:
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []
    return [element.text or "" for element in root.iter() if _local_name(element.tag) == "t"]


def _settings_update_fields(xml: str) -> bool:
    if not xml:
        return False
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return False
    for element in root.iter():
        if _local_name(element.tag) == "updateFields":
            return (_attribute(element, "val") or "true").lower() in {"true", "1", "on"}
    return False


def inspect(path: Path, expected_math: int = 0, font_profile: str | None = None) -> dict[str, object]:
    path = Path(path)
    errors: list[str] = []
    with ZipFile(path) as archive:
        names = archive.namelist()
        document = archive.read("word/document.xml").decode("utf-8")
        footers = "".join(
            archive.read(name).decode("utf-8")
            for name in names
            if name.startswith("word/footer") and name.endswith(".xml")
        )
        visible_xml = [
            archive.read(name).decode("utf-8")
            for name in names
            if name.startswith("word/")
            and name.endswith(".xml")
            and any(
                name.startswith(prefix)
                for prefix in (
                    "word/document.xml",
                    "word/header",
                    "word/footer",
                    "word/footnotes",
                    "word/endnotes",
                    "word/comments",
                )
            )
        ]
        settings = archive.read("word/settings.xml").decode("utf-8") if "word/settings.xml" in names else ""
        styles = archive.read("word/styles.xml").decode("utf-8") if "word/styles.xml" in names else ""
        media = [name for name in names if name.startswith("word/media/")]

    records = _field_records(document)
    toc_records = [
        record
        for record in records
        if record["instruction"].lstrip().startswith("TOC") and "Exhibit Title" not in record["instruction"]
    ]
    exhibit_records = [record for record in records if "Exhibit Title" in record["instruction"]]
    instructions = _field_instructions(document)
    toc = bool(toc_records) or any(
        instruction.lstrip().startswith("TOC") and "Exhibit Title" not in instruction
        for instruction in instructions
    )
    exhibit_toc = bool(exhibit_records) or any("Exhibit Title" in instruction for instruction in instructions)
    toc_cached = any(record["result"] for record in toc_records)
    exhibit_toc_cached = any(record["result"] for record in exhibit_records)
    page_field = any(
        instruction.lstrip().startswith("PAGE") for instruction in _field_instructions(footers)
    )
    visible_text = "\n".join(text for xml in visible_xml for text in _visible_text(xml))
    has_exhibits = bool(re.search(r"图表\s*\d+\s*：", visible_text))
    update_fields = _settings_update_fields(settings)
    omml_count = len(re.findall(r"<m:oMath(?:\s|>)", document))

    if omml_count < expected_math:
        errors.append(f"OMML 公式数量不足：expected>={expected_math}, actual={omml_count}")
    if not toc:
        errors.append("缺少 Word TOC 域（内容目录）")
    elif not toc_cached:
        errors.append("内容目录域没有可见缓存结果")
    if has_exhibits and not exhibit_toc:
        errors.append("存在图表但缺少图表目录 TOC 域")
    elif exhibit_toc and not exhibit_toc_cached:
        errors.append("图表目录域没有可见缓存结果")
    if not page_field:
        errors.append("页脚缺少 PAGE 域")
    if not update_fields:
        errors.append("缺少 updateFields 设置，Microsoft Word 打开时不会自动刷新域")
    if re.search(r"\$\$|\\(?:frac|sum|int|begin)\b", visible_text):
        errors.append("DOCX 正文仍含未转换的 LaTeX 源码")
    for banned in BANNED_VISIBLE_TEXT:
        if re.search(re.escape(banned), visible_text, re.IGNORECASE):
            errors.append(f"检测到禁止的可见品牌文字：{banned}")
    for name in media:
        if "logo" in name.casefold():
            errors.append(f"媒体文件名包含可疑 logo 标识：{name}")

    expected_fonts: list[str] = []
    if font_profile is not None:
        tokens = json.loads(TOKENS_PATH.read_text(encoding="utf-8"))["word"]
        if font_profile not in tokens["font_profiles"]:
            raise ValueError(f"未知 Word 字体档位：{font_profile}")
        font_spec = tokens["font_profiles"][font_profile]
        expected_fonts = sorted(
            {
                value
                for key, value in font_spec.items()
                if key.endswith("_zh") or key.endswith("_latin")
            }
        )
        for family in expected_fonts:
            if family not in styles:
                errors.append(f"样式表缺少 {font_profile} 字体：{family}")

    return {
        "path": str(path),
        "font_profile": font_profile,
        "expected_fonts": expected_fonts,
        "omml_count": omml_count,
        "toc": toc,
        "toc_cached": toc_cached,
        "exhibit_toc": exhibit_toc,
        "exhibit_toc_cached": exhibit_toc_cached,
        "page_field": page_field,
        "update_fields": update_fields,
        "media_files": len(media),
        "errors": errors,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--expected-math", type=int, default=0)
    parser.add_argument("--font-profile")
    args = parser.parse_args()
    try:
        result = inspect(args.docx, args.expected_math, args.font_profile)
    except ValueError as error:
        print(json.dumps({"error": str(error)}, ensure_ascii=False, indent=2))
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if result["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
