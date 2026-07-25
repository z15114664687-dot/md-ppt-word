#!/usr/bin/env python3
"""Inspect PPTX canvas, master geometry, fonts, and visible branding."""

from __future__ import annotations

import argparse
import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[2]
TOKENS_PATH = ROOT / "shared-assets" / "design-tokens.json"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU_PER_INCH = 914400
EMU_PER_POINT = 12700
BANNED_VISIBLE_TEXT = ("国金证券", "LOGO")


def p(tag: str) -> str:
    return f"{{{P_NS}}}{tag}"


def a(tag: str) -> str:
    return f"{{{A_NS}}}{tag}"


def inspect(path: Path, font_profile: str | None = None) -> dict[str, object]:
    path = Path(path)
    tokens = json.loads(TOKENS_PATH.read_text(encoding="utf-8"))["ppt"]
    font_profile = font_profile or tokens["default_font_profile"]
    if font_profile not in tokens["font_profiles"]:
        raise ValueError(f"未知 PPT 字体档位：{font_profile}")
    fonts = tokens["font_profiles"][font_profile]
    errors: list[str] = []
    visible_text: list[str] = []
    title_offsets: list[int] = []
    slide_title_offsets: list[dict[str, int | str]] = []
    minimum_y = round(tokens["logo_slot"]["content_top_min_pt"] * EMU_PER_POINT)

    with ZipFile(path) as archive:
        names = archive.namelist()
        presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
        slide_size = presentation.find(p("sldSz"))
        width = int(slide_size.get("cx", "0")) if slide_size is not None else 0
        height = int(slide_size.get("cy", "0")) if slide_size is not None else 0
        expected_width = round(tokens["canvas"]["width_in"] * EMU_PER_INCH)
        expected_height = round(tokens["canvas"]["height_in"] * EMU_PER_INCH)
        if (width, height) != (expected_width, expected_height):
            errors.append(
                f"画布尺寸不符合 tokens：actual={width}x{height}, expected={expected_width}x{expected_height}"
            )

        theme_xml = "".join(
            archive.read(name).decode("utf-8")
            for name in names
            if name.startswith("ppt/theme/theme") and name.endswith(".xml")
        )
        layout_xml = "".join(
            archive.read(name).decode("utf-8")
            for name in names
            if name.startswith("ppt/slideLayouts/slideLayout") and name.endswith(".xml")
        )
        for role in ("zh", "latin"):
            if fonts[role] not in theme_xml:
                errors.append(f"主题缺少 {font_profile} 字体：{fonts[role]}")
        hans_faces: list[str] = []
        for name in names:
            if not name.startswith("ppt/theme/theme") or not name.endswith(".xml"):
                continue
            root = ET.fromstring(archive.read(name))
            hans_faces.extend(
                element.get("typeface", "")
                for element in root.iter(a("font"))
                if element.get("script") == "Hans"
            )
        if not hans_faces or any(face != fonts["zh"] for face in hans_faces):
            errors.append(
                f"主题 Hans 字体映射不符合 {font_profile}："
                f"actual={hans_faces}, expected={fonts['zh']}"
            )
        if fonts["display_zh"] not in layout_xml:
            errors.append(f"版式缺少封面/章节字体：{fonts['display_zh']}")

        geometry_names = [
            name
            for name in names
            if (name.startswith("ppt/slideMasters/slideMaster") or name.startswith("ppt/slideLayouts/slideLayout"))
            and name.endswith(".xml")
        ]
        for name in geometry_names:
            root = ET.fromstring(archive.read(name))
            for shape in root.iter(p("sp")):
                placeholder = shape.find(f".//{p('ph')}")
                if placeholder is None or placeholder.get("type") != "title":
                    continue
                offset = shape.find(f".//{a('xfrm')}/{a('off')}")
                if offset is not None:
                    title_offsets.append(int(offset.get("y", "0")))

        for name in names:
            if not re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
                continue
            root = ET.fromstring(archive.read(name))
            for shape in root.iter(p("sp")):
                placeholder = shape.find(f".//{p('ph')}")
                if placeholder is None or placeholder.get("type") != "title":
                    continue
                # 没有显式 xfrm 的标题继承已检查的 layout/master；
                # 有 xfrm 时必须额外验证，防止渲染器覆盖母版安全区。
                offset = shape.find(f"{p('spPr')}/{a('xfrm')}/{a('off')}")
                if offset is not None:
                    slide_title_offsets.append({"slide": name, "y": int(offset.get("y", "0"))})

        for name in names:
            if not name.startswith("ppt/") or not name.endswith(".xml"):
                continue
            try:
                root = ET.fromstring(archive.read(name))
            except ET.ParseError:
                continue
            visible_text.extend(element.text or "" for element in root.iter(a("t")))

    if not title_offsets:
        errors.append("母版中未找到普通标题占位符")
    elif any(y < minimum_y for y in title_offsets):
        errors.append(f"普通标题进入 logo 安全区：minimum={minimum_y}, actual={title_offsets}")
    unsafe_slide_titles = [item for item in slide_title_offsets if item["y"] < minimum_y]
    if unsafe_slide_titles:
        errors.append(
            f"最终幻灯片标题进入 logo 安全区："
            f"minimum={minimum_y}, actual={unsafe_slide_titles}"
        )

    joined_text = "\n".join(visible_text)
    for banned in BANNED_VISIBLE_TEXT:
        if re.search(re.escape(banned), joined_text, re.IGNORECASE):
            errors.append(f"检测到禁止的可见品牌文字：{banned}")

    return {
        "path": str(path),
        "font_profile": font_profile,
        "canvas_emu": {"width": width, "height": height},
        "title_offsets_emu": title_offsets,
        "slide_title_offsets_emu": slide_title_offsets,
        "visible_text_items": len(visible_text),
        "errors": errors,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--font-profile")
    args = parser.parse_args()
    try:
        result = inspect(args.pptx, args.font_profile)
    except (ValueError, KeyError) as error:
        print(json.dumps({"error": str(error)}, ensure_ascii=False, indent=2))
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if result["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
