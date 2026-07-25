#!/usr/bin/env python3
"""Build the 16:9 research reference PPTX master used by Quarto/Pandoc.

Starts from Pandoc's built-in reference.pptx (so every layout Pandoc expects
exists), then rewrites theme fonts, colors, and text styles from
shared-assets/design-tokens.json. No third-party branding is copied in.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
TOKENS_PATH = ROOT / "shared-assets" / "design-tokens.json"
DEFAULT_OUTPUT = ROOT / "shared-assets" / "ppt" / "research-reference.pptx"

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
ET.register_namespace("a", A_NS)
ET.register_namespace("p", P_NS)
ET.register_namespace(
    "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)


def a(tag: str) -> str:
    return f"{{{A_NS}}}{tag}"


def p(tag: str) -> str:
    return f"{{{P_NS}}}{tag}"


def load_ppt_tokens() -> dict:
    return json.loads(TOKENS_PATH.read_text(encoding="utf-8"))["ppt"]


def sz(points: float) -> str:
    return str(round(points * 100))


def pandoc_command() -> list[str]:
    if shutil.which("pandoc"):
        return [shutil.which("pandoc") or "pandoc"]
    if shutil.which("quarto"):
        return [shutil.which("quarto") or "quarto", "pandoc"]
    raise RuntimeError("缺少 Pandoc/Quarto：无法获取内置 reference.pptx 底版。安装方式见仓库 README.md")


def fetch_base_template(target: Path) -> None:
    data = subprocess.run(
        [*pandoc_command(), "--print-default-data-file", "reference.pptx"],
        check=True,
        capture_output=True,
    ).stdout
    target.write_bytes(data)


EMU_PER_INCH = 914400
EMU_PER_POINT = 12700


def _set_theme_fonts_and_colors(theme_path: Path, tokens: dict, fonts: dict) -> None:
    tree = ET.parse(theme_path)
    root = tree.getroot()
    latin_face = fonts["latin"]
    ea_face = fonts["zh"]
    for font_scheme_tag in ("majorFont", "minorFont"):
        for scheme in root.iter(a(font_scheme_tag)):
            latin = scheme.find(a("latin"))
            if latin is not None:
                latin.set("typeface", latin_face)
            ea = scheme.find(a("ea"))
            if ea is not None:
                ea.set("typeface", ea_face)
            for script_font in scheme.findall(a("font")):
                if script_font.get("script") in {"Hans", "Hant"}:
                    script_font.set("typeface", ea_face)
    accents = {
        "accent1": tokens["colors"]["primary"].lstrip("#"),
        "accent2": tokens["colors"]["emphasis"].lstrip("#"),
        "accent3": tokens["colors"]["dark_red"].lstrip("#"),
    }
    for name, value in accents.items():
        for element in root.iter(a(name)):
            srgb = element.find(a("srgbClr"))
            if srgb is not None:
                srgb.set("val", value)
    tree.write(theme_path, xml_declaration=True, encoding="UTF-8")


def _style_def_rpr(
    def_rpr: ET.Element,
    size_pt: float | None,
    color: str | None,
    ea_typeface: str | None = None,
) -> None:
    if size_pt is not None:
        def_rpr.set("sz", sz(size_pt))
    if color:
        for existing in def_rpr.findall(a("solidFill")):
            def_rpr.remove(existing)
        fill = ET.Element(a("solidFill"))
        srgb = ET.SubElement(fill, a("srgbClr"))
        srgb.set("val", color.lstrip("#"))
        def_rpr.insert(0, fill)
    if ea_typeface:
        for existing in def_rpr.findall(a("ea")):
            def_rpr.remove(existing)
        ea_font = ET.SubElement(def_rpr, a("ea"))
        ea_font.set("typeface", ea_typeface)


def _first_level_def_rpr(style: ET.Element, level_tag: str) -> ET.Element | None:
    level = style.find(a(level_tag))
    if level is None:
        return None
    return level.find(a("defRPr"))


def _set_master_text_styles(master_path: Path, tokens: dict) -> None:
    tree = ET.parse(master_path)
    root = tree.getroot()
    sizes = tokens["font_sizes_pt"]
    styles = root.find(p("txStyles"))
    if styles is None:
        raise RuntimeError("slideMaster1.xml 缺少 txStyles")

    title_style = styles.find(p("titleStyle"))
    if title_style is not None:
        def_rpr = _first_level_def_rpr(title_style, "lvl1pPr")
        if def_rpr is not None:
            _style_def_rpr(def_rpr, sizes["slide_title"], tokens["colors"]["primary"])

    body_style = styles.find(p("bodyStyle"))
    if body_style is not None:
        body_sizes = {
            "lvl1pPr": sizes["body"],
            "lvl2pPr": sizes["compact_label"],
            "lvl3pPr": sizes["source_footnote"],
            "lvl4pPr": sizes["source_footnote"],
            "lvl5pPr": sizes["source_footnote"],
        }
        for level_tag, size_pt in body_sizes.items():
            def_rpr = _first_level_def_rpr(body_style, level_tag)
            if def_rpr is not None:
                _style_def_rpr(def_rpr, size_pt, tokens["colors"]["text"])
    tree.write(master_path, xml_declaration=True, encoding="UTF-8")


def _set_layout_placeholder_styles(layout_path: Path, tokens: dict, fonts: dict) -> None:
    sizes = tokens["font_sizes_pt"]
    primary = tokens["colors"]["primary"]
    display = fonts["display_zh"]
    # (size, color, ea_typeface)；方正兰亭粗黑简体只进封面主标题与章节页标题。
    placeholder_specs = {
        "ctrTitle": (sizes["cover_title"], primary, display),
        "subTitle": (sizes["cover_subtitle"], tokens["colors"]["secondary_text"], None),
    }
    tree = ET.parse(layout_path)
    root = tree.getroot()
    csld = root.find(p("cSld"))
    layout_name = csld.get("name", "") if csld is not None else ""
    is_section_layout = "Section" in layout_name
    changed = False
    for shape in root.iter(p("sp")):
        placeholder = shape.find(f".//{p('nvSpPr')}/{p('nvPr')}/{p('ph')}")
        if placeholder is None:
            continue
        ph_type = placeholder.get("type", "")
        spec = placeholder_specs.get(ph_type)
        if spec is None and is_section_layout and ph_type == "title":
            spec = (sizes["section_title"], primary, display)
        if spec is None:
            continue
        tx_body = shape.find(f"{p('txBody')}")
        if tx_body is None:
            continue
        lst_style = tx_body.find(a("lstStyle"))
        if lst_style is None:
            lst_style = ET.Element(a("lstStyle"))
            tx_body.insert(1, lst_style)
        level = lst_style.find(a("lvl1pPr"))
        if level is None:
            level = ET.SubElement(lst_style, a("lvl1pPr"))
        def_rpr = level.find(a("defRPr"))
        if def_rpr is None:
            def_rpr = ET.SubElement(level, a("defRPr"))
        _style_def_rpr(def_rpr, *spec)
        changed = True
    if changed:
        tree.write(layout_path, xml_declaration=True, encoding="UTF-8")


def _set_canvas(presentation_path: Path, tokens: dict) -> tuple[float, float]:
    tree = ET.parse(presentation_path)
    root = tree.getroot()
    slide_size = root.find(p("sldSz"))
    if slide_size is None:
        raise RuntimeError("presentation.xml 缺少 sldSz")
    old_width = int(slide_size.get("cx", "0"))
    old_height = int(slide_size.get("cy", "0"))
    new_width = round(tokens["canvas"]["width_in"] * EMU_PER_INCH)
    new_height = round(tokens["canvas"]["height_in"] * EMU_PER_INCH)
    if old_width <= 0 or old_height <= 0:
        raise RuntimeError("presentation.xml 的画布尺寸无效")
    slide_size.set("cx", str(new_width))
    slide_size.set("cy", str(new_height))
    tree.write(presentation_path, xml_declaration=True, encoding="UTF-8")
    return new_width / old_width, new_height / old_height


def _scale_layout_geometry(path: Path, scale_x: float, scale_y: float) -> None:
    tree = ET.parse(path)
    root = tree.getroot()
    for offset in root.iter(a("off")):
        if offset.get("x") is not None:
            offset.set("x", str(round(int(offset.get("x", "0")) * scale_x)))
        if offset.get("y") is not None:
            offset.set("y", str(round(int(offset.get("y", "0")) * scale_y)))
    for extent in root.iter(a("ext")):
        if extent.get("cx") is not None:
            extent.set("cx", str(round(int(extent.get("cx", "0")) * scale_x)))
        if extent.get("cy") is not None:
            extent.set("cy", str(round(int(extent.get("cy", "0")) * scale_y)))
    tree.write(path, xml_declaration=True, encoding="UTF-8")


def _enforce_title_safe_zone(path: Path, tokens: dict) -> None:
    tree = ET.parse(path)
    root = tree.getroot()
    minimum_y = round(tokens["logo_slot"]["content_top_min_pt"] * EMU_PER_POINT)
    changed = False
    for shape in root.iter(p("sp")):
        placeholder = shape.find(f".//{p('nvSpPr')}/{p('nvPr')}/{p('ph')}")
        if placeholder is None or placeholder.get("type") != "title":
            continue
        transform = shape.find(f"{p('spPr')}/{a('xfrm')}")
        if transform is None:
            continue
        offset = transform.find(a("off"))
        extent = transform.find(a("ext"))
        if offset is None or int(offset.get("y", "0")) >= minimum_y:
            continue
        old_y = int(offset.get("y", "0"))
        offset.set("y", str(minimum_y))
        if extent is not None:
            height = int(extent.get("cy", "0"))
            extent.set("cy", str(max(1, height - (minimum_y - old_y))))
        changed = True
    if changed:
        tree.write(path, xml_declaration=True, encoding="UTF-8")


def build(output: Path = DEFAULT_OUTPUT, font_profile: str | None = None) -> None:
    tokens = load_ppt_tokens()
    font_profile = font_profile or tokens["default_font_profile"]
    try:
        fonts = tokens["font_profiles"][font_profile]
    except KeyError as error:
        raise ValueError(f"未知 PPT 字体档位：{font_profile}") from error
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="qmd-ppt-reference-") as tmp:
        tmpdir = Path(tmp)
        base = tmpdir / "base.pptx"
        fetch_base_template(base)
        extract = tmpdir / "extract"
        with zipfile.ZipFile(base) as archive:
            names = archive.namelist()
            archive.extractall(extract)

        scale_x, scale_y = _set_canvas(extract / "ppt" / "presentation.xml", tokens)
        geometry_files = [
            *sorted(extract.glob("ppt/slideMasters/slideMaster*.xml")),
            *sorted(extract.glob("ppt/slideLayouts/slideLayout*.xml")),
        ]
        for geometry in geometry_files:
            _scale_layout_geometry(geometry, scale_x, scale_y)
            _enforce_title_safe_zone(geometry, tokens)
        for theme in sorted(extract.glob("ppt/theme/theme*.xml")):
            _set_theme_fonts_and_colors(theme, tokens, fonts)
        _set_master_text_styles(extract / "ppt" / "slideMasters" / "slideMaster1.xml", tokens)
        for layout in sorted(extract.glob("ppt/slideLayouts/slideLayout*.xml")):
            _set_layout_placeholder_styles(layout, tokens, fonts)

        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
            for name in names:
                archive.write(extract / name, name)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    tokens = load_ppt_tokens()
    parser.add_argument(
        "--font-profile",
        choices=tuple(tokens["font_profiles"]),
        default=tokens["default_font_profile"],
    )
    args = parser.parse_args()
    build(args.output, args.font_profile)
    print(f"reference master written: {args.output} (font_profile={args.font_profile})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
