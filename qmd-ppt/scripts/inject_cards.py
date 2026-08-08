#!/usr/bin/env python3
"""Inject native card shapes into a Quarto-rendered research PPTX.

Quarto/Pandoc's pptx writer only fills placeholder layouts — it can't draw the
高密度 conclusion bars / KPI blocks a research deck wants. This post-processor
reads card markers from the QMD (raw blocks Pandoc drops from the pptx) and
appends the matching **native, still-editable** shapes onto the slide, matched
by its `##` title. Runs after `quarto render`, before inspect.

Card markers (fenced raw blocks — Pandoc drops them for the pptx target):

    ```{=ppt-takeaway}
    一句结论
    来源：Wind，作者整理
    ```

    ```{=ppt-kpi}
    9.77% | 年化收益
    314%  | 累计收益
    55%   | 最大回撤
    ```

    ```{=ppt-cards}
    供需 | 收入增速由负转正，需求侧确认
    盈利 | 毛利率连续两季改善，盈利侧确认
    估值 | 处于近五年低位，估值侧留有空间
    ```

`{=ppt-kpi}` gives a row of KPI number cards; `{=ppt-cards}` a row of 2-4
content cards (`标题 | 描述`); `{=ppt-flow}` numbered step cards ①→②→③ joined
by arrows (筛选/流程图); `{=ppt-compare}` header-band cards side by side with a
VS badge — the LAST row is the emphasised (red-header) side, and its 描述 splits
into bullets on `；`. All take an optional trailing `来源：…` line, injected as
a caption under the cards.

Colours, fonts, and sizes come from shared-assets/design-tokens.json — the same
token source as the reference master, so the cards match the brand layer.
"""

from __future__ import annotations

import argparse
import json
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from validate_qmd import extract_card_blocks, split_slides

ROOT = Path(__file__).resolve().parents[2]
TOKENS_PATH = ROOT / "shared-assets" / "design-tokens.json"

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("a", A_NS)
ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS)

EMU_PER_INCH = 914400
SOURCE_LINE = re.compile(r"^(?:(?:资料)?来源|source)[ \t]*[:：][ \t]*(\S.*)$", re.IGNORECASE)


def p(tag: str) -> str:
    return f"{{{P_NS}}}{tag}"


def a(tag: str) -> str:
    return f"{{{A_NS}}}{tag}"


def load_tokens() -> dict:
    return json.loads(TOKENS_PATH.read_text(encoding="utf-8"))["ppt"]


def _parse_rows_block(block: str) -> tuple[list[tuple[str, str]], str]:
    """Parse a `left | right` rows block with an optional trailing 来源/source line."""
    rows: list[tuple[str, str]] = []
    source = ""
    for line in block.splitlines():
        stripped_line = line.strip()
        if not stripped_line:
            continue
        source_match = SOURCE_LINE.match(stripped_line)
        if source_match:
            source = source_match.group(1)
        elif "|" in stripped_line:
            left, _, right = stripped_line.partition("|")
            if left.strip() and right.strip():
                rows.append((left.strip(), right.strip()))
    return rows, source


def _parse_takeaway_block(block: str) -> tuple[str, str]:
    text_lines: list[str] = []
    source = ""
    for line in block.splitlines():
        stripped_line = line.strip()
        if not stripped_line:
            continue
        source_match = SOURCE_LINE.match(stripped_line)
        if source_match:
            source = source_match.group(1)
        else:
            text_lines.append(stripped_line)
    return "\n".join(text_lines), source


def parse_cards(qmd_path: Path) -> dict[str, dict[str, list]]:
    """Return parsed card text/rows and their optional source captions by slide."""
    text = Path(qmd_path).read_text(encoding="utf-8")
    frontmatter = re.match(r"\A---\s*\n.*?\n---\s*\n", text, re.DOTALL)
    body = text[frontmatter.end() :] if frontmatter else text
    cards: dict[str, dict[str, list]] = {}
    for title, _, content in split_slides(body):
        takeaways: list[tuple[str, str]] = []
        row_blocks: dict[str, list[tuple[list[tuple[str, str]], str]]] = {
            "kpi": [],
            "cards": [],
            "flow": [],
            "compare": [],
        }
        for kind, block in extract_card_blocks(content):
            if kind == "takeaway":
                parsed_takeaway = _parse_takeaway_block(block)
                if parsed_takeaway[0]:
                    takeaways.append(parsed_takeaway)
                continue
            parsed = _parse_rows_block(block)
            if parsed[0]:
                row_blocks[kind].append(parsed)
        kpi_blocks = row_blocks["kpi"]
        grid_blocks = row_blocks["cards"]
        flow_blocks = row_blocks["flow"]
        compare_blocks = row_blocks["compare"]
        if takeaways or kpi_blocks or grid_blocks or flow_blocks or compare_blocks:
            cards[title] = {
                "takeaway": takeaways,
                "kpi": kpi_blocks,
                "cards": grid_blocks,
                "flow": flow_blocks,
                "compare": compare_blocks,
            }
    return cards


def _slide_title(root: ET.Element) -> str:
    for shape in root.iter(p("sp")):
        placeholder = shape.find(f".//{p('ph')}")
        if placeholder is None or placeholder.get("type") not in {"title", "ctrTitle"}:
            continue
        return "".join(node.text or "" for node in shape.iter(a("t"))).strip()
    return ""


def _next_shape_id(root: ET.Element) -> int:
    ids = [int(node.get("id", "0")) for node in root.iter(p("cNvPr")) if node.get("id", "").isdigit()]
    return (max(ids) + 1) if ids else 100


def _rect(shape_id: int, name: str, x: int, y: int, w: int, h: int, fill: str, *, rounded: bool = False, stroke: str | None = None, geom: str | None = None) -> str:
    geom = geom or ("roundRect" if rounded else "rect")
    adj = '<a:gd name="adj" fmla="val 8000"/>' if geom == "roundRect" else ""
    line = (
        f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{stroke.lstrip("#")}"/></a:solidFill></a:ln>'
        if stroke
        else "<a:ln><a:noFill/></a:ln>"
    )
    return (
        f'<p:sp xmlns:p="{P_NS}" xmlns:a="{A_NS}">'
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst>{adj}</a:avLst></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{line}</p:spPr>'
        f"<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"
    )


def _run(text: str, size_pt: float, color: str, latin: str, ea: str, *, bold: bool = False) -> str:
    b = "1" if bold else "0"
    return (
        f'<a:r><a:rPr lang="zh-CN" sz="{round(size_pt * 100)}" b="{b}">'
        f'<a:solidFill><a:srgbClr val="{color.lstrip("#")}"/></a:solidFill>'
        f'<a:latin typeface="{latin}"/><a:ea typeface="{ea}"/></a:rPr>'
        f"<a:t>{escape(text)}</a:t></a:r>"
    )


def _textbox(shape_id: int, name: str, x: int, y: int, w: int, h: int, runs: str | list[str], *, align: str = "l", anchor: str = "ctr", lins: int = 91440) -> str:
    """runs: one paragraph's runs, or a list of them (multi-paragraph body)."""
    paras = runs if isinstance(runs, list) else [runs]
    body = "".join(
        f'<a:p><a:pPr algn="{align}">'
        + ('<a:spcBef><a:spcPts val="600"/></a:spcBef>' if index else "")
        + f"</a:pPr>{para}</a:p>"
        for index, para in enumerate(paras)
    )
    return (
        f'<p:sp xmlns:p="{P_NS}" xmlns:a="{A_NS}">'
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="square" lIns="{lins}" rIns="{lins}" tIns="0" bIns="0" anchor="{anchor}"/>'
        f"<a:lstStyle/>{body}</p:txBody></p:sp>"
    )


def _takeaway_shapes(base_id: int, tokens: dict, fonts: dict, text: str, source: str = "") -> list[str]:
    width = round(tokens["canvas"]["width_in"] * EMU_PER_INCH)
    margin = round(0.42 * EMU_PER_INCH)
    height = round(0.52 * EMU_PER_INCH)
    y = round(tokens["canvas"]["height_in"] * EMU_PER_INCH) - height - round(0.32 * EMU_PER_INCH)
    bar_w = round(0.07 * EMU_PER_INCH)
    card_w = width - 2 * margin
    size = tokens["font_sizes_pt"]["key_claim"]
    run = _run(text, size, tokens["colors"]["text"], fonts["latin"], fonts["zh"], bold=False)
    shapes = [
        _rect(base_id, "takeaway-card", margin, y, card_w, height, tokens["colors"]["pale_red"].lstrip("#"), rounded=True),
        _rect(base_id + 1, "takeaway-bar", margin, y, bar_w, height, tokens["colors"]["primary"].lstrip("#")),
        _textbox(base_id + 2, "takeaway-text", margin + round(0.28 * EMU_PER_INCH), y, card_w - round(0.4 * EMU_PER_INCH), height, run, align="l"),
    ]
    if source:
        source_run = _run(
            "来源：" + source,
            tokens["font_sizes_pt"]["source_footnote"],
            tokens["colors"]["secondary_text"],
            fonts["latin"],
            fonts["zh"],
        )
        shapes.append(
            _textbox(
                base_id + 3,
                "takeaway-source",
                margin,
                y + height + round(0.03 * EMU_PER_INCH),
                card_w,
                round(0.23 * EMU_PER_INCH),
                source_run,
                align="l",
                anchor="ctr",
            )
        )
    return shapes


def _kpi_shapes(base_id: int, tokens: dict, fonts: dict, rows: list[tuple[str, str]], source: str = "") -> list[str]:
    width = round(tokens["canvas"]["width_in"] * EMU_PER_INCH)
    margin = round(0.42 * EMU_PER_INCH)
    top = round(tokens["logo_slot"]["content_top_min_pt"] / 72 * EMU_PER_INCH) + round(0.72 * EMU_PER_INCH)
    height = round(1.15 * EMU_PER_INCH)
    gap = round(0.18 * EMU_PER_INCH)
    count = len(rows)
    cell_w = (width - 2 * margin - gap * (count - 1)) // count
    value_size = tokens["font_sizes_pt"]["closing"] * 0.7
    label_size = tokens["font_sizes_pt"]["body"]
    shapes: list[str] = []
    shape_id = base_id
    for index, (value, label) in enumerate(rows):
        x = margin + index * (cell_w + gap)
        shapes.append(_rect(shape_id, f"kpi-card-{index}", x, top, cell_w, height, tokens["colors"]["pale_red"].lstrip("#"), rounded=True))
        value_run = _run(value, value_size, tokens["colors"]["primary"], fonts["latin"], fonts["zh"], bold=True)
        shapes.append(_textbox(shape_id + 1, f"kpi-value-{index}", x, top + round(0.18 * EMU_PER_INCH), cell_w, round(0.6 * EMU_PER_INCH), value_run, align="ctr", anchor="ctr"))
        label_run = _run(label, label_size, tokens["colors"]["secondary_text"], fonts["latin"], fonts["zh"], bold=False)
        shapes.append(_textbox(shape_id + 2, f"kpi-label-{index}", x, top + round(0.72 * EMU_PER_INCH), cell_w, round(0.35 * EMU_PER_INCH), label_run, align="ctr", anchor="ctr"))
        shape_id += 3
    if source:
        source_run = _run("来源：" + source, tokens["font_sizes_pt"]["source_footnote"], tokens["colors"]["secondary_text"], fonts["latin"], fonts["zh"])
        shapes.append(_textbox(shape_id, "kpi-source", margin, top + height + round(0.1 * EMU_PER_INCH), width - 2 * margin, round(0.3 * EMU_PER_INCH), source_run, align="l", anchor="ctr"))
    return shapes


def _cardgrid_shapes(base_id: int, tokens: dict, fonts: dict, rows: list[tuple[str, str]], source: str = "") -> list[str]:
    """A row of 2-4 content cards (accent bar + bold title + wrapping body) — the
    high-density concept layout Quarto placeholders can't produce."""
    width = round(tokens["canvas"]["width_in"] * EMU_PER_INCH)
    margin = round(0.42 * EMU_PER_INCH)
    top = round(tokens["logo_slot"]["content_top_min_pt"] / 72 * EMU_PER_INCH) + round(0.72 * EMU_PER_INCH)
    height = round(2.35 * EMU_PER_INCH)
    gap = round(0.22 * EMU_PER_INCH)
    count = len(rows)
    cell_w = (width - 2 * margin - gap * (count - 1)) // count
    bar_w = round(0.06 * EMU_PER_INCH)
    pad = round(0.24 * EMU_PER_INCH)
    title_size = tokens["font_sizes_pt"]["key_claim"]
    body_size = tokens["font_sizes_pt"]["body"]
    shapes: list[str] = []
    shape_id = base_id
    for index, (heading, desc) in enumerate(rows):
        x = margin + index * (cell_w + gap)
        shapes.append(_rect(shape_id, f"cardg-{index}", x, top, cell_w, height, "FFFFFF", rounded=True, stroke=tokens["colors"]["border_red"]))
        shapes.append(_rect(shape_id + 1, f"cardg-bar-{index}", x, top, bar_w, height, tokens["colors"]["primary"].lstrip("#"), rounded=False))
        heading_run = _run(heading, title_size, tokens["colors"]["primary"], fonts["latin"], fonts["zh"], bold=True)
        shapes.append(_textbox(shape_id + 2, f"cardg-title-{index}", x + pad, top + round(0.24 * EMU_PER_INCH), cell_w - 2 * pad, round(0.5 * EMU_PER_INCH), heading_run, align="l", anchor="t"))
        desc_run = _run(desc, body_size, tokens["colors"]["text"], fonts["latin"], fonts["zh"], bold=False)
        shapes.append(_textbox(shape_id + 3, f"cardg-desc-{index}", x + pad, top + round(0.82 * EMU_PER_INCH), cell_w - 2 * pad, height - round(1.0 * EMU_PER_INCH), desc_run, align="l", anchor="t"))
        shape_id += 4
    if source:
        source_run = _run("来源：" + source, tokens["font_sizes_pt"]["source_footnote"], tokens["colors"]["secondary_text"], fonts["latin"], fonts["zh"])
        shapes.append(_textbox(shape_id, "cardg-source", margin, top + height + round(0.12 * EMU_PER_INCH), width - 2 * margin, round(0.3 * EMU_PER_INCH), source_run, align="l", anchor="ctr"))
    return shapes


def _flow_shapes(base_id: int, tokens: dict, fonts: dict, rows: list[tuple[str, str]], source: str = "") -> list[str]:
    """步骤①→②→③ flow: numbered step cards joined by brand-red arrows — the
    hand-drawn-flow-diagram look (boxes + arrows) research decks use for 筛选/流程."""
    width = round(tokens["canvas"]["width_in"] * EMU_PER_INCH)
    margin = round(0.42 * EMU_PER_INCH)
    top = round(tokens["logo_slot"]["content_top_min_pt"] / 72 * EMU_PER_INCH) + round(0.72 * EMU_PER_INCH)
    height = round(1.85 * EMU_PER_INCH)
    arrow_zone = round(0.5 * EMU_PER_INCH)
    arrow_w = round(0.3 * EMU_PER_INCH)
    arrow_h = round(0.26 * EMU_PER_INCH)
    pad = round(0.2 * EMU_PER_INCH)
    count = len(rows)
    cell_w = (width - 2 * margin - arrow_zone * (count - 1)) // count
    title_size = tokens["font_sizes_pt"]["key_claim"]
    body_size = tokens["font_sizes_pt"]["body"]
    shapes: list[str] = []
    shape_id = base_id
    for index, (heading, desc) in enumerate(rows):
        x = margin + index * (cell_w + arrow_zone)
        shapes.append(_rect(shape_id, f"flow-{index}", x, top, cell_w, height, tokens["colors"]["pale_red"].lstrip("#"), rounded=True))
        step = chr(0x2460 + index) if index < 20 else str(index + 1)
        heading_run = _run(f"{step} {heading}", title_size, tokens["colors"]["primary"], fonts["latin"], fonts["zh"], bold=True)
        shapes.append(_textbox(shape_id + 1, f"flow-title-{index}", x + pad, top + round(0.2 * EMU_PER_INCH), cell_w - 2 * pad, round(0.45 * EMU_PER_INCH), heading_run, align="l", anchor="t"))
        desc_run = _run(desc, body_size, tokens["colors"]["text"], fonts["latin"], fonts["zh"], bold=False)
        shapes.append(_textbox(shape_id + 2, f"flow-desc-{index}", x + pad, top + round(0.72 * EMU_PER_INCH), cell_w - 2 * pad, height - round(0.92 * EMU_PER_INCH), desc_run, align="l", anchor="t"))
        shape_id += 3
        if index < count - 1:
            arrow_x = x + cell_w + (arrow_zone - arrow_w) // 2
            arrow_y = top + (height - arrow_h) // 2
            shapes.append(_rect(shape_id, f"flow-arrow-{index}", arrow_x, arrow_y, arrow_w, arrow_h, tokens["colors"]["primary"].lstrip("#"), geom="rightArrow"))
            shape_id += 1
    if source:
        source_run = _run("来源：" + source, tokens["font_sizes_pt"]["source_footnote"], tokens["colors"]["secondary_text"], fonts["latin"], fonts["zh"])
        shapes.append(_textbox(shape_id, "flow-source", margin, top + height + round(0.12 * EMU_PER_INCH), width - 2 * margin, round(0.3 * EMU_PER_INCH), source_run, align="l", anchor="ctr"))
    return shapes


def _compare_shapes(base_id: int, tokens: dict, fonts: dict, rows: list[tuple[str, str]], source: str = "") -> list[str]:
    """左右对比卡 (改前 vs 改后、基准 vs 组合): header-band cards side by side; the
    LAST row is the emphasised side (red header). Body bullets split on ；/;。"""
    width = round(tokens["canvas"]["width_in"] * EMU_PER_INCH)
    margin = round(0.42 * EMU_PER_INCH)
    top = round(tokens["logo_slot"]["content_top_min_pt"] / 72 * EMU_PER_INCH) + round(0.72 * EMU_PER_INCH)
    height = round(2.6 * EMU_PER_INCH)
    header_h = round(0.5 * EMU_PER_INCH)
    gap = round(0.7 * EMU_PER_INCH)
    pad = round(0.24 * EMU_PER_INCH)
    count = len(rows)
    cell_w = (width - 2 * margin - gap * (count - 1)) // count
    title_size = tokens["font_sizes_pt"]["key_claim"]
    body_size = tokens["font_sizes_pt"]["body"]
    shapes: list[str] = []
    shape_id = base_id
    for index, (heading, desc) in enumerate(rows):
        x = margin + index * (cell_w + gap)
        band = tokens["colors"]["primary"] if index == count - 1 else tokens["colors"]["secondary_text"]
        shapes.append(_rect(shape_id, f"cmp-body-{index}", x, top, cell_w, height, "FFFFFF", stroke=tokens["colors"]["border_red"]))
        shapes.append(_rect(shape_id + 1, f"cmp-band-{index}", x, top, cell_w, header_h, band.lstrip("#")))
        heading_run = _run(heading, title_size, "#FFFFFF", fonts["latin"], fonts["zh"], bold=True)
        shapes.append(_textbox(shape_id + 2, f"cmp-title-{index}", x, top, cell_w, header_h, heading_run, align="ctr", anchor="ctr"))
        bullets = [part.strip() for part in re.split(r"[;；]", desc) if part.strip()]
        bullet_runs = [
            _run("· " + bullet, body_size, tokens["colors"]["text"], fonts["latin"], fonts["zh"], bold=False)
            for bullet in bullets
        ]
        shapes.append(_textbox(shape_id + 3, f"cmp-desc-{index}", x + pad, top + header_h + round(0.18 * EMU_PER_INCH), cell_w - 2 * pad, height - header_h - round(0.3 * EMU_PER_INCH), bullet_runs, align="l", anchor="t"))
        shape_id += 4
    if count == 2:
        badge = round(0.55 * EMU_PER_INCH)
        badge_x = margin + cell_w + (gap - badge) // 2
        badge_y = top + (height - badge) // 2
        shapes.append(_rect(shape_id, "cmp-vs", badge_x, badge_y, badge, badge, tokens["colors"]["primary"].lstrip("#"), geom="ellipse"))
        vs_run = _run("VS", tokens["font_sizes_pt"]["module_title"], "#FFFFFF", fonts["latin"], fonts["zh"], bold=True)
        shapes.append(_textbox(shape_id + 1, "cmp-vs-text", badge_x, badge_y, badge, badge, vs_run, align="ctr", anchor="ctr"))
        shape_id += 2
    if source:
        source_run = _run("来源：" + source, tokens["font_sizes_pt"]["source_footnote"], tokens["colors"]["secondary_text"], fonts["latin"], fonts["zh"])
        shapes.append(_textbox(shape_id, "cmp-source", margin, top + height + round(0.12 * EMU_PER_INCH), width - 2 * margin, round(0.3 * EMU_PER_INCH), source_run, align="l", anchor="ctr"))
    return shapes


def inject(pptx_path: Path, qmd_path: Path, font_profile: str | None = None) -> dict[str, object]:
    tokens = load_tokens()
    font_profile = font_profile or tokens["default_font_profile"]
    fonts = tokens["font_profiles"][font_profile]
    cards = parse_cards(Path(qmd_path))
    if not cards:
        return {"injected": 0, "unmatched": []}

    pptx_path = Path(pptx_path)
    injected = 0
    unmatched: list[str] = []
    matched_titles: set[str] = set()

    with zipfile.ZipFile(pptx_path) as archive:
        names = archive.namelist()
        blobs = {name: archive.read(name) for name in names}

    slide_names = sorted(
        (n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
        key=lambda n: int(re.search(r"(\d+)", n).group(1)),
    )
    for name in slide_names:
        root = ET.fromstring(blobs[name])
        title = _slide_title(root)
        spec = cards.get(title)
        if not spec:
            continue
        matched_titles.add(title)
        sp_tree = root.find(f"{p('cSld')}/{p('spTree')}")
        if sp_tree is None:
            continue
        shape_id = _next_shape_id(root)
        new_shapes: list[str] = []
        for rows, source in spec["kpi"]:
            new_shapes.extend(_kpi_shapes(shape_id, tokens, fonts, rows, source))
            shape_id += len(rows) * 3 + 1
        for rows, source in spec.get("cards", []):
            new_shapes.extend(_cardgrid_shapes(shape_id, tokens, fonts, rows, source))
            shape_id += len(rows) * 4 + 1
        for rows, source in spec.get("flow", []):
            new_shapes.extend(_flow_shapes(shape_id, tokens, fonts, rows, source))
            shape_id += len(rows) * 4 + 2
        for rows, source in spec.get("compare", []):
            new_shapes.extend(_compare_shapes(shape_id, tokens, fonts, rows, source))
            shape_id += len(rows) * 4 + 4
        for text, source in spec["takeaway"]:
            takeaway_shapes = _takeaway_shapes(shape_id, tokens, fonts, text, source)
            new_shapes.extend(takeaway_shapes)
            shape_id += len(takeaway_shapes)
        for shape_xml in new_shapes:
            sp_tree.append(ET.fromstring(shape_xml))
        blobs[name] = ET.tostring(root, xml_declaration=True, encoding="UTF-8")
        injected += 1

    unmatched = [title for title in cards if title not in matched_titles]

    with zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in names:
            archive.writestr(name, blobs[name])

    return {"injected": injected, "unmatched": unmatched}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("qmd", type=Path)
    parser.add_argument("--font-profile")
    args = parser.parse_args()
    result = inject(args.pptx, args.qmd, args.font_profile)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if result["unmatched"]:
        print("WARN: 以下页标题在 pptx 中未匹配（卡片未注入）：" + "；".join(result["unmatched"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
