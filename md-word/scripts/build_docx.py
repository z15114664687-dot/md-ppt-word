#!/usr/bin/env python3
"""Build a fixed-format research DOCX with native OMML equations."""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from build_reference_docx import build as build_reference, load_word_tokens, pandoc_command
from inspect_docx import inspect
from postprocess_docx import postprocess
from render_diagrams import render_mermaid_blocks
from validate_markdown import IMAGE, validate


ROOT = Path(__file__).resolve().parents[2]
SHARED_SCRIPTS = ROOT / "shared-assets" / "scripts"
if str(SHARED_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SHARED_SCRIPTS))

from font_preflight import check_font_profile, font_environment  # noqa: E402


def stage_source_with_png_figures(
    source: Path, tmpdir: Path, font_profile: str, text: str | None = None
) -> Path:
    """Convert local SVG figures to PNG so every renderer (Word/LibreOffice) shows them.

    Pandoc 能把 SVG 原样嵌进 DOCX，但旧版 Word 和 LibreOffice 渲染不稳定，
    因此在进入 Pandoc 前统一栅格化为高分辨率 PNG。

    text 由上游阶段（如 mermaid 渲染）传入时，相对图片路径仍按 source.parent 解析。
    """
    prestaged = text is not None
    text = source.read_text(encoding="utf-8") if text is None else text
    svg_resources = {
        resource
        for resource in IMAGE.findall(text)
        if resource.lower().endswith(".svg") and not re.match(r"https?://", resource)
    }
    if not svg_resources:
        if not prestaged:
            return source
        staged = tmpdir / source.name
        staged.write_text(text, encoding="utf-8")
        return staged
    for resource in sorted(svg_resources):
        svg_text = (source.parent / resource).read_text(encoding="utf-8")
        declared = re.search(r'data-font-profile=["\']([^"\']+)["\']', svg_text)
        if declared and declared.group(1) != font_profile:
            raise RuntimeError(
                f"SVG 字体档位与 DOCX 不一致：{resource}="
                f"{declared.group(1)}, requested={font_profile}。请按目标档位重新 materialize。"
            )
    converter = shutil.which("rsvg-convert")
    if not converter:
        raise RuntimeError("Markdown 引用了 SVG 图片但缺少 rsvg-convert（brew install librsvg），或先把图片改为 PNG")
    figures_dir = tmpdir / "figures"
    figures_dir.mkdir()
    timeout = load_word_tokens()["render_timeout_seconds"]
    with font_environment("word", font_profile) as env:
        for index, resource in enumerate(sorted(svg_resources)):
            png = figures_dir / f"figure-{index}.png"
            try:
                subprocess.run(
                    [converter, "--dpi-x", "300", "--dpi-y", "300", "--output", str(png), str(source.parent / resource)],
                    env=env,
                    check=True,
                    capture_output=True,
                    text=True,
                    timeout=timeout,
                )
            except subprocess.TimeoutExpired as error:
                raise RuntimeError(f"SVG 转 PNG 超时（{timeout} 秒）：{resource}") from error
            except subprocess.CalledProcessError as error:
                detail = (error.stderr or error.stdout or "").strip()
                suffix = f" - {detail}" if detail else ""
                raise RuntimeError(f"SVG 转 PNG 失败：{resource}{suffix}") from error
            if not png.exists():
                raise RuntimeError(f"SVG 转换器未生成 PNG：{resource}")
            text = text.replace(f"({resource})", f"({png})")
    staged = tmpdir / source.name
    staged.write_text(text, encoding="utf-8")
    return staged


def build(source: Path, output: Path, font_profile: str = "preview", reference: Path | None = None) -> dict[str, object]:
    source = source.resolve()
    output = output.resolve()
    result = validate(source)
    if result["errors"]:
        raise RuntimeError("Markdown 预检失败：" + "；".join(result["errors"]))

    pandoc = pandoc_command()
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="md-word-") as tmp:
        tmpdir = Path(tmp)
        if reference is None:
            reference = tmpdir / "reference.docx"
            build_reference(reference, font_profile)
        # mermaid 流程图先按 tokens 主题渲成 PNG，再和普通图片一起进 Pandoc。
        diagram_text, diagrams = render_mermaid_blocks(
            source.read_text(encoding="utf-8"), tmpdir / "diagrams", "word", font_profile
        )
        staged = stage_source_with_png_figures(
            source, tmpdir, font_profile, text=diagram_text if diagrams else None
        )
        html = tmpdir / "report.html"
        draft = tmpdir / "report.docx"
        # -implicit_figures：题注按契约写在对象上方，不要 Pandoc 用 alt 文本在图下再生成一份。
        # 哨兵 title：HTML 中转会把元数据 title（默认取文件名）落成 Title 段落，后处理按哨兵删除。
        # 不用 --embed-resources：data-URI 会被整段塞进 DOCX 的图片 descr 属性，
        # LibreOffice 渲染会静默丢图；图片由第二步按 --resource-path 从磁盘读取。
        subprocess.run(
            [*pandoc, str(staged), "--from", "markdown+tex_math_dollars+fenced_divs-implicit_figures", "--to", "html5", "--mathml", "--standalone", "--metadata", "title=__MD_WORD_DROP_TITLE__", "--resource-path", str(source.parent), "--output", str(html)],
            check=True,
        )
        # 不用 --toc/--number-sections：目录（内容目录+图表目录）和中式章节
        # 编号都在 postprocess 里完成，避免编号文本破坏样式匹配。
        subprocess.run(
            [*pandoc, str(html), "--from", "html", "--to", "docx", "--reference-doc", str(reference.resolve()), "--resource-path", str(source.parent), "--output", str(draft)],
            check=True,
        )
        postprocess(draft, font_profile)
        shutil.copy2(draft, output)

    font_report = check_font_profile("word", font_profile)
    inspection = inspect(output, int(result["math_expressions"]), font_profile)
    if inspection["errors"]:
        raise RuntimeError("DOCX 验收失败：" + "；".join(inspection["errors"]))
    return {
        "validation": result,
        "font_preflight": font_report,
        "inspection": inspection,
        "output": str(output),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("markdown", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    tokens = load_word_tokens()
    parser.add_argument(
        "--font-profile",
        choices=tuple(tokens["font_profiles"]),
        default=tokens["default_font_profile"],
        help="preview=内置 Noto 字体便于本机渲染质检；delivery=交付目标字体",
    )
    parser.add_argument("--reference", type=Path, default=None, help="自定义参考 DOCX；默认按字体档位现场生成")
    args = parser.parse_args()
    try:
        result = build(args.markdown, args.output, args.font_profile, args.reference)
        print(result)
        if not result["font_preflight"]["render_safe"]:
            print(
                "WARNING: DOCX 已写入目标字体名，但本机缺少精确字体，不能声明完成本地视觉验收。",
                file=sys.stderr,
            )
        return 0
    except RuntimeError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
