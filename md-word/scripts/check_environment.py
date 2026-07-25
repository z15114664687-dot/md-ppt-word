#!/usr/bin/env python3
"""Report dependency and exact-font readiness for the md-word workflow."""

from __future__ import annotations

import argparse
import importlib.util
import json
import shutil
import sys
from pathlib import Path
from typing import Callable


ROOT = Path(__file__).resolve().parents[2]
SHARED_SCRIPTS = ROOT / "shared-assets" / "scripts"
if str(SHARED_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SHARED_SCRIPTS))

from font_preflight import check_font_profile  # noqa: E402


def build_report(
    font_profile: str,
    tool_lookup: Callable[[str], str | None] = shutil.which,
    python_docx_available: bool | None = None,
    system_families: set[str] | None = None,
    bundled_family_scanner: Callable[[Path], set[str] | None] | None = None,
) -> dict[str, object]:
    tools = {
        name: tool_lookup(name)
        for name in ("pandoc", "quarto", "soffice", "pdftoppm", "rsvg-convert")
    }
    if python_docx_available is None:
        python_docx_available = importlib.util.find_spec("docx") is not None
    fonts = check_font_profile(
        "word",
        font_profile,
        system_families,
        bundled_family_scanner,
    )
    converter_ready = bool(tools["pandoc"] or tools["quarto"])
    return {
        "executables": tools,
        "python_docx": python_docx_available,
        "fonts": fonts,
        "ready": bool(converter_ready and python_docx_available and fonts["render_safe"]),
        "visual_ready": bool(tools["soffice"] and tools["pdftoppm"] and fonts["render_safe"]),
    }


def main() -> int:
    tokens = json.loads((ROOT / "shared-assets" / "design-tokens.json").read_text(encoding="utf-8"))["word"]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--font-profile",
        choices=tuple(tokens["font_profiles"]),
        default=tokens["default_font_profile"],
    )
    args = parser.parse_args()
    report = build_report(args.font_profile)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    tools = report["executables"]
    if not (tools["pandoc"] or tools["quarto"]):
        print("缺少 Pandoc（或 Quarto 内置 Pandoc）：只能预检 Markdown，不能生成 DOCX。")
    if not report["python_docx"]:
        print("缺少 python-docx：无法生成参考模板和执行 OOXML 后处理。")
    if not report["fonts"]["render_safe"]:
        print("字体档位无法在本机可靠渲染；请安装报告中的精确字体，或改用 preview 档位。")
    if not report["visual_ready"]:
        print("提示：缺少 LibreOffice/Poppler 或字体不完整时，不能完成逐页 PNG 视觉质检。")
    if not tools["rsvg-convert"]:
        print("提示：缺少 rsvg-convert 时，本地 SVG 图片无法稳定转换为 PNG。")
    return 0 if report["ready"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
