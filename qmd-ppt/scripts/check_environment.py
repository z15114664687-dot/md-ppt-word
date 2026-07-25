#!/usr/bin/env python3
"""Report executable and exact-font readiness for the qmd-ppt workflow."""

from __future__ import annotations

import argparse
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
    system_families: set[str] | None = None,
    bundled_family_scanner: Callable[[Path], set[str] | None] | None = None,
) -> dict[str, object]:
    tools = {name: tool_lookup(name) for name in ("quarto", "soffice", "pdftoppm")}
    fonts = check_font_profile(
        "ppt",
        font_profile,
        system_families,
        bundled_family_scanner,
    )
    return {
        "executables": tools,
        "fonts": fonts,
        "ready": bool(tools["quarto"] and fonts["render_safe"]),
        "visual_ready": bool(tools["soffice"] and tools["pdftoppm"] and fonts["render_safe"]),
    }


def main() -> int:
    tokens = json.loads((ROOT / "shared-assets" / "design-tokens.json").read_text(encoding="utf-8"))["ppt"]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--font-profile",
        choices=tuple(tokens["font_profiles"]),
        default=tokens["default_font_profile"],
    )
    args = parser.parse_args()
    report = build_report(args.font_profile)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if not report["executables"]["quarto"]:
        print("Quarto 未安装：可以编写和预检 QMD，但不能生成 PPTX。")
    if not report["fonts"]["render_safe"]:
        print("字体档位无法在本机可靠渲染；请安装报告中的精确字体，或改用 preview 档位。")
    if not report["visual_ready"]:
        print("提示：缺少 LibreOffice/Poppler 或字体不完整时，不能完成逐页 PNG 视觉质检。")
    return 0 if report["ready"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
