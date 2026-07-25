#!/usr/bin/env python3
"""Render DOCX pages with the project-local CJK font configuration."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SHARED_SCRIPTS = ROOT / "shared-assets" / "scripts"
if str(SHARED_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SHARED_SCRIPTS))

from font_preflight import font_environment  # noqa: E402


def require(name: str) -> str:
    executable = shutil.which(name)
    if not executable:
        raise RuntimeError(f"缺少预览依赖：{name}")
    return executable


def render(docx: Path, output_dir: Path, font_profile: str | None = None) -> list[Path]:
    soffice = require("soffice")
    pdftoppm = require("pdftoppm")
    tokens = json.loads((ROOT / "shared-assets" / "design-tokens.json").read_text(encoding="utf-8"))["word"]
    font_profile = font_profile or tokens["default_font_profile"]
    timeout = tokens["render_timeout_seconds"]
    docx = docx.resolve()
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    with font_environment("word", font_profile) as env:
        with tempfile.TemporaryDirectory(prefix="md-word-profile-") as profile, tempfile.TemporaryDirectory(
            prefix="md-word-pdf-"
        ) as pdf_dir, tempfile.TemporaryDirectory(prefix="md-word-pages-") as pages_dir:
            try:
                subprocess.run(
                    [
                        soffice,
                        f"-env:UserInstallation={Path(profile).as_uri()}",
                        "--invisible",
                        "--headless",
                        "--norestore",
                        "--nodefault",
                        "--nofirststartwizard",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        pdf_dir,
                        str(docx),
                    ],
                    env=env,
                    check=True,
                    capture_output=True,
                    text=True,
                    timeout=timeout,
                )
            except subprocess.TimeoutExpired as error:
                raise RuntimeError(f"LibreOffice 渲染超时（{timeout} 秒）：{docx}") from error
            except subprocess.CalledProcessError as error:
                detail = (error.stderr or error.stdout or "").strip()
                suffix = f" - {detail}" if detail else ""
                raise RuntimeError(f"LibreOffice 渲染失败：{docx}{suffix}") from error
            pdf = Path(pdf_dir) / f"{docx.stem}.pdf"
            if not pdf.exists():
                raise RuntimeError(f"LibreOffice 未生成 PDF：{pdf}")
            try:
                subprocess.run(
                    [pdftoppm, "-png", "-r", "160", str(pdf), str(Path(pages_dir) / "page")],
                    check=True,
                    capture_output=True,
                    text=True,
                    timeout=timeout,
                )
            except subprocess.TimeoutExpired as error:
                raise RuntimeError(f"PDF 转 PNG 超时（{timeout} 秒）：{pdf}") from error
            except subprocess.CalledProcessError as error:
                detail = (error.stderr or error.stdout or "").strip()
                suffix = f" - {detail}" if detail else ""
                raise RuntimeError(f"PDF 转 PNG 失败：{pdf}{suffix}") from error
            generated = sorted(Path(pages_dir).glob("page-*.png"))
            if not generated:
                raise RuntimeError(f"PDF 转换器未生成页面 PNG：{pdf}")
            for stale in output_dir.glob("page-*.png"):
                stale.unlink()
            rendered: list[Path] = []
            for page in generated:
                target = output_dir / page.name
                shutil.copy2(page, target)
                rendered.append(target)
    return rendered


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    tokens = json.loads((ROOT / "shared-assets" / "design-tokens.json").read_text(encoding="utf-8"))["word"]
    parser.add_argument(
        "--font-profile",
        choices=tuple(tokens["font_profiles"]),
        default=tokens["default_font_profile"],
    )
    args = parser.parse_args()
    try:
        pages = render(args.docx, args.output_dir, args.font_profile)
    except RuntimeError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    print(f"rendered_pages={len(pages)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
