#!/usr/bin/env python3
"""Validate, render, and create page previews for a QMD PowerPoint deck."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from build_reference_pptx import build as build_reference
from inject_cards import inject
from inspect_pptx import inspect
from validate_qmd import validate


ROOT = Path(__file__).resolve().parents[2]
SHARED_SCRIPTS = ROOT / "shared-assets" / "scripts"
if str(SHARED_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SHARED_SCRIPTS))

from font_preflight import font_environment  # noqa: E402


def require(name: str) -> str:
    executable = shutil.which(name)
    if not executable:
        raise RuntimeError(f"缺少运行依赖：{name}（安装方式见仓库 README.md）")
    return executable


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("qmd", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--preview-dir", type=Path)
    tokens = json.loads((ROOT / "shared-assets" / "design-tokens.json").read_text(encoding="utf-8"))["ppt"]
    parser.add_argument(
        "--font-profile",
        choices=tuple(tokens["font_profiles"]),
        default=tokens["default_font_profile"],
    )
    parser.add_argument(
        "--reference",
        type=Path,
        default=None,
        help="自定义参考母版；默认按当前 tokens 和字体档位现场生成，避免入库母版过期",
    )
    args = parser.parse_args()
    source = args.qmd.resolve()
    output = args.output.resolve()
    preview_dir = args.preview_dir.resolve() if args.preview_dir else None
    timeout = tokens["render_timeout_seconds"]

    result = validate(source)
    if result["errors"]:
        for error in result["errors"]:
            print(f"ERROR: {error}", file=sys.stderr)
        return 2

    try:
        quarto = require("quarto")
        with font_environment("ppt", args.font_profile) as render_env, tempfile.TemporaryDirectory(
            prefix="qmd-ppt-reference-"
        ) as reference_dir:
            if args.reference is None:
                reference = Path(reference_dir) / "reference.pptx"
                build_reference(reference, args.font_profile)
            else:
                reference = args.reference.resolve()
                if not reference.exists():
                    raise RuntimeError(f"参考母版不存在：{reference}")

            output.parent.mkdir(parents=True, exist_ok=True)
            staged_output = source.parent / output.name
            subprocess.run(
                [
                    quarto,
                    "render",
                    source.name,
                    "--to",
                    "pptx",
                    "--metadata",
                    f"reference-doc:{reference}",
                    "--output",
                    staged_output.name,
                ],
                cwd=source.parent,
                env=render_env,
                check=True,
                timeout=timeout,
            )
            if not staged_output.exists():
                raise RuntimeError(f"Quarto 未生成预期文件：{staged_output}")
            if staged_output != output:
                shutil.move(str(staged_output), str(output))

            inject(output, source, args.font_profile)

            inspection = inspect(output, args.font_profile)
            if inspection["errors"]:
                raise RuntimeError("PPTX 验收失败：" + "；".join(inspection["errors"]))

            if preview_dir:
                soffice = require("soffice")
                pdftoppm = require("pdftoppm")
                preview_dir.mkdir(parents=True, exist_ok=True)
                with tempfile.TemporaryDirectory(prefix="qmd-ppt-profile-") as profile, tempfile.TemporaryDirectory(
                    prefix="qmd-ppt-preview-"
                ) as preview_tmp:
                    preview_tmp_path = Path(preview_tmp)
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
                            str(preview_tmp_path),
                            str(output),
                        ],
                        env=render_env,
                        check=True,
                        capture_output=True,
                        text=True,
                        timeout=timeout,
                    )
                    pdf = preview_tmp_path / f"{output.stem}.pdf"
                    if not pdf.exists():
                        raise RuntimeError(f"LibreOffice 未生成 PDF：{pdf}")
                    subprocess.run(
                        [pdftoppm, "-png", "-r", "160", str(pdf), str(preview_tmp_path / "slide")],
                        env=render_env,
                        check=True,
                        capture_output=True,
                        text=True,
                        timeout=timeout,
                    )
                    pages = sorted(preview_tmp_path.glob("slide-*.png"))
                    if not pages:
                        raise RuntimeError(f"PDF 转换器未生成幻灯片 PNG：{pdf}")
                    for stale in preview_dir.glob("slide-*.png"):
                        stale.unlink()
                    shutil.copy2(pdf, preview_dir / pdf.name)
                    for page in pages:
                        shutil.copy2(page, preview_dir / page.name)
    except subprocess.TimeoutExpired:
        print(f"ERROR: 渲染超时（{timeout} 秒）", file=sys.stderr)
        return 2
    except (RuntimeError, subprocess.CalledProcessError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
