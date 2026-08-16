#!/usr/bin/env python3
"""Render fenced ```mermaid blocks to themed PNG via Quarto.

Word 线走 Pandoc，本身不认 mermaid；qmd-ppt 线由 Quarto 原生处理。
本模块把 Markdown 里的 ```mermaid 块在进入 Pandoc 之前渲染成 PNG，
主题色取自 design-tokens.json，字体取自当前字体档位——因此流程图与
正文、图表使用同一套品牌规则，不需要引入新的图形工具链。
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
TOKENS_PATH = ROOT / "shared-assets" / "design-tokens.json"
SHARED_SCRIPTS = ROOT / "shared-assets" / "scripts"
if str(SHARED_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SHARED_SCRIPTS))

from font_preflight import font_environment  # noqa: E402

MERMAID_BLOCK = re.compile(r"(?ms)^```[ \t]*mermaid[ \t]*\n(.*?)^```[ \t]*$")
INIT_DIRECTIVE = re.compile(r"^\s*%%\{\s*init\s*:", re.MULTILINE)
# 每个字体档位里承担图形标签的字体键，按目标区分。
LABEL_FONT_KEYS = {"word": ("heading_zh", "heading_latin"), "ppt": ("zh", "latin")}


def load_tokens(target: str) -> dict:
    if target not in {"word", "ppt"}:
        raise ValueError(f"未知目标：{target}，只支持 word / ppt")
    return json.loads(TOKENS_PATH.read_text(encoding="utf-8"))[target]


def label_font(tokens: dict, target: str, font_profile: str) -> str:
    profiles = tokens["font_profiles"]
    if font_profile not in profiles:
        raise ValueError(f"未知字体档位：{font_profile}，可选 {sorted(profiles)}")
    profile = profiles[font_profile]
    for key in LABEL_FONT_KEYS[target]:
        if profile.get(key):
            return str(profile[key])
    raise ValueError(f"字体档位 {font_profile} 未定义图形标签字体")


def init_directive(tokens: dict, target: str, font_profile: str) -> str:
    """Build the mermaid %%{init}%% directive from tokens + font profile."""
    diagram = tokens["diagram"]
    variables = dict(diagram["mermaid_theme_variables"])
    variables["fontFamily"] = label_font(tokens, target, font_profile)
    payload = {"theme": diagram["mermaid_theme"], "themeVariables": variables}
    return "%%{init: " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + "}%%"


def _render_one(source: str, out_png: Path, tokens: dict, target: str, font_profile: str, env: dict) -> None:
    """Render a single mermaid source string to out_png using Quarto."""
    quarto = shutil.which("quarto")
    if not quarto:
        raise RuntimeError(
            "Markdown 含 ```mermaid 块但缺少 Quarto（安装方式见仓库 README.md），"
            "或先把流程图改为已渲染好的 PNG/SVG 图片"
        )
    timeout = tokens["render_timeout_seconds"]
    body = source if INIT_DIRECTIVE.search(source) else f"{init_directive(tokens, target, font_profile)}\n{source}"
    with tempfile.TemporaryDirectory(prefix="md-word-mermaid-") as tmp:
        work = Path(tmp)
        qmd = work / "diagram.qmd"
        qmd.write_text(
            "---\ntitle: \"\"\nformat:\n  markdown:\n    mermaid-format: png\n---\n\n"
            "```{mermaid}\n" + body.rstrip("\n") + "\n```\n",
            encoding="utf-8",
        )
        try:
            subprocess.run(
                [quarto, "render", qmd.name],
                cwd=work,
                env=env,
                check=True,
                capture_output=True,
                text=True,
                timeout=timeout,
            )
        except subprocess.TimeoutExpired as error:
            raise RuntimeError(f"mermaid 渲染超时（{timeout} 秒）") from error
        except subprocess.CalledProcessError as error:
            detail = (error.stderr or error.stdout or "").strip().splitlines()
            suffix = f" - {detail[-1]}" if detail else ""
            raise RuntimeError(f"mermaid 渲染失败{suffix}") from error
        produced = sorted(work.glob("diagram_files/**/*.png"))
        if not produced:
            raise RuntimeError("Quarto 未生成 mermaid PNG；请检查图形语法")
        out_png.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(produced[0], out_png)


def render_mermaid_blocks(
    text: str,
    out_dir: Path,
    target: str = "word",
    font_profile: str = "preview",
) -> tuple[str, list[Path]]:
    """Replace every ```mermaid fence with an image reference to a rendered PNG.

    Returns the rewritten Markdown and the list of generated PNG paths.
    图片路径写成绝对路径，避免后续 Pandoc 阶段的相对路径解析歧义。
    """
    blocks = list(MERMAID_BLOCK.finditer(text))
    if not blocks:
        return text, []

    tokens = load_tokens(target)
    label_font(tokens, target, font_profile)  # 提前校验档位，避免渲染到一半才失败
    rendered: list[Path] = []
    with font_environment(target, font_profile) as env:
        for index, match in enumerate(blocks):
            png = out_dir / f"diagram-{index}.png"
            _render_one(match.group(1), png, tokens, target, font_profile, env)
            rendered.append(png)

    # 从后往前替换，保证前面的匹配位置不被改动。
    for match, png in zip(reversed(blocks), reversed(rendered)):
        text = text[: match.start()] + f"![]({png})" + text[match.end() :]
    return text, rendered


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("markdown", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--target", choices=("word", "ppt"), default="word")
    parser.add_argument("--font-profile", default=None)
    args = parser.parse_args()
    tokens = load_tokens(args.target)
    font_profile = args.font_profile or tokens["default_font_profile"]
    try:
        text, rendered = render_mermaid_blocks(
            args.markdown.read_text(encoding="utf-8"), args.out_dir, args.target, font_profile
        )
    except (RuntimeError, ValueError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    print(json.dumps({"diagrams": [str(p) for p in rendered]}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
