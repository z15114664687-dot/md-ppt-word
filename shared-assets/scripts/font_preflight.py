#!/usr/bin/env python3
"""Verify exact font families for portable preview and target delivery profiles."""

from __future__ import annotations

import argparse
from contextlib import contextmanager
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Callable


ROOT = Path(__file__).resolve().parents[2]
TOKENS_PATH = ROOT / "shared-assets" / "design-tokens.json"


def _normalize(family: str) -> str:
    return " ".join(family.split()).casefold()


def installed_font_families() -> set[str] | None:
    """Return exact Fontconfig family names, or None when they cannot be queried."""
    fc_list = shutil.which("fc-list")
    if not fc_list:
        return None
    cache = Path(tempfile.gettempdir()) / "research-output-font-cache"
    cache.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["XDG_CACHE_HOME"] = str(cache)
    completed = subprocess.run(
        [fc_list, "--format=%{family}\n"],
        capture_output=True,
        text=True,
        env=env,
        check=False,
    )
    if completed.returncode != 0:
        return None
    families: set[str] = set()
    for line in completed.stdout.splitlines():
        families.update(part.strip() for part in line.split(",") if part.strip())
    return families


def bundled_font_families(path: Path) -> set[str] | None:
    """Read exact family names declared inside one bundled font file."""
    fc_scan = shutil.which("fc-scan")
    if not fc_scan:
        return None
    cache = Path(tempfile.gettempdir()) / "research-output-font-cache"
    cache.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["XDG_CACHE_HOME"] = str(cache)
    completed = subprocess.run(
        [fc_scan, "--format=%{family}\n", str(path)],
        capture_output=True,
        text=True,
        env=env,
        check=False,
    )
    if completed.returncode != 0:
        return None
    families: set[str] = set()
    for line in completed.stdout.splitlines():
        families.update(part.strip() for part in line.split(",") if part.strip())
    return families or None


def check_font_profile(
    target: str,
    profile: str,
    system_families: set[str] | None = None,
    bundled_family_scanner: Callable[[Path], set[str] | None] | None = None,
) -> dict[str, object]:
    tokens = json.loads(TOKENS_PATH.read_text(encoding="utf-8"))
    if target not in tokens:
        raise ValueError(f"未知输出类型：{target}")
    profiles = tokens[target].get("font_profiles", {})
    if profile not in profiles:
        raise ValueError(f"未知字体档位：{target}.{profile}")
    spec = profiles[profile]

    bundled_family_scanner = bundled_family_scanner or bundled_font_families
    bundled = []
    missing: list[str] = []
    unverified: list[str] = []
    bundled_detected: set[str] = set()
    bundled_scans_complete = True
    for relative in spec.get("bundled_files", []):
        path = ROOT / "shared-assets" / relative
        exists = path.is_file() and path.stat().st_size > 0
        families = bundled_family_scanner(path) if exists else None
        bundled.append(
            {
                "path": str(path),
                "exists": exists,
                "families": sorted(families or []),
            }
        )
        if not exists:
            missing.append(f"bundled:{relative}")
        elif families is None:
            bundled_scans_complete = False
            unverified.append(f"bundled-family:{relative}")
        else:
            bundled_detected.update(families)

    required = list(spec.get("required_system_fonts", []))
    role_families = {
        value
        for key, value in spec.items()
        if key not in {"note", "bundled_files", "required_system_fonts"}
        and isinstance(value, str)
    }
    required_normalized = {_normalize(family) for family in required}
    bundled_expected = {
        family for family in role_families if _normalize(family) not in required_normalized
    }
    if bundled_scans_complete:
        bundled_normalized = {_normalize(family) for family in bundled_detected}
        for family in sorted(bundled_expected):
            if _normalize(family) not in bundled_normalized:
                missing.append(f"bundled-family:{family}")

    detected = installed_font_families() if system_families is None else system_families
    available: list[str] = []
    if detected is None:
        unverified.extend(required)
    else:
        normalized = {_normalize(family) for family in detected}
        for family in required:
            if _normalize(family) in normalized:
                available.append(family)
            else:
                missing.append(family)

    return {
        "target": target,
        "profile": profile,
        "fonts": {
            key: value
            for key, value in spec.items()
            if key not in {"note", "bundled_files", "required_system_fonts"}
        },
        "bundled": bundled,
        "required_system_fonts": required,
        "available": available,
        "missing": missing,
        "unverified": unverified,
        "render_safe": not missing and not unverified,
        "note": spec.get("note", ""),
    }


@contextmanager
def font_environment(target: str, profile: str):
    """Yield a rendering environment that exposes bundled fonts deterministically."""
    report = check_font_profile(target, profile)
    if not report["render_safe"]:
        details = report["missing"] or report["unverified"]
        raise RuntimeError(f"字体档位无法可靠渲染：{target}.{profile}，缺失或无法核验 {details}")
    env = os.environ.copy()
    bundled = report["bundled"]
    if not bundled:
        yield env
        return
    source_config = ROOT / "shared-assets" / "fonts" / "fonts.conf"
    if not source_config.is_file():
        raise RuntimeError(f"缺少字体配置：{source_config}")
    cache = Path(tempfile.gettempdir()) / "research-output-font-cache"
    cache.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="research-output-fontconfig-") as config_dir:
        portable = Path(config_dir) / "fonts.conf"
        portable.write_text(
            source_config.read_text(encoding="utf-8")
            .replace("__PROJECT_FONT_DIR__", str(ROOT / "shared-assets" / "fonts"))
            .replace("__FONT_CACHE_DIR__", str(cache)),
            encoding="utf-8",
        )
        env["FONTCONFIG_FILE"] = str(portable)
        env["XDG_CACHE_HOME"] = str(cache)
        yield env


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--target", choices=("ppt", "word"), required=True)
    parser.add_argument("--font-profile", required=True)
    args = parser.parse_args()
    try:
        report = check_font_profile(args.target, args.font_profile)
    except ValueError as error:
        print(json.dumps({"error": str(error)}, ensure_ascii=False, indent=2))
        return 2
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["render_safe"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
