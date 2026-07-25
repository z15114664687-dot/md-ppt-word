#!/usr/bin/env python3
"""Materialize validated research JSON as deterministic Markdown and SVG."""

from __future__ import annotations

import argparse
import html
import json
import math
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
TOKENS_PATH = ROOT / "shared-assets" / "design-tokens.json"
REQUIRED_FIELDS = (
    "title",
    "data_status",
    "summary",
    "method",
    "periods",
    "series",
    "metrics",
    "conclusion",
    "appendix",
    "sources",
)


def _text(value: object, field: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field} 必须是非空字符串")
    return " ".join(value.split())


def _text_list(value: object, field: str) -> list[str]:
    if not isinstance(value, list) or not value:
        raise ValueError(f"{field} 必须是非空字符串数组")
    return [_text(item, f"{field}[{index}]") for index, item in enumerate(value)]


def _validate(payload: object) -> dict[str, object]:
    if not isinstance(payload, dict):
        raise ValueError("顶层 JSON 必须是对象")
    missing = [field for field in REQUIRED_FIELDS if field not in payload]
    if missing:
        raise ValueError("缺少必需字段：" + ", ".join(missing))

    data_status = _text(payload["data_status"], "data_status")
    if data_status not in {"actual", "simulated"}:
        raise ValueError("data_status 只能是 actual 或 simulated")

    periods = _text_list(payload["periods"], "periods")
    if len(periods) < 2:
        raise ValueError("periods 至少需要两个时点")

    raw_series = payload["series"]
    if not isinstance(raw_series, list) or not raw_series:
        raise ValueError("series 必须是非空数组")
    series: list[dict[str, object]] = []
    labels: set[str] = set()
    for index, item in enumerate(raw_series):
        if not isinstance(item, dict):
            raise ValueError(f"series[{index}] 必须是对象")
        label = _text(item.get("label"), f"series[{index}].label")
        values = item.get("values")
        if label in labels:
            raise ValueError(f"series.label 不得重复：{label}")
        if not isinstance(values, list) or len(values) != len(periods):
            raise ValueError(f"series[{index}].values 长度必须与 periods 一致")
        numeric_values: list[float] = []
        for value_index, value in enumerate(values):
            if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
                raise ValueError(f"series[{index}].values[{value_index}] 必须是有限数值")
            numeric_values.append(float(value))
        labels.add(label)
        series.append({"label": label, "values": numeric_values})

    raw_metrics = payload["metrics"]
    if not isinstance(raw_metrics, list) or not raw_metrics:
        raise ValueError("metrics 必须是非空数组")
    metrics: list[dict[str, str]] = []
    for index, item in enumerate(raw_metrics):
        if not isinstance(item, dict):
            raise ValueError(f"metrics[{index}] 必须是对象")
        metrics.append(
            {
                "name": _text(item.get("name"), f"metrics[{index}].name"),
                "value": _text(item.get("value"), f"metrics[{index}].value"),
            }
        )

    return {
        "title": _text(payload["title"], "title"),
        "data_status": data_status,
        "summary": _text_list(payload["summary"], "summary"),
        "method": _text_list(payload["method"], "method"),
        "periods": periods,
        "series": series,
        "metrics": metrics,
        "conclusion": _text_list(payload["conclusion"], "conclusion"),
        "appendix": _text_list(payload["appendix"], "appendix"),
        "sources": _text_list(payload["sources"], "sources"),
    }


def _svg(payload: dict[str, object], tokens: dict[str, object], font_profile: str) -> str:
    width, height = 960, 480
    left, right, top, bottom = 76, 36, 86, 62
    plot_width = width - left - right
    plot_height = height - top - bottom
    periods = payload["periods"]
    series = payload["series"]
    assert isinstance(periods, list) and isinstance(series, list)
    all_values = [value for item in series for value in item["values"]]
    minimum, maximum = min(all_values), max(all_values)
    span = maximum - minimum
    pad = span * 0.08 if span else max(abs(maximum) * 0.08, 1.0)
    y_min, y_max = minimum - pad, maximum + pad

    colors = tokens["word"]["colors"]
    profile = tokens["word"]["font_profiles"][font_profile]
    palette = colors["chart_series"]
    font_family = html.escape(
        f'{profile["heading_zh"]}, {profile["heading_latin"]}',
        quote=True,
    )

    def x_position(index: int) -> float:
        return left + plot_width * index / (len(periods) - 1)

    def y_position(value: float) -> float:
        return top + plot_height * (y_max - value) / (y_max - y_min)

    lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}" data-font-profile="{font_profile}">',
        f'<rect width="{width}" height="{height}" fill="{colors["background"]}"/>',
        f'<g font-family="{font_family}" fill="{colors["text"]}">',
        f'<text x="{left}" y="36" font-size="24" font-weight="700">策略与基准序列</text>',
    ]
    for tick in range(5):
        ratio = tick / 4
        value = y_max - (y_max - y_min) * ratio
        y = top + plot_height * ratio
        lines.append(
            f'<line x1="{left}" y1="{y:.2f}" x2="{width - right}" y2="{y:.2f}" '
            f'stroke="{colors["chart_grid"]}" stroke-width="1"/>'
        )
        lines.append(
            f'<text x="{left - 10}" y="{y + 5:.2f}" font-size="13" text-anchor="end" '
            f'fill="{colors["secondary_text"]}">{value:.3g}</text>'
        )

    label_step = max(1, math.ceil(len(periods) / 7))
    label_indexes = set(range(0, len(periods), label_step)) | {len(periods) - 1}
    for index, period in enumerate(periods):
        x = x_position(index)
        if index in label_indexes:
            lines.append(
                f'<text x="{x:.2f}" y="{height - 25}" font-size="13" text-anchor="middle" '
                f'fill="{colors["secondary_text"]}">{html.escape(period)}</text>'
            )

    for series_index, item in enumerate(series):
        color = palette[series_index % len(palette)]
        points = " ".join(
            f"{x_position(index):.2f},{y_position(value):.2f}"
            for index, value in enumerate(item["values"])
        )
        lines.append(
            f'<polyline points="{points}" fill="none" stroke="{color}" stroke-width="3" '
            'stroke-linecap="round" stroke-linejoin="round"/>'
        )
        for index, value in enumerate(item["values"]):
            lines.append(
                f'<circle cx="{x_position(index):.2f}" cy="{y_position(value):.2f}" r="3.5" fill="{color}"/>'
            )
        legend_x = left + series_index * 190
        lines.append(f'<line x1="{legend_x}" y1="62" x2="{legend_x + 28}" y2="62" stroke="{color}" stroke-width="3"/>')
        lines.append(
            f'<text x="{legend_x + 36}" y="67" font-size="14">{html.escape(item["label"])}</text>'
        )
    lines.extend(["</g>", "</svg>", ""])
    return "\n".join(lines)


def _markdown(payload: dict[str, object]) -> str:
    source_items = list(payload["sources"])
    if payload["data_status"] == "simulated" and not any("模拟数据，仅作演示" in item for item in source_items):
        source_items.append("模拟数据，仅作演示")
    source = "来源：" + "；".join(source_items) + "。"

    def paragraphs(items: object) -> str:
        assert isinstance(items, list)
        return "\n\n".join(items)

    def table_cell(value: str) -> str:
        return value.replace("|", "\\|")

    metric_rows = [
        f'| {table_cell(item["name"])} | {table_cell(item["value"])} |'
        for item in payload["metrics"]
    ]
    blocks = [
        f'# {payload["title"]}',
        '## 摘要',
        paragraphs(payload["summary"]),
        '## 方法',
        paragraphs(payload["method"]),
        '## 回测',
        '图表1：策略与基准序列',
        '![策略与基准序列](figures/series.svg)',
        source,
        '图表2：核心指标',
        '\n'.join(['| 指标 | 数值 |', '|---|---:|', *metric_rows]),
        source,
        '## 结论',
        paragraphs(payload["conclusion"]),
        '# 附录',
        paragraphs(payload["appendix"]),
    ]
    return "\n\n".join(blocks) + "\n"


def materialize(source: Path, output: Path, font_profile: str | None = None) -> dict[str, object]:
    source = Path(source)
    output = Path(output)
    payload = _validate(json.loads(source.read_text(encoding="utf-8")))
    tokens = json.loads(TOKENS_PATH.read_text(encoding="utf-8"))
    word_tokens = tokens["word"]
    font_profile = font_profile or word_tokens["default_font_profile"]
    if font_profile not in word_tokens["font_profiles"]:
        raise ValueError(f"未知 Word 字体档位：{font_profile}")
    output.parent.mkdir(parents=True, exist_ok=True)
    figures = output.parent / "figures"
    figures.mkdir(parents=True, exist_ok=True)
    chart = figures / "series.svg"
    chart.write_text(_svg(payload, tokens, font_profile), encoding="utf-8")
    output.write_text(_markdown(payload), encoding="utf-8")
    return {
        "source": str(source),
        "output": str(output),
        "chart": str(chart),
        "data_status": payload["data_status"],
        "font_profile": font_profile,
        "exhibits": 2,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("json", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    word_tokens = json.loads(TOKENS_PATH.read_text(encoding="utf-8"))["word"]
    parser.add_argument(
        "--font-profile",
        choices=tuple(word_tokens["font_profiles"]),
        default=word_tokens["default_font_profile"],
    )
    args = parser.parse_args()
    try:
        print(
            json.dumps(
                materialize(args.json, args.output, args.font_profile),
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    except (OSError, json.JSONDecodeError, ValueError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
