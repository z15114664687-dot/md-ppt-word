# 研究报告 JSON 数据契约

这个中间层把“开放式研究/分析”与“确定性排版”分开。模型或上游程序只生成 JSON，`materialize_report.py` 负责验证、统一模拟数据标注，并生成可由 `validate_markdown.py` 检查的 Markdown 和 SVG。

## 必需字段

```json
{
  "title": "报告标题",
  "data_status": "actual",
  "summary": ["摘要段落"],
  "method": ["方法段落"],
  "periods": ["2025Q1", "2025Q2"],
  "series": [
    {"label": "策略净值", "values": [1.0, 1.05]}
  ],
  "metrics": [
    {"name": "累计收益", "value": "5.0%"}
  ],
  "conclusion": ["结论段落"],
  "appendix": ["口径与参数"],
  "sources": ["Wind，截至2025年6月30日"]
}
```

规则：

- `data_status` 只能是 `actual` 或 `simulated`。`simulated` 会强制补入“模拟数据，仅作演示”。
- 所有段落数组和 `sources` 不得为空。
- `periods` 至少两项；每个 `series.values` 的长度必须与它一致，且只能包含有限数值。
- `metrics` 的值使用已定稿展示字符串，避免下游脚本猜测百分比、单位或小数位。
- 输出固定使用两个连续展项：序列图和指标表。更复杂报告应扩展数据契约和对应测试，不要手改生成的 Markdown 来绕过校验。

## 运行

```bash
python3 md-word/scripts/materialize_report.py input.json \
  --output output/report.md --font-profile preview
python3 md-word/scripts/validate_markdown.py output/report.md
python3 md-word/scripts/build_docx.py output/report.md --output output/report.docx
```

`materialize_report.py` 和 `build_docx.py` 必须使用同一字体档位。生成的 SVG 会写入档位元数据，不一致时 DOCX 构建会拒绝继续，避免图表文字与正文字体错档。
