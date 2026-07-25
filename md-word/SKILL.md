---
name: md-word
description: Use when the user needs structured research JSON, Markdown, or an already-executed static QMD converted into an editable, fixed-layout Word DOCX with native OMML formulas, directories, captions, sources, tables, images, and appendices.
---

# Markdown 研究报告转 Word

## 目标

把结构化研究内容转为可继续编辑的 `.docx`。版式规则来自本地参考样本，但不得带入样本品牌标识或专属文字。公式必须是 Word 原生 OMML，不得以截图、普通文本或 OLE 对象冒充。

## 开始前必读

1. 读取 `../shared-assets/design-tokens.json`（全部页面、字号、颜色数值的唯一事实来源）。
2. 读取 `references/markdown-contract.md`、`references/word-rules.md` 和 `references/formula-support.md`。输入是研究 JSON 时还要读取 `references/report-data-contract.md`。
3. 运行 `python3 scripts/check_environment.py --font-profile preview` 确认依赖与可重现渲染字体。
4. 运行 `python3 scripts/validate_markdown.py input.md`。

## 固定实现

采用以下管线，不用 `python-docx` 单独解析 Markdown：

```text
JSON（可选）→ 受约束 Markdown + SVG
→ Markdown/QMD
→ 语义与资源预检
→ Pandoc Markdown → HTML5 + MathML
→ Pandoc HTML → DOCX（reference.docx 按字体档位现场生成）
→ python-docx/OOXML 后处理：样式、中式章节编号、内容目录+图表目录、页眉页脚、表格和分页
→ 解包检查 TOC、图表目录、PAGE、m:oMath
→ LibreOffice 渲染 PDF/PNG，逐页复核
```

目录与章节编号不在 Pandoc 阶段做（不用 `--toc`/`--number-sections`），统一由后处理完成，避免编号文本破坏样式匹配。

HTML5 + MathML 中间层用于提高 LaTeX 公式转为 Word 原生 OMML 的稳定性。不能把 `$...$` 原样写进 Word，也不能把公式栅格化。

## 版式规则

数值一律以 `../shared-assets/design-tokens.json` 为准，此处只列规则本身：

- A4 页面与页边距按 tokens 的 `word.page` / `word.margins_mm`。
- 字体走 tokens 的 `word.font_profiles`双档位：`preview` 使用项目内置字体做可重现的本机视觉验收；`delivery` 写入交付环境的目标字体。精确字体不完整时必须报告，不能把系统替换当成验收通过。
- 正文首行缩进取 tokens 的 `word.body_first_line_indent_chars`，且只加在 Body Text/First Paragraph 样式上，表格单元格不缩进。
- 标题主色与表头底色按 tokens 的 `word.colors`。
- 左上 logo 区域继续留空，不显示框或文字；右上不放红字。
- 章节编号为中式：一级标题"第X部分："，二级标题"1、"（摘要、附录、风险提示、参考文献不编号），由后处理自动添加，Markdown 里不要手写编号。
- 图表统一编号"图表N：题"（图和表共用一套序号、全角冒号），题注在对象上方；来源与附注在对象下方。
- 双目录：摘要之后自动插入"内容目录"（TOC 域）与"图表目录"（按 Exhibit Title 样式收集的 TOC 域）；页码为 PAGE 域。交付前提示用户在 Word 中全选并更新域。
- 表格首行重复、行不跨页拆分；超宽表格优先拆表或转横向附录。
- 附录独立分页；模拟数据必须标注"模拟数据，仅作演示"。
- 国金原版页眉含栏目名、页脚含免责声明文字；当前版本页眉留白、页脚只放页码，如需还原需用户明确确认后再加开关。

## Markdown 约定

使用标准标题层级，必须含摘要、方法、回测、结论和附录。图和表共用"图表N："统一编号：

```markdown
图表1：模拟组合净值

![模拟组合净值](figures/nav.png)

来源：模拟数据，仅作演示。

图表2：模拟回测指标

| 指标 | 数值 |
|---|---:|
| 年化收益 | 10.2% |

来源：模拟数据，仅作演示。
```

完整约定见 `references/markdown-contract.md`。

## 运行

依赖安装见仓库根目录 `README.md`（Pandoc 是独立应用，python-docx 是 Python 包）。

当上游是结构化研究数据时，先生成受约束中间层：

```bash
python3 md-word/scripts/materialize_report.py input.json \
  --output output/report.md --font-profile preview
python3 md-word/scripts/validate_markdown.py output/report.md
```

转换报告（参考模板按字体档位现场生成，无需预先构建）：

```bash
python3 md-word/scripts/build_docx.py input.md --output output/report.docx
# 交付 Windows Word 时：
python3 md-word/scripts/build_docx.py input.md --output output/report.docx --font-profile delivery
```

`delivery` 档位的字体边界：如果输入含需要栅格化的 SVG，本机必须精确安装目标字体，否则构建会拒绝继续。如果只写入 DOCX 字体名而本机没有目标字体，文档可用于目标环境打开，但不得声称已完成本地视觉验收。

由 JSON 生成报告时，`materialize_report.py` 与 `build_docx.py` 必须使用同一 `--font-profile`；脚本会检查 SVG 中的档位元数据并拒绝错档栅格化。

单独刷新共享参考模板（仅用于人工检查样式）：

```bash
python3 md-word/scripts/build_reference_docx.py \
  --output shared-assets/word/research-report-reference.docx
```

校验已有 DOCX：

```bash
python3 md-word/scripts/inspect_docx.py output/report.docx \
  --expected-math 3 --font-profile preview
```

使用项目内置中文字体逐页渲染预览：

```bash
python3 md-word/scripts/render_docx_preview.py output/report.docx \
  --output-dir output/report-pages
```

不要直接调用裸 `soffice`；本项目的 `fonts.conf` 会防止 LibreOffice 把中文字体错误替换成不含中文字形的字体。

## 交付门槛

- DOCX 可打开，内容目录、图表目录、页码域、标题层级、中式章节编号、图片、表格和附录完整。
- 源 Markdown 中每个支持范围内的公式都对应 OMML `<m:oMath>`。
- DOCX 中没有残留 `$$`、`\frac` 等可见 LaTeX 源码。
- logo 与右上红字不存在，页眉区域保持留白。
- 图、表、来源与附注不孤立跨页；表格重复表头。
- 用 `scripts/render_docx_preview.py` 逐页渲染检查；只通过 XML 检查不能代表视觉完成。
- 字体预检报告必须与输出档位一致；只有 `render_safe=true` 的档位才能做对应字体的本机视觉验收。

## 边界

复杂自定义宏、未展开的 `\newcommand`、部分 `align`/`cases`/宏包命令可能无法无损转为 OMML。遇到不支持语法时先改写为标准 LaTeX，再转换；不得静默降级为图片。包含可执行代码块的 QMD 必须先运行并把表格、图片和文本结果固化，再交给本 Skill；本 Skill 不负责执行 R/Python 代码。
