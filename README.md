# AI 研究输出双 Skill：md-word 与 qmd-ppt

把 AI 生成的研究内容转成两种可交付格式：

- **md-word/**：结构化 JSON/Markdown → 固定版式、可编辑的 Word 研究报告（借鉴本地版式样本，不带入品牌元素；公式为 Word 原生 OMML）。
- **qmd-ppt/**：Quarto QMD → 标准 16:9、可二次编辑的研究型 PPTX。
- **shared-assets/**：两个 Skill 共用的 `design-tokens.json`（所有字号/颜色/页面数值的唯一事实来源）、字体、Word 参考模板、示例与调研记录。

## 依赖安装

注意区分"独立应用"（装到系统里的命令行/桌面程序）和"Python 包"（pip 安装）：

| 依赖                                  | 类型                    | 哪条线需要                            | 安装（macOS）                       |
| ------------------------------------- | ----------------------- | ------------------------------------- | ----------------------------------- |
| Pandoc                                | 独立应用                | md→Word**必需**                | `brew install pandoc`             |
| python-docx                           | Python 包               | md→Word**必需**                | `pip3 install python-docx`        |
| Quarto                                | 独立应用（自带 Pandoc） | qmd→PPT**必需**；Word 线不需要 | `brew install quarto`             |
| LibreOffice（提供`soffice`）        | 桌面应用                | 仅本机渲染 PNG 质检，可选             | `brew install --cask libreoffice` |
| Poppler（提供`pdftoppm`）           | 命令行工具              | 仅 PDF→PNG 预览，可选                | `brew install poppler`            |
| librsvg（提供`rsvg-convert`）       | 命令行工具              | 仅当 Markdown 里有 SVG 图片时需要     | `brew install librsvg`            |
| Fontconfig（`fc-list`/`fc-scan`） | 命令行工具              | 字体 exact-family 预检必需            | `brew install fontconfig`         |

注意：Quarto 的 Homebrew cask 走 pkg 安装器、需要 sudo 密码；非交互环境可改用官方 tarball 解压到用户目录后把 `bin/quarto` 软链进 PATH（本仓库即此装法）。

说明：Word 线只要 Pandoc 就够（脚本会在缺 Pandoc 时退回使用 Quarto 内置的 Pandoc，但不要求安装 Quarto）——**例外：Markdown 里含 ```mermaid 流程图时需要 Quarto**，由它渲染成主题化 PNG；PPT 线必须用 Quarto，因为要执行 QMD 里的代码块并按任务要求原生渲染 PPTX。

环境自检：

```bash
python3 md-word/scripts/check_environment.py --font-profile preview
python3 qmd-ppt/scripts/check_environment.py --font-profile preview
```

字体分为两档：`preview` 使用项目内置字体，用于跨机器可重复质检；`delivery` 使用 `design-tokens.json` 定义的目标交付字体。字体预检按精确 family 名判断，不接受 Fontconfig 的静默替换。

## 运行命令

### Markdown → Word

```bash
# 可选：结构化数据先生成受约束 Markdown + SVG
python3 md-word/scripts/materialize_report.py \
  shared-assets/examples/sample-research-data.json \
  --output output/sample-materialized/report.md --font-profile preview

# 预检（纯标准库，无需任何依赖）
python3 md-word/scripts/validate_markdown.py shared-assets/examples/sample-research-report.md

# 转换（参考模板按字体档位现场生成）
python3 md-word/scripts/build_docx.py shared-assets/examples/sample-research-report.md \
  --output output/sample-report.docx

# 交付环境用 delivery 档位；先精确检查本机字体
python3 md-word/scripts/check_environment.py --font-profile delivery
python3 md-word/scripts/build_docx.py input.md --output output/report.docx --font-profile delivery

# 结构校验（OMML 公式、双目录、页码域、品牌禁令）
python3 md-word/scripts/inspect_docx.py output/report.docx \
  --expected-math 2 --font-profile preview

# 本机逐页渲染质检（需 LibreOffice + Poppler；使用项目内置中文字体）
python3 md-word/scripts/render_docx_preview.py output/report.docx --output-dir output/report-pages
```

### QMD → PPT

```bash
python3 qmd-ppt/scripts/validate_qmd.py shared-assets/examples/sample-deck.qmd

# 参考母版（已入库；tokens 改动后重新生成）
python3 qmd-ppt/scripts/build_reference_pptx.py

# 渲染（自动注入共享参考母版）
python3 qmd-ppt/scripts/render_qmd_ppt.py shared-assets/examples/sample-deck.qmd \
  --output output/sample-deck.pptx --preview-dir output/deck-pages \
  --font-profile preview

# 独立检查画布、母版字体、logo 安全区和禁止品牌文字
python3 qmd-ppt/scripts/inspect_pptx.py output/sample-deck.pptx \
  --font-profile preview
```

## 测试

```bash
python3 -m unittest discover -s tests -v
```

测试需要 python-docx；转换类端到端用例在缺少 Pandoc/Quarto 时自动跳过。LibreOffice 真实页面渲染默认不在单元测试中启动；需要时使用 `RUN_OFFICE_E2E=1` 运行 Word 端到端用例，并通过上述 `--preview-dir` 对 PPT 逐页验收。

## 约定

- 所有版式数值只写在 `shared-assets/design-tokens.json`，文档与脚本引用它，不重复抄写。
- 双 Skill 及 `shared-assets/` 不得出现绝对路径或仓库外运行时引用（`asset-manifest.json` 中的本地来源路径一律相对仓库根目录）。
- 交付物不含任何品牌标识：左上 logo 区留白、右上不放红字，`inspect_docx.py` 会强制检查。
- 模拟数据必须标注"模拟数据，仅作演示"。
