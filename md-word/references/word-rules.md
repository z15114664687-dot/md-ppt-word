# Word 版式规则

所有数值（页面、页边距、字号、颜色、行距、缩进）以 `../../shared-assets/design-tokens.json` 的 `word` 节为唯一事实来源，本文件只描述规则，不重复抄写数值；两处不一致时以 tokens 为准并修正文档。

## 页面和层级

- 页面与页边距：tokens `word.page` / `word.margins_mm`（对齐国金参考稿实测）。
- 字号层级：tokens `word.font_sizes_pt`（封面、一级、二级、正文、题注、来源）。
- 字体双档位：tokens `word.font_profiles`。`preview` 使用项目内置字体保证本机渲染可重现；`delivery` 使用目标交付字体。具体字体名和要求不在文档中重复维护。
- 正文首行缩进取 tokens `word.body_first_line_indent_chars`。缩进只配置在 Body Text / First Paragraph 样式上，Normal 不缩进，因此表格单元格保持顶格。
- 正文行距与段后间距：tokens `word.line_spacing_pt` / `word.space_after_pt`；标题与下一段保持同页。
- 一级标题段前分页；附录独立分页。

## 章节编号（中式，后处理添加）

- 一级标题："第一部分：xxx"；二级标题："1、xxx"，每个部分内重新从 1 计数。
- 摘要、附录、风险提示、参考文献、目录类标题不编号。
- 编号在 DOCX 后处理阶段追加到标题文本，Pandoc 不使用 `--number-sections`，Markdown 源不手写编号。

## 页眉页脚

页眉不显示 logo、tab、红字或占位框，顶部版心自然保留品牌安全区；页脚居中使用 PAGE 域，不写死页码。国金原版页眉含栏目名、页脚含免责声明文字——当前版本不还原这两处，如需还原须经用户确认后再实现为开关。

## 图片和表格

- 图和表统一"图表N：题"编号（Exhibit Title 样式，左对齐），来源和注释（Source Note 样式）在对象下方。
- 图片按版心等比例缩放，不拉伸；接受 PNG/JPEG/SVG。GIF/WebP 必须在输入前转为兼容 PNG，预检会提示。本地 SVG 由构建脚本用 `rsvg-convert`（librsvg）自动栅格化后再进 Pandoc；栅格化使用所选字体档位，缺少精确字体时拒绝继续。
- 表头底色深蓝、白字（tokens `word.colors.table_header`）；正文白底黑字；边框细而克制。
- 表头跨页重复，整行不得跨页拆分。
- 来源行与上一个对象保持同页，避免孤立到下一页。

## 目录和域

- 双目录：摘要之后独立分页插入"内容目录"（`TOC \o "1-3"`）与"图表目录"（`TOC \t "Exhibit Title,1"`，按题注样式收集），对齐国金参考稿的内容目录+图表目录结构。
- 页码和未来交叉引用都应使用 Word 域。自动生成后设置 `updateFields`，但不同 Word/LibreOffice 版本更新行为不完全一致，最终交付前仍需在 Microsoft Word 中更新全部域。
