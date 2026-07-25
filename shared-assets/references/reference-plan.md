# AI 研究输出双 Skill 参考方案

## 三层目录

```text
建投证金agent/
├── qmd-ppt/        QMD → 研究型可编辑 PPTX
├── md-word/        Markdown → 固定版式 Word 研究报告
└── shared-assets/  字体、设计 tokens、Word 参考模板、示例和调研记录
```

两个 Skill 共用 `design-tokens.json`，避免同一项目里出现两套字号、色彩和 logo 规则。

## QMD → PPT

- 画布比例取 `design-tokens.json` 的 `ppt.canvas`。
- 视觉参考层只继承本地培训稿的色彩角色与字体角色，实际值取 `ppt.colors` / `ppt.font_profiles`，不复制排版。
- 字号按研究报告/咨询风格的信息密度设定，具体值取 `ppt.font_sizes_pt`，不继承培训稿的演讲级大字号。
- 左上 logo 区域为空白安全区，无框无文字；未来只替换命名锚点。
- 输出坚持 Quarto/Pandoc 原生 PPTX；不采用整页图片式 PPT。
- 一页一个判断，按“结论 → 证据 → 解释/业务含义 → 来源”组织。
- 参考母版 `shared-assets/ppt/research-reference.pptx` 由 `build_reference_pptx.py` 从 Pandoc 内置模板 + tokens 生成，渲染脚本自动注入；已完成 Skill、静态预检、示例 QMD 与母版。

## Markdown → Word

- 页面、页边距、字号、颜色、行距、缩进全部以 `design-tokens.json` 的 `word` 节为唯一事实来源。
- 字体双档位取 `word.font_profiles`：`preview` 用内置字体做本机渲染质检，`delivery` 使用目标交付字体。
- 正文首行缩进取 tokens；页眉不含参考样本的品牌元素；页脚使用 PAGE 域。
- 图和表统一"图表N：题"编号（对齐国金参考稿）；题注在上、来源与附注在下；表头重复、行不拆分。
- 双目录：内容目录 + 图表目录，均为 Word TOC 域，由后处理在摘要之后插入；附录独立分页。
- 章节编号中式（第X部分：/ 1、），由后处理添加，不用 Pandoc `--number-sections`。
- 公式走 `Markdown → HTML5 + MathML → DOCX OMML`，并解包检查 `<m:oMath>`；不接受图片公式或残留 LaTeX 源码。
- 已完成参考 DOCX、静态验证器、OMML/双目录/PAGE 检查器、中文渲染脚本和示例报告。完整 Markdown → DOCX 转换需要安装 Pandoc（独立应用，见仓库 README.md）；Word 线不需要 Quarto。

## 外部 Skill 的复用原则

- 复用：action title、ghost deck、页面容量、证据/资产计划、真实母版布局、逐页视觉 QA、公式 XML 回归测试。
- 调整：把 HTML-first 或 PptxGenJS-first 做法改为 QMD/Quarto 原生输出。
- 拒绝：整页图片式 PPT、复制第三方模板、复制 AGPL/非商用资产、仅凭 X 帖子决定技术方案。

详细来源和许可证边界见 `external-skill-review.md`。
