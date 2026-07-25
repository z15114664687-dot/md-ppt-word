# 外部 Skill 设计研究

调研日期：2026-07-16。只吸收流程和可验证的设计约束，不复制第三方模板、品牌资产或受限制代码。

## 用户指定的五个参考

| 来源 | 值得复用 | 本项目取舍 |
|---|---|---|
| [ppt-master](https://github.com/hugohe3/ppt-master) | 设计规格锁定、资产清单、逐页 QA、先内容后视觉 | 复用分阶段流程，不采用整页 SVG 作为本项目主输出 |
| [huashu-design](https://github.com/alchaincyf/huashu-design) | 品牌资产协议、样张先行、反模板化、HTML 到可编辑 PPT 的约束 | 复用品牌与视觉验收思想，不采用其强制多风格确认和 HTML-first 管线 |
| [GordenPPTSkill](https://github.com/GordenSun/GordenPPTSkill) | 页面容量、同级字号一致、不能靠缩字塞内容、渲染 QA | 直接吸收容量门槛；不复制带非商用限制的模板资产 |
| [frontend-slides](https://github.com/zarazhangrui/frontend-slides) | 固定画布、密度模式、溢出检查、视觉预览 | 复用密度和视觉 QA；不输出单页 HTML 演示 |
| [guizang-ppt-skill](https://github.com/op7418/guizang-ppt-skill) | 叙事节奏、版式注册表、图片比例和 P0/P1 检查 | 仅借鉴思想；不复制 AGPL 代码或资产 |

## GitHub 扩展检索

| 来源 | 值得复用 | 本项目取舍 |
|---|---|---|
| [academic-pptx-skill](https://github.com/Gabberflast/academic-pptx-skill) | action title、ghost deck、论证优先于装饰、逐页引用 | 直接纳入研究型 PPT 规则 |
| [presentation-skill](https://github.com/sirilsengolraj-source/presentation-skill) | design/content/evidence/asset 计划分离、可重建工作区、可视化 QA | 简化为 QMD + tokens + 证据清单，不引入大型 JSON 渲染器 |
| [pptx-from-layouts-skill](https://github.com/tristan-mcinnis/pptx-from-layouts-skill) | 使用真实母版布局和占位符，不在模板截图上叠文字 | 纳入母版原则；模板从零构建，不复制附件布局 |
| [Anthropic PPTX skill](https://github.com/anthropics/skills/blob/main/skills/pptx/SKILL.md) | 内容 QA、逐页渲染、溢出/遮挡/模板残留检查 | 纳入交付门槛 |
| [codex-ppt-skill](https://github.com/ningzimu/codex-ppt-skill) | 大纲、风格、样张、全量四阶段确认 | 仅复用样张门槛；拒绝整页图片式 PPT 作为主输出 |

## Word 与公式实现

- [Pandoc](https://pandoc.org/MANUAL.html) 的 `reference-doc` 可继承样式、页面、页边距、页眉页脚，并将数学公式写为 Word 原生 OMML。
- [ChineseResearchLaTeX](https://github.com/huangwb8/ChineseResearchLaTeX) 使用 Markdown → HTML5 + MathML → DOCX 的两阶段管线，并对 `<m:oMath>` 做回归测试；这一点会直接用于 `md-word`。
- [docu.md](https://docu.md/) 的可取之处是将图片、表格、公式和 DOCX 导出拆成模块；本项目不复制其实现，采用可审计的本地脚本和 OOXML 验证。

## X 检索说明

已尝试在 X 公开搜索 Quarto/Pandoc 与 PPTX/DOCX 的工作流讨论，页面在隔离浏览器中停留在加载状态，未获得可核验样本。因此本版本没有从 X 帖子引入任何硬规则；后续若用户提供具体帖子或恢复登录态，可再补充使用反馈，但技术决策仍以源码和官方文档为准。
