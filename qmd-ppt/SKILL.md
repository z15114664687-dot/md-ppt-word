---
name: qmd-ppt
description: Use when the user needs Markdown or Quarto QMD research content converted into an editable native PowerPoint deck with conclusion-led titles, charts, tables, formulas, sources, appendices, and a reusable token-driven research style.
---

# QMD 研究型 PPT

## 目标

从结构化 `.qmd` 生成标准 16:9、可二次编辑的研究型 `.pptx`。文本和普通表格必须保留为 PowerPoint 原生对象；外部图表可作为矢量图或高分辨率图片嵌入，但不得把整页做成不可编辑截图。

## 开始前必读

1. 读取 `../shared-assets/design-tokens.json`。
2. 读取 `references/qmd-contract.md`、`references/ppt-rules.md` 与 `references/slide-quality.md`。
3. 若需要理解外部 Skill 的取舍，读取 `../shared-assets/references/external-skill-review.md`。
4. 运行 `python3 scripts/check_environment.py --font-profile preview`。缺少 Quarto 时可以继续编写和静态验证 QMD，但不可声称已完成 PPTX 渲染。

## 参考优先级

- 视觉参考层：`实习任务/证金业务常见问题专项培训PPT.pptx`只提供色彩角色与字体角色，具体值只读 tokens；不复制其排版、品牌元素或演讲级字号。
- 字号：按研究报告/咨询风格的信息密度自定，数值固化在 `../shared-assets/design-tokens.json` 的 `ppt.font_sizes_pt`。
- 页面语法：Quarto/Pandoc 原生 PPTX 布局与占位符。
- 研究表达：结论式标题、证据展项、来源和口径完整。

禁止带入附件中的公司标识、专属字体、装饰图形、页面构图或公司信息。

## 固定规则

- 当前左上角 logo 区域保持完全空白：不放 logo、不画占位框、不写 `LOGO`。
- 预留区位置读取 `design-tokens.json`；页面内容从 `content_top_min_pt` 以下开始。
- 字体只读 tokens 的 `ppt.font_profiles`：`preview` 使用项目内置字体做可重现的视觉质检；`delivery` 使用目标交付字体。缺失精确字体时必须中止对应档位渲染，不能静默替换。
- 红色只用于结论、关键数字、风险或关键曲线，灰阶承载其余数据。
- 一页一个主要判断；标题优先写可验证结论，不写“行业分析”“市场现状”等空泛栏目名。
- 字号与密度全部读取 tokens 的 `ppt.font_sizes_pt`；空间不足时按“删减 → 换版式 → 拆页”处理，不得靠缩字塞入。
- 来源和脚注不得低于 tokens 的 `ppt.font_sizes_pt.minimum`。
- 图表必须标注标题、单位、统计区间、来源；模拟数据必须标注“模拟数据，仅作演示”。
- LaTeX 公式保留为 `$...$` 或 `$$...$$`，交给 Quarto/Pandoc 渲染，不转成手写 Unicode 近似式。

## 工作流

1. 提取受众、页数、核心问题、结论和证据。
2. 先写 ghost deck：只读各页结论式标题时，必须能讲完整个论证。
3. 为每页选择语义角色：封面、章节、证据、对比、图表、表格、方法、决策、风险或附录。
4. 按 `references/qmd-contract.md` 编写 QMD；正文、图表和表格都从原始证据生成。
5. 运行静态预检：

```bash
python3 qmd-ppt/scripts/validate_qmd.py path/to/deck.qmd
```

6. 渲染（自动注入共享参考母版；母版缺失时先用 `build_reference_pptx.py` 生成）：

```bash
python3 qmd-ppt/scripts/render_qmd_ppt.py path/to/deck.qmd \
  --output output/deck.pptx --preview-dir output/deck-pages \
  --font-profile preview
```

交付档位在构建前先运行：

```bash
python3 qmd-ppt/scripts/check_environment.py --font-profile delivery
python3 qmd-ppt/scripts/render_qmd_ppt.py path/to/deck.qmd \
  --output output/deck-delivery.pptx --font-profile delivery
```

7. 逐页对照 `references/slide-quality.md`：先看 `validate_qmd.py` 的版式门禁有没有 error（顶层图未包 columns、columns 后跟段落等），再逐张 PNG 目检四种“简陋页”典型、溢出、字号、来源、logo 安全区、图表可读性和整套节奏。**门禁全绿 ≠ 版式合格**，逐页 PNG 目检是交付前必过项。

## 交付门槛

- 输出为标准 16:9 PPTX，可正常打开。
- 文字仍为文本对象，普通表格仍为表格对象。
- 没有模板残留、空占位符、可见 logo 占位框或未经授权的品牌标识。
- 标题字号、正文层级和最小字号符合共享 tokens。
- 所有数据页均有来源，所有公式可读。
- 至少完成一次渲染—检查—修复循环；未实际渲染时必须明确说明。
- 字体预检的 `render_safe` 必须为真，且必须与输出使用同一档位。
- `validate_qmd.py` 的版式门禁无 error（顶层图/表未包 columns、columns 后跟顶层段落等）；warning 已逐条评估。
- 逐页 PNG 已按 `references/slide-quality.md` 目检，无“简陋页”四种典型。

## 提升密度：卡片注入

Quarto 只填占位符版式，做不了高密度卡片。用 `references/qmd-contract.md`「卡片注入」的 raw block 在需要的页加**原生可编辑**的卡片：`{=ppt-kpi}` KPI 数字块、`{=ppt-cards}` 卡片网格（3–4 并列要点/策略/风险）、`{=ppt-flow}` 流程卡（①→②→③ 步骤+箭头，替代手画流程图/SmartArt）、`{=ppt-compare}` 对比卡（改前 vs 改后，末行红头强调）、`{=ppt-takeaway}` 结论条；`render_qmd_ppt.py` 渲染后自动调 `inject_cards.py` 注入。**写页前先查 `references/slide-quality.md` 第〇节「内容 → 版式决策表」对号入座**——选对版式比事后修图省力。数据图表页仍走 Quarto 原生（真图表、可编辑），卡片补在数字、并列项、结论这类 Quarto 做得寡淡的地方。示范见 `shared-assets/examples/sample-deck.qmd`。

> 定位：qmd-ppt 的强项是**数据图表为主的研究报告**（真图表 + 文本/表格原生可编辑 + 一条命令可重现）。高密度纯观点/框架页（多卡片网格、图标阵）仍是 ppt-master / 手画 SVG 的强项——两条链路分工，一份 deck 可混排，不要让 qmd-ppt 去追 ppt-master 的密度天花板。

## 原生能力边界

Quarto 生成的文本和表格可编辑，但 R/Python 绘制的图表通常作为图片或矢量图嵌入，不等于 PowerPoint 原生数据图表。若用户要求双击图表编辑数据系列，需要另加原生图表后处理器，不得假装 QMD 已实现。注入的卡片（kpi/cards/flow/compare/takeaway）是原生 shape、可编辑，但它是装饰/强调层，不承载可重算的数据系列。
