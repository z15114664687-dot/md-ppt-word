# LaTeX → Word 公式规则

支持优先使用标准 Pandoc/TeX 数学语法：分式、上下标、根式、求和、积分、希腊字母、矩阵和常用关系符号。

转换采用 Markdown → HTML5 + MathML → DOCX。验收不看截图，而是解包 `.docx` 并检查：

- 行内/陈列公式存在 `<m:oMath>`。
- 陈列公式可包含 `<m:oMathPara>`。
- Word 文本中不出现未处理的 `$$`、`\frac`、`\sum`。
- OMML 数量与源 Markdown 的公式数量一致或更多（部分复杂公式可能拆分）。

以下内容需先改写：自定义 `\newcommand`、依赖宏包的命令、复杂 `align` 标签、TikZ、手工 `\tag`。不可自动改写时应报错，不得转成图片后假装完成。
