import importlib.util
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "qmd-ppt" / "scripts"
SHARED_SCRIPTS = ROOT / "shared-assets" / "scripts"
SCRIPT = SCRIPTS / "validate_qmd.py"
TOKENS = ROOT / "shared-assets" / "design-tokens.json"
REFERENCE_PPTX = ROOT / "shared-assets" / "ppt" / "research-reference.pptx"


def load_validator():
    spec = importlib.util.spec_from_file_location("validate_qmd", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def load_qmd_script(name: str):
    path = SCRIPTS / f"{name}.py"
    if str(SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SCRIPTS))
    if str(SHARED_SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SHARED_SCRIPTS))
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def load_shared_module(name: str):
    path = SHARED_SCRIPTS / f"{name}.py"
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class QmdPptSkillTests(unittest.TestCase):
    def test_typography_follows_research_density(self):
        tokens = json.loads(TOKENS.read_text(encoding="utf-8"))
        sizes = tokens["ppt"]["font_sizes_pt"]
        self.assertEqual(sizes["cover_title"], 32)
        self.assertEqual(sizes["section_title"], 26)
        self.assertEqual(sizes["slide_title"], 20)
        self.assertEqual(sizes["body"], 11)
        self.assertEqual(sizes["minimum"], 9)
        # 研报/咨询密度：正文不得回到培训稿的演讲级大字号
        self.assertLessEqual(sizes["body"], 12)
        self.assertLessEqual(sizes["slide_title"], 24)

    def test_ppt_tokens_define_portable_and_delivery_font_profiles(self):
        ppt = json.loads(TOKENS.read_text(encoding="utf-8"))["ppt"]
        self.assertEqual(ppt["default_font_profile"], "preview")
        preview = ppt["font_profiles"]["preview"]
        delivery = ppt["font_profiles"]["delivery"]
        self.assertEqual(preview["zh"], "Noto Sans CJK SC")
        self.assertIn("fonts/NotoSansCJKsc-Regular.otf", preview["bundled_files"])
        self.assertEqual(delivery["zh"], "MiSans")
        self.assertEqual(delivery["latin"], "Arial")
        self.assertEqual(delivery["display_zh"], "方正兰亭粗黑简体")
        self.assertIn("MiSans", delivery["required_system_fonts"])

    def test_font_preflight_distinguishes_bundled_preview_from_missing_delivery_fonts(self):
        preflight = load_shared_module("font_preflight")
        preview = preflight.check_font_profile(
            "ppt",
            "preview",
            system_families=set(),
            bundled_family_scanner=lambda path: {"Noto Sans CJK SC"},
        )
        delivery = preflight.check_font_profile(
            "ppt",
            "delivery",
            system_families={"Arial"},
            bundled_family_scanner=lambda path: set(),
        )

        self.assertTrue(preview["render_safe"])
        self.assertEqual(preview["missing"], [])
        self.assertFalse(delivery["render_safe"])
        self.assertEqual(set(delivery["missing"]), {"MiSans", "方正兰亭粗黑简体"})

        mismatched = preflight.check_font_profile(
            "ppt",
            "preview",
            system_families=set(),
            bundled_family_scanner=lambda path: {"Unexpected Font"},
        )
        self.assertIn("bundled-family:Noto Sans CJK SC", mismatched["missing"])

    def test_environment_report_includes_selected_font_profile(self):
        checker = load_qmd_script("check_environment")
        report = checker.build_report(
            "preview",
            tool_lookup=lambda name: f"/usr/bin/{name}",
            system_families=set(),
            bundled_family_scanner=lambda path: {"Noto Sans CJK SC"},
        )
        self.assertEqual(report["fonts"]["profile"], "preview")
        self.assertTrue(report["fonts"]["render_safe"])
        self.assertTrue(report["ready"])

    def test_logo_slot_is_reserved_but_blank(self):
        tokens = json.loads(TOKENS.read_text(encoding="utf-8"))
        slot = tokens["ppt"]["logo_slot"]
        self.assertFalse(slot["enabled"])
        self.assertFalse(slot["visible_placeholder"])
        self.assertGreaterEqual(slot["content_top_min_pt"], slot["y_pt"] + slot["height_pt"])

    def test_validator_accepts_research_slide_with_source(self):
        validator = load_validator()
        qmd = """---
title: 行业研究示例
format: pptx
---

## 行业盈利拐点已在二季度出现

- 收入同比改善，毛利率企稳
- 关键驱动来自需求恢复与供给收缩

来源：Wind，公司公告
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.qmd"
            path.write_text(qmd, encoding="utf-8")
            result = validator.validate(path)
        self.assertEqual(result["errors"], [])

    def test_validator_rejects_wrong_format_and_overdense_slide(self):
        validator = load_validator()
        bullets = "\n".join(f"- 第 {idx} 条内容" for idx in range(1, 10))
        qmd = f"""---
title: 错误示例
format: html
---

## 分析

{bullets}
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.qmd"
            path.write_text(qmd, encoding="utf-8")
            result = validator.validate(path)
        joined = "\n".join(result["errors"])
        self.assertIn("format 必须为 pptx", joined)
        self.assertIn("9 个列表项", joined)

    def test_validator_requires_source_on_percentage_only_slide(self):
        """A slide whose only data is percentages still needs a 来源 line.

        Regression: the has_data pattern used to end in `\\b`, which can never
        hold after `%` (a percentage is always followed by 、。， space or
        end-of-line, all non-word). Percentage-only slides therefore skipped
        the source check entirely — the most common case in a research deck.
        """
        validator = load_validator()
        qmd = """---
title: 百分比数据页
format: pptx
---

## 毛利率连续两个季度改善

- 毛利率由 14.1% 改善到 18.7%
- 收入同比由 -8.2% 转正至 +6.4%
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.qmd"
            path.write_text(qmd, encoding="utf-8")
            result = validator.validate(path)
        self.assertIn("缺少来源", "\n".join(result["errors"]))

    def test_validator_requires_source_inside_percentage_card(self):
        validator = load_validator()
        cases = (("", True), ("来源：Wind，作者测算", False))
        for source_line, should_error in cases:
            with self.subTest(source_line=source_line), tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / "deck.qmd"
                path.write_text(
                    f"""---
title: 卡片数据页
format: pptx
---

## 组合三年表现一览

```{{=ppt-kpi}}
9.8% | 年化收益
55.2% | 最大回撤
{source_line}
```
""",
                    encoding="utf-8",
                )
                result = validator.validate(path)
            joined = "\n".join(result["errors"])
            if should_error:
                self.assertIn("缺少来源", joined)
            else:
                self.assertNotIn("缺少来源", joined)

    def test_validator_requires_card_source_on_its_own_line(self):
        validator = load_validator()
        qmd = """---
title: 卡片来源行
format: pptx
---

## 组合收益仍需补齐可核验来源

```{=ppt-kpi}
9.8% | 年化收益，来源：待补
```
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.qmd"
            path.write_text(qmd, encoding="utf-8")
            result = validator.validate(path)
        self.assertIn("缺少来源", "\n".join(result["errors"]))

    def test_validator_rejects_empty_card_source_line(self):
        validator = load_validator()
        qmd = """---
title: 空来源行
format: pptx
---

## 卡片数据需要非空来源

```{=ppt-kpi}
9.8% | 年化收益
来源：
备注 | 待补
```
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.qmd"
            path.write_text(qmd, encoding="utf-8")
            result = validator.validate(path)
        self.assertIn("缺少来源", "\n".join(result["errors"]))

    def test_validator_keeps_raw_cards_separate_on_duplicate_titles(self):
        validator = load_validator()
        qmd = """---
title: 同名页数据归属
format: pptx
---

## 组合表现一览

```{=ppt-kpi}
9.8% | 年化收益
```

## 组合表现一览

```{=ppt-kpi}
8.1% | 年化收益
来源：Wind，作者测算
```
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.qmd"
            path.write_text(qmd, encoding="utf-8")
            result = validator.validate(path)
        self.assertIn("第 1 页“组合表现一览”卡片包含数据但缺少来源", result["errors"])
        self.assertIn("第 2 页标题“组合表现一览”与前页重复，无法唯一匹配卡片", result["errors"])

    def test_validator_does_not_read_bpm_as_basis_points(self):
        """The ASCII suffix guard keeps 120bpm from becoming 120bp."""
        validator = load_validator()
        qmd = """---
title: 非金融单位
format: pptx
---

## 静息心率维持在正常区间

- 静息心率 120bpm 属于正常范围
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.qmd"
            path.write_text(qmd, encoding="utf-8")
            result = validator.validate(path)
        self.assertEqual(result["errors"], [])

    def test_validator_requires_source_on_basis_points_followed_by_chinese(self):
        validator = load_validator()
        for value in ("5bp回落", "25bps上升"):
            with self.subTest(value=value), tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / "deck.qmd"
                path.write_text(
                    f"""---
title: 基点数据页
format: pptx
---

## 利差变化需要来源验证

- 利差{value}
""",
                    encoding="utf-8",
                )
                result = validator.validate(path)
            self.assertIn("缺少来源", "\n".join(result["errors"]))

    def test_validator_ignores_slide_markers_inside_fenced_code(self):
        validator = load_validator()
        qmd = """---
title: 代码示例
format: pptx
---

## 只有这一页是真实幻灯片

```python
## 这是 Python 注释，不是幻灯片
print("hello")
```
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.qmd"
            path.write_text(qmd, encoding="utf-8")
            result = validator.validate(path)
        self.assertEqual(result["slides"], 1)

    def test_validator_ignores_card_examples_inside_longer_fence(self):
        validator = load_validator()
        qmd = """---
title: 卡片语法示例
format: pptx
---

## 以下代码只用于说明卡片语法

````markdown
## 这是文档中的伪幻灯片标题
```{=ppt-kpi}
9.8% | 年化收益
```
````
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.qmd"
            path.write_text(qmd, encoding="utf-8")
            result = validator.validate(path)
        self.assertNotIn("缺少来源", "\n".join(result["errors"]))

    def test_card_parser_ignores_examples_inside_longer_fence(self):
        injector = load_qmd_script("inject_cards")
        qmd = """---
title: 卡片语法示例
format: pptx
---

## 以下代码只用于说明卡片语法

````markdown
## 这是文档中的伪幻灯片标题
```{=ppt-kpi}
9.8% | 年化收益
```
````
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.qmd"
            path.write_text(qmd, encoding="utf-8")
            cards = injector.parse_cards(path)
        self.assertEqual(cards, {})

    def test_takeaway_parser_separates_source_line(self):
        injector = load_qmd_script("inject_cards")
        qmd = """---
title: 结论条来源
format: pptx
---

## 三重信号同向才能确认拐点

```{=ppt-takeaway}
供需、盈利、估值同向才能写成结论
来源：Wind，公司公告
```
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.qmd"
            path.write_text(qmd, encoding="utf-8")
            cards = injector.parse_cards(path)
        self.assertEqual(
            cards["三重信号同向才能确认拐点"]["takeaway"],
            [("供需、盈利、估值同向才能写成结论", "Wind，公司公告")],
        )


class ReferencePptxTests(unittest.TestCase):
    @unittest.skipUnless(shutil.which("pandoc") or shutil.which("quarto"), "需要 Pandoc/Quarto 构建参考母版")
    def test_built_reference_applies_token_canvas_and_title_safe_zone(self):
        builder = load_qmd_script("build_reference_pptx")
        tokens = json.loads(TOKENS.read_text(encoding="utf-8"))["ppt"]
        with tempfile.TemporaryDirectory() as tmp:
            reference = Path(tmp) / "reference.pptx"
            builder.build(reference, "preview")
            with zipfile.ZipFile(reference) as archive:
                presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
                geometry_xml = {
                    name: archive.read(name)
                    for name in archive.namelist()
                    if name.startswith("ppt/slideMasters/slideMaster")
                    or name.startswith("ppt/slideLayouts/slideLayout")
                }

        p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
        a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        slide_size = presentation.find(f"{{{p_ns}}}sldSz")
        self.assertIsNotNone(slide_size)
        self.assertEqual(int(slide_size.get("cx")), round(tokens["canvas"]["width_in"] * 914400))
        self.assertEqual(int(slide_size.get("cy")), round(tokens["canvas"]["height_in"] * 914400))
        self.assertAlmostEqual(
            int(slide_size.get("cx")) / int(slide_size.get("cy")),
            16 / 9,
            places=6,
        )

        minimum_y = round(tokens["logo_slot"]["content_top_min_pt"] * 12700)
        title_offsets = []
        for xml in geometry_xml.values():
            root = ET.fromstring(xml)
            for shape in root.iter(f"{{{p_ns}}}sp"):
                placeholder = shape.find(f".//{{{p_ns}}}ph")
                if placeholder is None or placeholder.get("type") != "title":
                    continue
                offset = shape.find(f".//{{{a_ns}}}xfrm/{{{a_ns}}}off")
                if offset is not None:
                    title_offsets.append(int(offset.get("y")))
        self.assertTrue(title_offsets)
        self.assertTrue(all(y >= minimum_y for y in title_offsets), title_offsets)

    @unittest.skipUnless(shutil.which("pandoc") or shutil.which("quarto"), "需要 Pandoc/Quarto 构建参考母版")
    def test_pptx_inspector_checks_profile_geometry_and_visible_branding(self):
        builder = load_qmd_script("build_reference_pptx")
        inspector = load_qmd_script("inspect_pptx")
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            clean = root / "clean.pptx"
            branded = root / "branded.pptx"
            builder.build(clean, "preview")
            clean_result = inspector.inspect(clean, "preview")
            with zipfile.ZipFile(clean) as source, zipfile.ZipFile(branded, "w") as target:
                for info in source.infolist():
                    target.writestr(info, source.read(info.filename))
                target.writestr(
                    "ppt/slides/slide999.xml",
                    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                    '<p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/>'
                    '</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="1"/>'
                    '<a:ext cx="100" cy="100"/></a:xfrm></p:spPr>'
                    '<p:txBody><a:p><a:r><a:t>国金证券</a:t>'
                    "</a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>",
                )
            branded_result = inspector.inspect(branded, "preview")

        self.assertEqual(clean_result["errors"], [])
        self.assertIn("检测到禁止的可见品牌文字：国金证券", branded_result["errors"])
        self.assertTrue(any("最终幻灯片标题进入 logo 安全区" in error for error in branded_result["errors"]))

    def test_reference_master_exists_and_follows_tokens(self):
        self.assertTrue(REFERENCE_PPTX.exists(), "缺少参考母版，先运行 build_reference_pptx.py")
        tokens = json.loads(TOKENS.read_text(encoding="utf-8"))["ppt"]
        profile = tokens["default_font_profile"]
        fonts = tokens["font_profiles"][profile]
        with zipfile.ZipFile(REFERENCE_PPTX) as archive:
            theme = archive.read("ppt/theme/theme1.xml").decode("utf-8")
            master = archive.read("ppt/slideMasters/slideMaster1.xml").decode("utf-8")
        theme_root = ET.fromstring(theme)
        a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        hans_faces = [
            element.get("typeface")
            for element in theme_root.iter(f"{{{a_ns}}}font")
            if element.get("script") == "Hans"
        ]
        self.assertIn(fonts["zh"], theme)
        self.assertIn(fonts["latin"], theme)
        self.assertTrue(hans_faces)
        self.assertEqual(set(hans_faces), {fonts["zh"]})
        self.assertIn(tokens["colors"]["primary"].lstrip("#"), theme)
        self.assertIn(f'sz="{round(tokens["font_sizes_pt"]["slide_title"] * 100)}"', master)
        self.assertIn(f'sz="{round(tokens["font_sizes_pt"]["body"] * 100)}"', master)
        self.assertNotIn("国金", theme + master)
        with zipfile.ZipFile(REFERENCE_PPTX) as archive:
            layouts = "".join(
                archive.read(name).decode("utf-8")
                for name in archive.namelist()
                if name.startswith("ppt/slideLayouts/slideLayout")
            )
        self.assertIn(fonts["display_zh"], layouts)
        self.assertIn(f'sz="{round(tokens["font_sizes_pt"]["cover_title"] * 100)}"', layouts)
        self.assertEqual(load_qmd_script("inspect_pptx").inspect(REFERENCE_PPTX, profile)["errors"], [])


@unittest.skipUnless(shutil.which("quarto"), "需要 Quarto 才能跑 QMD→PPTX 端到端渲染")
class EndToEndRenderTests(unittest.TestCase):
    def test_render_sample_deck_produces_editable_pptx(self):
        sample = ROOT / "shared-assets" / "examples" / "sample-deck.qmd"
        with tempfile.TemporaryDirectory() as tmp:
            workdir = Path(tmp)
            source = workdir / "deck.qmd"
            source.write_text(sample.read_text(encoding="utf-8"), encoding="utf-8")
            shutil.copytree(sample.parent / "figures", workdir / "figures")
            output = workdir / "deck.pptx"
            completed = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPTS / "render_qmd_ppt.py"),
                    str(source),
                    "--output",
                    str(output),
                    "--font-profile",
                    "preview",
                ],
                capture_output=True,
                text=True,
            )
            self.assertEqual(completed.returncode, 0, completed.stderr)
            with zipfile.ZipFile(output) as archive:
                slides = [n for n in archive.namelist() if n.startswith("ppt/slides/slide")]
                slide_xml = "".join(archive.read(n).decode("utf-8") for n in slides)
                theme = archive.read("ppt/theme/theme1.xml").decode("utf-8")
            inspection = load_qmd_script("inspect_pptx").inspect(output, "preview")
        self.assertGreaterEqual(len(slides), 2)
        self.assertIn("三年次新股策略研究", slide_xml)
        self.assertIn("takeaway-source", slide_xml)
        self.assertIn("Noto Sans CJK SC", theme)
        self.assertEqual(inspection["errors"], [])


if __name__ == "__main__":
    unittest.main()
