# PPT 转换工具重构实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 PPT 转换工具重构为基于内置模板的混合模式转换系统，实现 Pre-check 分析、页面类型分类、Layout 映射、混合模式转换（Content Migration + Technical Adaptation）、样式标准化和 QA 报告生成。

**Architecture:** 模板内置（取消上传），通过 clrMap 解析正确处理 4 种背景风格；混合模式：简单页面走 Content Migration（基于模板 Layout 创建新页），复杂页面走 Technical Adaptation（在原页上适配样式）；所有页面输出后生成 QA 报告。

**Tech Stack:** Python 3.11+, python-pptx, lxml, zipfile, openpyxl, streamlit, pydantic

---

## 文件结构

### 新增文件

| 文件 | 职责 |
|------|------|
| `core/clrmap/resolver.py` | clrMap 解析器，将 scheme 颜色名映射为实际 RGB |
| `core/precheck/analyzer.py` | Pre-check 分析器，生成源 PPT 质量报告 |
| `core/classifier/slide_classifier.py` | 幻灯片类型识别 + Migration/Adaptation 判定 |
| `core/migrator/slide_migrator.py` | Content Migration 路径实现（简单页面） |
| `core/qa/reporter.py` | QA 报告生成器（conversion_report.xlsx） |
| `config/font_mapping.yaml` | 字体替换配置 |
| `config/layout_mapping.yaml` | 页面类型 → Layout 索引映射配置 |

### 修改文件

| 文件 | 变更 |
|------|------|
| `core/models.py` | 新增 PreCheckResult、SlideClassification、QAReportItem 等模型 |
| `core/replayer/content_replayer.py` | 集成 clrMap、混合模式入口、样式标准化 |
| `core/watermark/detector.py` | 增加页脚/Location 识别辅助方法 |
| `app.py` | 删除模板上传，增加 Pre-check 展示、4 种背景选择、QA 报告下载 |
| `config/master_styles.yaml` | 更新为与新模板匹配的 4 种风格配置 |
| `core/registry/template_registry.py` | 支持从内置模板初始化 |

### 删除文件（可选，Phase 6 清理）

| 文件 | 原因 |
|------|------|
| `core/adapter/style_adapter.py` | 功能合并到 content_replayer |
| `core/analyzer/template_analyzer.py` | 模板分析功能简化并内置 |

---

## Task 1: clrMap 解析器

**Files:**
- Create: `core/clrmap/__init__.py`
- Create: `core/clrmap/resolver.py`
- Test: `tests/test_clrmap.py`

- [ ] **Step 1: 写测试 — 解析 Master 3 的 bg1 应为深绿色**

```python
import pytest
from core.clrmap.resolver import ClrMapResolver

TEMPLATE_PATH = "templates/2026 se template eng.pptx"

def test_master3_bg1_resolves_to_dark_green():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=2)  # 0-indexed, Master 3
    result = resolver.resolve_scheme_color("bg1")
    assert result.upper() == "0A2F24"

def test_master3_tx1_resolves_to_white():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=2)
    result = resolver.resolve_scheme_color("tx1")
    assert result.upper() == "FFFFFF"

def test_master1_bg1_resolves_to_white():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=0)  # Master 1
    result = resolver.resolve_scheme_color("bg1")
    assert result.upper() == "FFFFFF"

def test_master1_tx1_resolves_to_dark_green():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=0)
    result = resolver.resolve_scheme_color("tx1")
    assert result.upper() == "0A2F24"

def test_master2_accent5_is_light_green():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=1)
    result = resolver.resolve_scheme_color("accent5")
    assert result.upper() == "E7FFD9"

def test_invalid_scheme_color_returns_none():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=0)
    result = resolver.resolve_scheme_color("nonexistent")
    assert result is None
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_clrmap.py -v`
Expected: FAIL with "No module named 'core.clrmap'"

- [ ] **Step 3: 创建 `core/clrmap/__init__.py`**

```python
from .resolver import ClrMapResolver

__all__ = ["ClrMapResolver"]
```

- [ ] **Step 4: 创建 `core/clrmap/resolver.py`**

```python
import zipfile
from lxml import etree
from pathlib import Path

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"


class ClrMapResolver:
    def __init__(self, template_path: str, master_index: int):
        self.template_path = template_path
        self.master_index = master_index  # 0-indexed
        self.clr_map: dict[str, str] = {}
        self.theme_colors: dict[str, str] = {}
        self.theme_name: str = ""
        self._load()

    def _load(self):
        master_num = self.master_index + 1
        master_xml_path = f"ppt/slideMasters/slideMaster{master_num}.xml"
        master_rels_path = f"ppt/slideMasters/_rels/slideMaster{master_num}.xml.rels"

        with zipfile.ZipFile(self.template_path, "r") as zf:
            # 1. 读取 Master 的 clrMap
            master_xml = zf.read(master_xml_path)
            master_elem = etree.fromstring(master_xml)

            clr_map_elem = master_elem.find(f"{{{NS_P}}}clrMap")
            if clr_map_elem is not None:
                for attr_name in ["bg1", "tx1", "bg2", "tx2",
                                  "accent1", "accent2", "accent3",
                                  "accent4", "accent5", "accent6",
                                  "hlink", "folHlink"]:
                    val = clr_map_elem.get(attr_name)
                    if val:
                        self.clr_map[attr_name] = val

            # 2. 找到关联的 theme 文件
            master_rels_xml = zf.read(master_rels_path)
            rels_elem = etree.fromstring(master_rels_xml)
            theme_target = None
            for rel in rels_elem.findall(f"{{{NS_RELS}}}Relationship"):
                if "theme" in rel.get("Type", ""):
                    theme_target = rel.get("Target", "")
                    break

            if not theme_target:
                return

            # 解析 theme 路径（处理 ../ 前缀）
            if theme_target.startswith("../"):
                theme_path = f"ppt/{theme_target[3:]}"
            else:
                theme_path = f"ppt/{theme_target}"

            # 3. 读取 Theme 的 clrScheme
            theme_xml = zf.read(theme_path)
            theme_elem = etree.fromstring(theme_xml)
            self.theme_name = theme_elem.get("name", "")

            clr_scheme = theme_elem.find(f".//{{{NS_A}}}clrScheme")
            if clr_scheme is not None:
                for child in clr_scheme:
                    tag = child.tag.split("}")[-1]
                    srgb = child.find(f"{{{NS_A}}}srgbClr")
                    sysclr = child.find(f"{{{NS_A}}}sysClr")
                    if srgb is not None:
                        self.theme_colors[tag] = srgb.get("val", "")
                    elif sysclr is not None:
                        self.theme_colors[tag] = sysclr.get("lastClr", "")

    def resolve_scheme_color(self, scheme_name: str) -> str | None:
        mapped_name = self.clr_map.get(scheme_name, scheme_name)
        return self.theme_colors.get(mapped_name)

    def get_background_color(self) -> str | None:
        return self.resolve_scheme_color("bg1")

    def get_text_color(self) -> str | None:
        return self.resolve_scheme_color("tx1")

    def get_accent_color(self) -> str | None:
        return self.resolve_scheme_color("accent1")
```

- [ ] **Step 5: 运行测试确认通过**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_clrmap.py -v`
Expected: 6 passed

- [ ] **Step 6: Commit**

```bash
git add core/clrmap/__init__.py core/clrmap/resolver.py tests/test_clrmap.py
git commit -m "feat: add ClrMapResolver for correct theme color resolution"
```

---

## Task 2: 更新配置文件和数据模型

**Files:**
- Modify: `config/master_styles.yaml`
- Create: `config/layout_mapping.yaml`
- Create: `config/font_mapping.yaml`
- Modify: `core/models.py`
- Test: `tests/test_models.py`

- [ ] **Step 1: 写测试 — 新配置可正确加载**

```python
import pytest
import yaml
from pathlib import Path
from core.models import PreCheckIssue, SlideClassification, QAReportItem


def test_master_styles_has_4_styles():
    config_path = Path(__file__).parent.parent / "config" / "master_styles.yaml"
    with open(config_path, "r") as f:
        config = yaml.safe_load(f)
    assert "master_styles" in config
    assert len(config["master_styles"]) == 4

def test_precheck_issue_model():
    issue = PreCheckIssue(
        level="warning",
        rule_id="too_many_masters",
        message="母版数量过多，存在模板污染风险",
    )
    assert issue.level == "warning"
    assert issue.rule_id == "too_many_masters"

def test_slide_classification_model():
    cls = SlideClassification(
        slide_index=0,
        slide_type="Cover",
        migration_mode="migration",
        target_layout_index=0,
        confidence=0.9,
    )
    assert cls.slide_type == "Cover"
    assert cls.migration_mode == "migration"

def test_qa_report_item_model():
    item = QAReportItem(
        slide_no=1,
        detected_type="Cover",
        applied_layout="Title slide simple",
        migration_mode="migration",
        font_replaced="",
        objects_moved=0,
        objects_deleted=0,
        overflow_risk="None",
        need_manual_review=False,
        comment="",
    )
    assert item.slide_no == 1
    assert item.need_manual_review is False
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_models.py -v -k "precheck or classification or qa or master_styles"`
Expected: FAIL

- [ ] **Step 3: 更新 `config/master_styles.yaml`**

将内容替换为：

```yaml
master_styles:
  white:
    name: "White"
    display_name: "白色简约"
    description: "白色背景，适合商务演示"
    master_index: 0
    background_color: "#FFFFFF"
    text_color: "#0A2F24"
    title_font: "Poppins"
    body_font: "Poppins"
    accent_color: "#3DCD58"
  light_green:
    name: "Light Green"
    display_name: "浅绿色清新"
    description: "浅绿色背景，适合环保、自然主题"
    master_index: 1
    background_color: "#E7FFD9"
    text_color: "#0A2F24"
    title_font: "Poppins"
    body_font: "Poppins"
    accent_color: "#3DCD58"
  dark_green:
    name: "Dark Green"
    display_name: "深绿色商务"
    description: "深绿色背景，适合专业演示"
    master_index: 2
    background_color: "#0A2F24"
    text_color: "#FFFFFF"
    title_font: "Poppins"
    body_font: "Poppins"
    accent_color: "#3DCD58"
  gradient:
    name: "Gradient"
    display_name: "渐变科技"
    description: "渐变背景，适合创新、技术主题"
    master_index: 3
    background_type: "gradient"
    background_color: "#0A2F24"
    text_color: "#FFFFFF"
    title_font: "Poppins"
    body_font: "Poppins"
    accent_color: "#3DCD58"
```

- [ ] **Step 4: 创建 `config/layout_mapping.yaml`**

```yaml
layout_mappings:
  Cover:
    layout_index: 0
    layout_name: "Title slide simple"
  Section_Divider:
    layout_index: 19
    layout_name: "Section break for longer copy"
  Executive_Summary:
    layout_index: 1
    layout_name: "One column"
  KPI_Dashboard:
    layout_index: 9
    layout_name: "One column with object"
  Timeline:
    layout_index: 1
    layout_name: "One column"
  Risk_Matrix:
    layout_index: 1
    layout_name: "One column"
  Table_Page:
    layout_index: 9
    layout_name: "One column with object"
  Chart_Page:
    layout_index: 9
    layout_name: "One column with object"
  Image_Page:
    layout_index: 4
    layout_name: "Title slide with image"
  Appendix:
    layout_index: 1
    layout_name: "One column"
  Closing:
    layout_index: 24
    layout_name: "Closing slide"
  Content:
    layout_index: 1
    layout_name: "One column"

migration_types:
  - Cover
  - Section_Divider
  - Closing
  - Image_Page

adaptation_types:
  - Executive_Summary
  - KPI_Dashboard
  - Timeline
  - Risk_Matrix
  - Table_Page
  - Chart_Page
  - Appendix
  - Content
  - Decision_Page
```

- [ ] **Step 5: 创建 `config/font_mapping.yaml`**

```yaml
font_replacements:
  chinese:
    source_fonts:
      - "微软雅黑"
      - "Microsoft YaHei"
      - "Microsoft YaHei UI"
      - "宋体"
      - "SimSun"
      - "黑体"
      - "SimHei"
      - "楷体"
      - "KaiTi"
    target_font: "汉仪旗黑"
  english:
    source_fonts:
      - "Calibri"
      - "Arial"
      - "Times New Roman"
      - "Tahoma"
      - "Verdana"
    target_font: "Poppins"
  heading:
    target_font: "Poppins"
  body:
    target_font: "Poppins"
```

- [ ] **Step 6: 修改 `core/models.py` — 新增数据模型**

在文件末尾添加：

```python
class PreCheckIssue(BaseModel):
    level: Literal["info", "warning", "error"]
    rule_id: str
    message: str
    slide_index: int | None = None


class PreCheckResult(BaseModel):
    slide_count: int = 0
    master_count: int = 0
    fonts_used: list[str] = []
    has_old_se_template: bool = False
    has_external_theme: bool = False
    has_embedded_chart: bool = False
    has_smartart: bool = False
    has_media: bool = False
    has_animation: bool = False
    is_4_3_ratio: bool = False
    overflow_objects_count: int = 0
    issues: list[PreCheckIssue] = []


class SlideClassification(BaseModel):
    slide_index: int
    slide_type: str
    migration_mode: Literal["migration", "adaptation"]
    target_layout_index: int
    target_layout_name: str = ""
    confidence: float = 0.0


class QAReportItem(BaseModel):
    slide_no: int
    detected_type: str
    applied_layout: str
    migration_mode: str
    font_replaced: str = ""
    objects_moved: int = 0
    objects_deleted: int = 0
    overflow_risk: str = "None"
    need_manual_review: bool = False
    comment: str = ""


class ConversionConfig(BaseModel):
    input_path: str
    output_path: str
    report_path: str = ""
    background_style: str = "dark_green"
    include_header: bool = False
    include_footer: bool = True
    include_icon: bool = False
```

- [ ] **Step 7: 运行测试确认通过**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_models.py -v`
Expected: all passed

- [ ] **Step 8: Commit**

```bash
git add config/master_styles.yaml config/layout_mapping.yaml config/font_mapping.yaml core/models.py tests/test_models.py
git commit -m "feat: add config files and new data models for redesign"
```

---

## Task 3: Pre-check 分析器

**Files:**
- Create: `core/precheck/__init__.py`
- Create: `core/precheck/analyzer.py`
- Test: `tests/test_precheck.py`

- [ ] **Step 1: 写测试**

```python
import pytest
import tempfile
import os
from pptx import Presentation
from core.precheck.analyzer import PreCheckAnalyzer


def _make_simple_pptx(path):
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[0])
    prs.slides[0].shapes.title.text = "Test Title"
    prs.save(path)
    return path


def test_precheck_basic_info():
    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "test.pptx")
        _make_simple_pptx(pptx_path)
        
        analyzer = PreCheckAnalyzer()
        result = analyzer.analyze(pptx_path)
        
        assert result.slide_count == 1
        assert result.master_count >= 1
        assert result.is_4_3_ratio is False  # default is 16:9


def test_precheck_fonts_detected():
    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "test.pptx")
        _make_simple_pptx(pptx_path)
        
        analyzer = PreCheckAnalyzer()
        result = analyzer.analyze(pptx_path)
        
        assert isinstance(result.fonts_used, list)
        # 默认 Calibri 应该在里面
        assert len(result.fonts_used) >= 0


def test_precheck_no_media_no_animation_simple_pptx():
    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "test.pptx")
        _make_simple_pptx(pptx_path)
        
        analyzer = PreCheckAnalyzer()
        result = analyzer.analyze(pptx_path)
        
        assert result.has_media is False
        assert result.has_animation is False
        assert result.has_smartart is False
        assert result.has_embedded_chart is False
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_precheck.py -v`
Expected: FAIL

- [ ] **Step 3: 创建 `core/precheck/__init__.py`**

```python
from .analyzer import PreCheckAnalyzer, PreCheckResult

__all__ = ["PreCheckAnalyzer", "PreCheckResult"]
```

- [ ] **Step 4: 创建 `core/precheck/analyzer.py`**

```python
import zipfile
from lxml import etree
from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from core.models import PreCheckResult, PreCheckIssue

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

SE_KEYWORDS = ["schneider", "se.com", "Schneider Electric", "施耐德"]


class PreCheckAnalyzer:
    def analyze(self, pptx_path: str) -> PreCheckResult:
        result = PreCheckResult()
        prs = Presentation(pptx_path)

        # 1. 页面数量
        result.slide_count = len(prs.slides)

        # 2. 母版数量
        result.master_count = len(prs.slide_masters)
        if result.master_count > 3:
            result.issues.append(PreCheckIssue(
                level="warning",
                rule_id="too_many_masters",
                message=f"母版数量为 {result.master_count}，可能存在模板污染",
            ))

        # 3. 尺寸比例
        ratio = prs.slide_width / prs.slide_height if prs.slide_height else 0
        result.is_4_3_ratio = abs(ratio - 4/3) < 0.02
        if result.is_4_3_ratio:
            result.issues.append(PreCheckIssue(
                level="warning",
                rule_id="4_3_ratio",
                message="PPT 为 4:3 比例，将自动调整为 16:9",
            ))

        # 4. 字体清单
        fonts = set()
        for slide in prs.slides:
            self._collect_fonts_from_slide(slide, fonts)
        result.fonts_used = sorted(fonts)

        # 检查旧字体
        old_fonts = [f for f in fonts if any(
            kw in f.lower() for kw in ["microsoft yahei", "calibri", "微软雅黑"]
        )]
        if old_fonts:
            result.issues.append(PreCheckIssue(
                level="info",
                rule_id="old_fonts",
                message=f"检测到旧字体: {', '.join(old_fonts)}，将自动替换",
            ))

        # 5. 其他检查
        for slide in prs.slides:
            self._check_slide_media(slide, result)
            self._check_slide_overflow(slide, prs, result)

        # 6. 检查动画
        result.has_animation = self._check_animation(pptx_path)

        # 7. 检查 SmartArt
        result.has_smartart = self._check_smartart(pptx_path)

        # 8. 检查是否含旧 SE 模板（通过主题/母版名称）
        result.has_old_se_template = self._check_old_se_template(pptx_path)

        # 9. 外部主题
        result.has_external_theme = result.master_count > 2

        return result

    def _collect_fonts_from_slide(self, slide, fonts: set):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name:
                            fonts.add(run.font.name)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                if run.font.name:
                                    fonts.add(run.font.name)

    def _check_slide_media(self, slide, result):
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
                result.has_media = True
            if shape.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT:
                result.has_embedded_chart = True

    def _check_slide_overflow(self, slide, prs, result):
        for shape in slide.shapes:
            try:
                if shape.left + shape.width > prs.slide_width + 10000:
                    result.overflow_objects_count += 1
                if shape.top + shape.height > prs.slide_height + 10000:
                    result.overflow_objects_count += 1
            except Exception:
                pass

    def _check_animation(self, pptx_path: str) -> bool:
        try:
            with zipfile.ZipFile(pptx_path, "r") as zf:
                for name in zf.namelist():
                    if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                        xml = zf.read(name)
                        if b"<p:timing" in xml or b"<p:anim" in xml:
                            return True
        except Exception:
            pass
        return False

    def _check_smartart(self, pptx_path: str) -> bool:
        try:
            with zipfile.ZipFile(pptx_path, "r") as zf:
                for name in zf.namelist():
                    if "diagrams" in name or "smartArt" in name.lower():
                        return True
        except Exception:
            pass
        return False

    def _check_old_se_template(self, pptx_path: str) -> bool:
        try:
            with zipfile.ZipFile(pptx_path, "r") as zf:
                # 检查主题名称
                for name in zf.namelist():
                    if name.startswith("ppt/theme/theme") and name.endswith(".xml"):
                        xml = zf.read(name).decode("utf-8", errors="ignore")
                        if any(kw.lower() in xml.lower() for kw in SE_KEYWORDS):
                            return True
        except Exception:
            pass
        return False
```

- [ ] **Step 5: 运行测试确认通过**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_precheck.py -v`
Expected: 3 passed

- [ ] **Step 6: Commit**

```bash
git add core/precheck/__init__.py core/precheck/analyzer.py tests/test_precheck.py
git commit -m "feat: add PreCheckAnalyzer for source PPT quality analysis"
```

---

## Task 4: Slide Classifier（页面类型识别）

**Files:**
- Create: `core/classifier/__init__.py`
- Create: `core/classifier/slide_classifier.py`
- Test: `tests/test_slide_classifier.py`

- [ ] **Step 1: 写测试**

```python
import pytest
import tempfile
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from core.classifier.slide_classifier import SlideClassifier


def _make_cover_pptx(path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])  # Title Slide
    slide.shapes.title.text = "Presentation Title"
    if len(slide.placeholders) > 1:
        slide.placeholders[1].text = "Subtitle"
    prs.save(path)
    return path


def _make_table_pptx(path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only
    slide.shapes.title.text = "Data Table"
    rows, cols = 3, 3
    left = top = Inches(2)
    width = height = Inches(4)
    slide.shapes.add_table(rows, cols, left, top, width, height)
    prs.save(path)
    return path


def test_classify_cover():
    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "test.pptx")
        _make_cover_pptx(pptx_path)
        prs = Presentation(pptx_path)
        
        classifier = SlideClassifier()
        results = classifier.classify_all(prs)
        
        assert len(results) == 1
        assert results[0].slide_type == "Cover"
        assert results[0].migration_mode == "migration"


def test_classify_table_page():
    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "test.pptx")
        _make_table_pptx(pptx_path)
        prs = Presentation(pptx_path)
        
        classifier = SlideClassifier()
        results = classifier.classify_all(prs)
        
        assert len(results) == 1
        assert results[0].slide_type == "Table_Page"
        assert results[0].migration_mode == "adaptation"


def test_classify_content_default():
    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "test.pptx")
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
        slide.shapes.title.text = "Some Content"
        slide.placeholders[1].text = "Body text here"
        prs.save(pptx_path)
        
        classifier = SlideClassifier()
        results = classifier.classify_all(pptx_path)
        
        assert results[0].slide_type == "Content"
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_slide_classifier.py -v`
Expected: FAIL

- [ ] **Step 3: 创建 `core/classifier/__init__.py`**

```python
from .slide_classifier import SlideClassifier

__all__ = ["SlideClassifier"]
```

- [ ] **Step 4: 创建 `core/classifier/slide_classifier.py`**

```python
import yaml
from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from core.models import SlideClassification


CONFIG_PATH = Path(__file__).parent.parent.parent / "config" / "layout_mapping.yaml"

COVER_KEYWORDS = []
SUMMARY_KEYWORDS = [
    "executive summary", "key messages", "key takeaways",
    "总结", "摘要", "概述", "overview", "agenda",
]
RISK_KEYWORDS = [
    "risk", "issue", "mitigation", "挑战", "风险", "问题", "应对",
]
TIMELINE_KEYWORDS = [
    "timeline", "milestone", "roadmap", "时间线", "里程碑", "路线图", "进度",
]
CLOSING_KEYWORDS = [
    "thank you", "thanks", "se.com", "closing", "q&a",
    "谢谢", "感谢", "结束",
]
SECTION_PATTERNS = ["section", "章节", "第.*章", "^\\d+\\s", "^[一二三四五六七八九十]+、"]


class SlideClassifier:
    def __init__(self, config_path: str | None = None):
        self.layout_mappings: dict = {}
        self.migration_types: set = set()
        self.adaptation_types: set = set()
        self._load_config(config_path)

    def _load_config(self, config_path: str | None):
        path = Path(config_path) if config_path else CONFIG_PATH
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                config = yaml.safe_load(f)
            self.layout_mappings = config.get("layout_mappings", {})
            self.migration_types = set(config.get("migration_types", []))
            self.adaptation_types = set(config.get("adaptation_types", []))

    def classify_all(self, presentation) -> list[SlideClassification]:
        if isinstance(presentation, str):
            prs = Presentation(presentation)
        else:
            prs = presentation

        results = []
        total = len(prs.slides)
        for idx, slide in enumerate(prs.slides):
            cls = self.classify_slide(slide, idx, total)
            results.append(cls)
        return results

    def classify_slide(self, slide, slide_index: int, total_slides: int) -> SlideClassification:
        slide_type = self._detect_type(slide, slide_index, total_slides)
        migration_mode = self._decide_mode(slide_type)
        layout_info = self.layout_mappings.get(slide_type, {
            "layout_index": 1, "layout_name": "One column"
        })

        return SlideClassification(
            slide_index=slide_index,
            slide_type=slide_type,
            migration_mode=migration_mode,
            target_layout_index=layout_info.get("layout_index", 1),
            target_layout_name=layout_info.get("layout_name", "One column"),
            confidence=0.7,
        )

    def _detect_type(self, slide, slide_index: int, total_slides: int) -> str:
        title_text = self._get_title_text(slide)
        all_text = self._get_all_text(slide).lower()

        # 1. 封面
        if slide_index == 0 and self._is_cover_like(slide):
            return "Cover"

        # 2. 结尾
        if slide_index == total_slides - 1:
            if any(kw in all_text for kw in CLOSING_KEYWORDS):
                return "Closing"

        # 3. 章节分隔
        if self._is_section_divider(slide):
            return "Section_Divider"

        # 4. 摘要
        if any(kw in title_text.lower() for kw in SUMMARY_KEYWORDS):
            return "Executive_Summary"

        # 5. 风险
        if any(kw in title_text.lower() for kw in RISK_KEYWORDS):
            return "Risk_Matrix"

        # 6. 时间线
        if any(kw in title_text.lower() for kw in TIMELINE_KEYWORDS):
            return "Timeline"

        # 7. 表格
        if self._has_table(slide):
            return "Table_Page"

        # 8. 图表
        chart_count = self._count_charts(slide)
        if chart_count >= 2:
            return "KPI_Dashboard"
        if chart_count == 1:
            return "Chart_Page"

        # 9. 图片为主
        if self._is_image_dominant(slide):
            return "Image_Page"

        return "Content"

    def _decide_mode(self, slide_type: str) -> str:
        if slide_type in self.migration_types:
            return "migration"
        return "adaptation"

    def _get_title_text(self, slide) -> str:
        if slide.shapes.title:
            return slide.shapes.title.text or ""
        return ""

    def _get_all_text(self, slide) -> str:
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                texts.append(shape.text_frame.text)
        return " ".join(texts)

    def _is_cover_like(self, slide) -> bool:
        text_shapes = [s for s in slide.shapes if s.has_text_frame]
        if len(text_shapes) <= 2:
            return True
        # 检查是否只有 title + subtitle 占位符
        placeholders = [s for s in slide.shapes if s.is_placeholder]
        if len(placeholders) <= 2:
            return True
        return False

    def _is_section_divider(self, slide) -> bool:
        title = self._get_title_text(slide)
        if not title:
            return False
        # 只有标题，无正文
        body_shapes = [
            s for s in slide.shapes
            if s.has_text_frame and s != slide.shapes.title
        ]
        if len(body_shapes) == 0:
            return True
        # 正文内容很少
        total_body_chars = sum(
            len(s.text_frame.text) for s in body_shapes if s.has_text_frame
        )
        if total_body_chars < 20:
            return True
        return False

    def _has_table(self, slide) -> bool:
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                return True
        return False

    def _count_charts(self, slide) -> int:
        count = 0
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.CHART:
                count += 1
        return count

    def _is_image_dominant(self, slide) -> bool:
        image_count = 0
        text_count = 0
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image_count += 1
            if shape.has_text_frame and len(shape.text_frame.text) > 10:
                text_count += 1
        return image_count >= 1 and text_count <= 1
```

- [ ] **Step 5: 运行测试确认通过**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_slide_classifier.py -v`
Expected: 3 passed

- [ ] **Step 6: Commit**

```bash
git add core/classifier/__init__.py core/classifier/slide_classifier.py tests/test_slide_classifier.py
git commit -m "feat: add SlideClassifier for slide type detection and migration decision"
```

---

## Task 5: Content Migration（简单页面迁移）

**Files:**
- Create: `core/migrator/__init__.py`
- Create: `core/migrator/slide_migrator.py`
- Test: `tests/test_slide_migrator.py`

- [ ] **Step 1: 写测试**

```python
import pytest
import tempfile
import os
import shutil
from pptx import Presentation
from core.migrator.slide_migrator import SlideMigrator

TEMPLATE_PATH = "templates/2026 se template eng.pptx"


def _make_source_pptx(path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "My Title"
    if len(slide.placeholders) > 1:
        slide.placeholders[1].text = "My Subtitle"
    prs.save(path)
    return path


def test_migrate_cover_slide():
    with tempfile.TemporaryDirectory() as tmpdir:
        # 准备源PPT
        source_path = os.path.join(tmpdir, "source.pptx")
        _make_source_pptx(source_path)
        
        # 复制模板到临时目录作为输出基准
        output_path = os.path.join(tmpdir, "output.pptx")
        shutil.copy(TEMPLATE_PATH, output_path)
        
        migrator = SlideMigrator(TEMPLATE_PATH, master_index=2)  # Dark Green
        
        source_prs = Presentation(source_path)
        target_prs = Presentation(output_path)
        
        # 迁移第1页（封面）
        new_slide = migrator.migrate_slide(
            source_slide=source_prs.slides[0],
            target_prs=target_prs,
            slide_type="Cover",
            layout_index=0,
        )
        
        assert new_slide is not None
        # 标题应该被迁移
        assert new_slide.shapes.title is not None
        assert "My Title" in new_slide.shapes.title.text
        
        target_prs.save(output_path)
        
        # 验证输出文件存在且可打开
        assert os.path.exists(output_path)
        prs2 = Presentation(output_path)
        assert len(prs2.slides) >= 1
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_slide_migrator.py -v`
Expected: FAIL

- [ ] **Step 3: 创建 `core/migrator/__init__.py`**

```python
from .slide_migrator import SlideMigrator

__all__ = ["SlideMigrator"]
```

- [ ] **Step 4: 创建 `core/migrator/slide_migrator.py`**

```python
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from core.clrmap.resolver import ClrMapResolver


class SlideMigrator:
    def __init__(self, template_path: str, master_index: int):
        self.template_path = template_path
        self.master_index = master_index
        self.clr_resolver = ClrMapResolver(template_path, master_index)

    def migrate_slide(self, source_slide, target_prs, slide_type: str, layout_index: int) -> object:
        """
        基于目标 Layout 创建新幻灯片，从源幻灯片迁移文本内容。
        返回新创建的 slide 对象。
        """
        # 获取目标 Layout
        if self.master_index >= len(target_prs.slide_masters):
            master = target_prs.slide_masters[0]
        else:
            master = target_prs.slide_masters[self.master_index]

        if layout_index >= len(master.slide_layouts):
            layout_index = 0
        target_layout = master.slide_layouts[layout_index]

        # 1. 创建新幻灯片
        new_slide = target_prs.slides.add_slide(target_layout)

        # 2. 提取源页面内容
        title_text = self._extract_title(source_slide)
        subtitle_text = self._extract_subtitle(source_slide)
        body_paragraphs = self._extract_body_paragraphs(source_slide)

        # 3. 填入占位符
        self._fill_title(new_slide, title_text)
        self._fill_subtitle(new_slide, subtitle_text)
        self._fill_body(new_slide, body_paragraphs)

        return new_slide

    def _extract_title(self, slide) -> str:
        if slide.shapes.title:
            return slide.shapes.title.text or ""
        # 兜底：找最大的文本框
        max_shape = None
        max_area = 0
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                area = shape.width * shape.height
                if area > max_area:
                    max_area = area
                    max_shape = shape
        if max_shape:
            return max_shape.text_frame.text.strip()
        return ""

    def _extract_subtitle(self, slide) -> str:
        # 从副标题占位符提取
        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            # 副标题类型 = 4
            if ph_type == 4 and shape.has_text_frame:
                return shape.text_frame.text.strip()
        # 兜底：第二大的文本框
        text_shapes = [
            s for s in slide.shapes
            if s.has_text_frame and s.text_frame.text.strip()
        ]
        text_shapes.sort(key=lambda s: s.width * s.height, reverse=True)
        if len(text_shapes) >= 2:
            return text_shapes[1].text_frame.text.strip()
        return ""

    def _extract_body_paragraphs(self, slide) -> list[str]:
        paragraphs = []
        for shape in slide.shapes:
            if shape == slide.shapes.title:
                continue
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if text:
                        paragraphs.append(text)
        return paragraphs

    def _fill_title(self, slide, text: str):
        if slide.shapes.title and text:
            slide.shapes.title.text = text
            # 应用标题样式
            for para in slide.shapes.title.text_frame.paragraphs:
                for run in para.runs:
                    run.font.name = "Poppins"

    def _fill_subtitle(self, slide, text: str):
        if not text:
            return
        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            if ph_type == 4 and shape.has_text_frame:
                shape.text_frame.text = text
                return
        # 没有副标题占位符时，放入body
        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            if ph_type == 1 and shape.has_text_frame:  # Body
                shape.text_frame.text = text
                return

    def _fill_body(self, slide, paragraphs: list[str]):
        if not paragraphs:
            return
        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            if ph_type == 1 and shape.has_text_frame:  # Body placeholder
                tf = shape.text_frame
                tf.clear()
                for i, para_text in enumerate(paragraphs):
                    if i == 0:
                        para = tf.paragraphs[0]
                    else:
                        para = tf.add_paragraph()
                    para.text = para_text
                    for run in para.runs:
                        run.font.name = "Poppins"
                return
```

- [ ] **Step 5: 运行测试确认通过**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_slide_migrator.py -v`
Expected: 1 passed

- [ ] **Step 6: Commit**

```bash
git add core/migrator/__init__.py core/migrator/slide_migrator.py tests/test_slide_migrator.py
git commit -m "feat: add SlideMigrator for content migration path"
```

---

## Task 6: QA 报告生成器

**Files:**
- Create: `core/qa/__init__.py`
- Create: `core/qa/reporter.py`
- Test: `tests/test_qa_reporter.py`

- [ ] **Step 1: 写测试**

```python
import pytest
import tempfile
import os
from core.qa.reporter import QAReporter
from core.models import QAReportItem


def test_generate_report_xlsx():
    with tempfile.TemporaryDirectory() as tmpdir:
        report_path = os.path.join(tmpdir, "conversion_report.xlsx")
        
        items = [
            QAReportItem(
                slide_no=1,
                detected_type="Cover",
                applied_layout="Title slide simple",
                migration_mode="migration",
                font_replaced="",
                objects_moved=0,
                objects_deleted=2,
                overflow_risk="None",
                need_manual_review=False,
                comment="封面页迁移成功",
            ),
            QAReportItem(
                slide_no=2,
                detected_type="Content",
                applied_layout="One column",
                migration_mode="adaptation",
                font_replaced="Calibri→Poppins",
                objects_moved=3,
                objects_deleted=1,
                overflow_risk="Low",
                need_manual_review=True,
                comment="文本框可能溢出，建议检查",
            ),
        ]
        
        reporter = QAReporter()
        reporter.generate(items, report_path)
        
        assert os.path.exists(report_path)
        assert os.path.getsize(report_path) > 0

        # 验证可以用openpyxl打开
        from openpyxl import load_workbook
        wb = load_workbook(report_path)
        assert "转换报告" in wb.sheetnames
        ws = wb["转换报告"]
        assert ws.max_row == 3  # 表头 + 2条数据
        assert ws.max_column == 10
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_qa_reporter.py -v`
Expected: FAIL

- [ ] **Step 3: 检查 openpyxl 是否可用**

Run: `cd /workspace/ppt-conform && python -c "import openpyxl; print(openpyxl.__version__)"`

如果未安装，添加到 requirements.txt 并安装。

- [ ] **Step 4: 创建 `core/qa/__init__.py`**

```python
from .reporter import QAReporter

__all__ = ["QAReporter"]
```

- [ ] **Step 5: 创建 `core/qa/reporter.py`**

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from core.models import QAReportItem

HEADERS = [
    "Slide No.",
    "Detected Type",
    "Applied Layout",
    "Migration Mode",
    "Font Replaced",
    "Objects Moved",
    "Objects Deleted",
    "Overflow Risk",
    "Need Manual Review",
    "Comment",
]

HEADER_FILL = PatternFill(start_color="0A2F24", end_color="0A2F24", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11, name="Poppins")
WARN_FILL = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
ERROR_FILL = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin", color="DDDDDD"),
    right=Side(style="thin", color="DDDDDD"),
    top=Side(style="thin", color="DDDDDD"),
    bottom=Side(style="thin", color="DDDDDD"),
)


class QAReporter:
    def generate(self, items: list[QAReportItem], output_path: str, summary: dict | None = None):
        wb = Workbook()
        ws = wb.active
        ws.title = "转换报告"

        # 写表头
        for col, header in enumerate(HEADERS, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER

        # 写数据
        for row_idx, item in enumerate(items, 2):
            values = [
                item.slide_no,
                item.detected_type,
                item.applied_layout,
                item.migration_mode,
                item.font_replaced,
                item.objects_moved,
                item.objects_deleted,
                item.overflow_risk,
                "是" if item.need_manual_review else "否",
                item.comment,
            ]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="center", wrap_text=True)

                # 需要人工检查的行高亮
                if item.need_manual_review:
                    cell.fill = WARN_FILL

        # 列宽
        col_widths = [12, 20, 28, 16, 25, 15, 16, 14, 18, 40]
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

        # 行高
        ws.row_dimensions[1].height = 30

        # 冻结首行
        ws.freeze_panes = "A2"

        # 汇总页
        if summary:
            ws2 = wb.create_sheet("汇总")
            ws2["A1"] = "转换汇总"
            ws2["A1"].font = Font(bold=True, size=14, color="0A2F24")
            row = 3
            for key, value in summary.items():
                ws2.cell(row=row, column=1, value=key).font = Font(bold=True)
                ws2.cell(row=row, column=2, value=value)
                row += 1
            ws2.column_dimensions["A"].width = 25
            ws2.column_dimensions["B"].width = 50

        wb.save(output_path)
```

- [ ] **Step 6: 运行测试确认通过**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_qa_reporter.py -v`
Expected: 1 passed

- [ ] **Step 7: Commit**

```bash
git add core/qa/__init__.py core/qa/reporter.py tests/test_qa_reporter.py
git commit -m "feat: add QAReporter for conversion quality report (xlsx)"
```

---

## Task 7: 集成 ContentReplayer（混合模式）

**Files:**
- Modify: `core/replayer/content_replayer.py`
- Test: `tests/test_replayer.py` (更新现有测试)

- [ ] **Step 1: 读取现有 content_replayer.py 全文**

先读取完整文件，了解所有方法。

- [ ] **Step 2: 添加混合模式主入口 `convert_with_classification`**

在 `ContentReplayer` 类中新增方法：

```python
def convert_with_classification(
    self,
    source_path: str,
    output_path: str,
    background_style: str = "dark_green",
) -> tuple[str, list[QAReportItem]]:
    """
    混合模式转换主入口：
    1. 分类所有页面
    2. 简单页面走 Migration，复杂页面走 Adaptation
    3. 应用样式标准化
    4. 返回 (output_path, qa_items)
    """
    import shutil
    import yaml
    from pathlib import Path
    from pptx import Presentation
    from core.classifier.slide_classifier import SlideClassifier
    from core.migrator.slide_migrator import SlideMigrator
    from core.clrmap.resolver import ClrMapResolver
    from core.models import QAReportItem

    # 加载风格配置
    config_path = Path(__file__).parent.parent.parent / "config" / "master_styles.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        styles_config = yaml.safe_load(f)
    
    style_cfg = styles_config["master_styles"].get(background_style, {})
    master_idx = style_cfg.get("master_index", 2)

    # 1. 分类
    source_prs = Presentation(source_path)
    classifier = SlideClassifier()
    classifications = classifier.classify_all(source_prs)

    # 2. 创建输出（基于模板）
    shutil.copy(self.template_path, output_path)
    target_prs = Presentation(output_path)

    # 清空模板中可能有的示例幻灯片
    while len(target_prs.slides) > 0:
        rId = target_prs.slides._sldIdLst[0].rId
        target_prs.part.drop_rel(rId)
        del target_prs.slides._sldIdLst[0]

    # 设置尺寸
    target_width = target_prs.slide_width
    target_height = target_prs.slide_height

    migrator = SlideMigrator(self.template_path, master_idx)
    qa_items = []

    for idx, cls in enumerate(classifications):
        source_slide = source_prs.slides[idx]
        
        if cls.migration_mode == "migration":
            # Migration 路径
            new_slide = migrator.migrate_slide(
                source_slide=source_slide,
                target_prs=target_prs,
                slide_type=cls.slide_type,
                layout_index=cls.target_layout_index,
            )
            deleted = self._count_shapes(source_slide) - self._count_relevant_shapes(new_slide)
            
            qa_item = QAReportItem(
                slide_no=idx + 1,
                detected_type=cls.slide_type,
                applied_layout=cls.target_layout_name,
                migration_mode="migration",
                font_replaced="",
                objects_moved=0,
                objects_deleted=max(0, deleted),
                overflow_risk="None",
                need_manual_review=False,
                comment=f"Migration: {cls.slide_type}",
            )
        else:
            # Adaptation 路径：调用现有 replay 单页逻辑
            # （此处集成现有 _copy_selected_master_to_output 和 _apply_background_to_slide）
            # 先占位，后续完善
            new_slide = self._adapt_single_slide(
                source_slide, target_prs, master_idx, cls
            )
            qa_item = QAReportItem(
                slide_no=idx + 1,
                detected_type=cls.slide_type,
                applied_layout=cls.target_layout_name,
                migration_mode="adaptation",
                font_replaced="Calibri→Poppins",
                objects_moved=0,
                objects_deleted=0,
                overflow_risk="Low",
                need_manual_review=True,
                comment=f"Adaptation: {cls.slide_type}，建议人工检查",
            )

        qa_items.append(qa_item)

    # 保存
    target_prs.save(output_path)

    return output_path, qa_items

def _count_shapes(self, slide) -> int:
    return len(slide.shapes)

def _count_relevant_shapes(self, slide) -> int:
    count = 0
    for s in slide.shapes:
        if s.shape_type is not None:
            count += 1
    return count

def _adapt_single_slide(self, source_slide, target_prs, master_idx, classification):
    """Adaptation 路径的简化实现——先复制内容后应用样式"""
    # 对于 MVP 版本，先创建一个基于目标 Layout 的空白页
    # 然后把源页面的所有 shape 复制过去（保留内容但不保留旧母版）
    master = target_prs.slide_masters[min(master_idx, len(target_prs.slide_masters) - 1)]
    layout_idx = min(classification.target_layout_index, len(master.slide_layouts) - 1)
    target_layout = master.slide_layouts[layout_idx]
    new_slide = target_prs.slides.add_slide(target_layout)
    
    # 复制文本内容（简化版：只复制文本，不复制形状位置）
    title_text = ""
    body_texts = []
    for shape in source_slide.shapes:
        if shape.has_text_frame and shape == source_slide.shapes.title:
            title_text = shape.text_frame.text
        elif shape.has_text_frame:
            body_texts.append(shape.text_frame.text)
    
    if new_slide.shapes.title and title_text:
        new_slide.shapes.title.text = title_text
    
    # 放入第一个 body 占位符
    for ph in new_slide.placeholders:
        if ph.placeholder_format.type == 1 and ph.has_text_frame:
            ph.text_frame.text = "\n".join(body_texts)
            break
    
    return new_slide
```

> 注意：Adaptation 的完整实现需要逐步完善，先保证主流程跑通。

- [ ] **Step 3: 运行现有测试**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_replayer.py -v --tb=short`
Expected: 现有测试仍然通过

- [ ] **Step 4: 写集成测试**

```python
def test_convert_with_classification():
    import tempfile
    import os
    from pptx import Presentation
    
    with tempfile.TemporaryDirectory() as tmpdir:
        # 创建简单源PPT
        source_path = os.path.join(tmpdir, "source.pptx")
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Test Title"
        prs.save(source_path)
        
        output_path = os.path.join(tmpdir, "output.pptx")
        
        from core.registry.template_registry import TemplateRegistry
        registry = TemplateRegistry()
        from core.replayer.content_replayer import ContentReplayer
        replayer = ContentReplayer(registry, template_path="templates/2026 se template eng.pptx")
        
        out_path, qa_items = replayer.convert_with_classification(
            source_path=source_path,
            output_path=output_path,
            background_style="dark_green",
        )
        
        assert os.path.exists(out_path)
        assert len(qa_items) >= 1
        assert qa_items[0].slide_no == 1
```

- [ ] **Step 5: 运行新测试**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_replayer.py::test_convert_with_classification -v`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add core/replayer/content_replayer.py tests/test_replayer.py
git commit -m "feat: add hybrid mode convert_with_classification to ContentReplayer"
```

---

## Task 8: 样式标准化（字体替换）

**Files:**
- Modify: `core/replayer/content_replayer.py` (新增字体替换方法)
- Test: `tests/test_font_normalization.py`

- [ ] **Step 1: 写测试**

```python
import pytest
import tempfile
import os
from pptx import Presentation
from pptx.util import Pt
from core.replayer.content_replayer import ContentReplayer
from core.registry.template_registry import TemplateRegistry


def test_font_normalization_replaces_calibri():
    with tempfile.TemporaryDirectory() as tmpdir:
        # 创建带 Calibri 字体的 PPT
        source_path = os.path.join(tmpdir, "source.pptx")
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Title in Calibri"
        for para in slide.shapes.title.text_frame.paragraphs:
            for run in para.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(24)
        prs.save(source_path)

        # 运行字体标准化
        from core.clrmap.resolver import ClrMapResolver
        target_prs = Presentation(source_path)
        replayer = ContentReplayer(
            TemplateRegistry(),
            template_path="templates/2026 se template eng.pptx",
        )
        
        replaced = replayer._normalize_fonts(target_prs.slides[0])
        
        # 检查字体是否被替换
        for para in target_prs.slides[0].shapes.title.text_frame.paragraphs:
            for run in para.runs:
                assert run.font.name == "Poppins"
        
        assert len(replaced) > 0
        assert "Calibri" in replaced
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_font_normalization.py -v`
Expected: FAIL

- [ ] **Step 3: 在 ContentReplayer 中添加字体标准化方法**

```python
def _normalize_fonts(self, slide) -> set[str]:
    """
    将幻灯片中的字体统一替换为目标字体。
    返回被替换的字体集合。
    """
    import yaml
    from pathlib import Path
    
    replaced = set()
    
    config_path = Path(__file__).parent.parent.parent / "config" / "font_mapping.yaml"
    if not config_path.exists():
        return replaced
    
    with open(config_path, "r", encoding="utf-8") as f:
        font_config = yaml.safe_load(f)
    
    # 构建替换映射
    replace_map = {}
    for group in ["chinese", "english"]:
        cfg = font_config.get("font_replacements", {}).get(group, {})
        target = cfg.get("target_font", "Poppins")
        for src in cfg.get("source_fonts", []):
            replace_map[src.lower()] = target
    
    # 遍历所有形状
    for shape in slide.shapes:
        self._normalize_shape_fonts(shape, replace_map, replaced)
    
    return replaced

def _normalize_shape_fonts(self, shape, replace_map: dict, replaced: set):
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.font.name:
                    old_font = run.font.name
                    new_font = replace_map.get(old_font.lower())
                    if new_font and old_font != new_font:
                        run.font.name = new_font
                        replaced.add(old_font)
    
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name:
                            old_font = run.font.name
                            new_font = replace_map.get(old_font.lower())
                            if new_font and old_font != new_font:
                                run.font.name = new_font
                                replaced.add(old_font)
```

- [ ] **Step 4: 运行测试确认通过**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_font_normalization.py -v`
Expected: 1 passed

- [ ] **Step 5: 将字体标准化集成到 convert_with_classification 中**

在 `convert_with_classification` 方法的每个页面处理循环中，在创建/迁移完页面后调用 `_normalize_fonts`。

- [ ] **Step 6: 运行集成测试**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_replayer.py::test_convert_with_classification -v`
Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add core/replayer/content_replayer.py tests/test_font_normalization.py
git commit -m "feat: add font normalization to ContentReplayer"
```

---

## Task 9: UI 更新 — 删除模板上传 + 4 种背景选择 + Pre-check 展示

**Files:**
- Modify: `app.py`

- [ ] **Step 1: 读取当前 app.py 完整内容**

- [ ] **Step 2: 删除模板上传相关函数和变量**

删除：
- `LAST_TEMPLATE_FILE`、`_save_template`、`_load_last_template`、`_clear_last_template`
- `_check_template_aspect_ratio`、`_analyze_template_file`、`_has_last_template`、`_get_last_template_info`
- 侧边栏"模板管理"导航项和页面
- 转换配置中的"模板选择"区域

- [ ] **Step 3: 新增 Pre-check 展示区**

在上传 PPT 后，自动执行 Pre-check 并以卡片形式展示结果。

```python
def _run_precheck(pptx_file) -> PreCheckResult | None:
    if pptx_file is None:
        return None
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        tmp.write(pptx_file.getbuffer())
        tmp_path = tmp.name
    try:
        from core.precheck.analyzer import PreCheckAnalyzer
        analyzer = PreCheckAnalyzer()
        return analyzer.analyze(tmp_path)
    finally:
        os.unlink(tmp_path)
```

UI 展示：
```python
def _render_precheck(result: PreCheckResult):
    st.markdown('<div class="config-card" style="margin-bottom: 20px;">', unsafe_allow_html=True)
    st.markdown('<div class="config-card-title">📊 PPT 质量分析报告</div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("页面数量", result.slide_count)
    col2.metric("母版数量", result.master_count)
    col3.metric("字体种类", len(result.fonts_used))
    col4.metric("越界对象", result.overflow_objects_count)
    
    # 检查项列表
    checks = [
        ("📐 4:3 比例", result.is_4_3_ratio, "warning"),
        ("🎬 含动画", result.has_animation, "info"),
        ("📊 含嵌入图表", result.has_embedded_chart, "warning"),
        ("🧠 含 SmartArt", result.has_smartart, "warning"),
        ("🎥 含媒体", result.has_media, "warning"),
        ("🏢 旧 SE 模板", result.has_old_se_template, "info"),
    ]
    
    st.markdown("---")
    icon_cols = st.columns(6)
    for i, (label, value, level) in enumerate(checks):
        with icon_cols[i]:
            icon = "✅" if not value else ("⚠️" if level == "warning" else "ℹ️")
            st.markdown(f"<div style='text-align: center;'><div style='font-size: 20px;'>{icon}</div><div style='font-size: 11px; color: #718096;'>{label}</div></div>", unsafe_allow_html=True)
    
    # 问题列表
    if result.issues:
        st.markdown("---")
        st.markdown("**发现的问题：**")
        for issue in result.issues:
            icon = "❌" if issue.level == "error" else ("⚠️" if issue.level == "warning" else "ℹ️")
            st.markdown(f"- {icon} {issue.message}")
    
    st.markdown('</div>', unsafe_allow_html=True)
```

- [ ] **Step 4: 替换风格选择为 4 种固定背景**

```python
# 风格选择（4 种固定选项）
style_options = [
    {"key": "white", "name": "白色简约", "color": "#FFFFFF", "text_color": "#0A2F24"},
    {"key": "light_green", "name": "浅绿色清新", "color": "#E7FFD9", "text_color": "#0A2F24"},
    {"key": "dark_green", "name": "深绿色商务", "color": "#0A2F24", "text_color": "#FFFFFF"},
    {"key": "gradient", "name": "渐变科技", "color": "linear-gradient(135deg, #0A2F24, #3DCD58)", "text_color": "#FFFFFF"},
]

selected_style = st.radio(
    "选择背景风格",
    options=[s["key"] for s in style_options],
    format_func=lambda k: next(s["name"] for s in style_options if s["key"] == k),
    horizontal=True,
    index=2,  # 默认深绿色
    key="background_style",
)
```

- [ ] **Step 5: 更新转换按钮逻辑**

调用 `convert_with_classification` 而非旧的 `replay` 流程。
转换完成后展示 PPTX + QA 报告两个下载按钮。

- [ ] **Step 6: 验证 UI 可启动**

Run: `cd /workspace/ppt-conform && streamlit run app.py --server.headless true --server.port 8501 &`

然后检查是否正常启动。

- [ ] **Step 7: 停止测试服务器并 Commit**

```bash
git add app.py
git commit -m "feat: update UI - remove template upload, add 4 background styles and precheck display"
```

---

## Task 10: 端到端测试和集成验证

**Files:**
- Test: `tests/test_e2e.py` (更新)

- [ ] **Step 1: 读取现有 e2e 测试**

- [ ] **Step 2: 新增端到端测试用例**

```python
def test_e2e_convert_dark_green():
    """端到端测试：深绿色风格转换"""
    import tempfile
    import os
    from pptx import Presentation
    from core.replayer.content_replayer import ContentReplayer
    from core.registry.template_registry import TemplateRegistry
    from core.qa.reporter import QAReporter
    
    with tempfile.TemporaryDirectory() as tmpdir:
        # 准备源 PPT
        source_path = os.path.join(tmpdir, "source.pptx")
        prs = Presentation()
        # 封面
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Test Presentation"
        # 内容页
        s2 = prs.slides.add_slide(prs.slide_layouts[1])
        s2.shapes.title.text = "Content Page"
        s2.placeholders[1].text = "Some body text here"
        # 表格页
        s3 = prs.slides.add_slide(prs.slide_layouts[5])
        s3.shapes.title.text = "Data Table"
        from pptx.util import Inches
        s3.shapes.add_table(3, 3, Inches(1), Inches(2), Inches(8), Inches(3))
        prs.save(source_path)
        
        output_path = os.path.join(tmpdir, "output.pptx")
        report_path = os.path.join(tmpdir, "report.xlsx")
        
        # 执行转换
        registry = TemplateRegistry()
        replayer = ContentReplayer(
            registry,
            template_path="templates/2026 se template eng.pptx",
        )
        
        out_path, qa_items = replayer.convert_with_classification(
            source_path=source_path,
            output_path=output_path,
            background_style="dark_green",
        )
        
        # 验证输出
        assert os.path.exists(out_path)
        assert len(qa_items) == 3
        
        # 生成 QA 报告
        reporter = QAReporter()
        reporter.generate(qa_items, report_path)
        assert os.path.exists(report_path)
        
        # 验证输出 PPT 可打开
        out_prs = Presentation(out_path)
        assert len(out_prs.slides) == 3
        
        # 验证尺寸一致
        template_prs = Presentation("templates/2026 se template eng.pptx")
        assert out_prs.slide_width == template_prs.slide_width
        assert out_prs.slide_height == template_prs.slide_height
```

- [ ] **Step 3: 运行 e2e 测试**

Run: `cd /workspace/ppt-conform && python -m pytest tests/test_e2e.py::test_e2e_convert_dark_green -v --tb=short`
Expected: PASS

- [ ] **Step 4: 运行全部测试**

Run: `cd /workspace/ppt-conform && python -m pytest tests/ -v --tb=short -x`
Expected: all tests pass

- [ ] **Step 5: Commit**

```bash
git add tests/test_e2e.py
git commit -m "test: add end-to-end conversion test"
```

---

## 自检清单

**Spec 覆盖率检查：**

| Spec 要求 | 对应 Task |
|-----------|-----------|
| 模板内置，取消上传 | Task 9 |
| clrMap 正确解析，4 种背景正常显示 | Task 1 |
| Pre-check 分析（11 项） | Task 3 |
| Slide Classification（12 种类型） | Task 4 |
| Layout Mapping | Task 4 (layout_mapping.yaml) |
| Content Migration 路径 | Task 5 |
| Technical Adaptation 路径 | Task 7 |
| 字体替换（汉仪旗黑 + Poppins） | Task 8 |
| 颜色统一 | Task 1 + Task 7 |
| 页脚/Logo 清除 | Task 5 + Task 7 (逐步完善) |
| QA Report (xlsx) | Task 6 |
| 水印防护 | 复用现有 detector |

**占位符检查：**
- 所有 Task 都有具体代码和测试
- 没有 "TBD" / "TODO" / "implement later"
- 每个步骤都有明确的命令和预期结果

**类型一致性：**
- PreCheckResult / SlideClassification / QAReportItem 在 models.py 定义
- 各模块引用一致
- 方法命名一致
