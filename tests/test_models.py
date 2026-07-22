import pytest
from pydantic import ValidationError
from core.models import (
    ContentBlock, SlideContentModel,
    MasterStyle, LayoutType, Placeholder,
    TemplateRegistry, TemplateOptions,
    UserConfig, ValidationIssue, ValidationReport,
    WatermarkType, WatermarkElement, WatermarkReport
)


def test_content_block():
    block = ContentBlock(type="paragraph", text="hello", level=0)
    assert block.type == "paragraph"
    assert block.text == "hello"


def test_slide_content_model():
    blocks = [ContentBlock(type="paragraph", text="test")]
    model = SlideContentModel(slide_index=0, title="Title", body_blocks=blocks)
    assert model.slide_index == 0
    assert model.title == "Title"
    assert len(model.body_blocks) == 1


def test_watermark_report():
    elements = [WatermarkElement(
        type=WatermarkType.TEXT,
        slide_index=0,
        text_content="TRAE AI 生成",
        confidence=0.95
    )]
    report = WatermarkReport(detected=True, elements=elements, summary="Found 1 watermark")
    assert report.detected is True
    assert len(report.elements) == 1


def test_user_config():
    config = UserConfig(
        input_path="/path/input.pptx",
        output_path="/path/output.pptx",
        master_style="F2",
        include_header=True,
        include_footer=True
    )
    assert config.master_style == "F2"
    assert config.include_header is True


def test_master_styles_has_4_styles():
    import yaml
    from pathlib import Path
    config_path = Path(__file__).parent.parent / "config" / "master_styles.yaml"
    with open(config_path, "r") as f:
        config = yaml.safe_load(f)
    assert "master_styles" in config
    assert len(config["master_styles"]) == 4


def test_precheck_issue_model():
    from core.models import PreCheckIssue
    issue = PreCheckIssue(
        level="warning",
        rule_id="too_many_masters",
        message="母版数量过多，存在模板污染风险",
    )
    assert issue.level == "warning"
    assert issue.rule_id == "too_many_masters"


def test_slide_classification_model():
    from core.models import SlideClassification
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
    from core.models import QAReportItem
    item = QAReportItem(
        slide_no=1,
        detected_type="Cover",
        applied_layout="Title slide simple",
        migration_mode="migration",
    )
    assert item.slide_no == 1
    assert item.need_manual_review is False


def test_conversion_config_model():
    from core.models import ConversionConfig
    config = ConversionConfig(
        input_path="/tmp/input.pptx",
        output_path="/tmp/output.pptx",
    )
    assert config.background_style == "dark_green"
    assert config.include_footer is True