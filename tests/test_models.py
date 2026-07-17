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