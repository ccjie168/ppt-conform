from enum import Enum
from typing import Any, Literal
from pydantic import BaseModel


class WatermarkType(str, Enum):
    TEXT = "text"
    IMAGE = "image"
    SHAPE = "shape"


class WatermarkElement(BaseModel):
    type: WatermarkType
    slide_index: int
    shape_id: int | None = None
    text_content: str | None = None
    bounding_box: tuple[float, float, float, float] | None = None
    confidence: float


class WatermarkReport(BaseModel):
    detected: bool
    elements: list[WatermarkElement]
    summary: str


class TextFormat(BaseModel):
    """文本格式信息（从原PPT提取，作为模板格式的兜底）"""
    font_name: str | None = None
    font_size: float | None = None
    font_color: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    alignment: int | None = None
    line_spacing: float | None = None


class ShapeFormat(BaseModel):
    """形状格式信息（从原PPT提取，作为模板格式的兜底）"""
    fill_color: str | None = None
    fill_type: str | None = None
    line_color: str | None = None
    line_width: int | None = None
    line_style: str | None = None
    left: int | None = None
    top: int | None = None
    width: int | None = None
    height: int | None = None
    rotation: float | None = None
    shape_type: str | None = None


class ContentBlock(BaseModel):
    type: Literal["paragraph", "list", "image", "table", "chart", "other"]
    text: str | None = None
    level: int = 0
    content: Any | None = None
    semantic_role: Literal[
        "title", "subtitle", "body_main", "body_sidebar",
        "footer", "decoration", "caption", "unknown"
    ] = "unknown"
    original_placeholder_type: int | None = None
    original_placeholder_idx: int | None = None
    source_shape_id: int | None = None
    # 新增：与raw_shapes中对应形状的shape_id（用于精确去重）
    raw_shape_id: int | None = None
    # 新增：格式信息
    text_format: TextFormat | None = None
    shape_format: ShapeFormat | None = None


class SlideContentModel(BaseModel):
    slide_index: int
    title: str | None
    subtitle: str | None = None
    body_blocks: list[ContentBlock]
    notes: str | None = None
    original_layout_type: str | None = None
    raw_shapes: list[dict] = []
    extra_images: list[dict] = []
    extra_tables: list[dict] = []
    extra_text_shapes: list[dict] = []
    extra_autoshapes: list[dict] = []
    layout_features: dict = {}
    title_source: dict = {}
    # 新增：标题格式信息
    title_format: TextFormat | None = None


class MasterStyle(BaseModel):
    name: str
    id: str
    display_name: str
    thumbnail: bytes | None = None


class Placeholder(BaseModel):
    name: str
    type: Literal["title", "body", "subtitle", "footer", "slide_number", "header"]
    idx: int


class LayoutType(BaseModel):
    name: str
    master_id: str
    layout_id: str
    placeholders: list[Placeholder]


class TemplateOptions(BaseModel):
    headers: dict[str, str] = {}
    footers: dict[str, str] = {}
    icons: dict[str, bytes] = {}


class TemplateRegistry(BaseModel):
    masters: dict[str, MasterStyle] = {}
    layouts: dict[str, list[LayoutType]] = {}
    options: TemplateOptions = TemplateOptions()


class UserConfig(BaseModel):
    input_path: str
    output_path: str
    master_style: str
    include_header: bool = False
    include_footer: bool = False
    include_icon: bool = False
    per_page_layouts: dict[int, str] | None = None


class ValidationIssue(BaseModel):
    level: Literal["pass", "warning", "fail"]
    rule_id: str
    message: str
    slide_index: int | None = None


class ValidationReport(BaseModel):
    passed: bool
    issues: list[ValidationIssue]
    summary: str