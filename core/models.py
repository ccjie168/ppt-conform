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


class ContentBlock(BaseModel):
    type: Literal["paragraph", "list", "image", "table", "chart", "other"]
    text: str | None = None
    level: int = 0
    content: Any | None = None
    # 语义角色：标记内容在幻灯片中的结构角色，用于回填时匹配模板占位符
    semantic_role: Literal[
        "title", "subtitle", "body_main", "body_sidebar",
        "footer", "decoration", "caption", "unknown"
    ] = "unknown"
    # 原PPT中该内容来自的占位符类型（1=title, 2=body, 4=subtitle, 7=text, 等）
    original_placeholder_type: int | None = None
    # 原PPT中该内容来自的占位符索引
    original_placeholder_idx: int | None = None
    # 来源形状ID，用于建立内容与结构的映射关系
    source_shape_id: int | None = None


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
    # 布局特征：记录原PPT的布局结构特征，用于回填时选择模板布局
    layout_features: dict = {}
    # 标题的来源占位符信息（用于回填时匹配模板占位符）
    title_source: dict = {}


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