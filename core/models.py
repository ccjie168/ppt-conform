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


class SlideContentModel(BaseModel):
    slide_index: int
    title: str | None
    body_blocks: list[ContentBlock]
    notes: str | None = None
    original_layout_type: str | None = None


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