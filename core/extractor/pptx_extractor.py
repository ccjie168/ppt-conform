from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from core.models import SlideContentModel, ContentBlock, WatermarkReport
from core.watermark.detector import WatermarkDetector


class PptxExtractor:
    """PPT 内容抽取器：从源 PPT 抽取内容模型，自动过滤水印"""

    def __init__(self):
        self.watermark_detector = WatermarkDetector()

    def extract(self, pptx_path: str) -> list[SlideContentModel]:
        if not Path(pptx_path).exists():
            raise FileNotFoundError(f"File not found: {pptx_path}")

        prs = Presentation(pptx_path)
        models = []

        for idx, slide in enumerate(prs.slides):
            model = self._extract_slide(slide, idx)
            models.append(model)

        return models

    def _extract_slide(self, slide, slide_index: int) -> SlideContentModel:
        title = None
        body_blocks = []
        notes_text = None

        # 提取备注
        try:
            if slide.has_notes_slide:
                notes_text = slide.notes_slide.notes_text_frame.text.strip() or None
        except Exception:
            pass

        for shape in slide.shapes:
            # 跳过占位符中的标题（单独处理）
            if shape == slide.shapes.title:
                title_text = self._get_shape_text(shape)
                if title_text:
                    watermark_report = self.watermark_detector.detect_text(title_text, slide_index)
                    if not watermark_report.detected:
                        title = title_text
                continue

            # 文本框 / 占位符
            if shape.has_text_frame:
                blocks = self._extract_text_blocks(shape, slide_index)
                body_blocks.extend(blocks)
                continue

            # 表格
            if shape.has_table:
                table_data = self._extract_table(shape.table)
                if table_data:
                    body_blocks.append(ContentBlock(
                        type="table",
                        text=None,
                        content=table_data,
                        level=0
                    ))
                continue

            # 图片：保留 blob 引用，后续重放时复制
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    image = shape.image
                    body_blocks.append(ContentBlock(
                        type="image",
                        text=None,
                        content={
                            "blob": image.blob,
                            "ext": image.ext,
                            "left": shape.left,
                            "top": shape.top,
                            "width": shape.width,
                            "height": shape.height,
                        },
                        level=0
                    ))
                except Exception:
                    pass
                continue

            # 组合形状：递归处理
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                blocks = self._extract_group(shape, slide_index)
                body_blocks.extend(blocks)
                continue

        return SlideContentModel(
            slide_index=slide_index,
            title=title,
            body_blocks=body_blocks,
            notes=notes_text,
            original_layout_type=self._detect_layout_type(slide, slide_index)
        )

    def _extract_text_blocks(self, shape, slide_index: int) -> list[ContentBlock]:
        """从文本框/占位符抽取段落，保留层级"""
        blocks = []
        tf = shape.text_frame

        for paragraph in tf.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue

            # 水印检测
            watermark_report = self.watermark_detector.detect_text(text, slide_index)
            if watermark_report.detected:
                continue

            blocks.append(ContentBlock(
                type="paragraph",
                text=text,
                level=paragraph.level or 0
            ))
        return blocks

    def _extract_group(self, group, slide_index: int) -> list[ContentBlock]:
        """递归处理组合形状中的子形状"""
        blocks = []
        for shape in group.shapes:
            if shape.has_text_frame:
                blocks.extend(self._extract_text_blocks(shape, slide_index))
        return blocks

    def _extract_table(self, table) -> list[list[str]]:
        """抽取表格内容为二维列表"""
        data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text.strip())
            data.append(row_data)
        return data

    def _get_shape_text(self, shape) -> str:
        """获取形状中的纯文本"""
        try:
            return shape.text_frame.text.strip()
        except Exception:
            return ""

    def _detect_layout_type(self, slide, slide_index: int) -> str:
        """根据 slide 的 layout 名称推断布局类型"""
        try:
            layout_name = (slide.slide_layout.name or "").lower()
            if "封面" in layout_name or "cover" in layout_name or "title" in layout_name:
                return "cover"
            if "章节" in layout_name or "section" in layout_name:
                return "section"
            if "结尾" in layout_name or "结束" in layout_name or "closing" in layout_name:
                return "closing"
        except Exception:
            pass
        if slide_index == 0:
            return "cover"
        return "content"
