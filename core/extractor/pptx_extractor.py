from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from core.models import SlideContentModel, ContentBlock, WatermarkReport
from core.watermark.detector import WatermarkDetector


class PptxExtractor:
    """PPT 内容抽取器：从源 PPT 抽取内容模型，保留原始格式，自动过滤水印"""

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
        raw_shapes = []

        try:
            if slide.has_notes_slide:
                notes_text = slide.notes_slide.notes_text_frame.text.strip() or None
        except Exception:
            pass

        for shape in slide.shapes:
            shape_data = self._extract_shape(shape, slide_index)
            if shape_data:
                raw_shapes.append(shape_data)

            if shape == slide.shapes.title:
                title_text = self._get_shape_text(shape)
                if title_text:
                    watermark_report = self.watermark_detector.detect_text(title_text, slide_index)
                    if not watermark_report.detected:
                        title = title_text
                continue

            if shape.has_text_frame:
                blocks = self._extract_text_blocks(shape, slide_index)
                body_blocks.extend(blocks)
                continue

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

            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                blocks = self._extract_group(shape, slide_index)
                body_blocks.extend(blocks)
                continue

        return SlideContentModel(
            slide_index=slide_index,
            title=title,
            body_blocks=body_blocks,
            notes=notes_text,
            original_layout_type=self._detect_layout_type(slide, slide_index),
            raw_shapes=raw_shapes,
        )

    def _extract_shape(self, shape, slide_index: int) -> dict | None:
        try:
            if shape == slide_index:
                pass
            shape_type = shape.shape_type

            if shape.has_text_frame:
                return self._extract_text_shape(shape, slide_index)
            elif shape_type == MSO_SHAPE_TYPE.PICTURE:
                return self._extract_image_shape(shape)
            elif shape.has_table:
                return self._extract_table_shape(shape)
            elif shape_type == MSO_SHAPE_TYPE.GROUP:
                return self._extract_group_shape(shape, slide_index)
            else:
                return None
        except Exception:
            return None

    def _extract_text_shape(self, shape, slide_index: int) -> dict:
        paragraphs = []
        tf = shape.text_frame

        for paragraph in tf.paragraphs:
            para_data = {
                "text": paragraph.text,
                "level": paragraph.level or 0,
                "alignment": str(paragraph.alignment) if paragraph.alignment else None,
                "runs": [],
            }

            for run in paragraph.runs:
                run_data = {
                    "text": run.text,
                    "font_name": run.font.name,
                    "font_size": run.font.size,
                    "bold": run.font.bold,
                    "italic": run.font.italic,
                    "underline": run.font.underline,
                    "color": None,
                }
                try:
                    if run.font.color and run.font.color.rgb:
                        run_data["color"] = str(run.font.color.rgb)
                except Exception:
                    pass
                para_data["runs"].append(run_data)

            paragraphs.append(para_data)

        return {
            "type": "text",
            "left": shape.left,
            "top": shape.top,
            "width": shape.width,
            "height": shape.height,
            "paragraphs": paragraphs,
            "shape_name": shape.name,
        }

    def _extract_image_shape(self, shape) -> dict:
        try:
            image = shape.image
            return {
                "type": "image",
                "left": shape.left,
                "top": shape.top,
                "width": shape.width,
                "height": shape.height,
                "blob": image.blob,
                "ext": image.ext,
            }
        except Exception:
            return None

    def _extract_table_shape(self, shape) -> dict:
        try:
            table_data = self._extract_table(shape.table)
            return {
                "type": "table",
                "left": shape.left,
                "top": shape.top,
                "width": shape.width,
                "height": shape.height,
                "data": table_data,
            }
        except Exception:
            return None

    def _extract_group_shape(self, group, slide_index: int) -> dict:
        shapes_data = []
        for shape in group.shapes:
            shape_data = self._extract_shape(shape, slide_index)
            if shape_data:
                shapes_data.append(shape_data)
        return {
            "type": "group",
            "left": group.left,
            "top": group.top,
            "width": group.width,
            "height": group.height,
            "shapes": shapes_data,
        }

    def _extract_text_blocks(self, shape, slide_index: int) -> list[ContentBlock]:
        blocks = []
        tf = shape.text_frame

        for paragraph in tf.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue

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
        blocks = []
        for shape in group.shapes:
            if shape.has_text_frame:
                blocks.extend(self._extract_text_blocks(shape, slide_index))
        return blocks

    def _extract_table(self, table) -> list[list[str]]:
        data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text.strip())
            data.append(row_data)
        return data

    def _get_shape_text(self, shape) -> str:
        try:
            return shape.text_frame.text.strip()
        except Exception:
            return ""

    def _detect_layout_type(self, slide, slide_index: int) -> str:
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
