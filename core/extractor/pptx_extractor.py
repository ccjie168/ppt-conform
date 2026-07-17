from pathlib import Path
from pptx import Presentation
from core.models import SlideContentModel, ContentBlock, WatermarkReport
from core.watermark.detector import WatermarkDetector


class PptxExtractor:
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

        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text.strip()
                if not text:
                    continue

                watermark_report = self.watermark_detector.detect_text(text, slide_index)
                if watermark_report.detected:
                    continue

                if shape == slide.shapes.title:
                    title = text
                else:
                    body_blocks.append(ContentBlock(
                        type="paragraph",
                        text=text,
                        level=0
                    ))

            elif shape.shape_type == 13:
                pass

        return SlideContentModel(
            slide_index=slide_index,
            title=title,
            body_blocks=body_blocks
        )