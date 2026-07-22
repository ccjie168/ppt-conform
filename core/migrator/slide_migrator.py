from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from core.clrmap.resolver import ClrMapResolver


class SlideMigrator:
    """Content Migration: create new slide from template layout, migrate text content only."""

    def __init__(self, template_path: str, master_index: int):
        self.template_path = template_path
        self.master_index = master_index
        self.clr_resolver = ClrMapResolver(template_path, master_index)

    def migrate_slide(self, source_slide, target_prs, slide_type: str, layout_index: int):
        """
        Create a new slide based on target layout, migrate text content from source.
        Returns the new slide object.
        """
        # Get target layout
        if self.master_index >= len(target_prs.slide_masters):
            master = target_prs.slide_masters[0]
        else:
            master = target_prs.slide_masters[self.master_index]

        if layout_index >= len(master.slide_layouts):
            layout_index = 0
        target_layout = master.slide_layouts[layout_index]

        # 1. Create new slide
        new_slide = target_prs.slides.add_slide(target_layout)

        # 2. Extract source content
        title_text = self._extract_title(source_slide)
        subtitle_text = self._extract_subtitle(source_slide)
        body_paragraphs = self._extract_body_paragraphs(source_slide)

        # 3. Fill placeholders
        self._fill_title(new_slide, title_text)
        self._fill_subtitle(new_slide, subtitle_text)
        self._fill_body(new_slide, body_paragraphs)

        return new_slide

    def _extract_title(self, slide) -> str:
        if slide.shapes.title:
            return slide.shapes.title.text or ""
        # Fallback: find largest text shape
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
        # Try subtitle placeholder (type 4 in python-pptx)
        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            # PP_PLACEHOLDER.SUBTITLE = 4
            if ph_type == 4 and shape.has_text_frame:
                return shape.text_frame.text.strip()
        # Fallback: second largest text shape
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
        title_shape = slide.shapes.title
        for shape in slide.shapes:
            if title_shape and shape == title_shape:
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
        # No subtitle placeholder, try body placeholder
        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            if ph_type == 2 and shape.has_text_frame:  # BODY = 2
                shape.text_frame.text = text
                return

    def _fill_body(self, slide, paragraphs: list[str]):
        if not paragraphs:
            return
        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            if ph_type == 2 and shape.has_text_frame:  # BODY = 2
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
