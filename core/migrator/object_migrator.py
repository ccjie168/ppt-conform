"""
ObjectMigrator: migrate non-placeholder shapes from source slide to target slide.
Handles pictures, tables, text boxes, auto-shapes, lines, etc.
Filters out decorative objects (old logos, watermarks, footers).

NEW FEATURES:
1. Text box semantic classification (title/subtitle/body/footer/watermark)
2. Format normalization (font size mapping, color mapping to theme)
3. Returns semantic classification for caller to fill placeholders
"""

from io import BytesIO
from copy import deepcopy
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
from pptx.dml.color import RGBColor


class ObjectMigrator:
    """Migrate non-placeholder objects while filtering decorative ones."""

    FOOTER_RATIO = 0.88
    LOGO_MAX_W, LOGO_MAX_H = 500_000, 200_000
    TINY_SIZE = 50_000

    FOOTER_KEYWORDS = ["se.com", "schneider", "copyright", "\u00a9", "page", "footer"]
    WATERMARK_KEYWORDS = ["trademark", "confidential", "draft", "watermark", "trae ai", "ai generated", "ai 生成"]

    # Font size semantics (pt)
    SIZE_THRESHOLDS = {
        "title": (28, float("inf")),
        "subtitle": (20, 27),
        "body": (14, 19),
        "caption": (10, 13),
        "footer": (0, 9),
    }

    # Font size normalization mapping (pt)
    SIZE_NORMALIZATION = {
        "title": 24,
        "subtitle": 16,
        "body": 11,
        "caption": 9,
        "footer": 8,
    }

    # Color mapping: original color -> normalized color
    COLOR_MAPPING = {
        "#1f3a68": "#0A2F24",
        "#000000": "#0A2F24",
        "#666666": "#718096",
        "#ffffff": "#FFFFFF",
        "#cccccc": "#A0AEC0",
        "#3dcd58": "#3DCD58",
        "#f39200": "#3DCD58",
        "#cc0000": "#DC2626",
    }

    def __init__(self, text_color="#0A2F24", bg_dark=False):
        self.text_color = text_color
        self.bg_dark = bg_dark

    def migrate_objects(self, source_slide, new_slide) -> tuple[int, int, dict]:
        """
        Migrate all non-placeholder, non-decorative objects.
        Returns (migrated_count, skipped_count, semantic_info).
        semantic_info = {"title_text": str, "subtitle_text": str, "body_texts": list}
        """
        migrated = 0
        skipped = 0
        semantic_info = {"title_text": "", "subtitle_text": "", "body_texts": []}

        classified_text_boxes = self._classify_text_boxes(source_slide)

        for shape in source_slide.shapes:
            if shape.is_placeholder:
                continue
            if self.is_decorative(shape, source_slide):
                skipped += 1
                continue

            try:
                if self._migrate_single(shape, new_slide, classified_text_boxes):
                    migrated += 1
                else:
                    skipped += 1
            except Exception:
                skipped += 1

        semantic_info.update(classified_text_boxes)
        return migrated, skipped, semantic_info

    def _classify_text_boxes(self, source_slide) -> dict:
        """Classify text boxes by semantics and extract content."""
        title_text = ""
        subtitle_text = ""
        body_texts = []

        for shape in source_slide.shapes:
            if shape.is_placeholder:
                continue
            if not shape.has_text_frame:
                continue

            text = shape.text_frame.text.strip()
            if not text:
                continue

            semantics = self._analyze_text_box_semantics(shape, source_slide)

            if semantics == "title" and not title_text:
                title_text = text
            elif semantics == "subtitle" and not subtitle_text:
                subtitle_text = text
            elif semantics == "body":
                body_texts.append(text)
            elif semantics == "caption":
                body_texts.append(text)

        return {
            "title_text": title_text,
            "subtitle_text": subtitle_text,
            "body_texts": body_texts,
        }

    def _analyze_text_box_semantics(self, shape, source_slide) -> str:
        """Analyze text box semantics based on position, size, and font."""
        try:
            prs = source_slide.part.package.presentation_part.presentation
            slide_height = prs.slide_height
            slide_width = prs.slide_width
        except Exception:
            slide_height = 6858000
            slide_width = 12192000

        text = shape.text_frame.text.lower()

        if any(kw in text for kw in self.WATERMARK_KEYWORDS):
            return "watermark"

        if shape.top > slide_height * self.FOOTER_RATIO:
            return "footer"

        font_size = self._get_shape_font_size(shape)

        for semantic, (min_size, max_size) in self.SIZE_THRESHOLDS.items():
            if min_size <= font_size <= max_size:
                return semantic

        if shape.top < slide_height * 0.25:
            return "title"
        elif shape.top < slide_height * 0.5:
            return "subtitle"
        else:
            return "body"

    def _get_shape_font_size(self, shape) -> float:
        """Get the font size of a text box (average or largest)."""
        sizes = []
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.size:
                        sizes.append(run.font.size.pt)
        return max(sizes) if sizes else 11

    def _get_shape_font_color(self, shape) -> str:
        """Get the font color of a text box as hex string."""
        colors = []
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.color and run.font.color.rgb:
                        colors.append("#" + run.font.color.rgb.hex)
        return colors[0] if colors else "#000000"

    def _migrate_single(self, shape, new_slide, classified_boxes) -> bool:
        """Dispatch to type-specific migration with format normalization."""
        st = shape.shape_type
        if st == MSO_SHAPE_TYPE.PICTURE:
            self.migrate_picture(shape, new_slide)
            return True
        if st == MSO_SHAPE_TYPE.TABLE:
            self.migrate_table(shape, new_slide)
            return True
        if st == MSO_SHAPE_TYPE.TEXT_BOX:
            semantics = self._analyze_text_box_semantics(shape, new_slide)
            self.migrate_text_box(shape, new_slide, semantics)
            return True
        if st in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.FREEFORM):
            self.migrate_shape_by_xml(shape, new_slide)
            return True
        if st == MSO_SHAPE_TYPE.LINE:
            self.migrate_line(shape, new_slide)
            return True
        return False

    def is_decorative(self, shape, source_slide) -> bool:
        """True if shape is an old logo, watermark, footer, etc."""
        try:
            prs = source_slide.part.package.presentation_part.presentation
            slide_height = prs.slide_height
        except Exception:
            slide_height = 6858000

        if shape.top > slide_height * self.FOOTER_RATIO:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if shape.width < self.LOGO_MAX_W and shape.height < self.LOGO_MAX_H:
                    return True
            if shape.has_text_frame:
                text = shape.text_frame.text.lower()
                if any(kw in text for kw in self.FOOTER_KEYWORDS):
                    return True

        if shape.has_text_frame:
            text = shape.text_frame.text.lower()
            if any(kw in text for kw in self.WATERMARK_KEYWORDS):
                return True

        try:
            if shape.width < self.TINY_SIZE and shape.height < self.TINY_SIZE:
                if not (shape.has_text_frame and shape.text_frame.text.strip()):
                    return True
        except Exception:
            pass

        return False

    def migrate_picture(self, shape, new_slide):
        image_stream = BytesIO(shape.image.blob)
        new_slide.shapes.add_picture(
            image_stream, shape.left, shape.top, shape.width, shape.height
        )

    def migrate_table(self, shape, new_slide):
        table = shape.table
        rows, cols = len(table.rows), len(table.columns)
        new_table = new_slide.shapes.add_table(
            rows, cols, shape.left, shape.top, shape.width, shape.height
        ).table
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                new_table.cell(i, j).text = cell.text

    def migrate_text_box(self, shape, new_slide, semantics="body"):
        """Migrate text box with format normalization based on semantics."""
        txBox = new_slide.shapes.add_textbox(
            shape.left, shape.top, shape.width, shape.height
        )
        tf = txBox.text_frame
        tf.word_wrap = shape.text_frame.word_wrap

        target_size = self.SIZE_NORMALIZATION.get(semantics, 11)
        target_color = self.text_color

        for i, para in enumerate(shape.text_frame.paragraphs):
            new_para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            new_para.text = para.text
            new_para.level = para.level
            new_para.font.size = Pt(target_size)
        try:
            new_para.font.color.rgb = RGBColor.from_string(target_color.lstrip("#"))
        except Exception:
            pass

    def migrate_shape_by_xml(self, shape, new_slide):
        new_sp = deepcopy(shape._element)
        nsmap = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        }
        nvSpPr = new_sp.find(".//p:nvSpPr", nsmap)
        if nvSpPr is not None:
            ph = nvSpPr.find(".//p:ph", nsmap)
            if ph is not None:
                nvSpPr.remove(ph)
        new_slide.shapes._spTree.insert_element_before(new_sp, "p:extLst")

    def migrate_line(self, shape, new_slide):
        from pptx.enum.shapes import MSO_SHAPE
        new_slide.shapes.add_shape(
            MSO_SHAPE.LINE_INVERSE, shape.left, shape.top, shape.width, shape.height
        )
