"""
ObjectMigrator: migrate non-placeholder shapes from source slide to target slide.
Handles pictures, tables, text boxes, auto-shapes, lines, etc.
Filters out decorative objects (old logos, watermarks, footers).
"""

from io import BytesIO
from copy import deepcopy
from pptx.enum.shapes import MSO_SHAPE_TYPE


class ObjectMigrator:
    """Migrate non-placeholder objects while filtering decorative ones."""

    # Footer-area threshold (88% of slide height)
    FOOTER_RATIO = 0.88
    # Small-logo size threshold (EMU)
    LOGO_MAX_W, LOGO_MAX_H = 500_000, 200_000
    # Tiny decoration threshold (EMU)
    TINY_SIZE = 50_000

    FOOTER_KEYWORDS = [
        "se.com", "schneider", "copyright", "\u00a9", "page", "footer",
    ]
    WATERMARK_KEYWORDS = [
        "trademark", "confidential", "draft", "watermark",
        "trae ai", "ai generated",
    ]

    def migrate_objects(self, source_slide, new_slide) -> tuple[int, int]:
        """
        Migrate all non-placeholder, non-decorative objects.
        Returns (migrated_count, skipped_count).
        """
        migrated = 0
        skipped = 0

        for shape in source_slide.shapes:
            if shape.is_placeholder:
                continue
            if self.is_decorative(shape, source_slide):
                skipped += 1
                continue

            try:
                if self._migrate_single(shape, new_slide):
                    migrated += 1
                else:
                    skipped += 1
            except Exception:
                skipped += 1

        return migrated, skipped

    def _migrate_single(self, shape, new_slide) -> bool:
        """Dispatch to type-specific migration. Returns True if handled."""
        st = shape.shape_type
        if st == MSO_SHAPE_TYPE.PICTURE:
            self.migrate_picture(shape, new_slide)
            return True
        if st == MSO_SHAPE_TYPE.TABLE:
            self.migrate_table(shape, new_slide)
            return True
        if st == MSO_SHAPE_TYPE.TEXT_BOX:
            self.migrate_text_box(shape, new_slide)
            return True
        if st in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.FREEFORM):
            self.migrate_shape_by_xml(shape, new_slide)
            return True
        if st == MSO_SHAPE_TYPE.LINE:
            self.migrate_line(shape, new_slide)
            return True
        # CHART, GROUP, MEDIA, OLE_OBJECT intentionally not handled
        return False

    # ------------------------------------------------------------------
    #  Decorative-object detection
    # ------------------------------------------------------------------

    def is_decorative(self, shape, source_slide) -> bool:
        """True if shape is an old logo, watermark, footer, etc."""
        # 1. Footer-area objects
        try:
            prs = source_slide.part.package.presentation_part.presentation
            slide_height = prs.slide_height
            if shape.top > slide_height * self.FOOTER_RATIO:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    if shape.width < self.LOGO_MAX_W and shape.height < self.LOGO_MAX_H:
                        return True
                if shape.has_text_frame:
                    text = shape.text_frame.text.lower()
                    if any(kw in text for kw in self.FOOTER_KEYWORDS):
                        return True
        except Exception:
            pass

        # 2. Watermark / confidential text
        if shape.has_text_frame:
            text = shape.text_frame.text.lower()
            if any(kw in text for kw in self.WATERMARK_KEYWORDS):
                return True

        # 3. Tiny decoration shapes without text
        try:
            if shape.width < self.TINY_SIZE and shape.height < self.TINY_SIZE:
                if not (shape.has_text_frame and shape.text_frame.text.strip()):
                    return True
        except Exception:
            pass

        return False

    # ------------------------------------------------------------------
    #  Per-type migration
    # ------------------------------------------------------------------

    def migrate_picture(self, shape, new_slide):
        image_stream = BytesIO(shape.image.blob)
        new_slide.shapes.add_picture(
            image_stream,
            shape.left, shape.top,
            shape.width, shape.height,
        )

    def migrate_table(self, shape, new_slide):
        table = shape.table
        rows, cols = len(table.rows), len(table.columns)
        new_table = new_slide.shapes.add_table(
            rows, cols,
            shape.left, shape.top,
            shape.width, shape.height,
        ).table
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                new_table.cell(i, j).text = cell.text

    def migrate_text_box(self, shape, new_slide):
        txBox = new_slide.shapes.add_textbox(
            shape.left, shape.top,
            shape.width, shape.height,
        )
        tf = txBox.text_frame
        tf.word_wrap = shape.text_frame.word_wrap
        for i, para in enumerate(shape.text_frame.paragraphs):
            new_para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            new_para.text = para.text
            new_para.level = para.level

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
            MSO_SHAPE.LINE_INVERSE,
            shape.left, shape.top,
            shape.width, shape.height,
        )
