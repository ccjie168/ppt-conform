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
                table_data = self._extract_table_full(shape.table)
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
            shape_type = shape.shape_type

            if shape.has_text_frame:
                return self._extract_text_shape(shape, slide_index)
            elif shape_type == MSO_SHAPE_TYPE.PICTURE:
                return self._extract_image_shape(shape)
            elif shape.has_table:
                return self._extract_table_shape(shape)
            elif shape_type == MSO_SHAPE_TYPE.GROUP:
                return self._extract_group_shape(shape, slide_index)
            elif shape_type == MSO_SHAPE_TYPE.CHART:
                return self._extract_chart_shape(shape)
            elif shape_type == MSO_SHAPE_TYPE.OLE_OBJECT:
                return self._extract_ole_shape(shape)
            else:
                # 尝试提取自选图形的几何信息
                return self._extract_auto_shape(shape)
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
                "line_spacing": None,
                "space_before": None,
                "space_after": None,
            }

            try:
                if paragraph.line_spacing is not None:
                    para_data["line_spacing"] = paragraph.line_spacing
                if paragraph.space_before is not None:
                    para_data["space_before"] = paragraph.space_before
                if paragraph.space_after is not None:
                    para_data["space_after"] = paragraph.space_after
            except Exception:
                pass

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
                "crop_left": getattr(shape, "crop_left", None),
                "crop_top": getattr(shape, "crop_top", None),
                "crop_right": getattr(shape, "crop_right", None),
                "crop_bottom": getattr(shape, "crop_bottom", None),
            }
        except Exception:
            return None

    def _extract_table_shape(self, shape) -> dict:
        try:
            table_data = self._extract_table_full(shape.table)
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

    def _extract_table_full(self, table) -> dict:
        """完整提取表格：单元格文本、背景色、字体格式、合并单元格"""
        cells = []
        col_widths = []
        row_heights = []

        try:
            for col_idx in range(len(table.columns)):
                col_widths.append(table.columns[col_idx].width)
        except Exception:
            pass

        try:
            for row_idx in range(len(table.rows)):
                row_heights.append(table.rows[row_idx].height)
        except Exception:
            pass

        merged_cells = []
        for r, row in enumerate(table.rows):
            for c, cell in enumerate(row.cells):
                cell_data = {
                    "row": r,
                    "col": c,
                    "text": cell.text.strip(),
                    "fill_color": None,
                    "font_name": None,
                    "font_size": None,
                    "bold": None,
                    "italic": None,
                    "color": None,
                    "alignment": None,
                    "vertical_anchor": None,
                    "margin_left": None,
                    "margin_right": None,
                    "margin_top": None,
                    "margin_bottom": None,
                }

                try:
                    fill = cell.fill
                    if fill.type == 1:  # solid
                        if fill.fore_color and fill.fore_color.rgb:
                            cell_data["fill_color"] = str(fill.fore_color.rgb)
                except Exception:
                    pass

                try:
                    if cell.vertical_anchor is not None:
                        cell_data["vertical_anchor"] = str(cell.vertical_anchor)
                    cell_data["margin_left"] = cell.margin_left
                    cell_data["margin_right"] = cell.margin_right
                    cell_data["margin_top"] = cell.margin_top
                    cell_data["margin_bottom"] = cell.margin_bottom
                except Exception:
                    pass

                try:
                    tf = cell.text_frame
                    if tf.paragraphs:
                        p = tf.paragraphs[0]
                        if p.alignment is not None:
                            cell_data["alignment"] = str(p.alignment)
                        if p.runs:
                            run = p.runs[0]
                            font = run.font
                            if font.name:
                                cell_data["font_name"] = font.name
                            if font.size:
                                cell_data["font_size"] = font.size
                            if font.bold is not None:
                                cell_data["bold"] = font.bold
                            if font.italic is not None:
                                cell_data["italic"] = font.italic
                            try:
                                if font.color and font.color.rgb:
                                    cell_data["color"] = str(font.color.rgb)
                            except Exception:
                                pass
                except Exception:
                    pass

                # 检测合并单元格
                try:
                    if cell.span_h > 1 or cell.span_v > 1:
                        merged_cells.append({
                            "start_row": r,
                            "start_col": c,
                            "row_span": cell.span_v,
                            "col_span": cell.span_h,
                        })
                except Exception:
                    pass

                cells.append(cell_data)

        return {
            "cells": cells,
            "rows": len(table.rows),
            "cols": len(table.columns),
            "col_widths": col_widths,
            "row_heights": row_heights,
            "merged_cells": merged_cells,
        }

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

    def _extract_chart_shape(self, shape) -> dict:
        """完整提取图表：数据、类别、系列颜色、图例、坐标轴"""
        try:
            chart = shape.chart
            chart_data = []
            categories = []
            series_colors = []
            chart_title = None
            has_legend = False
            legend_position = None

            try:
                if chart.has_title and chart.chart_title:
                    chart_title = chart.chart_title.text_frame.text
            except Exception:
                pass

            try:
                if chart.has_legend:
                    has_legend = True
                    legend_position = str(chart.legend.position) if chart.legend else None
            except Exception:
                pass

            try:
                for series in chart.series:
                    series_data = {
                        "name": series.name,
                        "values": [],
                        "color": None,
                    }
                    try:
                        values = series.values
                        if values:
                            series_data["values"] = list(values)
                    except Exception:
                        for point in series.points:
                            try:
                                series_data["values"].append(point.value)
                            except Exception:
                                series_data["values"].append(None)

                    try:
                        if series.format.fill.type == 1:
                            if series.format.fill.fore_color and series.format.fill.fore_color.rgb:
                                series_data["color"] = str(series.format.fill.fore_color.rgb)
                    except Exception:
                        pass

                    chart_data.append(series_data)
            except Exception:
                pass

            try:
                if chart.plots and chart.plots[0].categories:
                    for cat in chart.plots[0].categories:
                        try:
                            categories.append(str(cat))
                        except Exception:
                            pass
            except Exception:
                pass

            if not categories and chart_data:
                max_len = max(len(s.get("values", [])) for s in chart_data) if chart_data else 0
                categories = [f"类别{i+1}" for i in range(max_len)]

            # 提取坐标轴信息
            x_axis = {}
            y_axis = {}
            try:
                if chart.category_axis:
                    ca = chart.category_axis
                    x_axis = {
                        "has_title": ca.has_title,
                        "title": ca.axis_title.text_frame.text if ca.has_title and ca.axis_title else None,
                        "visible": ca.visible,
                    }
            except Exception:
                pass
            try:
                if chart.value_axis:
                    va = chart.value_axis
                    y_axis = {
                        "has_title": va.has_title,
                        "title": va.axis_title.text_frame.text if va.has_title and va.axis_title else None,
                        "visible": va.visible,
                        "minimum_scale": va.minimum_scale,
                        "maximum_scale": va.maximum_scale,
                    }
            except Exception:
                pass

            return {
                "type": "chart",
                "left": shape.left,
                "top": shape.top,
                "width": shape.width,
                "height": shape.height,
                "chart_type": str(chart.chart_type),
                "data": chart_data,
                "categories": categories,
                "chart_title": chart_title,
                "has_legend": has_legend,
                "legend_position": legend_position,
                "x_axis": x_axis,
                "y_axis": y_axis,
            }
        except Exception:
            return None

    def _extract_ole_shape(self, shape) -> dict:
        try:
            ole_format = shape.ole_format
            if ole_format and ole_format.binary:
                return {
                    "type": "ole",
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height,
                    "prog_id": ole_format.prog_id,
                    "blob": ole_format.binary,
                }
        except Exception:
            pass
        return None

    def _extract_auto_shape(self, shape) -> dict | None:
        """提取自选图形的几何信息"""
        try:
            shape_type = shape.shape_type
            return {
                "type": "autoshape",
                "shape_type": str(shape_type),
                "left": shape.left,
                "top": shape.top,
                "width": shape.width,
                "height": shape.height,
                "rotation": getattr(shape, "rotation", 0),
                "fill_color": None,
                "line_color": None,
                "text": shape.text_frame.text if shape.has_text_frame else None,
            }
        except Exception:
            return None

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
        """简化版表格提取（向后兼容）"""
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
