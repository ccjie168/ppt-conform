from pathlib import Path
from io import BytesIO
from pptx import Presentation
from pptx.util import Emu
from pptx.dml.color import RGBColor
from core.models import SlideContentModel, UserConfig
from core.registry.template_registry import TemplateRegistry


class ContentReplayer:
    """内容重放器：将抽取的内容模型重放到目标模板，保留源格式，只替换背景"""

    def __init__(self, registry: TemplateRegistry, template_path: str | None = None):
        self.registry = registry
        self.template_path = template_path

    def replay(self, content_models: list[SlideContentModel], config: UserConfig) -> str:
        if self.template_path and Path(self.template_path).exists():
            output_prs = Presentation(self.template_path)
            self._clear_slides(output_prs)
        else:
            output_prs = Presentation()

        selected_master_index = int(config.master_style) if config.master_style.isdigit() else 0

        if selected_master_index >= len(output_prs.slide_masters):
            selected_master_index = 0

        selected_master = output_prs.slide_masters[selected_master_index]

        for slide_idx, model in enumerate(content_models):
            if selected_master:
                layout_name = self._determine_layout(model)
                layout_index = self._resolve_layout_index(layout_name, selected_master, config.master_style)
                slide_layouts = list(selected_master.slide_layouts)
                if layout_index >= len(slide_layouts):
                    layout_index = 0
                slide = output_prs.slides.add_slide(slide_layouts[layout_index])
            else:
                slide = output_prs.slides.add_slide(output_prs.slide_layouts[1])

            self._clear_placeholders(slide)
            self._copy_slide_content(slide, model)
            self._check_and_fix_overflow(slide)

        output_prs.save(config.output_path)
        return config.output_path

    def _clear_slides(self, prs) -> None:
        sldIdLst = prs.slides._sldIdLst
        rIds = [sldId.rId for sldId in list(sldIdLst)]
        for rId in rIds:
            try:
                prs.part.drop_rel(rId)
            except KeyError:
                pass
        for child in list(sldIdLst):
            sldIdLst.remove(child)

    def _clear_placeholders(self, slide) -> None:
        """清空幻灯片上的内容占位符，保留页眉页脚和日期/页码占位符"""
        header_footer_types = (13, 14, 15, 16)
        
        shapes_to_remove = []
        for shape in slide.shapes:
            try:
                if shape.is_placeholder:
                    phf = shape.placeholder_format
                    if phf.type not in header_footer_types:
                        shapes_to_remove.append(shape)
            except Exception:
                continue

        spTree = slide.shapes._spTree
        for shape in shapes_to_remove:
            try:
                spTree.remove(shape._element)
            except Exception:
                pass

    def _build_layout_map_from_master(self, master) -> dict:
        mapping = {}
        for idx, layout in enumerate(master.slide_layouts):
            name = (layout.name or "").lower()
            mapping[name] = idx
            if "封面" in name or "cover" in name or "title" in name:
                mapping["cover"] = idx
            elif "章节" in name or "section" in name or "节" in name:
                mapping["section"] = idx
            elif "内容" in name or "content" in name:
                mapping["content"] = idx
            elif "结尾" in name or "结束" in name or "closing" in name or "end" in name:
                mapping["closing"] = idx
        return mapping

    def _resolve_layout_index(
        self, layout_name: str, master, master_style: str
    ) -> int:
        layout_map = self._build_layout_map_from_master(master)
        if layout_name in layout_map:
            return layout_map[layout_name]

        layout_info = self.registry.get_layout_by_name(master_style, layout_name)
        if layout_info:
            idx = layout_info.get("index", 0)
            if idx < len(list(master.slide_layouts)):
                return idx

        if layout_name == "cover":
            return 0
        return min(1, len(list(master.slide_layouts)) - 1) if master else 1

    def _copy_slide_content(self, slide, model: SlideContentModel) -> None:
        if not model.raw_shapes:
            self._fill_simple_content(slide, model)
            return

        for shape_data in model.raw_shapes:
            self._add_shape_from_data(slide, shape_data)

    def _fill_simple_content(self, slide, model: SlideContentModel) -> None:
        if model.title:
            self._fill_title(slide, model.title)

        text_blocks = [b for b in model.body_blocks if b.type == "paragraph" and b.text]
        other_blocks = [b for b in model.body_blocks if b.type in ("table", "image")]

        if text_blocks:
            body_placeholder = self._find_body_placeholder(slide)
            if body_placeholder:
                self._fill_text_into_placeholder(body_placeholder, text_blocks)
            else:
                self._add_text_as_new_shape(slide, text_blocks)

        for block in other_blocks:
            if block.type == "image":
                self._add_image(slide, block.content)
            elif block.type == "table":
                self._add_table(slide, block.content)

    def _add_shape_from_data(self, slide, shape_data: dict) -> None:
        shape_type = shape_data.get("type")
        try:
            if shape_type == "text":
                self._add_text_shape(slide, shape_data)
            elif shape_type == "image":
                self._add_image_shape(slide, shape_data)
            elif shape_type == "table":
                self._add_table_shape(slide, shape_data)
            elif shape_type == "chart":
                self._add_chart_shape(slide, shape_data)
            elif shape_type == "ole":
                self._add_ole_shape(slide, shape_data)
        except Exception:
            pass

    def _add_text_shape(self, slide, shape_data: dict) -> None:
        left = shape_data.get("left", Emu(914400))
        top = shape_data.get("top", Emu(914400))
        width = shape_data.get("width", Emu(914400 * 8))
        height = shape_data.get("height", Emu(914400 * 2))

        textbox = slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        tf.word_wrap = True
        tf.clear()

        paragraphs = shape_data.get("paragraphs", [])
        for i, para_data in enumerate(paragraphs):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            runs = para_data.get("runs", [])
            if runs:
                for run_data in runs:
                    run = p.add_run()
                    run.text = run_data.get("text", "")
                    self._apply_run_format(run, run_data)
            else:
                p.text = para_data.get("text", "")

            p.level = para_data.get("level", 0)
            if para_data.get("alignment") is not None:
                p.alignment = para_data["alignment"]

    def _apply_run_format(self, run, run_data: dict) -> None:
        font = run.font
        if run_data.get("font_name"):
            font.name = run_data["font_name"]
        if run_data.get("font_size") is not None:
            font.size = run_data["font_size"]
        if run_data.get("bold") is not None:
            font.bold = run_data["bold"]
        if run_data.get("italic") is not None:
            font.italic = run_data["italic"]
        if run_data.get("underline") is not None:
            font.underline = run_data["underline"]
        if run_data.get("color"):
            try:
                font.color.rgb = RGBColor.from_string(run_data["color"])
            except Exception:
                pass

    def _add_image_shape(self, slide, shape_data: dict) -> None:
        blob = shape_data.get("blob")
        if not blob:
            return
        stream = BytesIO(blob)
        slide.shapes.add_picture(
            stream,
            left=shape_data.get("left", Emu(914400)),
            top=shape_data.get("top", Emu(914400)),
            width=shape_data.get("width"),
            height=shape_data.get("height"),
        )

    def _add_table_shape(self, slide, shape_data: dict) -> None:
        data = shape_data.get("data", [])
        if not data:
            return
        rows = len(data)
        cols = len(data[0]) if data[0] else 1
        table_shape = slide.shapes.add_table(
            rows, cols,
            shape_data.get("left", Emu(914400)),
            shape_data.get("top", Emu(914400 * 3)),
            shape_data.get("width", Emu(914400 * 8)),
            shape_data.get("height", Emu(914400 * 2)),
        )
        for r, row_data in enumerate(data):
            for c, cell_text in enumerate(row_data):
                if c < cols:
                    table_shape.table.cell(r, c).text = cell_text

    def _add_chart_shape(self, slide, shape_data: dict) -> None:
        try:
            chart_type = shape_data.get("chart_type", "")
            chart_data = shape_data.get("data", [])
            categories = shape_data.get("categories", [])
            
            from pptx.chart.data import CategoryChartData
            from pptx.enum.chart import XL_CHART_TYPE
            
            if not chart_data or not categories:
                return

            chart_data_obj = CategoryChartData()
            chart_data_obj.categories = categories
            
            for series in chart_data:
                series_name = series.get("name", "Series")
                values = series.get("values", [])
                if values:
                    chart_data_obj.add_series(series_name, values)

            left = shape_data.get("left", Emu(914400))
            top = shape_data.get("top", Emu(914400))
            width = shape_data.get("width", Emu(914400 * 8))
            height = shape_data.get("height", Emu(914400 * 4))

            slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED,
                left, top, width, height,
                chart_data_obj
            )
        except Exception:
            pass

    def _add_ole_shape(self, slide, shape_data: dict) -> None:
        try:
            blob = shape_data.get("blob")
            if not blob:
                return
            stream = BytesIO(blob)
            left = shape_data.get("left", Emu(914400))
            top = shape_data.get("top", Emu(914400))
            width = shape_data.get("width", Emu(914400 * 6))
            height = shape_data.get("height", Emu(914400 * 4))
            slide.shapes.add_ole_object(stream, left, top, width, height)
        except Exception:
            pass

    def _fill_title(self, slide, title_text: str) -> None:
        try:
            if slide.shapes.title is not None:
                slide.shapes.title.text = title_text
                return
        except Exception:
            pass

        for shape in slide.placeholders:
            try:
                phf = shape.placeholder_format
                if phf.type == 0:
                    shape.text_frame.text = title_text
                    return
            except Exception:
                continue

    def _find_body_placeholder(self, slide) -> object | None:
        for shape in slide.placeholders:
            try:
                phf = shape.placeholder_format
                ph_type = phf.type

                if ph_type == 2:
                    return shape
                if ph_type in (7, 8, 9, 10):
                    return shape
                if ph_type not in (0, 1, 3, 4):
                    return shape
            except Exception:
                continue

        for shape in slide.placeholders:
            try:
                phf = shape.placeholder_format
                ph_type = phf.type
                if ph_type != 0 and ph_type != 1 and shape.has_text_frame:
                    return shape
            except Exception:
                continue

        return None

    def _fill_text_into_placeholder(self, placeholder, text_blocks) -> None:
        try:
            tf = placeholder.text_frame
            tf.clear()
            tf.word_wrap = True

            first = True
            for block in text_blocks:
                if first:
                    p = tf.paragraphs[0]
                    first = False
                else:
                    p = tf.add_paragraph()
                p.text = block.text
                p.level = block.level
        except Exception:
            pass

    def _add_text_as_new_shape(self, slide, text_blocks) -> None:
        try:
            left = Emu(914400)
            top = Emu(914400 * 3)
            width = Emu(914400 * 10)
            height = Emu(914400 * 6)

            textbox = slide.shapes.add_textbox(left, top, width, height)
            textbox.text_frame.word_wrap = True
            tf = textbox.text_frame

            for i, block in enumerate(text_blocks):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = block.text
                p.level = block.level
        except Exception:
            pass

    def _add_image(self, slide, content: dict) -> None:
        try:
            blob = content.get("blob")
            if not blob:
                return
            stream = BytesIO(blob)
            pic = slide.shapes.add_picture(
                stream,
                left=content.get("left") or Emu(914400),
                top=content.get("top") or Emu(914400),
                width=content.get("width"),
                height=content.get("height"),
            )
        except Exception:
            pass

    def _add_table(self, slide, data: list[list[str]]) -> None:
        try:
            if not data:
                return
            rows = len(data)
            cols = len(data[0]) if data[0] else 1
            table_shape = slide.shapes.add_table(rows, cols, Emu(914400), Emu(914400 * 3), Emu(914400 * 8), Emu(914400 * 2))
            for r, row_data in enumerate(data):
                for c, cell_text in enumerate(row_data):
                    if c < cols:
                        table_shape.table.cell(r, c).text = cell_text
        except Exception:
            pass

    def _check_and_fix_overflow(self, slide) -> None:
        try:
            slide_width = slide.part.ppt.slide_width
            slide_height = slide.part.ppt.slide_height
        except Exception:
            slide_width = Emu(914400 * 13.333)
            slide_height = Emu(914400 * 7.5)

        for shape in slide.shapes:
            try:
                if shape.left < 0:
                    shape.left = Emu(91440)
                if shape.top < 0:
                    shape.top = Emu(91440)
                if shape.left + shape.width > slide_width:
                    shape.width = slide_width - shape.left - Emu(91440)
                if shape.top + shape.height > slide_height:
                    shape.height = slide_height - shape.top - Emu(91440)
            except Exception:
                pass

    def _determine_layout(self, model: SlideContentModel) -> str:
        if model.original_layout_type:
            return model.original_layout_type
        if model.slide_index == 0:
            return "cover"
        if model.slide_index == len(model.body_blocks) - 1:
            return "closing"
        return "content"
