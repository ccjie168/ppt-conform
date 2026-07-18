from pathlib import Path
from io import BytesIO
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from core.models import SlideContentModel, UserConfig
from core.registry.template_registry import TemplateRegistry
from core.analyzer.template_format_extractor import TemplateFormatExtractor


class ContentReplayer:
    """内容重放器：将抽取的内容重放到目标模板，标题/正文字体遵循模板要求，保留源内容格式"""

    def __init__(self, registry: TemplateRegistry, template_path: str | None = None):
        self.registry = registry
        self.template_path = template_path
        self.template_formats: dict = {}
        self.theme_fonts: dict = {}

        if template_path and Path(template_path).exists():
            extractor = TemplateFormatExtractor()
            try:
                self.template_formats = extractor.extract_placeholder_formats(template_path)
                self.theme_fonts = extractor.extract_theme_fonts(template_path)
            except Exception:
                pass

    # 标准16:9尺寸 (13.333 x 7.5 英寸)
    WIDESCREEN_16_9 = Emu(12192000)
    WIDESCREEN_16_9_H = Emu(6858000)

    def replay(self, content_models: list[SlideContentModel], config: UserConfig) -> str:
        if self.template_path and Path(self.template_path).exists():
            output_prs = Presentation(self.template_path)
            self._clear_slides(output_prs)
        else:
            output_prs = Presentation()

        # 强制输出为16:9宽屏
        output_prs.slide_width = self.WIDESCREEN_16_9
        output_prs.slide_height = self.WIDESCREEN_16_9_H

        # 计算源和目标的尺寸比例，用于尺寸适配
        source_prs = Presentation(config.input_path)
        self.source_width = source_prs.slide_width
        self.source_height = source_prs.slide_height
        self.target_width = output_prs.slide_width
        self.target_height = output_prs.slide_height
        self.scale_x = self.target_width / self.source_width if self.source_width else 1.0
        self.scale_y = self.target_height / self.source_height if self.source_height else 1.0

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
            self._copy_slide_content(slide, model, slide_idx)
            self._check_and_fix_overflow(slide)

        output_prs.save(config.output_path)
        return config.output_path

    def _adapt_position(self, value, axis: str = "x") -> int:
        """根据源/目标尺寸比例适配位置或大小"""
        if value is None:
            return None
        scale = self.scale_x if axis == "x" else self.scale_y
        if abs(scale - 1.0) < 0.01:
            return value
        return int(value * scale)

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

    def _resolve_layout_index(self, layout_name: str, master, master_style: str) -> int:
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

    def _copy_slide_content(self, slide, model: SlideContentModel, slide_idx: int) -> None:
        if not model.raw_shapes:
            self._fill_simple_content(slide, model)
            return

        is_title_shape = True
        for shape_data in model.raw_shapes:
            # 标题形状使用模板格式
            if is_title_shape and shape_data.get("type") == "text" and model.title:
                self._add_text_shape(slide, shape_data, is_title=True)
                is_title_shape = False
            else:
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
                self._add_table_from_content(slide, block.content)

    def _add_shape_from_data(self, slide, shape_data: dict) -> None:
        shape_type = shape_data.get("type")
        try:
            if shape_type == "text":
                self._add_text_shape(slide, shape_data, is_title=False)
            elif shape_type == "image":
                self._add_image_shape(slide, shape_data)
            elif shape_type == "table":
                self._add_table_shape(slide, shape_data)
            elif shape_type == "chart":
                self._add_chart_shape(slide, shape_data)
            elif shape_type == "ole":
                self._add_ole_shape(slide, shape_data)
            elif shape_type == "autoshape":
                self._add_auto_shape(slide, shape_data)
            elif shape_type == "group":
                for sub_shape in shape_data.get("shapes", []):
                    self._add_shape_from_data(slide, sub_shape)
        except Exception:
            pass

    def _add_text_shape(self, slide, shape_data: dict, is_title: bool = False) -> None:
        left = self._adapt_position(shape_data.get("left", Emu(914400)), "x")
        top = self._adapt_position(shape_data.get("top", Emu(914400)), "y")
        width = self._adapt_position(shape_data.get("width", Emu(914400 * 8)), "x")
        height = self._adapt_position(shape_data.get("height", Emu(914400 * 2)), "y")

        textbox = slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        tf.word_wrap = True
        tf.clear()

        # 应用文本框填充色和边框色
        fill_color = shape_data.get("fill_color")
        if fill_color:
            try:
                textbox.fill.solid()
                textbox.fill.fore_color.rgb = RGBColor.from_string(fill_color)
            except Exception:
                pass

        line_color = shape_data.get("line_color")
        line_width = shape_data.get("line_width")
        if line_color:
            try:
                textbox.line.color.rgb = RGBColor.from_string(line_color)
                if line_width:
                    textbox.line.width = line_width
            except Exception:
                pass

        # 获取模板格式规范
        template_fmt = self._get_template_format(is_title)

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
                    self._apply_run_format(run, run_data, template_fmt, is_title)
            else:
                p.text = para_data.get("text", "")
                # 应用模板格式到整段
                if is_title and template_fmt:
                    self._apply_template_format_to_paragraph(p, template_fmt)

            p.level = para_data.get("level", 0)
            if para_data.get("alignment") is not None:
                p.alignment = self._parse_alignment(para_data["alignment"])
            elif template_fmt and template_fmt.get("alignment"):
                p.alignment = self._parse_alignment(template_fmt["alignment"])

            try:
                if para_data.get("line_spacing") is not None:
                    p.line_spacing = para_data["line_spacing"]
                if para_data.get("space_before") is not None:
                    p.space_before = para_data["space_before"]
                if para_data.get("space_after") is not None:
                    p.space_after = para_data["space_after"]
            except Exception:
                pass

    def _get_template_format(self, is_title: bool) -> dict:
        """获取模板中标题或正文的格式规范"""
        if is_title:
            return self.template_formats.get("title", {})
        return self.template_formats.get("body", {})

    def _apply_run_format(self, run, run_data: dict, template_fmt: dict | None = None, is_title: bool = False) -> None:
        """应用run格式，对于标题/正文使用模板要求的字体和大小"""
        font = run.font

        # 字体名称：优先使用模板要求
        font_name = run_data.get("font_name")
        if template_fmt and template_fmt.get("font_name"):
            font_name = template_fmt["font_name"]
        elif not font_name and self.theme_fonts:
            font_name = self.theme_fonts.get("major" if is_title else "minor")
        if font_name:
            font.name = font_name

        # 字号：标题/正文使用模板要求的大小
        font_size = run_data.get("font_size")
        if template_fmt and template_fmt.get("font_size"):
            font_size = template_fmt["font_size"]
        if font_size is not None:
            font.size = font_size

        # 粗体：标题强制使用模板要求，正文保留原格式
        if is_title and template_fmt and template_fmt.get("bold") is not None:
            font.bold = template_fmt["bold"]
        elif run_data.get("bold") is not None:
            font.bold = run_data["bold"]

        # 斜体
        if run_data.get("italic") is not None:
            font.italic = run_data["italic"]

        # 下划线
        if run_data.get("underline") is not None:
            font.underline = run_data["underline"]

        # 颜色：标题使用模板要求的颜色，正文保留原颜色
        color = run_data.get("color")
        if is_title and template_fmt and template_fmt.get("color"):
            color = template_fmt["color"]
        if color:
            try:
                font.color.rgb = RGBColor.from_string(color)
            except Exception:
                pass

    def _apply_template_format_to_paragraph(self, paragraph, template_fmt: dict) -> None:
        """将模板格式应用到整个段落"""
        if template_fmt.get("alignment"):
            paragraph.alignment = self._parse_alignment(template_fmt["alignment"])
        for run in paragraph.runs:
            font = run.font
            if template_fmt.get("font_name"):
                font.name = template_fmt["font_name"]
            if template_fmt.get("font_size"):
                font.size = template_fmt["font_size"]
            if template_fmt.get("bold") is not None:
                font.bold = template_fmt["bold"]
            if template_fmt.get("color"):
                try:
                    font.color.rgb = RGBColor.from_string(template_fmt["color"])
                except Exception:
                    pass

    def _parse_alignment(self, align_str: str):
        """将字符串对齐方式转为PP_ALIGN枚举"""
        if not align_str:
            return None
        align_map = {
            "LEFT": PP_ALIGN.LEFT,
            "CENTER": PP_ALIGN.CENTER,
            "RIGHT": PP_ALIGN.RIGHT,
            "JUSTIFY": PP_ALIGN.JUSTIFY,
            "DISTRIBUTE": PP_ALIGN.DISTRIBUTE,
        }
        for key, val in align_map.items():
            if key in align_str.upper():
                return val
        return None

    def _add_image_shape(self, slide, shape_data: dict) -> None:
        blob = shape_data.get("blob")
        if not blob:
            return
        stream = BytesIO(blob)
        pic = slide.shapes.add_picture(
            stream,
            left=self._adapt_position(shape_data.get("left", Emu(914400)), "x"),
            top=self._adapt_position(shape_data.get("top", Emu(914400)), "y"),
            width=self._adapt_position(shape_data.get("width"), "x"),
            height=self._adapt_position(shape_data.get("height"), "y"),
        )
        # 应用裁剪
        try:
            if shape_data.get("crop_left"):
                pic.crop_left = shape_data["crop_left"]
            if shape_data.get("crop_top"):
                pic.crop_top = shape_data["crop_top"]
            if shape_data.get("crop_right"):
                pic.crop_right = shape_data["crop_right"]
            if shape_data.get("crop_bottom"):
                pic.crop_bottom = shape_data["crop_bottom"]
        except Exception:
            pass

    def _add_table_shape(self, slide, shape_data: dict) -> None:
        table_data = shape_data.get("data", {})
        if not table_data:
            return

        # 兼容新旧格式
        if isinstance(table_data, dict):
            self._add_table_full(slide, shape_data)
        else:
            # 旧格式：二维数组
            self._add_table_simple(slide, shape_data, table_data)

    def _add_table_full(self, slide, shape_data: dict) -> None:
        """完整表格重放：保留单元格背景色、字体格式、合并单元格"""
        table_data = shape_data.get("data", {})
        rows = table_data.get("rows", 0)
        cols = table_data.get("cols", 0)
        cells_data = table_data.get("cells", [])

        if rows == 0 or cols == 0:
            return

        table_shape = slide.shapes.add_table(
            rows, cols,
            shape_data.get("left", Emu(914400)),
            shape_data.get("top", Emu(914400 * 3)),
            shape_data.get("width", Emu(914400 * 8)),
            shape_data.get("height", Emu(914400 * 2)),
        )
        table = table_shape.table

        # 设置列宽
        col_widths = table_data.get("col_widths", [])
        for i, width in enumerate(col_widths):
            if i < len(table.columns) and width:
                table.columns[i].width = width

        # 设置行高
        row_heights = table_data.get("row_heights", [])
        for i, height in enumerate(row_heights):
            if i < len(table.rows) and height:
                table.rows[i].height = height

        # 填充单元格内容和格式
        for cell_data in cells_data:
            r = cell_data["row"]
            c = cell_data["col"]
            if r >= rows or c >= cols:
                continue

            cell = table.cell(r, c)
            cell.text = cell_data.get("text", "")

            # 背景色
            if cell_data.get("fill_color"):
                try:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor.from_string(cell_data["fill_color"])
                except Exception:
                    pass

            # 对齐和锚点
            try:
                if cell_data.get("alignment"):
                    align = self._parse_alignment(cell_data["alignment"])
                    if align and cell.text_frame.paragraphs:
                        cell.text_frame.paragraphs[0].alignment = align
            except Exception:
                pass

            try:
                if cell_data.get("vertical_anchor"):
                    anchor_map = {
                        "TOP": MSO_ANCHOR.TOP,
                        "MIDDLE": MSO_ANCHOR.MIDDLE,
                        "BOTTOM": MSO_ANCHOR.BOTTOM,
                    }
                    for key, val in anchor_map.items():
                        if key in cell_data["vertical_anchor"].upper():
                            cell.vertical_anchor = val
                            break
            except Exception:
                pass

            # 边距
            try:
                if cell_data.get("margin_left") is not None:
                    cell.margin_left = cell_data["margin_left"]
                if cell_data.get("margin_right") is not None:
                    cell.margin_right = cell_data["margin_right"]
                if cell_data.get("margin_top") is not None:
                    cell.margin_top = cell_data["margin_top"]
                if cell_data.get("margin_bottom") is not None:
                    cell.margin_bottom = cell_data["margin_bottom"]
            except Exception:
                pass

            # 字体格式
            try:
                if cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
                    run = cell.text_frame.paragraphs[0].runs[0]
                    font = run.font
                    if cell_data.get("font_name"):
                        font.name = cell_data["font_name"]
                    if cell_data.get("font_size"):
                        font.size = cell_data["font_size"]
                    if cell_data.get("bold") is not None:
                        font.bold = cell_data["bold"]
                    if cell_data.get("italic") is not None:
                        font.italic = cell_data["italic"]
                    if cell_data.get("color"):
                        try:
                            font.color.rgb = RGBColor.from_string(cell_data["color"])
                        except Exception:
                            pass
            except Exception:
                pass

        # 合并单元格
        merged_cells = table_data.get("merged_cells", [])
        for merge_info in merged_cells:
            try:
                start_r = merge_info["start_row"]
                start_c = merge_info["start_col"]
                row_span = merge_info.get("row_span", 1)
                col_span = merge_info.get("col_span", 1)

                if row_span > 1 or col_span > 1:
                    end_r = start_r + row_span - 1
                    end_c = start_c + col_span - 1
                    cell_a = table.cell(start_r, start_c)
                    cell_b = table.cell(end_r, end_c)
                    cell_a.merge(cell_b)
            except Exception:
                pass

    def _add_table_simple(self, slide, shape_data: dict, data: list) -> None:
        """简化版表格重放（向后兼容）"""
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
                    table_shape.table.cell(r, c).text = str(cell_text)

    def _add_chart_shape(self, slide, shape_data: dict) -> None:
        """完整图表重放：保留图例、坐标轴、系列颜色"""
        try:
            from pptx.chart.data import CategoryChartData
            from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

            chart_data_list = shape_data.get("data", [])
            categories = shape_data.get("categories", [])

            if not chart_data_list or not categories:
                return

            chart_data_obj = CategoryChartData()
            chart_data_obj.categories = categories

            for series in chart_data_list:
                series_name = series.get("name", "Series")
                values = series.get("values", [])
                if values:
                    chart_data_obj.add_series(series_name, values)

            left = shape_data.get("left", Emu(914400))
            top = shape_data.get("top", Emu(914400))
            width = shape_data.get("width", Emu(914400 * 8))
            height = shape_data.get("height", Emu(914400 * 4))

            # 解析图表类型
            chart_type_str = shape_data.get("chart_type", "")
            chart_type = self._parse_chart_type(chart_type_str)

            chart_frame = slide.shapes.add_chart(
                chart_type,
                left, top, width, height,
                chart_data_obj
            )
            chart = chart_frame.chart

            # 设置图表标题
            if shape_data.get("chart_title"):
                chart.has_title = True
                chart.chart_title.text_frame.text = shape_data["chart_title"]
            else:
                chart.has_title = False

            # 设置图例
            if shape_data.get("has_legend"):
                chart.has_legend = True
                pos_str = shape_data.get("legend_position", "")
                if "BOTTOM" in pos_str:
                    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                elif "TOP" in pos_str:
                    chart.legend.position = XL_LEGEND_POSITION.TOP
                elif "LEFT" in pos_str:
                    chart.legend.position = XL_LEGEND_POSITION.LEFT
                elif "RIGHT" in pos_str:
                    chart.legend.position = XL_LEGEND_POSITION.RIGHT
            else:
                chart.has_legend = False

            # 设置系列颜色
            for i, series in enumerate(chart.series):
                if i < len(chart_data_list):
                    color = chart_data_list[i].get("color")
                    if color:
                        try:
                            series.format.fill.solid()
                            series.format.fill.fore_color.rgb = RGBColor.from_string(color)
                        except Exception:
                            pass

            # 设置坐标轴
            x_axis_data = shape_data.get("x_axis", {})
            y_axis_data = shape_data.get("y_axis", {})
            try:
                if chart.category_axis and x_axis_data:
                    ca = chart.category_axis
                    if x_axis_data.get("has_title") and x_axis_data.get("title"):
                        ca.has_title = True
                        ca.axis_title.text_frame.text = x_axis_data["title"]
                    if x_axis_data.get("visible") is not None:
                        ca.visible = x_axis_data["visible"]
            except Exception:
                pass
            try:
                if chart.value_axis and y_axis_data:
                    va = chart.value_axis
                    if y_axis_data.get("has_title") and y_axis_data.get("title"):
                        va.has_title = True
                        va.axis_title.text_frame.text = y_axis_data["title"]
                    if y_axis_data.get("visible") is not None:
                        va.visible = y_axis_data["visible"]
                    if y_axis_data.get("minimum_scale") is not None:
                        va.minimum_scale = y_axis_data["minimum_scale"]
                    if y_axis_data.get("maximum_scale") is not None:
                        va.maximum_scale = y_axis_data["maximum_scale"]
            except Exception:
                pass

        except Exception:
            pass

    def _parse_chart_type(self, chart_type_str: str):
        """将图表类型字符串转为XL_CHART_TYPE枚举"""
        from pptx.enum.chart import XL_CHART_TYPE
        type_map = {
            "COLUMN_CLUSTERED": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "COLUMN_STACKED": XL_CHART_TYPE.COLUMN_STACKED,
            "BAR_CLUSTERED": XL_CHART_TYPE.BAR_CLUSTERED,
            "LINE": XL_CHART_TYPE.LINE,
            "LINE_MARKERS": XL_CHART_TYPE.LINE_MARKERS,
            "PIE": XL_CHART_TYPE.PIE,
            "DOUGHNUT": XL_CHART_TYPE.DOUGHNUT,
            "AREA": XL_CHART_TYPE.AREA,
            "XY_SCATTER": XL_CHART_TYPE.XY_SCATTER,
            "BUBBLE": XL_CHART_TYPE.BUBBLE,
            "RADAR": XL_CHART_TYPE.RADAR,
        }
        for key, val in type_map.items():
            if key in chart_type_str.upper():
                return val
        return XL_CHART_TYPE.COLUMN_CLUSTERED

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

    def _add_auto_shape(self, slide, shape_data: dict) -> None:
        """重放自选图形：保留填充色、边框色等样式"""
        try:
            from pptx.enum.shapes import MSO_SHAPE
            shape_type_str = shape_data.get("shape_type", "")
            mso_shape = MSO_SHAPE.RECTANGLE
            try:
                if "OVAL" in shape_type_str or "ELLIPSE" in shape_type_str:
                    mso_shape = MSO_SHAPE.OVAL
                elif "ROUNDED_RECTANGLE" in shape_type_str:
                    mso_shape = MSO_SHAPE.ROUNDED_RECTANGLE
                elif "LINE" in shape_type_str:
                    mso_shape = MSO_SHAPE.LINE
                elif "ARROW" in shape_type_str:
                    mso_shape = MSO_SHAPE.RIGHT_ARROW
                elif "DIAMOND" in shape_type_str:
                    mso_shape = MSO_SHAPE.DIAMOND
                elif "TRIANGLE" in shape_type_str:
                    mso_shape = MSO_SHAPE.RIGHT_TRIANGLE
                elif "STAR" in shape_type_str:
                    mso_shape = MSO_SHAPE.FIVE_POINTED_STAR
            except Exception:
                pass

            left = self._adapt_position(shape_data.get("left", Emu(914400)), "x")
            top = self._adapt_position(shape_data.get("top", Emu(914400)), "y")
            width = self._adapt_position(shape_data.get("width", Emu(914400 * 4)), "x")
            height = self._adapt_position(shape_data.get("height", Emu(914400 * 2)), "y")

            shape = slide.shapes.add_shape(
                mso_shape, left, top, width, height
            )

            try:
                shape.rotation = shape_data.get("rotation", 0)
            except Exception:
                pass

            fill_color = shape_data.get("fill_color")
            if fill_color:
                try:
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = RGBColor.from_string(fill_color)
                except Exception:
                    pass

            line_color = shape_data.get("line_color")
            line_width = shape_data.get("line_width")
            if line_color:
                try:
                    shape.line.color.rgb = RGBColor.from_string(line_color)
                    if line_width:
                        shape.line.width = line_width
                except Exception:
                    pass

            if shape_data.get("text") and shape.has_text_frame:
                shape.text_frame.text = shape_data["text"]
        except Exception:
            pass

    def _fill_title(self, slide, title_text: str) -> None:
        try:
            if slide.shapes.title is not None:
                slide.shapes.title.text = title_text
                # 应用模板标题格式
                template_fmt = self.template_formats.get("title", {})
                if template_fmt and slide.shapes.title.text_frame.paragraphs:
                    p = slide.shapes.title.text_frame.paragraphs[0]
                    self._apply_template_format_to_paragraph(p, template_fmt)
                return
        except Exception:
            pass

        for shape in slide.placeholders:
            try:
                phf = shape.placeholder_format
                if phf.type == 0:
                    shape.text_frame.text = title_text
                    template_fmt = self.template_formats.get("title", {})
                    if template_fmt and shape.text_frame.paragraphs:
                        p = shape.text_frame.paragraphs[0]
                        self._apply_template_format_to_paragraph(p, template_fmt)
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
        return None

    def _fill_text_into_placeholder(self, placeholder, text_blocks) -> None:
        try:
            tf = placeholder.text_frame
            tf.clear()
            tf.word_wrap = True

            # 获取模板正文格式
            template_fmt = self.template_formats.get("body", {})

            first = True
            for block in text_blocks:
                if first:
                    p = tf.paragraphs[0]
                    first = False
                else:
                    p = tf.add_paragraph()
                p.text = block.text
                p.level = block.level

                # 应用模板正文格式
                if template_fmt:
                    self._apply_template_format_to_paragraph(p, template_fmt)
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

            template_fmt = self.template_formats.get("body", {})

            for i, block in enumerate(text_blocks):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = block.text
                p.level = block.level

                if template_fmt:
                    self._apply_template_format_to_paragraph(p, template_fmt)
        except Exception:
            pass

    def _add_image(self, slide, content: dict) -> None:
        try:
            blob = content.get("blob")
            if not blob:
                return
            stream = BytesIO(blob)
            slide.shapes.add_picture(
                stream,
                left=content.get("left") or Emu(914400),
                top=content.get("top") or Emu(914400),
                width=content.get("width"),
                height=content.get("height"),
            )
        except Exception:
            pass

    def _add_table_from_content(self, slide, content) -> None:
        """从ContentBlock.content添加表格"""
        if isinstance(content, dict) and "cells" in content:
            shape_data = {"data": content, "left": Emu(914400), "top": Emu(914400 * 3)}
            self._add_table_full(slide, shape_data)
        elif isinstance(content, list):
            shape_data = {"data": content, "left": Emu(914400), "top": Emu(914400 * 3)}
            self._add_table_simple(slide, shape_data, content)

    def _add_table(self, slide, data: list[list[str]]) -> None:
        shape_data = {"data": data}
        self._add_table_simple(slide, shape_data, data)

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
        return "content"
