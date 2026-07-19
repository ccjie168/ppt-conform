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
        self.default_text_color: str | None = None
        self.title_text_color: str | None = None
        self.template_width: int = 0
        self.template_height: int = 0
        self.footer_shapes: list[dict] = []
        self.background_image: dict | None = None

        if template_path and Path(template_path).exists():
            extractor = TemplateFormatExtractor()
            try:
                self.template_formats = extractor.extract_placeholder_formats(template_path)
                self.theme_fonts = extractor.extract_theme_fonts(template_path)
                self._analyze_template_footer(template_path)
            except Exception:
                pass

    def _analyze_template_footer(self, template_path: str, master_index: int = 0) -> None:
        """分析指定Master中的footer元素（图标、页脚文本等），用于复制到新slide"""
        self.footer_shapes = []
        self.background_image = None
        try:
            from pptx import Presentation
            from pptx.enum.shapes import MSO_SHAPE_TYPE

            prs = Presentation(template_path)
            self.template_width = prs.slide_width
            self.template_height = prs.slide_height

            if not prs.slide_masters or master_index >= len(prs.slide_masters):
                return

            master = prs.slide_masters[master_index]
            footer_threshold = self.template_height * 0.85

            for shape in master.shapes:
                try:
                    is_footer = False
                    shape_type = shape.shape_type

                    # 检查是否是背景图（铺满整个页面的图片）
                    if shape_type == MSO_SHAPE_TYPE.PICTURE:
                        try:
                            left = shape.left or 0
                            top = shape.top or 0
                            width = shape.width or 0
                            height = shape.height or 0
                            is_full_bg = (abs(left) < Emu(20000) and abs(top) < Emu(20000) and
                                          abs(width - self.template_width) < Emu(50000) and
                                          abs(height - self.template_height) < Emu(50000))
                            if is_full_bg:
                                # 保存背景图
                                self.background_image = {
                                    "image_blob": shape.image.blob,
                                    "image_ext": shape.image.ext,
                                    "left": left,
                                    "top": top,
                                    "width": width,
                                    "height": height,
                                }
                                continue
                        except Exception:
                            pass

                    # 底部区域的图片（施耐德图标等）
                    if shape_type == MSO_SHAPE_TYPE.PICTURE:
                        if shape.top and shape.top > footer_threshold:
                            is_footer = True

                    # 底部区域的文本框（页脚文本等）
                    if shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                        if shape.top and shape.top > footer_threshold:
                            is_footer = True

                    # 底部区域的占位符（页脚、页码等）
                    if shape.is_placeholder:
                        try:
                            ph_type = int(shape.placeholder_format.type)
                            # 13=SLIDE_NUMBER, 14=HEADER, 15=FOOTER, 16=DATE
                            if ph_type in (13, 14, 15, 16) and shape.top and shape.top > footer_threshold:
                                is_footer = True
                        except Exception:
                            pass

                    if is_footer:
                        right_margin = self.template_width - shape.left - shape.width
                        bottom_margin = self.template_height - shape.top - shape.height

                        shape_data = {
                            "type": shape_type,
                            "left": shape.left,
                            "top": shape.top,
                            "width": shape.width,
                            "height": shape.height,
                            "right_margin": right_margin,
                            "bottom_margin": bottom_margin,
                        }

                        # 图片：保存图片二进制数据
                        if shape_type == MSO_SHAPE_TYPE.PICTURE:
                            try:
                                shape_data["image_blob"] = shape.image.blob
                                shape_data["image_ext"] = shape.image.ext
                            except Exception:
                                continue

                        # 文本框：保存文本内容
                        if shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                            try:
                                shape_data["text"] = shape.text_frame.text
                                # 保存字体样式
                                if shape.text_frame.paragraphs and shape.text_frame.paragraphs[0].runs:
                                    run = shape.text_frame.paragraphs[0].runs[0]
                                    shape_data["font_name"] = run.font.name
                                    shape_data["font_size"] = run.font.size
                                    shape_data["font_bold"] = run.font.bold
                                    try:
                                        if run.font.color and run.font.color.rgb:
                                            shape_data["font_color"] = run.font.color.rgb
                                    except (AttributeError, TypeError):
                                        pass
                            except Exception:
                                shape_data["text"] = ""

                        # 占位符（页脚、页码等）：保存文本内容和类型
                        if shape.is_placeholder:
                            try:
                                shape_data["text"] = shape.text_frame.text
                                shape_data["ph_type"] = int(shape.placeholder_format.type)
                                # 保存字体样式
                                if shape.text_frame.paragraphs and shape.text_frame.paragraphs[0].runs:
                                    run = shape.text_frame.paragraphs[0].runs[0]
                                    shape_data["font_name"] = run.font.name
                                    shape_data["font_size"] = run.font.size
                                    shape_data["font_bold"] = run.font.bold
                                    try:
                                        if run.font.color and run.font.color.rgb:
                                            shape_data["font_color"] = run.font.color.rgb
                                    except (AttributeError, TypeError):
                                        pass
                            except Exception:
                                shape_data["text"] = ""

                        self.footer_shapes.append(shape_data)
                except Exception:
                    continue
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

        # 根据所选Master设置默认文字颜色（深色背景用白色文字，浅色背景用深色文字）
        # 同时重新分析所选Master的footer元素
        if self.template_path and Path(self.template_path).exists():
            try:
                extractor = TemplateFormatExtractor()
                self.default_text_color = extractor.get_text_color_for_master(
                    self.template_path, selected_master_index
                )
                
                # 标题颜色：浅色背景(Master 0/1)用dark green background 2，深色背景(Master 2/3)用白色
                colors = extractor.extract_theme_colors(self.template_path)
                if selected_master_index in (0, 1):
                    # 白色和浅绿模板：标题用 dark green background 2
                    self.title_text_color = colors.get("dk2", "3DCD58")
                else:
                    # 渐变和深绿模板：标题用白色
                    self.title_text_color = colors.get("lt1", "FFFFFF")
                
                # 更新template_formats中的颜色
                if self.default_text_color:
                    if "body" not in self.template_formats:
                        self.template_formats["body"] = {}
                    if not self.template_formats["body"].get("color"):
                        self.template_formats["body"]["color"] = self.default_text_color
                if self.title_text_color:
                    if "title" not in self.template_formats:
                        self.template_formats["title"] = {}
                    self.template_formats["title"]["color"] = self.title_text_color
                
                # 更新主题字体（使用对应Master的主题）
                self.theme_fonts = extractor.extract_theme_fonts(
                    self.template_path, selected_master_index
                )
                # 重新分析所选Master的footer元素
                self._analyze_template_footer(self.template_path, selected_master_index)
            except Exception:
                pass

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
            self._add_background_image(slide)
            self._add_footer_shapes(slide)
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
        """清空幻灯片上的内容占位符的文字内容，保留占位符的样式（位置、字体、颜色）
        
        - 页眉/页脚/页码占位符保留不动
        - 标题/正文/副标题占位符只清空文字，保留样式
        - 非占位符形状都删掉（背景图和footer后面手动加）
        """
        header_footer_types = (13, 14, 15, 16)  # slide_number, header, footer, date

        shapes_to_remove = []
        for shape in slide.shapes:
            try:
                if shape.is_placeholder:
                    phf = shape.placeholder_format
                    if phf.type in header_footer_types:
                        # 页眉页脚页码占位符保留不动
                        continue
                    # 标题/正文/副标题占位符：只清空文字内容，保留样式
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                run.text = ""
                else:
                    # 非占位符形状都删掉（背景图和footer后面手动加）
                    shapes_to_remove.append(shape)
            except Exception:
                shapes_to_remove.append(shape)

        spTree = slide.shapes._spTree
        for shape in shapes_to_remove:
            try:
                spTree.remove(shape._element)
            except Exception:
                pass

    def _add_background_image(self, slide) -> None:
        """将背景图添加到slide（最底层），并适配16:9尺寸"""
        if not self.background_image:
            return

        try:
            from io import BytesIO

            target_width = self.target_width
            target_height = self.target_height

            image_blob = self.background_image["image_blob"]
            stream = BytesIO(image_blob)

            # 添加背景图（铺满整个页面）
            pic = slide.shapes.add_picture(
                stream,
                left=0,
                top=0,
                width=target_width,
                height=target_height,
            )

            # 将背景图移到最底层
            spTree = slide.shapes._spTree
            spTree.remove(pic._element)
            spTree.insert(2, pic._element)  # 索引0是nvGrpSpPr，1是grpSpPr，2开始是形状
        except Exception:
            pass

    def _add_footer_shapes(self, slide) -> None:
        """将footer元素（施耐德图标、页脚文本等）添加到slide上，并适配16:9位置"""
        if not self.footer_shapes:
            return

        try:
            target_width = self.target_width
            target_height = self.target_height
        except Exception:
            target_width = Emu(12192000)
            target_height = Emu(6858000)

        # 判断是否需要适配（尺寸不同时按比例缩放位置）
        needs_adapt = (self.template_width > 0 and self.template_height > 0 and
                       (target_width != self.template_width or target_height != self.template_height))

        for footer_info in self.footer_shapes:
            try:
                from pptx.enum.shapes import MSO_SHAPE_TYPE
                from io import BytesIO

                shape_type = footer_info["type"]
                shape_w = footer_info["width"]
                shape_h = footer_info["height"]
                orig_right = footer_info["right_margin"]
                orig_bottom = footer_info["bottom_margin"]

                if needs_adapt:
                    # 计算缩放比例
                    scale_x = target_width / self.template_width
                    scale_y = target_height / self.template_height
                    # 保持右边距和下边距的比例
                    new_right = orig_right * scale_x
                    new_bottom = orig_bottom * scale_y
                    new_w = shape_w * scale_x
                    new_h = shape_h * scale_y
                    new_left = target_width - new_w - new_right
                    new_top = target_height - new_h - new_bottom
                else:
                    new_left = footer_info["left"]
                    new_top = footer_info["top"]
                    new_w = shape_w
                    new_h = shape_h

                # 确保位置合理
                if new_left < 0:
                    new_left = 0
                if new_top < 0:
                    new_top = 0

                if shape_type == MSO_SHAPE_TYPE.PICTURE:
                    # 添加图片
                    image_blob = footer_info.get("image_blob")
                    if image_blob:
                        stream = BytesIO(image_blob)
                        slide.shapes.add_picture(
                            stream,
                            left=int(new_left),
                            top=int(new_top),
                            width=int(new_w),
                            height=int(new_h),
                        )
                else:
                    # 文本框或占位符（页脚、页码等），都用文本框添加
                    textbox = slide.shapes.add_textbox(
                        int(new_left), int(new_top), int(new_w), int(new_h)
                    )
                    text = footer_info.get("text", "")
                    if text:
                        textbox.text_frame.text = text
                        # 应用字体样式：所有页脚文本强制使用 Poppins 6号字
                        try:
                            font_name = "Poppins"
                            font_size = Pt(6)
                            font_bold = footer_info.get("font_bold")
                            font_color = footer_info.get("font_color")
                            
                            para = textbox.text_frame.paragraphs[0]
                            if para.runs:
                                run = para.runs[0]
                                run.font.name = font_name
                                run.font.size = font_size
                                if font_bold is not None:
                                    run.font.bold = font_bold
                                if font_color:
                                    run.font.color.rgb = font_color
                        except Exception:
                            pass
            except Exception:
                continue

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
        """将内容回填到模板slide中，所有样式遵循模板
        
        核心原则：
        1. 标题填入模板的title占位符（保留模板样式）
        2. 正文填入模板的body占位符（保留模板样式）
        3. 额外元素（图片、表格、装饰形状）按位置适配，样式用模板配色
        """
        # 1. 填入标题（使用模板标题占位符的样式）
        if model.title:
            self._fill_title_into_placeholder(slide, model.title)
        
        # 2. 填入正文文本（使用模板正文占位符的样式）
        text_blocks = [b for b in model.body_blocks if b.type == "paragraph" and b.text]
        if text_blocks:
            self._fill_body_into_placeholder(slide, text_blocks)
        
        # 3. 处理额外元素（图片、表格等），位置按比例适配，样式用模板配色
        for block in model.body_blocks:
            if block.type == "image":
                self._add_image_from_block(slide, block.content)
            elif block.type == "table":
                self._add_table_from_block(slide, block.content)
        
        # 4. 处理额外的自选图形（装饰性元素），样式用模板配色
        if model.raw_shapes:
            # 收集已经被占位符处理的文本，避免重复
            processed_texts = set()
            if model.title:
                processed_texts.add(model.title.strip())
            for block in model.body_blocks:
                if block.type == "paragraph" and block.text:
                    processed_texts.add(block.text.strip())
            
            for shape_data in model.raw_shapes:
                shape_type = shape_data.get("type")
                
                # 跳过标题文本框（已经被填进标题占位符了）
                if shape_type == "text" and model.title:
                    shape_text = ""
                    for para in shape_data.get("paragraphs", []):
                        for run in para.get("runs", []):
                            shape_text += run.get("text", "")
                    if shape_text.strip() == model.title.strip():
                        continue
                
                # 跳过正文文本框（已经被填进正文占位符了）
                if shape_type == "text":
                    shape_text = ""
                    for para in shape_data.get("paragraphs", []):
                        for run in para.get("runs", []):
                            shape_text += run.get("text", "")
                    shape_text = shape_text.strip()
                    if shape_text and shape_text in processed_texts:
                        continue
                
                # 跳过图片/表格（已经在body_blocks中处理了）
                if shape_type in ("image", "table"):
                    continue
                
                if shape_type == "text":
                    # 额外文本框：位置适配，字体用模板的正文字体，颜色用模板的正文颜色
                    self._add_extra_text_shape(slide, shape_data)
                elif shape_type == "autoshape":
                    # 装饰形状：位置适配，颜色用模板的强调色
                    self._add_extra_autoshape(slide, shape_data)
                elif shape_type == "image":
                    # 额外图片：位置适配
                    self._add_image_shape(slide, shape_data)

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
                # 判断是否是标题：位于页面顶部15%区域的文本
                is_title = False
                top = shape_data.get("top", 0)
                height = shape_data.get("height", 0)
                if top and self.target_height > 0:
                    if top + height < self.target_height * 0.2:
                        is_title = True
                self._add_text_shape(slide, shape_data, is_title=is_title)
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
                # 深色背景模板：装饰性形状/文本框改为白色
                if self.default_text_color == "FFFFFF":
                    textbox.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
                else:
                    textbox.fill.fore_color.rgb = RGBColor.from_string(fill_color)
            except Exception:
                pass

        line_color = shape_data.get("line_color")
        line_width = shape_data.get("line_width")
        if line_color:
            try:
                # 深色背景模板：装饰性线条改为白色
                if self.default_text_color == "FFFFFF":
                    textbox.line.color.rgb = RGBColor.from_string("FFFFFF")
                else:
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
        """应用run格式：完全遵循模板样式，只保留文字内容
        
        核心原则：
        - 字体：模板的主题字体（标题用major，正文用minor）
        - 字号：模板定义的字号
        - 颜色：模板定义的颜色
        - 粗体/斜体：模板要求的优先，否则保留原格式（装饰性）
        """
        font = run.font
        
        # 字体名称：使用模板的主题字体
        font_name = None
        if self.theme_fonts:
            font_name = self.theme_fonts.get("major" if is_title else "minor")
        if not font_name and template_fmt and template_fmt.get("font_name"):
            tmpl_font = template_fmt["font_name"]
            if tmpl_font.startswith("+mj") or tmpl_font.startswith("+mn"):
                if self.theme_fonts:
                    font_name = self.theme_fonts.get("major" if "mj" in tmpl_font else "minor")
            else:
                font_name = tmpl_font
        if font_name:
            font.name = font_name
        
        # 字号：使用模板要求的大小
        if template_fmt and template_fmt.get("font_size"):
            font.size = template_fmt["font_size"]
        
        # 粗体：模板要求优先
        if template_fmt and template_fmt.get("bold") is not None:
            font.bold = template_fmt["bold"]
        
        # 颜色：使用模板要求的颜色
        if template_fmt and template_fmt.get("color"):
            try:
                font.color.rgb = RGBColor.from_string(template_fmt["color"])
            except Exception:
                pass
        elif self.default_text_color:
            try:
                font.color.rgb = RGBColor.from_string(self.default_text_color)
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

    def _is_light_color(self, hex_color: str) -> bool:
        """判断颜色是否为浅色（亮度 > 0.5 视为浅色）"""
        if not hex_color or len(hex_color) != 6:
            return False
        try:
            r = int(hex_color[0:2], 16) / 255.0
            g = int(hex_color[2:4], 16) / 255.0
            b = int(hex_color[4:6], 16) / 255.0
            # 相对亮度计算
            luminance = 0.299 * r + 0.587 * g + 0.114 * b
            return luminance > 0.5
        except Exception:
            return False

    def _get_text_color_for_bg(self, bg_color: str | None) -> str:
        """根据背景颜色获取匹配的文字颜色（color pairing）
        
        - 浅色背景 → 深色文字（dk1: #0A2F24）
        - 深色背景 → 浅色文字（白色: #FFFFFF）
        - 无背景 → 使用默认文字颜色
        """
        if bg_color and self._is_light_color(bg_color):
            # 浅色背景：用深色文字
            return "0A2F24"
        elif bg_color and not self._is_light_color(bg_color):
            # 深色背景：用白色文字
            return "FFFFFF"
        else:
            # 无背景：使用默认文字颜色
            return self.default_text_color or "000000"

    def _is_decoration_shape(self, shape_data: dict) -> bool:
        """判断形状是否为装饰性形状（竖杆、小色块等）
        
        判断规则：
        - 宽度很小（< 0.5in）且高宽比 > 3 → 竖杆类装饰
        - 没有文字内容 → 装饰
        - 面积很小 → 装饰
        """
        try:
            width = shape_data.get("width", 0) or 0
            height = shape_data.get("height", 0) or 0
            width_in = width / 914400
            height_in = height / 914400
            
            # 检查是否有文字
            has_text = False
            for para in shape_data.get("paragraphs", []):
                for run in para.get("runs", []):
                    if run.get("text", "").strip():
                        has_text = True
                        break
                if has_text:
                    break
            
            # 窄竖条（宽度<0.5in，高宽比>3）→ 装饰
            if width_in < 0.5 and height_in > 0 and height_in / max(width_in, 0.01) > 3:
                return True
            
            # 没有文字的小形状 → 装饰
            if not has_text and width_in < 2 and height_in < 2:
                return True
            
            return False
        except Exception:
            return False

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
                    # 深色背景模板：装饰性形状改为白色
                    if self.default_text_color == "FFFFFF":
                        shape.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
                    else:
                        shape.fill.fore_color.rgb = RGBColor.from_string(fill_color)
                except Exception:
                    pass

            line_color = shape_data.get("line_color")
            line_width = shape_data.get("line_width")
            if line_color:
                try:
                    # 深色背景模板：装饰性线条改为白色
                    if self.default_text_color == "FFFFFF":
                        shape.line.color.rgb = RGBColor.from_string("FFFFFF")
                    else:
                        shape.line.color.rgb = RGBColor.from_string(line_color)
                    if line_width:
                        shape.line.width = line_width
                except Exception:
                    pass

            if shape_data.get("text") and shape.has_text_frame:
                shape.text_frame.text = shape_data["text"]
        except Exception:
            pass

    def _fill_title_into_placeholder(self, slide, title_text: str) -> None:
        """将标题填入模板的标题占位符，并应用模板的标题样式"""
        try:
            if slide.shapes.title is not None:
                tf = slide.shapes.title.text_frame
                # 应用模板标题格式
                template_fmt = self._get_template_format(True)
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    # 清空原有run
                    for run in p.runs:
                        run.text = ""
                    # 添加新run并设置内容和样式
                    if p.runs:
                        run = p.runs[0]
                        run.text = title_text
                        self._apply_run_format(run, {}, template_fmt, True)
                    else:
                        run = p.add_run()
                        run.text = title_text
                        self._apply_run_format(run, {}, template_fmt, True)
                else:
                    tf.text = title_text
                return
        except Exception:
            pass
        
        # 备用：通过placeholder查找
        for shape in slide.placeholders:
            try:
                phf = shape.placeholder_format
                if phf.type in (0, 3):  # title, ctrTitle
                    tf = shape.text_frame
                    template_fmt = self._get_template_format(True)
                    if tf.paragraphs:
                        p = tf.paragraphs[0]
                        for run in p.runs:
                            run.text = ""
                        if p.runs:
                            run = p.runs[0]
                            run.text = title_text
                            self._apply_run_format(run, {}, template_fmt, True)
                        else:
                            run = p.add_run()
                            run.text = title_text
                            self._apply_run_format(run, {}, template_fmt, True)
                    return
            except Exception:
                continue

    def _fill_body_into_placeholder(self, slide, text_blocks: list) -> None:
        """将正文填入模板的正文占位符，保留模板原有样式"""
        body_placeholder = self._find_body_placeholder(slide)
        if not body_placeholder:
            return
        
        try:
            tf = body_placeholder.text_frame
            template_fmt = self.template_formats.get("body", {})
            
            # 清空现有内容（但保留第一段的样式）
            # 先获取第一段的样式参考
            ref_para = tf.paragraphs[0] if tf.paragraphs else None
            
            # 用clear清空所有段落，然后重新添加
            tf.clear()
            
            # 填入新内容，每段应用模板格式
            first = True
            for block in text_blocks:
                if first:
                    p = tf.paragraphs[0]
                    first = False
                else:
                    p = tf.add_paragraph()
                p.text = block.text
                p.level = block.level
                if template_fmt:
                    self._apply_template_format_to_paragraph(p, template_fmt)
        except Exception:
            pass

    def _add_extra_text_shape(self, slide, shape_data: dict) -> None:
        """添加额外的文本框：位置适配，样式用模板正文样式
        
        Color Pairing 规则：
        - 装饰形状（竖杆、小色块）：用模板强调色（浅色背景→深绿，深色背景→白色）
        - 内容框（有文字的大形状）：保留原背景色，文字根据背景色深浅自动匹配
        """
        try:
            left = self._adapt_position(shape_data.get("left", Emu(914400)), "x")
            top = self._adapt_position(shape_data.get("top", Emu(914400)), "y")
            width = self._adapt_position(shape_data.get("width", Emu(914400 * 4)), "x")
            height = self._adapt_position(shape_data.get("height", Emu(914400 * 2)), "y")
            
            textbox = slide.shapes.add_textbox(left, top, width, height)
            tf = textbox.text_frame
            tf.word_wrap = True
            tf.clear()
            
            fill_color = shape_data.get("fill_color")
            is_decoration = self._is_decoration_shape(shape_data)
            
            if is_decoration:
                # 装饰形状：用模板强调色
                textbox.fill.solid()
                if self.default_text_color == "FFFFFF":
                    textbox.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
                else:
                    textbox.fill.fore_color.rgb = RGBColor.from_string("3DCD58")
            elif fill_color:
                # 内容框：保留原背景色（因为背景色决定了内容框的视觉效果）
                textbox.fill.solid()
                textbox.fill.fore_color.rgb = RGBColor.from_string(fill_color)
            
            # 处理线条颜色
            line_color = shape_data.get("line_color")
            if line_color:
                if is_decoration:
                    if self.default_text_color == "FFFFFF":
                        textbox.line.color.rgb = RGBColor.from_string("FFFFFF")
                    else:
                        textbox.line.color.rgb = RGBColor.from_string("3DCD58")
                else:
                    textbox.line.color.rgb = RGBColor.from_string(line_color)
            
            # 判断是否是标题区域的文本
            is_title = top < self.target_height * 0.2
            template_fmt = self._get_template_format(is_title)
            
            # 确定文字颜色：根据背景色做 color pairing
            actual_bg_color = None
            if is_decoration:
                actual_bg_color = "FFFFFF" if self.default_text_color == "FFFFFF" else "3DCD58"
            elif fill_color:
                actual_bg_color = fill_color
            
            text_color_override = self._get_text_color_for_bg(actual_bg_color) if actual_bg_color else None
            
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
                        # 应用 color pairing 的文字颜色（覆盖模板样式）
                        if text_color_override and run.text.strip():
                            try:
                                run.font.color.rgb = RGBColor.from_string(text_color_override)
                            except Exception:
                                pass
                else:
                    p.text = para_data.get("text", "")
                    if template_fmt:
                        self._apply_template_format_to_paragraph(p, template_fmt)
                    if text_color_override and p.text.strip():
                        try:
                            for run in p.runs:
                                run.font.color.rgb = RGBColor.from_string(text_color_override)
                        except Exception:
                            pass
                
                p.level = para_data.get("level", 0)
        except Exception:
            pass

    def _add_extra_autoshape(self, slide, shape_data: dict) -> None:
        """添加额外的装饰形状：位置适配，颜色用模板配色
        
        Color Pairing 规则：
        - 装饰形状（小、无文字）：用模板强调色
        - 内容框（大、有文字）：保留原背景色，文字根据背景色匹配
        """
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
            except Exception:
                pass
            
            left = self._adapt_position(shape_data.get("left", Emu(914400)), "x")
            top = self._adapt_position(shape_data.get("top", Emu(914400)), "y")
            width = self._adapt_position(shape_data.get("width", Emu(914400 * 4)), "x")
            height = self._adapt_position(shape_data.get("height", Emu(914400 * 2)), "y")
            
            shape = slide.shapes.add_shape(mso_shape, left, top, width, height)
            
            try:
                shape.rotation = shape_data.get("rotation", 0)
            except Exception:
                pass
            
            fill_color = shape_data.get("fill_color")
            is_decoration = self._is_decoration_shape(shape_data)
            
            # 确定背景色
            actual_bg_color = None
            if is_decoration:
                # 装饰形状：用模板强调色
                if self.default_text_color == "FFFFFF":
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
                    actual_bg_color = "FFFFFF"
                else:
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = RGBColor.from_string("3DCD58")
                    actual_bg_color = "3DCD58"
                try:
                    shape.line.color.rgb = RGBColor.from_string(actual_bg_color)
                except Exception:
                    pass
            elif fill_color:
                # 内容框：保留原背景色
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor.from_string(fill_color)
                actual_bg_color = fill_color
                line_color = shape_data.get("line_color")
                if line_color:
                    try:
                        shape.line.color.rgb = RGBColor.from_string(line_color)
                    except Exception:
                        pass
            
            # 如果有文字，应用模板样式 + color pairing
            if shape_data.get("text") and shape.has_text_frame:
                shape.text_frame.text = shape_data["text"]
                template_fmt = self._get_template_format(False)
                if template_fmt and shape.text_frame.paragraphs:
                    for p in shape.text_frame.paragraphs:
                        self._apply_template_format_to_paragraph(p, template_fmt)
                # color pairing：根据背景色调整文字颜色
                if actual_bg_color:
                    text_color = self._get_text_color_for_bg(actual_bg_color)
                    try:
                        for p in shape.text_frame.paragraphs:
                            for run in p.runs:
                                run.font.color.rgb = RGBColor.from_string(text_color)
                    except Exception:
                        pass
        except Exception:
            pass

    def _add_image_from_block(self, slide, content: dict) -> None:
        """从body block添加图片：位置适配"""
        try:
            from io import BytesIO
            stream = BytesIO(content.get("blob", b""))
            left = self._adapt_position(content.get("left", Emu(914400)), "x")
            top = self._adapt_position(content.get("top", Emu(914400)), "y")
            width = self._adapt_position(content.get("width", Emu(914400 * 4)), "x")
            height = self._adapt_position(content.get("height", Emu(914400 * 3)), "y")
            slide.shapes.add_picture(stream, left, top, width, height)
        except Exception:
            pass

    def _add_table_from_block(self, slide, content: dict) -> None:
        """从body block添加表格"""
        try:
            self._add_table_full(slide, {
                "left": content.get("left", Emu(914400)),
                "top": content.get("top", Emu(914400)),
                "width": content.get("width", Emu(914400 * 8)),
                "height": content.get("height", Emu(914400 * 4)),
                "data": content,
            })
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
            slide_width = self.target_width
            slide_height = self.target_height
        except Exception:
            slide_width = Emu(12192000)
            slide_height = Emu(6858000)

        for shape in slide.shapes:
            try:
                # 跳过满屏背景图
                if (shape.shape_type == 13 and  # MSO_SHAPE_TYPE.PICTURE
                    shape.left == 0 and shape.top == 0 and
                    abs(shape.width - slide_width) < Emu(100000) and
                    abs(shape.height - slide_height) < Emu(100000)):
                    continue

                if shape.left < 0:
                    shape.left = Emu(91440)
                if shape.top < 0:
                    shape.top = Emu(91440)
                if shape.left + shape.width > slide_width + Emu(10000):
                    shape.width = slide_width - shape.left - Emu(91440)
                if shape.top + shape.height > slide_height + Emu(10000):
                    shape.height = slide_height - shape.top - Emu(91440)
            except Exception:
                pass

    def _determine_layout(self, model: SlideContentModel) -> str:
        if model.original_layout_type:
            return model.original_layout_type
        if model.slide_index == 0:
            return "cover"
        return "content"
