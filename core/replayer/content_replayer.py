from pathlib import Path
from io import BytesIO
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
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
        self.background_color: str | None = None
        self.background_theme_color: str | None = None
        self.placeholder_mapping: dict = {}  # 模板占位符语义映射

        if template_path and Path(template_path).exists():
            extractor = TemplateFormatExtractor()
            try:
                self.template_formats = extractor.extract_placeholder_formats(template_path)
                self.theme_fonts = extractor.extract_theme_fonts(template_path)
                self._analyze_template_footer(template_path)
            except Exception:
                pass

    def _analyze_template_footer(self, template_path: str, master_index: int = 0) -> None:
        """分析指定Master中的footer元素（图标、页脚文本等）和背景信息，用于复制到新slide"""
        self.footer_shapes = []
        self.master_brand_images = []
        self.background_image = None
        self.background_color = None
        self.background_theme_color = None
        try:
            from pptx import Presentation
            from pptx.enum.shapes import MSO_SHAPE_TYPE
            from lxml import etree

            prs = Presentation(template_path)
            self.template_width = prs.slide_width
            self.template_height = prs.slide_height

            if not prs.slide_masters or master_index >= len(prs.slide_masters):
                return

            master = prs.slide_masters[master_index]
            footer_threshold = self.template_height * 0.85

            nsmap = {
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            }

            # 分析Master背景色
            bg = master._element.find('.//p:bg', nsmap)
            if bg is not None:
                bgPr = bg.find('.//p:bgPr', nsmap)
                if bgPr is not None:
                    solidFill = bgPr.find('.//a:solidFill', nsmap)
                    if solidFill is not None:
                        schemeClr = solidFill.find('.//a:schemeClr', nsmap)
                        if schemeClr is not None:
                            self.background_theme_color = schemeClr.get("val")
                        srgbClr = solidFill.find('.//a:srgbClr', nsmap)
                        if srgbClr is not None:
                            self.background_color = srgbClr.get("val")

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

                    # 底部区域的图片（施耐德图标等）- 需要从母版删除，但保留在footer_shapes中以便重新添加
                    if shape_type == MSO_SHAPE_TYPE.PICTURE:
                        if shape.top and shape.top > footer_threshold:
                            is_footer = True
                            # 记录母版中的品牌图片，后续需要删除
                            self.master_brand_images.append({
                                "name": shape.name,
                                "left": shape.left,
                                "top": shape.top,
                                "width": shape.width,
                                "height": shape.height,
                            })

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

    def _remove_master_brand_images(self, prs) -> None:
        """从母版中删除品牌图片，避免与代码添加的页脚图标重复"""
        if not self.master_brand_images:
            return

        try:
            for master in prs.slide_masters:
                shapes_to_remove = []

                for brand_img in self.master_brand_images:
                    brand_name = brand_img.get("name", "")
                    brand_left = brand_img.get("left")
                    brand_top = brand_img.get("top")
                    brand_width = brand_img.get("width")
                    brand_height = brand_img.get("height")

                    for shape in master.shapes:
                        try:
                            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                                if brand_name and shape.name == brand_name:
                                    shapes_to_remove.append(shape)
                                    break
                                if (brand_left is not None and brand_top is not None and
                                    brand_width is not None and brand_height is not None):
                                    if (abs(shape.left - brand_left) < Emu(1000) and
                                        abs(shape.top - brand_top) < Emu(1000) and
                                        abs(shape.width - brand_width) < Emu(1000) and
                                        abs(shape.height - brand_height) < Emu(1000)):
                                        shapes_to_remove.append(shape)
                                        break
                        except Exception:
                            continue

                for shape in shapes_to_remove:
                    try:
                        sp = shape._element
                        sp.getparent().remove(sp)
                    except Exception:
                        continue

                # 同时检查所有幻灯片布局
                for layout in master.slide_layouts:
                    layout_shapes_to_remove = []
                    for brand_img in self.master_brand_images:
                        brand_name = brand_img.get("name", "")
                        brand_left = brand_img.get("left")
                        brand_top = brand_img.get("top")
                        brand_width = brand_img.get("width")
                        brand_height = brand_img.get("height")

                        for shape in layout.shapes:
                            try:
                                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                                    if brand_name and shape.name == brand_name:
                                        layout_shapes_to_remove.append(shape)
                                        break
                                    if (brand_left is not None and brand_top is not None and
                                        brand_width is not None and brand_height is not None):
                                        if (abs(shape.left - brand_left) < Emu(1000) and
                                            abs(shape.top - brand_top) < Emu(1000) and
                                            abs(shape.width - brand_width) < Emu(1000) and
                                            abs(shape.height - brand_height) < Emu(1000)):
                                            layout_shapes_to_remove.append(shape)
                                            break
                            except Exception:
                                continue

                    for shape in layout_shapes_to_remove:
                        try:
                            sp = shape._element
                            sp.getparent().remove(sp)
                        except Exception:
                            continue
        except Exception:
            pass

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
                
                # 根据master_style设置颜色配置（关键修复）
                # F1/F2: 浅色背景，使用深色文字
                # F3/F4: 深色背景，使用白色文字
                master_style = config.master_style.upper()
                self.master_style = master_style  # 保存为实例变量，供其他方法使用
                colors = extractor.extract_theme_colors(self.template_path)
                
                if master_style in ("F1", "F2"):
                    # 浅色背景模板：正文用深色，标题用dark green background 2
                    self.title_text_color = colors.get("dk2", "3DCD58")
                    # 确保正文颜色为深色
                    if not self.default_text_color or self._is_light_color(self.default_text_color):
                        self.default_text_color = "333333"
                else:
                    # 深色背景模板(F3/F4)：正文和标题都用白色
                    self.title_text_color = colors.get("lt1", "FFFFFF")
                    self.default_text_color = "FFFFFF"
                
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
                
                # 从母版中删除品牌图片，避免与代码添加的页脚图标重复
                self._remove_master_brand_images(output_prs)
                
                # 提取所选Master的占位符语义映射
                self.placeholder_mapping = extractor.extract_placeholder_mapping(
                    self.template_path, selected_master_index
                )
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
                
                # 动态获取当前layout的占位符映射（关键修复）
                try:
                    extractor = TemplateFormatExtractor()
                    self.placeholder_mapping = extractor.extract_placeholder_mapping(
                        self.template_path, selected_master_index, layout_index
                    )
                except Exception:
                    pass
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
        """添加背景：优先使用背景图，否则使用纯色背景（与模板Master一致）"""
        if self.background_image:
            try:
                from io import BytesIO

                target_width = self.target_width
                target_height = self.target_height

                image_blob = self.background_image["image_blob"]
                stream = BytesIO(image_blob)

                pic = slide.shapes.add_picture(
                    stream,
                    left=0,
                    top=0,
                    width=target_width,
                    height=target_height,
                )

                spTree = slide.shapes._spTree
                spTree.remove(pic._element)
                spTree.insert(2, pic._element)
            except Exception:
                pass
        elif self.background_theme_color:
            try:
                from pptx.enum.dml import MSO_THEME_COLOR
                theme_color_map = {
                    'bg1': MSO_THEME_COLOR.BACKGROUND_1,
                    'bg2': MSO_THEME_COLOR.BACKGROUND_2,
                    'tx1': MSO_THEME_COLOR.TEXT_1,
                    'tx2': MSO_THEME_COLOR.TEXT_2,
                    'accent1': MSO_THEME_COLOR.ACCENT_1,
                    'accent2': MSO_THEME_COLOR.ACCENT_2,
                    'accent3': MSO_THEME_COLOR.ACCENT_3,
                    'accent4': MSO_THEME_COLOR.ACCENT_4,
                    'accent5': MSO_THEME_COLOR.ACCENT_5,
                    'accent6': MSO_THEME_COLOR.ACCENT_6,
                    'lt1': MSO_THEME_COLOR.LIGHT_1,
                    'lt2': MSO_THEME_COLOR.LIGHT_2,
                    'dk1': MSO_THEME_COLOR.DARK_1,
                    'dk2': MSO_THEME_COLOR.DARK_2,
                }
                mso_color = theme_color_map.get(self.background_theme_color.lower())
                if mso_color is not None:
                    bg = slide.background
                    fill = bg.fill
                    fill.solid()
                    fill.fore_color.theme_color = mso_color
            except Exception:
                pass
        elif self.background_color:
            try:
                bg = slide.background
                fill = bg.fill
                fill.solid()
                from pptx.dml.color import RGBColor
                fill.fore_color.rgb = RGBColor.from_string(self.background_color)
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
        """从Master中构建布局名称到索引的映射
        
        智能识别各种布局类型，优先选择最常用的内容页布局。
        """
        mapping = {}
        layout_names = []
        
        for idx, layout in enumerate(master.slide_layouts):
            name = (layout.name or "").lower()
            layout_names.append((idx, name, layout))
            
            # 直接用名称映射
            mapping[name] = idx
        
        # 按优先级匹配各种布局类型
        # 封面布局
        cover_priorities = [
            "title slide simple", "title slide", "cover", "封面",
            "title slide with image", "title slide offer",
        ]
        for kw in cover_priorities:
            for idx, name, layout in layout_names:
                if kw in name and "cover" not in mapping:
                    mapping["cover"] = idx
                    break
            if "cover" in mapping:
                break
        
        # 章节分隔页
        section_priorities = [
            "section break for longer copy", "section break with subhead",
            "section break", "section", "章节",
        ]
        for kw in section_priorities:
            for idx, name, layout in layout_names:
                if kw in name and "section" not in mapping:
                    mapping["section"] = idx
                    break
            if "section" in mapping:
                break
        
        # 结尾页
        closing_priorities = [
            "closing slide", "closing", "end", "结尾", "结束", "thank you",
        ]
        for kw in closing_priorities:
            for idx, name, layout in layout_names:
                if kw in name and "closing" not in mapping:
                    mapping["closing"] = idx
                    break
            if "closing" in mapping:
                break
        
        # 议程页
        agenda_priorities = ["agenda", "议程"]
        for kw in agenda_priorities:
            for idx, name, layout in layout_names:
                if kw in name and "agenda" not in mapping:
                    mapping["agenda"] = idx
                    break
            if "agenda" in mapping:
                break
        
        # 内容页布局（最常用的标准内容页）
        # 优先级：one column > blank slide with title > 其他单列布局
        content_priorities = [
            "one column", "one column two line headline",
            "one column with sketches", "blank slide with title",
            "content", "正文",
        ]
        for kw in content_priorities:
            for idx, name, layout in layout_names:
                if kw in name and "content" not in mapping:
                    mapping["content"] = idx
                    break
            if "content" in mapping:
                break
        
        # 双列布局
        two_col_priorities = [
            "two columns", "two column", "2 column",
            "two columns with subtitle",
        ]
        for kw in two_col_priorities:
            for idx, name, layout in layout_names:
                if kw in name and "two_column" not in mapping:
                    mapping["two_column"] = idx
                    break
            if "two_column" in mapping:
                break
        
        # 三列布局
        three_col_priorities = ["three columns", "three column", "3 column"]
        for kw in three_col_priorities:
            for idx, name, layout in layout_names:
                if kw in name and "three_column" not in mapping:
                    mapping["three_column"] = idx
                    break
            if "three_column" in mapping:
                break
        
        # 表格页（使用内容页布局即可）
        if "table" not in mapping and "content" in mapping:
            mapping["table"] = mapping["content"]
        
        # 图片页
        if "image" not in mapping and "content" in mapping:
            mapping["image"] = mapping["content"]
        
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
        """将内容回填到模板slide中，模板格式优先，原格式兜底
        
        核心原则（解决内容与格式割裂）：
        1. 标题填入模板的title占位符（保留模板样式）
        2. 副标题填入模板的subtitle占位符
        3. 主正文填入模板的body占位符（如果有且只有一个主正文形状）
        4. 其他内容作为额外形状按原位置回填，保留原PPT布局结构
        5. 用 raw_shape_id 精确去重，避免内容重复处理
        """
        # 已处理的 raw_shape_id 集合（用于精确去重）
        processed_shape_ids = set()
        
        # 如果没有占位符标题，从 body_blocks 中找 semantic_role="title" 的 block 作为标题
        title_text = model.title
        title_format = model.title_format
        # 记录被用作标题的 block（只取第一个，其余的作为副标题处理）
        title_block_processed = None
        if not title_text:
            title_blocks = [b for b in model.body_blocks if b.semantic_role == "title" and b.text]
            if title_blocks:
                title_text = title_blocks[0].text
                title_format = title_blocks[0].text_format
                title_block_processed = title_blocks[0]
                # 标记这个 block 对应的 raw_shape_id 已处理
                if title_blocks[0].raw_shape_id is not None:
                    processed_shape_ids.add(title_blocks[0].raw_shape_id)
        
        # 1. 填入标题（使用模板标题占位符的样式，原格式兜底）
        if title_text:
            title_placeholder = self._find_placeholder_by_role(slide, "title")
            if title_placeholder:
                self._fill_title_into_placeholder(slide, title_text, title_placeholder, title_format)
            else:
                # 备用：使用 slide.shapes.title
                self._fill_title_into_placeholder(slide, title_text, None, title_format)
        
        # 2. 根据语义角色分组正文内容
        # 剩余的 semantic_role="title" block（未被用作标题的）作为副标题处理
        body_main_blocks = []
        body_sidebar_blocks = []
        subtitle_blocks = []
        other_blocks = []
        
        for block in model.body_blocks:
            # 跳过被用作标题的那一个 block
            if title_block_processed is not None and block is title_block_processed:
                continue
            # 剩余的 role="title" block 作为副标题
            if block.semantic_role == "title" and block.text:
                subtitle_blocks.append(block)
                continue
            if block.type == "paragraph" and block.text:
                if block.semantic_role == "subtitle":
                    subtitle_blocks.append(block)
                elif block.semantic_role == "body_sidebar":
                    body_sidebar_blocks.append(block)
                elif block.semantic_role in ("body_main", "title", "unknown"):
                    body_main_blocks.append(block)
                else:
                    body_main_blocks.append(block)
            else:
                other_blocks.append(block)
        
        # 3. 填入副标题（如果有，原格式兜底）
        if subtitle_blocks:
            subtitle_placeholder = self._find_placeholder_by_role(slide, "subtitle")
            if subtitle_placeholder:
                self._fill_body_into_placeholder(slide, subtitle_blocks, subtitle_placeholder)
                # 标记这些 block 对应的 raw_shape_id 已处理
                for block in subtitle_blocks:
                    if block.raw_shape_id is not None:
                        processed_shape_ids.add(block.raw_shape_id)
            else:
                # 模板没有副标题占位符：将副标题合并到主正文的最前面
                body_main_blocks = subtitle_blocks + body_main_blocks
        
        # 4. 填入主正文（使用模板正文占位符的样式）
        # 判断是否应该用占位符填入：只有当主正文来自同一个 shape 时才用占位符
        # 如果主正文来自多个不同 shape，说明原PPT有多个文本框，应该按原位置回填
        # 注意：副标题合并进来时不影响多shape判断（副标题应该和正文一起填入占位符）
        main_block_shape_ids = set()
        for block in body_main_blocks:
            if block.raw_shape_id is not None and block.semantic_role != "title":
                main_block_shape_ids.add(block.raw_shape_id)
        
        # 如果主正文来自同一个 shape（或没有 shape_id），填入占位符
        # 如果来自多个 shape，则全部按原位置回填，保留原PPT布局
        if body_main_blocks and len(main_block_shape_ids) <= 1:
            body_placeholder = self._find_placeholder_by_role(slide, "body_main")
            if body_placeholder:
                self._fill_body_into_placeholder(slide, body_main_blocks, body_placeholder)
                # 标记已处理
                for block in body_main_blocks:
                    if block.raw_shape_id is not None:
                        processed_shape_ids.add(block.raw_shape_id)
        
        # 5. 填入侧边栏内容（如果有对应的占位符）
        if body_sidebar_blocks:
            sidebar_placeholder = self._find_placeholder_by_role(slide, "body_sidebar")
            if sidebar_placeholder:
                self._fill_body_into_placeholder(slide, body_sidebar_blocks, sidebar_placeholder)
                for block in body_sidebar_blocks:
                    if block.raw_shape_id is not None:
                        processed_shape_ids.add(block.raw_shape_id)
            # 如果没有侧边栏占位符，不合并到主正文，而是按原位置回填（保留布局）
        
        # 6. 处理额外元素（图片、表格等），位置按比例适配，样式用模板配色
        for block in other_blocks:
            if block.type == "image":
                self._add_image_from_block(slide, block.content)
                if block.raw_shape_id is not None:
                    processed_shape_ids.add(block.raw_shape_id)
            elif block.type == "table":
                self._add_table_from_block(slide, block.content)
                if block.raw_shape_id is not None:
                    processed_shape_ids.add(block.raw_shape_id)
        
        # 7. 处理额外的自选图形和文本框（装饰性元素、未被占位符处理的内容）
        # 用 raw_shape_id 精确去重，避免内容重复处理
        if model.raw_shapes:
            for shape_data in model.raw_shapes:
                shape_id = shape_data.get("shape_id")
                
                # 精确去重：跳过已处理的 shape
                if shape_id is not None and shape_id in processed_shape_ids:
                    continue
                
                shape_type = shape_data.get("type")
                
                # 跳过图片/表格（已经在body_blocks中处理了）
                if shape_type in ("image", "table"):
                    continue
                
                if shape_type == "text":
                    self._add_extra_text_shape(slide, shape_data)
                elif shape_type == "autoshape":
                    self._add_extra_autoshape(slide, shape_data)
                elif shape_type == "image":
                    self._add_image_shape(slide, shape_data)

    def _find_placeholder_by_role(self, slide, role: str) -> object | None:
        """根据语义角色查找模板占位符
        
        匹配优先级（从高到低）：
        1. placeholder_mapping 中类型和索引都匹配
        2. placeholder_mapping 中类型匹配
        3. 占位符名称模糊匹配
        4. 角色特定的回退逻辑（title用slide.shapes.title，body找最大的文本占位符）
        """
        # 标题特殊处理：优先用 slide.shapes.title
        if role == "title":
            try:
                if slide.shapes.title is not None:
                    return slide.shapes.title
            except Exception:
                pass
        
        if self.placeholder_mapping:
            ph_info = self.placeholder_mapping.get(role)
            if ph_info:
                target_type = ph_info.get("placeholder_type")
                target_idx = ph_info.get("placeholder_idx")
                target_name = ph_info.get("name", "")
                
                # 第一优先级：类型和索引都匹配
                for shape in slide.placeholders:
                    try:
                        phf = shape.placeholder_format
                        if phf.type == target_type and phf.idx == target_idx:
                            return shape
                    except Exception:
                        continue
                
                # 第二优先级：类型匹配
                for shape in slide.placeholders:
                    try:
                        phf = shape.placeholder_format
                        if phf.type == target_type:
                            return shape
                    except Exception:
                        continue
                
                # 第三优先级：名称匹配
                if target_name:
                    for shape in slide.placeholders:
                        try:
                            if shape.name == target_name:
                                return shape
                        except Exception:
                            continue
        
        # 第四优先级：角色特定的回退逻辑
        if role == "title":
            # 找类型为1(TITLE)或3(CENTER_TITLE)的占位符
            for shape in slide.placeholders:
                try:
                    phf = shape.placeholder_format
                    if phf.type in (1, 3):
                        return shape
                except Exception:
                    continue
            # 最后尝试 slide.shapes.title（前面已经试过了，这里作为双保险）
            try:
                if slide.shapes.title is not None:
                    return slide.shapes.title
            except Exception:
                pass
        elif role == "body_main":
            # 找最大的文本/内容占位符（优先左半部分）
            best_placeholder = None
            best_area = 0
            slide_width = self.target_width or Emu(12192000)
            for shape in slide.placeholders:
                try:
                    phf = shape.placeholder_format
                    # 跳过页眉页脚
                    if phf.type in (13, 14, 15, 16):
                        continue
                    # 跳过标题
                    if phf.type in (1, 3):
                        continue
                    # 只考虑有文本框的占位符
                    if shape.has_text_frame:
                        area = (shape.width or 0) * (shape.height or 0)
                        # 主正文优先左半部分
                        left = shape.left or 0
                        if left < slide_width * 0.5:
                            area *= 1.5  # 左半部分加权
                        if area > best_area:
                            best_area = area
                            best_placeholder = shape
                except Exception:
                    continue
            if best_placeholder:
                return best_placeholder
        elif role == "body_sidebar":
            # 侧边栏：优先找右半部分的占位符
            slide_width = self.target_width or Emu(12192000)
            best_placeholder = None
            best_area = 0
            for shape in slide.placeholders:
                try:
                    phf = shape.placeholder_format
                    # 跳过页眉页脚
                    if phf.type in (13, 14, 15, 16):
                        continue
                    # 跳过标题
                    if phf.type in (1, 3):
                        continue
                    # 只考虑有文本框的占位符
                    if shape.has_text_frame:
                        left = shape.left or 0
                        # 侧边栏优先右半部分
                        if left >= slide_width * 0.5:
                            area = (shape.width or 0) * (shape.height or 0)
                            if area > best_area:
                                best_area = area
                                best_placeholder = shape
                except Exception:
                    continue
            if best_placeholder:
                return best_placeholder
        elif role == "subtitle":
            # 找类型为4(SUBTITLE)的占位符
            for shape in slide.placeholders:
                try:
                    phf = shape.placeholder_format
                    if phf.type == 4:
                        return shape
                except Exception:
                    continue
        
        return None

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

        # 应用文本框填充色和边框色（保留原格式）
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
                    # 从run_data中提取原格式信息
                    original_format = None
                    if run_data:
                        from core.models import TextFormat
                        original_format = TextFormat(
                            font_name=run_data.get("font_name"),
                            font_size=run_data.get("font_size").pt if run_data.get("font_size") else None,
                            font_color=run_data.get("color"),
                            bold=run_data.get("bold"),
                            italic=run_data.get("italic"),
                            underline=run_data.get("underline"),
                        )
                    self._apply_run_format(run, run_data, template_fmt, is_title, original_format)
            else:
                p.text = para_data.get("text", "")
                # 应用模板格式到整段（原格式兜底）
                if template_fmt:
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

        # 自动调整文本框大小以适应内容，避免文本溢出
        self._auto_fit_textbox(textbox)

    def _get_template_format(self, is_title: bool) -> dict:
        """获取模板中标题或正文的格式规范"""
        if is_title:
            return self.template_formats.get("title", {})
        return self.template_formats.get("body", {})

    def _auto_fit_textbox(self, textbox) -> None:
        """自动调整文本框大小以适应内容，避免文本溢出
        
        策略：估算文本所需高度，如果超出当前容器高度，则扩大容器高度。
        """
        try:
            tf = textbox.text_frame
            if not tf.word_wrap:
                return

            box_width = textbox.width
            box_height = textbox.height
            if box_width <= 0 or box_height <= 0:
                return

            # 估算文本总高度
            total_text_height = 0
            for para in tf.paragraphs:
                text = para.text
                if not text:
                    total_text_height += Emu(914400 * 0.2)
                    continue

                # 获取字号
                font_size = Emu(914400 * 0.18)  # 默认14pt
                for run in para.runs:
                    if run.font.size:
                        font_size = run.font.size
                        break

                # 估算行数和行高
                char_width = font_size * 0.6
                chars_per_line = max(1, int(box_width / char_width))
                num_lines = max(1, (len(text) + chars_per_line - 1) // chars_per_line)
                line_height = font_size * 1.2
                total_text_height += num_lines * line_height

            # 如果文本高度超出容器高度，扩大容器
            if total_text_height > box_height:
                # 增加10%的余量
                new_height = int(total_text_height * 1.1)
                # 不超过页面高度的限制
                max_height = self.template_height * 0.8 if self.template_height else Emu(914400 * 7)
                if new_height > max_height:
                    new_height = max_height
                textbox.height = new_height
        except Exception:
            pass

    def _apply_run_format(self, run, run_data: dict, template_fmt: dict | None = None, is_title: bool = False, original_format=None) -> None:
        """应用run格式：模板格式优先，原格式兜底
        
        核心原则：
        - 字体：模板的主题字体优先，否则使用原格式的字体
        - 字号：模板定义的字号优先，否则使用原格式的字号
        - 颜色：模板定义的颜色优先，否则使用原格式的颜色
        - 粗体/斜体：模板要求的优先，否则保留原格式
        """
        font = run.font
        
        # 字体名称：模板主题字体优先，原格式兜底
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
        if not font_name and original_format and original_format.font_name:
            font_name = original_format.font_name
        if font_name:
            font.name = font_name
        
        # 字号：模板要求优先，原格式兜底
        if template_fmt and template_fmt.get("font_size"):
            font.size = template_fmt["font_size"]
        elif original_format and original_format.font_size:
            font.size = Pt(original_format.font_size)
        
        # 粗体：模板要求优先，原格式兜底
        if template_fmt and template_fmt.get("bold") is not None:
            font.bold = template_fmt["bold"]
        elif original_format and original_format.bold is not None:
            font.bold = original_format.bold
        
        # 斜体：模板要求优先，原格式兜底
        if template_fmt and template_fmt.get("italic") is not None:
            font.italic = template_fmt["italic"]
        elif original_format and original_format.italic is not None:
            font.italic = original_format.italic
        
        # 下划线：模板要求优先，原格式兜底
        if template_fmt and template_fmt.get("underline") is not None:
            font.underline = template_fmt["underline"]
        elif original_format and original_format.underline is not None:
            font.underline = original_format.underline
        
        # 颜色：原格式优先（保留原设计意图），模板格式兜底，最后使用默认颜色
        color_applied = False
        if original_format and original_format.font_color:
            try:
                font.color.rgb = RGBColor.from_string(original_format.font_color)
                color_applied = True
            except Exception:
                pass
        elif template_fmt and template_fmt.get("color"):
            try:
                font.color.rgb = RGBColor.from_string(template_fmt["color"])
                color_applied = True
            except Exception:
                pass
        elif self.default_text_color:
            try:
                font.color.rgb = RGBColor.from_string(self.default_text_color)
                color_applied = True
            except Exception:
                pass

    def _apply_template_format_to_paragraph(self, paragraph, template_fmt: dict, original_format=None) -> None:
        """将模板格式应用到整个段落：模板格式优先，原格式兜底"""
        # 对齐：模板优先，原格式兜底
        if template_fmt.get("alignment"):
            paragraph.alignment = self._parse_alignment(template_fmt["alignment"])
        elif original_format and original_format.alignment is not None:
            from pptx.enum.text import PP_ALIGN
            alignment_map = {0: PP_ALIGN.LEFT, 1: PP_ALIGN.CENTER, 2: PP_ALIGN.RIGHT, 3: PP_ALIGN.JUSTIFY}
            paragraph.alignment = alignment_map.get(original_format.alignment)
        
        # 行距：模板优先，原格式兜底
        if original_format and original_format.line_spacing:
            try:
                paragraph.line_spacing = Pt(original_format.line_spacing)
            except Exception:
                pass
        
        for run in paragraph.runs:
            font = run.font
            # 字体名称：模板优先，原格式兜底
            if template_fmt.get("font_name"):
                font.name = template_fmt["font_name"]
            elif original_format and original_format.font_name:
                font.name = original_format.font_name
            
            # 字号：模板优先，原格式兜底
            if template_fmt.get("font_size"):
                font.size = template_fmt["font_size"]
            elif original_format and original_format.font_size:
                font.size = Pt(original_format.font_size)
            
            # 粗体：模板优先，原格式兜底
            if template_fmt.get("bold") is not None:
                font.bold = template_fmt["bold"]
            elif original_format and original_format.bold is not None:
                font.bold = original_format.bold
            
            # 斜体：模板优先，原格式兜底
            if template_fmt.get("italic") is not None:
                font.italic = template_fmt["italic"]
            elif original_format and original_format.italic is not None:
                font.italic = original_format.italic
            
            # 颜色：原格式优先（保留原设计意图），模板格式兜底
            if original_format and original_format.font_color:
                try:
                    font.color.rgb = RGBColor.from_string(original_format.font_color)
                except Exception:
                    pass
            elif template_fmt.get("color"):
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
        
        # Color Pairing：确保表格文字在背景上可见
        try:
            if self.default_text_color:
                for r in range(rows):
                    for c in range(cols):
                        cell = table.cell(r, c)
                        # 检查单元格是否有背景色
                        cell_fill_color = None
                        try:
                            if cell.fill.type == 1:  # solid
                                cell_fill_color = str(cell.fill.fore_color.rgb)
                        except:
                            pass
                        
                        # 决定文字颜色
                        if cell_fill_color:
                            # 有单元格背景色：根据单元格背景色调整
                            text_color = self._get_text_color_for_bg(cell_fill_color)
                        else:
                            # 没有单元格背景色：用默认文字颜色
                            text_color = self.default_text_color
                        
                        # 应用文字颜色
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                try:
                                    run.font.color.rgb = RGBColor.from_string(text_color)
                                except Exception:
                                    pass
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
        table = table_shape.table
        for r, row_data in enumerate(data):
            for c, cell_text in enumerate(row_data):
                if c < cols:
                    table.cell(r, c).text = str(cell_text)
        
        # Color Pairing：确保表格文字在背景上可见
        try:
            if self.default_text_color:
                for r in range(rows):
                    for c in range(cols):
                        cell = table.cell(r, c)
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                try:
                                    run.font.color.rgb = RGBColor.from_string(self.default_text_color)
                                except Exception:
                                    pass
        except Exception:
            pass

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
                # 自动调整自选图形大小以适应文本
                self._auto_fit_textbox(shape)
        except Exception:
            pass

    def _fill_title_into_placeholder(self, slide, title_text: str, title_placeholder=None, original_format=None) -> None:
        """将标题填入模板的标题占位符，并应用模板的标题样式
        
        如果指定了 title_placeholder，则填入该占位符；
        否则查找 slide.shapes.title 或第一个标题占位符。
        
        original_format: 原PPT中标题的格式信息（TextFormat对象），用于模板未定义时兜底
        """
        try:
            # 优先使用传入的占位符
            if title_placeholder is not None:
                tf = title_placeholder.text_frame
                template_fmt = self._get_template_format(True)
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    for run in p.runs:
                        run.text = ""
                    if p.runs:
                        run = p.runs[0]
                        run.text = title_text
                        self._apply_run_format(run, {}, template_fmt, True, original_format)
                    else:
                        run = p.add_run()
                        run.text = title_text
                        self._apply_run_format(run, {}, template_fmt, True, original_format)
                else:
                    tf.text = title_text
                return
            
            # 备用：使用 slide.shapes.title
            if slide.shapes.title is not None:
                tf = slide.shapes.title.text_frame
                template_fmt = self._get_template_format(True)
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    for run in p.runs:
                        run.text = ""
                    if p.runs:
                        run = p.runs[0]
                        run.text = title_text
                        self._apply_run_format(run, {}, template_fmt, True, original_format)
                    else:
                        run = p.add_run()
                        run.text = title_text
                        self._apply_run_format(run, {}, template_fmt, True, original_format)
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
                            self._apply_run_format(run, {}, template_fmt, True, original_format)
                        else:
                            run = p.add_run()
                            run.text = title_text
                            self._apply_run_format(run, {}, template_fmt, True, original_format)
                    return
            except Exception:
                continue

    def _fill_body_into_placeholder(self, slide, text_blocks: list, body_placeholder=None) -> None:
        """将正文填入模板的正文占位符：模板格式优先，原格式兜底
        
        如果指定了 body_placeholder，则填入该占位符；
        否则使用 _find_body_placeholder 查找。
        """
        if not body_placeholder:
            body_placeholder = self._find_body_placeholder(slide)
        if not body_placeholder:
            return
        
        try:
            tf = body_placeholder.text_frame
            template_fmt = self.template_formats.get("body", {})
            
            # 清空现有内容（但保留第一段的样式）
            tf.clear()
            
            # 填入新内容，每段应用模板格式 + 原格式兜底
            first = True
            for block in text_blocks:
                if first:
                    p = tf.paragraphs[0]
                    first = False
                else:
                    p = tf.add_paragraph()
                p.text = block.text
                p.level = block.level
                
                # 获取该block的原始格式信息
                original_format = block.text_format if hasattr(block, 'text_format') else None
                
                # 应用格式：模板格式优先，原格式兜底
                if template_fmt or original_format:
                    self._apply_template_format_to_paragraph(p, template_fmt, original_format)
        except Exception:
            pass

    def _add_extra_text_shape(self, slide, shape_data: dict) -> None:
        """添加额外的文本框：位置适配，模板格式优先，原格式兜底
        
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
            else:
                # 没有填充色的文本框：根据master_style确定背景色
                # F1/F2: 浅色背景，F3/F4: 深色背景
                master_style = getattr(self, 'master_style', 'F1').upper()
                if master_style in ("F3", "F4"):
                    # 深色背景模板：给文本框加浅色背景，文字用深色
                    textbox.fill.solid()
                    textbox.fill.fore_color.rgb = RGBColor.from_string("E7FFD9")
                    actual_bg_color = "E7FFD9"
                else:
                    # 浅色背景模板：使用白色背景，文字用深色
                    actual_bg_color = "FFFFFF"
            
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
                        # 从run_data中提取原格式信息
                        original_format = None
                        if run_data:
                            from core.models import TextFormat
                            original_format = TextFormat(
                                font_name=run_data.get("font_name"),
                                font_size=run_data.get("font_size").pt if run_data.get("font_size") else None,
                                font_color=run_data.get("color"),
                                bold=run_data.get("bold"),
                                italic=run_data.get("italic"),
                                underline=run_data.get("underline"),
                            )
                        self._apply_run_format(run, run_data, template_fmt, is_title, original_format)
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
            
            # 自动调整文本框大小以适应内容，避免文本溢出
            self._auto_fit_textbox(textbox)
        except Exception:
            pass

    def _add_extra_autoshape(self, slide, shape_data: dict) -> None:
        """添加额外的装饰形状：位置适配，模板格式优先，原格式兜底
        
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
            
            # 确定背景色：保留原格式优先
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
            else:
                # 内容框没有填充色：根据整体背景色决定是否加背景
                if self.default_text_color == "FFFFFF":
                    # 深色背景下，给内容框加浅色背景，确保文字可见
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = RGBColor.from_string("E7FFD9")
                    actual_bg_color = "E7FFD9"
                else:
                    # 浅色背景下，内容框透明也没关系，文字是深色的
                    actual_bg_color = None
            
            # 如果有文字，应用模板样式 + 原格式兜底 + color pairing
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
            
            # 自动调整文本框大小以适应内容，避免文本溢出
            self._auto_fit_textbox(textbox)
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
        """根据内容模型和布局特征确定布局类型
        
        优先级：
        1. original_layout_type（如果是封面/章节/结尾等特殊布局）
        2. layout_features 中的特征（表格/图片/多列等）
        3. 默认内容页布局
        
        注意：普通内容页的 original_layout_type 应该是 "content"，
        不要因为布局名称包含 "title" 就误判为 "cover"
        """
        # 特殊布局类型：封面、章节、结尾、议程 - 保留原始布局
        special_layouts = {"cover", "section", "closing", "agenda"}
        if model.original_layout_type and model.original_layout_type in special_layouts:
            return model.original_layout_type
        
        # 根据布局特征判断
        features = model.layout_features or {}
        
        # 多列布局
        if features.get("columns", 1) >= 2:
            return "two_column"
        
        # 表格页
        if features.get("has_table"):
            return "table"
        
        # 图片页（有图片且文字少）
        if features.get("has_image") and features.get("text_density", 1) < 0.3:
            return "image"
        
        # 默认内容页
        return "content"
