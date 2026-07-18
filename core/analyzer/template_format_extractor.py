"""模板格式提取器：从模板的Master/Layout中提取占位符的格式规范"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.enum.text import PP_ALIGN
from lxml import etree


class TemplateFormatExtractor:
    """从模板提取标题/正文占位符的字体、字号、颜色规范"""

    _NSMAP = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    }

    def extract_placeholder_formats(self, template_path: str) -> dict:
        """提取模板中各占位符的格式规范
        
        返回结构:
        {
            "title": {"font_name": "...", "font_size": Pt, "color": "RRGGBB", "bold": bool, "alignment": ...},
            "body": {"font_name": "...", "font_size": Pt, "color": "RRGGBB", ...},
            "subtitle": {...},
        }
        """
        prs = Presentation(template_path)
        formats = {}

        # 优先从第一个Master提取
        if prs.slide_masters:
            master = prs.slide_masters[0]
            master_formats = self._extract_formats_from_shapes(master.shapes)
            for k, v in master_formats.items():
                formats.setdefault(k, v)

        # 从Layouts补充（覆盖Master的设置）
        for layout in prs.slide_layouts:
            layout_formats = self._extract_formats_from_shapes(layout.shapes)
            for k, v in layout_formats.items():
                formats[k] = v

        return formats

    def _extract_formats_from_shapes(self, shapes) -> dict:
        formats = {}
        for shape in shapes:
            try:
                if not shape.is_placeholder:
                    continue
                phf = shape.placeholder_format
                ph_type = phf.type

                if ph_type == 0 or (phf.idx == 0 and ph_type is None):
                    formats["title"] = self._extract_text_format(shape)
                elif ph_type == 2 or (phf.idx == 1 and "title" not in formats):
                    formats["body"] = self._extract_text_format(shape)
                elif ph_type == 4:
                    formats["subtitle"] = self._extract_text_format(shape)
            except Exception:
                continue
        return formats

    def _extract_text_format(self, shape) -> dict:
        """从形状提取文本格式，包括继承自XML的格式"""
        fmt = {
            "font_name": None,
            "font_size": None,
            "bold": None,
            "italic": None,
            "color": None,
            "alignment": None,
        }

        try:
            tf = shape.text_frame
            if tf.paragraphs:
                p = tf.paragraphs[0]
                if p.alignment is not None:
                    fmt["alignment"] = str(p.alignment)

                if p.runs:
                    run = p.runs[0]
                    font = run.font
                    if font.name:
                        fmt["font_name"] = font.name
                    if font.size:
                        fmt["font_size"] = font.size
                    if font.bold is not None:
                        fmt["bold"] = font.bold
                    if font.italic is not None:
                        fmt["italic"] = font.italic
                    try:
                        if font.color and font.color.rgb:
                            fmt["color"] = str(font.color.rgb)
                    except Exception:
                        pass
        except Exception:
            pass

        # 从XML深层解析（处理继承的格式）
        try:
            xml_format = self._parse_format_from_xml(shape)
            for key, value in xml_format.items():
                if value is not None and fmt.get(key) is None:
                    fmt[key] = value
        except Exception:
            pass

        return fmt

    def _parse_format_from_xml(self, shape) -> dict:
        """直接从XML元素解析格式，包括布局样式中的默认格式"""
        fmt = {}
        try:
            xml_elem = shape.element

            # 查找 txBody 中的段落属性和run属性
            lst_style = xml_elem.find(".//p:txBody/a:lstStyle", self._NSMAP)
            if lst_style is not None:
                # 标题级别样式
                for level in range(5):
                    lvl_style = lst_style.find(f"a:lvl{level+1}pPr", self._NSMAP)
                    if lvl_style is not None:
                        def_rpr = lvl_style.find("a:defRPr", self._NSMAP)
                        if def_rpr is not None:
                            if fmt.get("font_name") is None:
                                latin = def_rpr.find("a:latin", self._NSMAP)
                                if latin is not None:
                                    typeface = latin.get("typeface")
                                    if typeface:
                                        fmt["font_name"] = typeface
                            if fmt.get("font_size") is None:
                                sz = def_rpr.get("sz")
                                if sz:
                                    fmt["font_size"] = Pt(int(sz) / 100)
                            if fmt.get("bold") is None:
                                b = def_rpr.get("b")
                                if b:
                                    fmt["bold"] = b == "1"
                            if fmt.get("italic") is None:
                                i = def_rpr.get("i")
                                if i:
                                    fmt["italic"] = i == "1"
                            if fmt.get("color") is None:
                                solid_fill = def_rpr.find("a:solidFill", self._NSMAP)
                                if solid_fill is not None:
                                    srgb = solid_fill.find("a:srgbClr", self._NSMAP)
                                    if srgb is not None:
                                        fmt["color"] = srgb.get("val")
                            break
        except Exception:
            pass
        return fmt

    def extract_theme_fonts(self, template_path: str) -> dict:
        """从模板主题提取字体方案"""
        prs = Presentation(template_path)
        theme_part = prs.slide_masters[0].part if prs.slide_masters else None
        if not theme_part:
            return {}

        theme_fonts = {"major": None, "minor": None}
        try:
            # 查找主题关联
            for rel in theme_part.rels.values():
                if "theme" in rel.reltype:
                    theme_element = etree.fromstring(rel.target_part.blob)
                    font_scheme = theme_element.find(
                        ".//a:fontScheme", self._NSMAP
                    )
                    if font_scheme is not None:
                        major = font_scheme.find("a:majorFont", self._NSMAP)
                        minor = font_scheme.find("a:minorFont", self._NSMAP)
                        if major is not None:
                            latin = major.find("a:latin", self._NSMAP)
                            if latin is not None:
                                theme_fonts["major"] = latin.get("typeface")
                        if minor is not None:
                            latin = minor.find("a:latin", self._NSMAP)
                            if latin is not None:
                                theme_fonts["minor"] = latin.get("typeface")
                    break
        except Exception:
            pass

        return theme_fonts
