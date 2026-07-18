"""模板分析器：分析 PPT 模板中的 master 和 layout，识别风格特征"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Emu
from lxml import etree
from typing import Any, Dict


NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


class TemplateAnalyzer:
    """分析 PPT 模板的结构与风格"""

    def __init__(self):
        self._theme_colors = {}

    def analyze(self, template_path: str) -> dict:
        if not Path(template_path).exists():
            raise FileNotFoundError(f"Template not found: {template_path}")

        prs = Presentation(template_path)
        self._theme_colors = self._extract_theme_colors(prs)

        masters_info = self._analyze_masters(prs)

        return {
            "file": str(template_path),
            "theme_colors": self._theme_colors,
            "masters": masters_info,
            "total_layouts": sum(len(m["layouts"]) for m in masters_info),
            "total_slides": len(prs.slides),
            "style_matches": self._match_styles(prs),
        }

    def _extract_theme_colors(self, prs) -> Dict[str, str]:
        """从主题 XML 提取所有颜色定义"""
        colors = {}
        try:
            theme_part = None
            for rel in prs.part.rels.values():
                if "theme" in rel.reltype:
                    theme_part = rel.target_part
                    break

            if theme_part is None:
                return colors

            theme_xml = etree.fromstring(theme_part.blob)
            clr_scheme = theme_xml.find(".//a:clrScheme", NSMAP)
            if clr_scheme is None:
                return colors

            color_names = {
                "a:dk1": "Dark1",
                "a:lt1": "Light1",
                "a:dk2": "Dark2",
                "a:lt2": "Light2",
                "a:accent1": "Accent1",
                "a:accent2": "Accent2",
                "a:accent3": "Accent3",
                "a:accent4": "Accent4",
                "a:accent5": "Accent5",
                "a:accent6": "Accent6",
                "a:hlink": "Hyperlink",
                "a:folHlink": "FollowedHyperlink",
            }

            for elem in clr_scheme:
                tag_name = etree.QName(elem.tag).localname
                full_tag = f"a:{tag_name}"
                if full_tag in color_names:
                    rgb = self._parse_color_elem(elem)
                    if rgb:
                        colors[color_names[full_tag]] = rgb

        except Exception as e:
            pass
        return colors

    def _parse_color_elem(self, elem) -> str | None:
        """解析颜色元素，返回 RGB 字符串"""
        try:
            rgb_elem = elem.find(".//a:rgbClr", NSMAP)
            if rgb_elem is not None:
                val = rgb_elem.get("val", "")
                if val:
                    return val.upper()

            sys_clr = elem.find(".//a:sysClr", NSMAP)
            if sys_clr is not None:
                last_clr = sys_clr.get("lastClr", "")
                if last_clr:
                    return last_clr.upper()

            scheme_clr = elem.find(".//a:schemeClr", NSMAP)
            if scheme_clr is not None:
                val = scheme_clr.get("val", "")
                if val:
                    mapped = {
                        "dk1": "202020",
                        "lt1": "FFFFFF",
                        "dk2": "404040",
                        "lt2": "F0F0F0",
                    }
                    return mapped.get(val, None)
        except Exception:
            pass
        return None

    def _analyze_masters(self, prs) -> list[dict]:
        masters = []
        for idx, master in enumerate(prs.slide_masters):
            master_info = {
                "index": idx,
                "name": self._get_master_name(master),
                "background": self._get_background(master),
                "fonts": self._get_fonts(master),
                "layouts": self._analyze_layouts(master),
            }
            masters.append(master_info)
        return masters

    def _analyze_layouts(self, master) -> list[dict]:
        layouts = []
        for idx, layout in enumerate(master.slide_layouts):
            layout_info = {
                "index": idx,
                "name": layout.name or "(未命名)",
                "background": self._get_background(layout),
                "placeholders": self._get_placeholders(layout),
                "type_guess": self._guess_layout_type(layout.name or ""),
            }
            layouts.append(layout_info)
        return layouts

    def _get_master_name(self, master) -> str:
        try:
            return master.name or f"Master_{id(master)}"
        except Exception:
            return "(未命名)"

    def _get_background(self, element) -> dict:
        """获取背景信息，解析主题颜色为实际 RGB"""
        try:
            bg = element.background
            fill = bg.fill
            if fill.type is None:
                return {"type": "inherit", "color": None, "gradient": None, "theme_color": None}

            from pptx.enum.dml import MSO_FILL
            if fill.type == MSO_FILL.SOLID:
                color = fill.fore_color
                try:
                    rgb = str(color.rgb)
                    return {"type": "solid", "color": rgb, "gradient": None, "theme_color": None}
                except Exception:
                    try:
                        theme_color = color.theme_color
                        theme_color_str = str(theme_color)
                        actual_rgb = self._resolve_theme_color(theme_color_str)
                        return {
                            "type": "solid",
                            "color": actual_rgb,
                            "gradient": None,
                            "theme_color": theme_color_str,
                        }
                    except Exception:
                        return {"type": "unknown", "color": None, "gradient": None, "theme_color": None}
            elif fill.type == MSO_FILL.GRADIENT:
                stops = []
                theme_stops = []
                try:
                    for stop in fill.gradient_stops:
                        try:
                            stops.append(str(stop.color.rgb))
                            theme_stops.append(None)
                        except Exception:
                            try:
                                tc = str(stop.color.theme_color)
                                rgb = self._resolve_theme_color(tc)
                                stops.append(rgb if rgb else "?")
                                theme_stops.append(tc)
                            except Exception:
                                stops.append("?")
                                theme_stops.append(None)
                except Exception:
                    pass
                return {
                    "type": "gradient",
                    "color": None,
                    "gradient": stops,
                    "theme_color": theme_stops if any(theme_stops) else None,
                }
            elif fill.type == MSO_FILL.PICTURE:
                return {"type": "picture", "color": None, "gradient": None, "theme_color": None}
            else:
                return {"type": str(fill.type), "color": None, "gradient": None, "theme_color": None}
        except Exception as e:
            return {"type": "error", "color": None, "gradient": None, "theme_color": None, "error": str(e)}

    def _resolve_theme_color(self, theme_color_str: str) -> str | None:
        """将主题颜色名解析为实际 RGB"""
        mapping = {
            "BACKGROUND_1": self._theme_colors.get("Light1"),
            "BACKGROUND_2": self._theme_colors.get("Light2"),
            "TEXT_1": self._theme_colors.get("Dark1"),
            "TEXT_2": self._theme_colors.get("Dark2"),
            "ACCENT_1": self._theme_colors.get("Accent1"),
            "ACCENT_2": self._theme_colors.get("Accent2"),
            "ACCENT_3": self._theme_colors.get("Accent3"),
            "ACCENT_4": self._theme_colors.get("Accent4"),
            "ACCENT_5": self._theme_colors.get("Accent5"),
            "ACCENT_6": self._theme_colors.get("Accent6"),
            "HYPERLINK": self._theme_colors.get("Hyperlink"),
            "FOLLOWED_HYPERLINK": self._theme_colors.get("FollowedHyperlink"),
        }
        return mapping.get(theme_color_str)

    def _get_fonts(self, element) -> dict:
        try:
            master_part = element.part if hasattr(element, "part") else None
            if master_part is None:
                return {"major": None, "minor": None}

            theme_part = None
            try:
                for rel in master_part.rels.values():
                    if "theme" in rel.reltype:
                        theme_part = rel.target_part
                        break
            except Exception:
                pass

            if theme_part is None:
                return {"major": None, "minor": None}

            theme_xml = etree.fromstring(theme_part.blob)
            font_scheme = theme_xml.find(".//a:fontScheme", NSMAP)
            if font_scheme is None:
                return {"major": None, "minor": None}

            major = font_scheme.find(".//a:majorFont/a:latin", NSMAP)
            minor = font_scheme.find(".//a:minorFont/a:latin", NSMAP)

            return {
                "major": major.get("typeface") if major is not None else None,
                "minor": minor.get("typeface") if minor is not None else None,
            }
        except Exception as e:
            return {"major": None, "minor": None, "error": str(e)}

    def _get_placeholders(self, layout) -> list[dict]:
        placeholders = []
        try:
            for ph in layout.placeholders:
                placeholders.append({
                    "idx": ph.placeholder_format.idx,
                    "type": str(ph.placeholder_format.type),
                    "name": ph.name,
                })
        except Exception:
            pass
        return placeholders

    def _guess_layout_type(self, name: str) -> str:
        name_lower = name.lower()
        if "封面" in name or "cover" in name_lower or "title" in name_lower:
            return "cover"
        if "章节" in name or "section" in name_lower:
            return "section"
        if "内容" in name or "content" in name_lower:
            return "content"
        if "结尾" in name or "结束" in name or "closing" in name_lower or "end" in name_lower:
            return "closing"
        return "unknown"

    def _match_styles(self, prs) -> list[dict]:
        matches = []
        for idx, master in enumerate(prs.slide_masters):
            bg = self._get_background(master)
            matched = self._match_single_style(bg)
            matches.append({
                "master_index": idx,
                "master_name": self._get_master_name(master),
                "background": bg,
                "matched_style": matched["style_id"],
                "matched_name": matched["style_name"],
                "confidence": matched["confidence"],
            })
        return matches

    def _match_single_style(self, bg: dict) -> dict:
        bg_type = bg.get("type")
        color = bg.get("color")
        gradient = bg.get("gradient")

        if bg_type == "gradient" and gradient and len(gradient) >= 2:
            return {
                "style_id": "F4",
                "style_name": "渐变科技",
                "confidence": "高",
            }

        if (bg_type == "solid" or bg_type == "theme") and color:
            color_lower = color.lower()
            if self._is_white(color_lower):
                return {"style_id": "F1", "style_name": "白色简约", "confidence": "高"}
            elif self._is_light_green(color_lower):
                return {"style_id": "F2", "style_name": "浅绿色清新", "confidence": "高"}
            elif self._is_dark_green(color_lower):
                return {"style_id": "F3", "style_name": "深绿色商务", "confidence": "高"}
            else:
                return {
                    "style_id": "?",
                    "style_name": f"未匹配（颜色 #{color}）",
                    "confidence": "低",
                }

        return {
            "style_id": "?",
            "style_name": "需人工确认",
            "confidence": "低",
        }

    def _is_white(self, color: str) -> bool:
        try:
            r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
            return r >= 240 and g >= 240 and b >= 240
        except Exception:
            return False

    def _is_light_green(self, color: str) -> bool:
        try:
            r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
            return g > r and g > b and r > 200 and g > 220
        except Exception:
            return False

    def _is_dark_green(self, color: str) -> bool:
        try:
            r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
            return g > r and g > b and r < 100 and g < 150 and b < 100
        except Exception:
            return False

    def _color_match(self, color1: str, color2: str, tolerance: int) -> bool:
        try:
            r1, g1, b1 = int(color1[0:2], 16), int(color1[2:4], 16), int(color1[4:6], 16)
            r2, g2, b2 = int(color2[0:2], 16), int(color2[2:4], 16), int(color2[4:6], 16)
            return (
                abs(r1 - r2) <= tolerance
                and abs(g1 - g2) <= tolerance
                and abs(b1 - b2) <= tolerance
            )
        except Exception:
            return False
