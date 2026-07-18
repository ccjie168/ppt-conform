"""模板分析器：分析 PPT 模板中的 master 和 layout，识别风格特征"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Emu
from lxml import etree
from typing import Any


# 命名空间
NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


# F1-F4 风格定义（背景色 HEX）
F_STYLE_COLORS = {
    "F1": {"name": "白色简约", "expected_bg": "FFFFFF", "tolerance": 10},
    "F2": {"name": "浅绿色清新", "expected_bg": "E8F5E9", "tolerance": 30},
    "F3": {"name": "深绿色商务", "expected_bg": "1B5E20", "tolerance": 40},
    "F4": {"name": "渐变科技", "expected_bg": None, "tolerance": 0},  # 渐变色特殊处理
}


class TemplateAnalyzer:
    """分析 PPT 模板的结构与风格"""

    def analyze(self, template_path: str) -> dict:
        """分析模板，返回结构化信息"""
        if not Path(template_path).exists():
            raise FileNotFoundError(f"Template not found: {template_path}")

        prs = Presentation(template_path)
        return {
            "file": str(template_path),
            "masters": self._analyze_masters(prs),
            "total_layouts": sum(
                len(m["layouts"]) for m in self._analyze_masters(prs)
            ),
            "total_slides": len(prs.slides),
            "style_matches": self._match_styles(prs),
        }

    def _analyze_masters(self, prs) -> list[dict]:
        """分析所有 slide master"""
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
        """分析 master 下的所有 layout"""
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
        """获取 master 名称"""
        try:
            return master.name or f"Master_{id(master)}"
        except Exception:
            return "(未命名)"

    def _get_background(self, element) -> dict:
        """获取背景信息：纯色/渐变/图片"""
        try:
            bg = element.background
            fill = bg.fill
            if fill.type is None:
                return {"type": "inherit", "color": None, "gradient": None}

            from pptx.enum.dml import MSO_FILL
            if fill.type == MSO_FILL.SOLID:
                color = fill.fore_color
                try:
                    rgb = str(color.rgb)
                    return {"type": "solid", "color": rgb, "gradient": None}
                except Exception:
                    try:
                        theme_color = color.theme_color
                        return {"type": "theme", "color": str(theme_color), "gradient": None}
                    except Exception:
                        return {"type": "unknown", "color": None, "gradient": None}
            elif fill.type == MSO_FILL.GRADIENT:
                stops = []
                try:
                    for stop in fill.gradient_stops:
                        try:
                            stops.append(str(stop.color.rgb))
                        except Exception:
                            stops.append("theme")
                except Exception:
                    pass
                return {"type": "gradient", "color": None, "gradient": stops}
            elif fill.type == MSO_FILL.PICTURE:
                return {"type": "picture", "color": None, "gradient": None}
            else:
                return {"type": str(fill.type), "color": None, "gradient": None}
        except Exception as e:
            return {"type": "error", "color": None, "gradient": None, "error": str(e)}

    def _get_fonts(self, element) -> dict:
        """获取 master 的主题字体"""
        try:
            theme = element.slide_master.element if hasattr(element, "slide_master") else element.element
            # 尝试从 theme 获取
            master_part = element.part if hasattr(element, "part") else None
            if master_part is None:
                return {"major": None, "minor": None}

            # 通过 XML 查找字体定义
            theme_part = None
            try:
                # master 关联的 theme
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
        """获取 layout 的占位符"""
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
        """根据 layout 名称推测类型"""
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
        """将模板中的 master 与 F1-F4 风格进行匹配"""
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
        """根据背景信息匹配 F1-F4 风格"""
        bg_type = bg.get("type")
        color = bg.get("color")
        gradient = bg.get("gradient")

        # F4 渐变科技：优先匹配渐变
        if bg_type == "gradient" and gradient and len(gradient) >= 2:
            return {
                "style_id": "F4",
                "style_name": F_STYLE_COLORS["F4"]["name"],
                "confidence": "高",
            }

        # F1-F3 通过纯色匹配
        if bg_type == "solid" and color:
            for style_id, info in F_STYLE_COLORS.items():
                if info["expected_bg"] is None:
                    continue
                if self._color_match(color, info["expected_bg"], info["tolerance"]):
                    return {
                        "style_id": style_id,
                        "style_name": info["name"],
                        "confidence": "高",
                    }
            # 颜色不匹配任何已知风格
            return {
                "style_id": "?",
                "style_name": f"未匹配（背景色 {color}）",
                "confidence": "低",
            }

        # theme 继承或图片背景：低置信度
        return {
            "style_id": "?",
            "style_name": "需人工确认",
            "confidence": "低",
        }

    def _color_match(self, color1: str, color2: str, tolerance: int) -> bool:
        """比较两个 HEX 颜色是否在容差内"""
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
