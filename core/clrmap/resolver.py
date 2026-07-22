import zipfile
from lxml import etree
from pathlib import Path

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"


class ClrMapResolver:
    def __init__(self, template_path: str, master_index: int):
        self.template_path = template_path
        self.master_index = master_index  # 0-indexed
        self.clr_map: dict[str, str] = {}
        self.theme_colors: dict[str, str] = {}
        self.theme_name: str = ""
        self._load()

    def _load(self):
        master_num = self.master_index + 1
        master_xml_path = f"ppt/slideMasters/slideMaster{master_num}.xml"
        master_rels_path = f"ppt/slideMasters/_rels/slideMaster{master_num}.xml.rels"

        with zipfile.ZipFile(self.template_path, "r") as zf:
            # 1. Read Master's clrMap
            master_xml = zf.read(master_xml_path)
            master_elem = etree.fromstring(master_xml)

            clr_map_elem = master_elem.find(f"{{{NS_P}}}clrMap")
            if clr_map_elem is not None:
                for attr_name in ["bg1", "tx1", "bg2", "tx2",
                                  "accent1", "accent2", "accent3",
                                  "accent4", "accent5", "accent6",
                                  "hlink", "folHlink"]:
                    val = clr_map_elem.get(attr_name)
                    if val:
                        self.clr_map[attr_name] = val

            # 2. Find the associated theme file
            master_rels_xml = zf.read(master_rels_path)
            rels_elem = etree.fromstring(master_rels_xml)
            theme_target = None
            for rel in rels_elem.findall(f"{{{NS_RELS}}}Relationship"):
                if "theme" in rel.get("Type", ""):
                    theme_target = rel.get("Target", "")
                    break

            if not theme_target:
                return

            # Resolve theme path (handle ../ prefix)
            if theme_target.startswith("../"):
                theme_path = f"ppt/{theme_target[3:]}"
            else:
                theme_path = f"ppt/{theme_target}"

            # 3. Read Theme's clrScheme
            theme_xml = zf.read(theme_path)
            theme_elem = etree.fromstring(theme_xml)
            self.theme_name = theme_elem.get("name", "")

            clr_scheme = theme_elem.find(f".//{{{NS_A}}}clrScheme")
            if clr_scheme is not None:
                for child in clr_scheme:
                    tag = child.tag.split("}")[-1]
                    srgb = child.find(f"{{{NS_A}}}srgbClr")
                    sysclr = child.find(f"{{{NS_A}}}sysClr")
                    if srgb is not None:
                        self.theme_colors[tag] = srgb.get("val", "")
                    elif sysclr is not None:
                        self.theme_colors[tag] = sysclr.get("lastClr", "")

    def resolve_scheme_color(self, scheme_name: str) -> str | None:
        """Resolve a scheme color name (like 'bg1') to actual RGB hex string."""
        mapped_name = self.clr_map.get(scheme_name, scheme_name)
        return self.theme_colors.get(mapped_name)

    def get_background_color(self) -> str | None:
        return self.resolve_scheme_color("bg1")

    def get_text_color(self) -> str | None:
        return self.resolve_scheme_color("tx1")

    def get_accent_color(self) -> str | None:
        return self.resolve_scheme_color("accent1")
