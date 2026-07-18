from pathlib import Path
from io import BytesIO
from pptx import Presentation
from pptx.util import Emu
from core.models import SlideContentModel, UserConfig
from core.registry.template_registry import TemplateRegistry


class ContentReplayer:
    """内容重放器：将抽取的内容模型重放到目标模板"""

    def __init__(self, registry: TemplateRegistry, template_path: str | None = None):
        self.registry = registry
        self.template_path = template_path

    def replay(self, content_models: list[SlideContentModel], config: UserConfig) -> str:
        if self.template_path and Path(self.template_path).exists():
            prs = Presentation(self.template_path)
            self._clear_slides(prs)
        else:
            prs = Presentation()

        selected_master_index = int(config.master_style) if config.master_style.isdigit() else 0

        if selected_master_index >= len(prs.slide_masters):
            selected_master_index = 0

        selected_master = prs.slide_masters[selected_master_index]

        layout_map = self._build_layout_map_from_master(selected_master)

        for model in content_models:
            layout_name = self._determine_layout(model)
            layout_index = self._resolve_layout_index(layout_name, layout_map, config.master_style)

            slide_layouts = list(selected_master.slide_layouts)
            if layout_index >= len(slide_layouts):
                layout_index = 0

            slide = prs.slides.add_slide(slide_layouts[layout_index])

            if model.title:
                try:
                    if slide.shapes.title is not None:
                        slide.shapes.title.text = model.title
                except Exception:
                    pass

            self._fill_body(slide, model.body_blocks)

        prs.save(config.output_path)
        return config.output_path

    def _clear_slides(self, prs) -> None:
        """清空演示文稿中的所有幻灯片，保留 master 和 layouts"""
        sldIdLst = prs.slides._sldIdLst
        rIds = [sldId.rId for sldId in list(sldIdLst)]
        for rId in rIds:
            try:
                prs.part.drop_rel(rId)
            except KeyError:
                pass
        for child in list(sldIdLst):
            sldIdLst.remove(child)

    def _build_layout_map_from_master(self, master) -> dict:
        """扫描 master 的所有 slide_layouts，构建 {名称小写: 索引} 映射"""
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
        self, layout_name: str, layout_map: dict, master_style: str
    ) -> int:
        """根据 layout 名称解析索引：先查模板实际布局，再查配置，最后回退 0"""
        if layout_name in layout_map:
            return layout_map[layout_name]

        layout_info = self.registry.get_layout_by_name(master_style, layout_name)
        if layout_info:
            idx = layout_info.get("index", 0)
            if layout_map and idx < len(layout_map):
                return idx

        if layout_name == "cover":
            return 0
        return min(1, max(layout_map.values()) if layout_map else 0)

    def _fill_body(self, slide, body_blocks) -> None:
        """填充正文：文本进占位符，表格和图片新增到 slide"""
        if not body_blocks:
            return

        # 分离文本块与其他类型
        text_blocks = [b for b in body_blocks if b.type == "paragraph" and b.text]
        other_blocks = [b for b in body_blocks if b.type in ("table", "image")]

        # 1. 文本块填充到正文占位符
        body_placeholder = None
        for shape in slide.placeholders:
            try:
                if shape.placeholder_format.idx == 1:
                    body_placeholder = shape
                    break
            except Exception:
                continue

        if body_placeholder and text_blocks:
            tf = body_placeholder.text_frame
            tf.clear()
            first = True
            for block in text_blocks:
                if first:
                    p = tf.paragraphs[0]
                    first = False
                else:
                    p = tf.add_paragraph()
                p.text = block.text
                p.level = block.level

        # 2. 表格和图片直接添加到 slide
        for block in other_blocks:
            if block.type == "image":
                self._add_image(slide, block.content)
            elif block.type == "table":
                self._add_table(slide, block.content)

    def _add_image(self, slide, content: dict) -> None:
        """将图片添加到 slide"""
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
        """将表格添加到 slide"""
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

    def _determine_layout(self, model: SlideContentModel) -> str:
        """根据内容模型推断布局类型"""
        if model.original_layout_type:
            return model.original_layout_type
        if model.slide_index == 0:
            return "cover"
        if model.slide_index == len(model.body_blocks) - 1:
            return "closing"
        return "content"
