from pathlib import Path
from pptx import Presentation
from core.models import SlideContentModel, UserConfig
from core.registry.template_registry import TemplateRegistry


class ContentReplayer:
    def __init__(self, registry: TemplateRegistry, template_path: str | None = None):
        self.registry = registry
        self.template_path = template_path

    def replay(self, content_models: list[SlideContentModel], config: UserConfig) -> str:
        if self.template_path and Path(self.template_path).exists():
            prs = Presentation(self.template_path)
            for slide in list(prs.slides):
                rId = prs.slides._sldIdLst[0].rId
                prs.part.drop_rel(rId)
                prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])
        else:
            prs = Presentation()

        for model in content_models:
            layout_name = self._determine_layout(model)
            layout_info = self.registry.get_layout_by_name(config.master_style, layout_name)

            if layout_info:
                layout_index = layout_info.get("index", 0)
            else:
                layout_index = 0

            slide = prs.slides.add_slide(prs.slide_layouts[layout_index])

            if model.title:
                if slide.shapes.title:
                    slide.shapes.title.text = model.title

            body_placeholder = None
            for shape in slide.placeholders:
                if shape.placeholder_format.idx == 1:
                    body_placeholder = shape
                    break

            if body_placeholder and model.body_blocks:
                tf = body_placeholder.text_frame
                tf.clear()
                for block in model.body_blocks:
                    if block.text:
                        p = tf.add_paragraph()
                        p.text = block.text
                        p.level = block.level

        prs.save(config.output_path)
        return config.output_path

    def _determine_layout(self, model: SlideContentModel) -> str:
        if model.original_layout_type:
            return model.original_layout_type
        if model.slide_index == 0:
            return "cover"
        if model.slide_index == len(model.body_blocks) - 1:
            return "closing"
        return "content"