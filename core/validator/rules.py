import yaml
import re
from pathlib import Path
from pptx import Presentation
from core.models import ValidationIssue


class ValidationRule:
    def __init__(self, rule_id: str, name: str, level: str, description: str):
        self.rule_id = rule_id
        self.name = name
        self.level = level
        self.description = description

    def check(self, pptx_path: str) -> list[ValidationIssue]:
        raise NotImplementedError


class WatermarkTextRule(ValidationRule):
    def __init__(self):
        super().__init__("R040", "水印文本检查", "fail", "产物中不得存在已知水印文本")
        self.keywords = []
        self._load_keywords()

    def _load_keywords(self):
        config_path = Path(__file__).parent.parent.parent / "config" / "validation_rules.yaml"
        if config_path.exists():
            with open(config_path, "r", encoding="utf-8") as f:
                config = yaml.safe_load(f)
                self.keywords = config.get("watermark_keywords", [])

    def check(self, pptx_path: str) -> list[ValidationIssue]:
        issues = []
        prs = Presentation(pptx_path)

        for slide_idx, slide in enumerate(prs.slides):
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text = shape.text
                    for keyword in self.keywords:
                        if keyword in text:
                            issues.append(ValidationIssue(
                                level=self.level,
                                rule_id=self.rule_id,
                                message=f"发现水印文本: {keyword}",
                                slide_index=slide_idx
                            ))

        return issues


class FontWhitelistRule(ValidationRule):
    def __init__(self):
        super().__init__("R010", "字体白名单检查", "fail", "仅允许使用模板主题字体")
        self.font_whitelist = []
        self._load_fonts()

    def _load_fonts(self):
        config_path = Path(__file__).parent.parent.parent / "config" / "validation_rules.yaml"
        if config_path.exists():
            with open(config_path, "r", encoding="utf-8") as f:
                config = yaml.safe_load(f)
                self.font_whitelist = config.get("font_whitelist", [])

    def check(self, pptx_path: str) -> list[ValidationIssue]:
        issues = []
        prs = Presentation(pptx_path)

        for slide_idx, slide in enumerate(prs.slides):
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            font_name = run.font.name or "Unknown"
                            if font_name not in self.font_whitelist:
                                issues.append(ValidationIssue(
                                    level=self.level,
                                    rule_id=self.rule_id,
                                    message=f"使用非白名单字体: {font_name}",
                                    slide_index=slide_idx
                                ))

        return issues