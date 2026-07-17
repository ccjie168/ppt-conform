import re
import yaml
from pathlib import Path
from core.models import WatermarkReport, WatermarkElement, WatermarkType


class WatermarkDetector:
    def __init__(self, config_path: str | None = None):
        self.keywords = []
        self.patterns = []
        self._load_config(config_path)

    def _load_config(self, config_path: str | None):
        if config_path is None:
            config_path = Path(__file__).parent.parent.parent / "config" / "watermark_blacklist.yaml"
        if Path(config_path).exists():
            with open(config_path, "r", encoding="utf-8") as f:
                config = yaml.safe_load(f)
                self.keywords = config.get("keywords", [])
                self.patterns = config.get("patterns", [])

    def detect_text(self, text: str, slide_index: int = 0) -> WatermarkReport:
        if not text:
            return WatermarkReport(detected=False, elements=[], summary="No text")

        # Check keywords first (prefer longest match)
        matched_keyword = None
        for keyword in self.keywords:
            if keyword in text:
                if matched_keyword is None or len(keyword) > len(matched_keyword):
                    matched_keyword = keyword

        if matched_keyword:
            return WatermarkReport(
                detected=True,
                elements=[WatermarkElement(
                    type=WatermarkType.TEXT,
                    slide_index=slide_index,
                    text_content=matched_keyword,
                    confidence=0.95
                )],
                summary=f"Found text watermark: {matched_keyword}"
            )

        # Check patterns only if no keyword matched
        for pattern in self.patterns:
            if re.match(pattern, text, re.IGNORECASE):
                return WatermarkReport(
                    detected=True,
                    elements=[WatermarkElement(
                        type=WatermarkType.TEXT,
                        slide_index=slide_index,
                        text_content=text[:100],
                        confidence=0.85
                    )],
                    summary=f"Found pattern watermark"
                )

        return WatermarkReport(detected=False, elements=[], summary="No watermarks detected")

    def detect_image_watermark(self, image_bytes: bytes) -> bool:
        try:
            from PIL import Image
            import io
            img = Image.open(io.BytesIO(image_bytes))
            if img.mode in ("RGBA", "LA"):
                alpha = img.getchannel("A")
                avg_alpha = alpha.mean()
                if 0 < avg_alpha < 200:
                    return True
            return False
        except Exception:
            return False