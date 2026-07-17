import pytest
from core.watermark.detector import WatermarkDetector
from core.models import WatermarkReport, WatermarkType


def test_load_blacklist():
    detector = WatermarkDetector()
    assert len(detector.keywords) > 0
    assert "TRAE AI 生成" in detector.keywords


def test_detect_text_watermark():
    detector = WatermarkDetector()
    report = detector.detect_text("TRAE AI 生成")
    assert report.detected is True
    assert len(report.elements) == 1
    assert report.elements[0].type == WatermarkType.TEXT
    assert report.elements[0].text_content == "TRAE AI 生成"


def test_no_watermark_text():
    detector = WatermarkDetector()
    report = detector.detect_text("这是正常的内容")
    assert report.detected is False
    assert len(report.elements) == 0