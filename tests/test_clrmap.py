import pytest
from core.clrmap.resolver import ClrMapResolver

TEMPLATE_PATH = "templates/2026 se template eng.pptx"


def test_master3_bg1_resolves_to_dark_green():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=2)  # 0-indexed, Master 3
    result = resolver.resolve_scheme_color("bg1")
    assert result.upper() == "0A2F24"


def test_master3_tx1_resolves_to_white():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=2)
    result = resolver.resolve_scheme_color("tx1")
    assert result.upper() == "FFFFFF"


def test_master1_bg1_resolves_to_white():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=0)  # Master 1
    result = resolver.resolve_scheme_color("bg1")
    assert result.upper() == "FFFFFF"


def test_master1_tx1_resolves_to_dark_green():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=0)
    result = resolver.resolve_scheme_color("tx1")
    assert result.upper() == "0A2F24"


def test_master2_accent5_is_light_green():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=1)
    result = resolver.resolve_scheme_color("accent5")
    assert result.upper() == "E7FFD9"


def test_invalid_scheme_color_returns_none():
    resolver = ClrMapResolver(TEMPLATE_PATH, master_index=0)
    result = resolver.resolve_scheme_color("nonexistent")
    assert result is None
