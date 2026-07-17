"""
Template Registry Tests
测试模板注册表功能
"""

import pytest
from pathlib import Path

from core.registry.template_registry import TemplateRegistry


class TestTemplateRegistry:
    """TemplateRegistry 测试类"""

    def test_load_master_styles(self):
        """测试加载主样式"""
        registry = TemplateRegistry()

        # 验证加载了4种主样式
        assert len(registry.master_styles) == 4
        assert "F1" in registry.master_styles
        assert "F2" in registry.master_styles
        assert "F3" in registry.master_styles
        assert "F4" in registry.master_styles

        # 验证F1样式属性
        f1_style = registry.master_styles["F1"]
        assert f1_style["name"] == "白色简约"
        assert f1_style["background_color"] == "#FFFFFF"
        assert f1_style["text_color"] == "#333333"

        # 验证F2样式属性
        f2_style = registry.master_styles["F2"]
        assert f2_style["name"] == "浅绿色清新"
        assert f2_style["background_color"] == "#E8F5E9"

        # 验证F3样式属性
        f3_style = registry.master_styles["F3"]
        assert f3_style["name"] == "深绿色商务"
        assert f3_style["background_color"] == "#1B5E20"

        # 验证F4渐变样式
        f4_style = registry.master_styles["F4"]
        assert f4_style["name"] == "渐变科技"
        assert "background" in f4_style
        assert f4_style["background"]["type"] == "gradient"

    def test_get_layout_by_name(self):
        """测试通过名称获取布局"""
        registry = TemplateRegistry()

        # 测试获取封面布局
        cover_layout = registry.get_layout_by_name("F1", "cover")
        assert cover_layout is not None
        assert cover_layout["name"] == "封面"
        assert cover_layout["index"] == 0

        # 测试获取章节布局
        section_layout = registry.get_layout_by_name("F1", "section")
        assert section_layout is not None
        assert section_layout["name"] == "章节"
        assert section_layout["index"] == 1

        # 测试获取内容布局
        content_layout = registry.get_layout_by_name("F1", "content")
        assert content_layout is not None
        assert content_layout["name"] == "内容"
        assert content_layout["index"] == 2

        # 测试获取结尾布局
        closing_layout = registry.get_layout_by_name("F1", "closing")
        assert closing_layout is not None
        assert closing_layout["name"] == "结尾"
        assert closing_layout["index"] == 3

    def test_get_layout_by_index(self):
        """测试通过索引获取布局"""
        registry = TemplateRegistry()

        # 测试通过索引获取布局
        layout_0 = registry.get_layout_by_index("F1", 0)
        assert layout_0 is not None
        assert layout_0["name"] == "封面"

        layout_1 = registry.get_layout_by_index("F1", 1)
        assert layout_1 is not None
        assert layout_1["name"] == "章节"

        layout_2 = registry.get_layout_by_index("F1", 2)
        assert layout_2 is not None
        assert layout_2["name"] == "内容"

        layout_3 = registry.get_layout_by_index("F1", 3)
        assert layout_3 is not None
        assert layout_3["name"] == "结尾"

    def test_invalid_master_style(self):
        """测试无效的主样式ID"""
        registry = TemplateRegistry()

        # 测试不存在的主样式
        result = registry.get_layout_by_name("F99", "cover")
        assert result is None

    def test_invalid_layout_name(self):
        """测试无效的布局名称"""
        registry = TemplateRegistry()

        # 测试不存在的布局名称
        result = registry.get_layout_by_name("F1", "invalid")
        assert result is None