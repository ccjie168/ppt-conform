"""
Template Registry Module
模板注册表，用于管理主样式和布局映射
"""

import yaml
from pathlib import Path
from typing import Dict, Any, Optional


class TemplateRegistry:
    """模板注册表类，管理主样式和布局映射"""

    def __init__(self, config_dir: Optional[Path] = None):
        """
        初始化模板注册表

        Args:
            config_dir: 配置文件目录路径，默认为 config/
        """
        if config_dir is None:
            config_dir = Path(__file__).parent.parent.parent / "config"

        self.config_dir = Path(config_dir)
        self.master_styles: Dict[str, Dict[str, Any]] = {}
        self.layout_mappings: Dict[str, Dict[str, Dict[str, Any]]] = {}

        # 加载配置
        self._load_master_styles()
        self._load_layout_mappings()

    def _load_master_styles(self) -> None:
        """从YAML文件加载主样式配置"""
        master_styles_path = self.config_dir / "master_styles.yaml"

        if not master_styles_path.exists():
            raise FileNotFoundError(
                f"Master styles config file not found: {master_styles_path}"
            )

        with open(master_styles_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)

        self.master_styles = data.get("master_styles", {})

    def _load_layout_mappings(self) -> None:
        """从YAML文件加载布局映射配置"""
        layout_mappings_path = self.config_dir / "layout_mappings.yaml"

        if not layout_mappings_path.exists():
            raise FileNotFoundError(
                f"Layout mappings config file not found: {layout_mappings_path}"
            )

        with open(layout_mappings_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)

        self.layout_mappings = data.get("layouts", {})

    def get_layout_by_name(
        self, master_id: str, layout_name: str
    ) -> Optional[Dict[str, Any]]:
        """
        通过布局名称获取布局信息

        Args:
            master_id: 主样式ID (F1-F4)
            layout_name: 布局名称 (cover, section, content, closing)

        Returns:
            布局信息字典，如果不存在则返回None
        """
        if master_id not in self.layout_mappings:
            return None

        layouts = self.layout_mappings[master_id]
        return layouts.get(layout_name)

    def get_layout_by_index(
        self, master_id: str, layout_index: int
    ) -> Optional[Dict[str, Any]]:
        """
        通过布局索引获取布局信息

        Args:
            master_id: 主样式ID (F1-F4)
            layout_index: 布局索引 (0-3)

        Returns:
            布局信息字典，如果不存在则返回None
        """
        if master_id not in self.layout_mappings:
            return None

        layouts = self.layout_mappings[master_id]

        # 查找匹配的索引
        for layout_data in layouts.values():
            if layout_data.get("index") == layout_index:
                return layout_data

        return None

    def get_master_style(self, master_id: str) -> Optional[Dict[str, Any]]:
        """
        获取主样式信息

        Args:
            master_id: 主样式ID (F1-F4)

        Returns:
            主样式信息字典，如果不存在则返回None
        """
        return self.master_styles.get(master_id)