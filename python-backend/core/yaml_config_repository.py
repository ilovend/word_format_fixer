import yaml
import os
from typing import Dict, Any, Optional
from core.config_repository import IConfigRepository

import sys

class YamlConfigRepository(IConfigRepository):
    """YAML文件配置仓库 - 实现配置持久化"""

    def __init__(self, config_path: str = None):
        if config_path:
            self.config_path = config_path
        else:
            # Determine base directory
            if getattr(sys, 'frozen', False):
                # If packaged, use the directory of the executable
                base_dir = os.path.dirname(sys.executable)
            else:
                # If running from source, use the project root (relative to this file)
                base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
            self.config_path = os.path.join(base_dir, 'config', 'presets.yaml')

    def load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        with open(self.config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}

    def save_config(self, config: Dict[str, Any]) -> None:
        """保存配置文件"""
        with open(self.config_path, 'w', encoding='utf-8') as f:
            yaml.dump(config, f, allow_unicode=True, default_flow_style=False)

    def get_preset(self, preset_name: str) -> Optional[Dict[str, Any]]:
        """获取指定预设的配置"""
        config = self.load_config()
        presets = config.get('presets', {})
        return presets.get(preset_name)

    def delete_preset(self, preset_id: str) -> None:
        """删除预设（纯数据操作）"""
        config = self.load_config()

        # 确保presets键存在
        if 'presets' in config and preset_id in config['presets']:
            del config['presets'][preset_id]
            self.save_config(config)

    def save_preset(self, preset_id: str, preset_data: Dict[str, Any]) -> None:
        """保存预设（纯数据操作）"""
        config = self.load_config()

        # 确保presets键存在
        if 'presets' not in config:
            config['presets'] = {}

        # 保存预设
        config['presets'][preset_id] = preset_data
        self.save_config(config)
