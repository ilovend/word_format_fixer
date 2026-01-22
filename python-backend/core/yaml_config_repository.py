import yaml
import os
from typing import Dict, Any, Optional
from core.config_repository import IConfigRepository

class YamlConfigRepository(IConfigRepository):
    """YAML文件配置仓库 - 实现配置持久化"""

    def __init__(self, config_path: str = None):
        self.config_path = config_path or os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'config', 'presets.yaml'
        )

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
