from typing import Dict, Any, Optional
from core.yaml_config_repository import YamlConfigRepository
from core.config_repository import IConfigRepository

class ConfigLoader:
    """配置加载器 - 负责配置的业务逻辑"""

    def __init__(self, repository: IConfigRepository = None):
        self.repository = repository or YamlConfigRepository()
        self.config = self.repository.load_config()

    def get_preset(self, preset_name: str) -> Optional[Dict[str, Any]]:
        """获取指定预设的配置"""
        presets = self.config.get('presets', {})
        return presets.get(preset_name)

    def get_all_presets(self) -> Dict[str, Any]:
        """获取所有预设"""
        return self.config.get('presets', {})

    def get_rule_defaults(self) -> Dict[str, Any]:
        """获取规则默认配置"""
        return self.config.get('rule_defaults', {})

    def load_preset_config(self, preset_name: str) -> Dict[str, Any]:
        """加载预设配置，合并默认配置"""
        preset = self.get_preset(preset_name)
        if not preset:
            return {}

        rule_defaults = self.get_rule_defaults()
        preset_rules = preset.get('rules', {})

        # 处理预设配置
        config = {}
        for rule_id, rule_config in preset_rules.items():
            if rule_config.get('enabled', True):
                # 获取规则默认配置
                default_config = rule_defaults.get(rule_id, {})
                # 获取预设特定配置
                preset_params = rule_config.get('parameters', {})
                # 合并配置
                final_config = {**default_config, **preset_params}
                # 移除enabled字段，因为它已经被处理
                final_config.pop('enabled', None)
                config[rule_id] = final_config

        return config

    def get_enabled_rules(self, preset_name: str) -> Dict[str, Dict[str, Any]]:
        """获取预设中启用的规则及其配置"""
        preset = self.get_preset(preset_name)
        if not preset:
            return {}

        enabled_rules = {}
        for rule_id, rule_config in preset.get('rules', {}).items():
            if rule_config.get('enabled', True):
                enabled_rules[rule_id] = rule_config.get('parameters', {})

        return enabled_rules

    def save_preset(self, preset_id: str, preset_data: Dict[str, Any]):
        """保存预设（委托给持久化层）"""
        self.repository.save_preset(preset_id, preset_data)
        # 更新内存中的配置
        self.config = self.repository.load_config()

    def delete_preset(self, preset_id: str):
        """删除预设（委托给持久化层）"""
        self.repository.delete_preset(preset_id)
        # 更新内存中的配置
        self.config = self.repository.load_config()
