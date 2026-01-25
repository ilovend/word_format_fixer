from typing import Dict, Any, List
from core.engine import RuleEngine
from core.config_loader import ConfigLoader


class ServiceContainer:
    """
    服务容器 - 单例模式管理共享资源
    确保 RuleEngine 等核心组件只创建一次，避免重复初始化开销
    """
    _instance = None
    _engine = None
    _config_loader = None
    
    @classmethod
    def get_instance(cls) -> 'ServiceContainer':
        """获取单例实例"""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance
    
    @classmethod
    def get_engine(cls) -> RuleEngine:
        """获取共享的规则引擎实例"""
        if cls._engine is None:
            cls._engine = RuleEngine()
        return cls._engine
    
    @classmethod
    def get_config_loader(cls) -> ConfigLoader:
        """获取共享的配置加载器实例"""
        if cls._config_loader is None:
            cls._config_loader = ConfigLoader()
        return cls._config_loader


class DocumentProcessingService:
    """文档处理服务 - 封装文档处理相关的业务逻辑"""

    def __init__(self):
        # 使用共享的规则引擎实例
        self.engine = ServiceContainer.get_engine()

    def process_document(self, document_path: str, active_rules: List[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        处理文档
        :param document_path: 文档路径
        :param active_rules: 激活的规则列表
        :return: 处理结果
        """
        if not document_path:
            raise ValueError("Missing document_path")

        # 调用规则引擎执行规则
        result = self.engine.execute(document_path, active_rules)
        return result


class ConfigManagementService:
    """配置管理服务 - 封装配置管理相关的业务逻辑"""

    def __init__(self):
        # 使用共享的配置加载器实例
        self.config_loader = ServiceContainer.get_config_loader()

    def get_all_presets(self) -> Dict[str, Any]:
        """
        获取所有预设
        :return: 预设字典
        """
        return self.config_loader.get_all_presets()

    def save_preset(self, preset_id: str, preset_data: Dict[str, Any]) -> Dict[str, str]:
        """
        保存预设
        :param preset_id: 预设ID
        :param preset_data: 预设数据
        :return: 操作结果
        """
        # 业务规则验证
        if not preset_id:
            raise ValueError("Missing preset_id")

        if not preset_data:
            raise ValueError("Missing preset_data")

        # 调用持久化层保存
        self.config_loader.save_preset(preset_id, preset_data)
        return {"status": "success", "message": "Preset saved successfully"}

    def delete_preset(self, preset_id: str) -> Dict[str, str]:
        """
        删除预设
        :param preset_id: 预设ID
        :return: 操作结果
        """
        # 业务规则验证
        if not preset_id:
            raise ValueError("Missing preset_id")

        # 业务规则：不允许删除默认预设
        if preset_id == 'default':
            raise ValueError("Cannot delete default preset")

        # 调用持久化层删除
        self.config_loader.delete_preset(preset_id)
        return {"status": "success", "message": "Preset deleted successfully"}


class RuleManagementService:
    """规则管理服务 - 封装规则管理相关的业务逻辑"""

    def __init__(self):
        # 使用共享的规则引擎实例
        self.engine = ServiceContainer.get_engine()

    def get_all_rules(self) -> List[Dict[str, Any]]:
        """
        获取所有规则信息
        :return: 规则信息列表
        """
        return self.engine.get_rules_info()

    def update_rule_config(self, rule_id: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        更新单个规则的配置
        :param rule_id: 规则ID
        :param params: 新的配置参数
        :return: 操作结果
        """
        if not rule_id:
            return {"rule_id": rule_id, "success": False, "error": "Missing rule_id"}
        
        rule = self.engine.get_rule_by_id(rule_id)
        if not rule:
            return {"rule_id": rule_id, "success": False, "error": f"Rule not found: {rule_id}"}
        
        errors = rule.update_config(params)
        if errors:
            return {"rule_id": rule_id, "success": False, "errors": errors}
        return {"rule_id": rule_id, "success": True}
