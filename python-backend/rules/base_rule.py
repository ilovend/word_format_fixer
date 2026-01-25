from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional, Type

from schemas.rule_params import RuleConfigSchema, ParamSchema


class RuleResult:
    """规则执行结果的标准化数据"""
    def __init__(self, rule_id: str, success: bool, fixed_count: int, details: List[str]):
        self.rule_id = rule_id
        self.success = success
        self.fixed_count = fixed_count
        self.details = details
    
    def dict(self):
        """返回字典表示"""
        return {
            "rule_id": self.rule_id,
            "success": self.success,
            "fixed_count": self.fixed_count,
            "details": self.details
        }


class BaseRule(ABC):
    """
    规则基类 - 所有规则必须继承
    
    子类需要定义：
    - display_name: 规则显示名称
    - category: 规则类别
    - param_schema: 参数 Schema 定义（RuleConfigSchema 实例）
    - apply(): 规则执行逻辑
    """
    
    # 子类需要覆盖的类属性
    display_name: str = "未命名规则"
    category: str = "其他规则"
    description: str = "无描述"
    param_schema: Optional[RuleConfigSchema] = None
    
    def __init__(self, config: Dict[str, Any] = None):
        # 使用 schema 的默认值初始化配置
        if self.param_schema:
            self.config = self.param_schema.get_defaults()
            # 用传入的 config 覆盖默认值
            if config:
                self.config.update(config)
        else:
            self.config = config or {}
        
        self.rule_id = self.__class__.__name__
        self.enabled = True
    
    @abstractmethod
    def apply(self, doc_context: Any) -> RuleResult:
        """
        核心执行逻辑
        :param doc_context: python-docx 的文档对象或其封装
        """
        pass
    
    def validate_config(self) -> List[str]:
        """
        验证当前配置是否有效
        :return: 错误信息列表，空列表表示验证通过
        """
        errors = []
        if not self.param_schema:
            return errors
        
        for param in self.param_schema.params:
            value = self.config.get(param.name)
            
            # 检查必填项
            if value is None and param.default is None:
                errors.append(f"参数 '{param.display_name}' 不能为空")
                continue
            
            # 检查数值范围
            if param.min_value is not None and value is not None:
                if value < param.min_value:
                    errors.append(f"参数 '{param.display_name}' 不能小于 {param.min_value}")
            
            if param.max_value is not None and value is not None:
                if value > param.max_value:
                    errors.append(f"参数 '{param.display_name}' 不能大于 {param.max_value}")
            
            # 检查枚举值
            if param.options is not None and value is not None:
                valid_values = [opt["value"] for opt in param.options]
                if value not in valid_values:
                    errors.append(f"参数 '{param.display_name}' 的值 '{value}' 不在有效选项中")
        
        return errors
    
    def get_metadata(self) -> Dict[str, Any]:
        """返回给前端渲染 UI 使用的元数据"""
        metadata = {
            "id": self.rule_id,
            "name": self.display_name,
            "description": self.description,
            "category": self.category,
            "enabled": self.enabled,
            "params": self.config,  # 当前配置值
        }
        
        # 添加 Schema 信息供前端动态渲染 UI
        if self.param_schema:
            metadata["param_schema"] = self.param_schema.to_ui_schema()
        
        return metadata
    
    def explain(self) -> str:
        """返回规则的可解释描述"""
        return self.description or self.__doc__ or "无描述"
    
    def update_config(self, new_config: Dict[str, Any]) -> List[str]:
        """
        更新配置并验证
        :param new_config: 新配置值
        :return: 验证错误列表，空列表表示成功
        """
        # 暂存旧配置
        old_config = self.config.copy()
        
        # 更新配置
        self.config.update(new_config)
        
        # 验证
        errors = self.validate_config()
        
        # 如果验证失败，回滚
        if errors:
            self.config = old_config
        
        return errors

