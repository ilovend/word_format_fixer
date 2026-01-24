from abc import ABC, abstractmethod
from typing import Any, Dict, List

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
    """规则基类 - 所有规则必须继承"""
    
    def __init__(self, config: Dict[str, Any] = None):
        self.config = config or {}
        self.rule_id = self.__class__.__name__
        self.enabled = True
        self.display_name = getattr(self, "display_name", self.rule_id)
    
    @abstractmethod
    def apply(self, doc_context: Any) -> RuleResult:
        """
        核心执行逻辑
        :param doc_context: python-docx 的文档对象或其封装
        """
        pass
    
    def get_metadata(self):
        """返回给前端渲染 UI 使用的元数据"""
        return {
            "id": self.rule_id,
            "name": getattr(self, "display_name", self.rule_id),
            "description": self.__doc__ or "无描述",
            "category": getattr(self, "category", "其他规则"),  # 添加类别字段
            "params": self.config
        }
    
    def explain(self) -> str:
        """返回规则的可解释描述"""
        return self.__class__.__doc__ or "无描述"
