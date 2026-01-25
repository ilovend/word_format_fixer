# -*- coding: utf-8 -*-
"""
规则参数 Schema 定义

提供统一的参数类型定义，用于：
1. 前端 UI 自动生成输入控件
2. 后端参数验证
3. 文档生成
"""

from enum import Enum
from typing import Any, Dict, List, Optional, Union


class ParamType(str, Enum):
    """参数类型枚举 - 决定前端渲染什么类型的控件"""
    STRING = "string"      # 文本输入框
    NUMBER = "number"      # 数字输入框
    INTEGER = "integer"    # 整数输入框
    BOOLEAN = "boolean"    # 开关按钮
    ENUM = "enum"          # 下拉选择框
    COLOR = "color"        # 颜色选择器
    FONT = "font"          # 字体选择器
    RANGE = "range"        # 滑块


class ParamSchema:
    """单个参数的 Schema 定义（不使用 pydantic，保持简单）"""
    
    def __init__(
        self,
        name: str,
        display_name: str,
        param_type: ParamType,
        default: Any = None,
        description: str = "",
        min_value: Optional[Union[int, float]] = None,
        max_value: Optional[Union[int, float]] = None,
        step: Optional[Union[int, float]] = None,
        options: Optional[List[Dict[str, Any]]] = None,
        placeholder: Optional[str] = None,
        unit: Optional[str] = None
    ):
        self.name = name
        self.display_name = display_name
        self.param_type = param_type.value if isinstance(param_type, ParamType) else param_type
        self.default = default
        self.description = description
        self.min_value = min_value
        self.max_value = max_value
        self.step = step
        self.options = options
        self.placeholder = placeholder
        self.unit = unit
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典，排除 None 值"""
        result = {
            "name": self.name,
            "display_name": self.display_name,
            "param_type": self.param_type,
            "default": self.default,
            "description": self.description,
        }
        
        if self.min_value is not None:
            result["min_value"] = self.min_value
        if self.max_value is not None:
            result["max_value"] = self.max_value
        if self.step is not None:
            result["step"] = self.step
        if self.options is not None:
            result["options"] = self.options
        if self.placeholder is not None:
            result["placeholder"] = self.placeholder
        if self.unit is not None:
            result["unit"] = self.unit
            
        return result


class RuleConfigSchema:
    """规则配置 Schema - 包含多个参数定义"""
    
    def __init__(self, params: List[ParamSchema] = None):
        self.params = params or []
    
    def to_ui_schema(self) -> List[Dict[str, Any]]:
        """转换为前端 UI 可用的格式"""
        return [param.to_dict() for param in self.params]
    
    def get_defaults(self) -> Dict[str, Any]:
        """获取所有参数的默认值"""
        return {param.name: param.default for param in self.params}


# ============== 便捷的参数工厂函数 ==============

def FontParam(
    name: str,
    display_name: str,
    default: str = "宋体",
    description: str = "字体名称"
) -> ParamSchema:
    """创建字体参数"""
    return ParamSchema(
        name=name,
        display_name=display_name,
        param_type=ParamType.FONT,
        default=default,
        description=description,
        options=[
            {"value": "宋体", "label": "宋体"},
            {"value": "黑体", "label": "黑体"},
            {"value": "楷体", "label": "楷体"},
            {"value": "仿宋", "label": "仿宋"},
            {"value": "微软雅黑", "label": "微软雅黑"},
            {"value": "Arial", "label": "Arial"},
            {"value": "Times New Roman", "label": "Times New Roman"},
        ]
    )


def SizeParam(
    name: str,
    display_name: str,
    default: Union[int, float] = 12,
    min_value: Union[int, float] = 8,
    max_value: Union[int, float] = 72,
    step: Union[int, float] = 0.5,
    unit: str = "pt",
    description: str = "字号大小"
) -> ParamSchema:
    """创建字号参数"""
    return ParamSchema(
        name=name,
        display_name=display_name,
        param_type=ParamType.RANGE,
        default=default,
        min_value=min_value,
        max_value=max_value,
        step=step,
        unit=unit,
        description=description
    )


def ColorParam(
    name: str,
    display_name: str,
    default: str = "#000000",
    description: str = "颜色值"
) -> ParamSchema:
    """创建颜色参数"""
    return ParamSchema(
        name=name,
        display_name=display_name,
        param_type=ParamType.COLOR,
        default=default,
        description=description
    )


def BoolParam(
    name: str,
    display_name: str,
    default: bool = True,
    description: str = ""
) -> ParamSchema:
    """创建布尔参数"""
    return ParamSchema(
        name=name,
        display_name=display_name,
        param_type=ParamType.BOOLEAN,
        default=default,
        description=description
    )


def EnumParam(
    name: str,
    display_name: str,
    options: List[Dict[str, str]],
    default: str = None,
    description: str = ""
) -> ParamSchema:
    """创建枚举参数"""
    return ParamSchema(
        name=name,
        display_name=display_name,
        param_type=ParamType.ENUM,
        default=default or (options[0]["value"] if options else None),
        options=options,
        description=description
    )


def RangeParam(
    name: str,
    display_name: str,
    default: Union[int, float] = 0,
    min_value: Union[int, float] = 0,
    max_value: Union[int, float] = 100,
    step: Union[int, float] = 1,
    unit: str = "",
    description: str = ""
) -> ParamSchema:
    """创建范围参数（滑块）"""
    return ParamSchema(
        name=name,
        display_name=display_name,
        param_type=ParamType.RANGE,
        default=default,
        min_value=min_value,
        max_value=max_value,
        step=step,
        unit=unit,
        description=description
    )
