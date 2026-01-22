from rules.base_rule import BaseRule, RuleResult
import re


class HorizontalRuleRemovalRule(BaseRule):
    """横线移除规则 - 移除从markdown通过pandoc转换到word时产生的横线"""

    display_name = "横线移除"
    category = "段落规则"
    
    def __init__(self, config=None):
        default_params = {
            'remove_horizontal_rules': True,
        }
        super().__init__({**default_params, **(config or {})})
    
    def apply(self, doc_context) -> RuleResult:
        """
        应用横线移除规则
        :param doc_context: 文档上下文对象
        :return: 规则执行结果
        """
        document = doc_context.get_document()
        fixed_count = 0
        details = []
        
        # 横线模式：匹配由连字符、星号或下划线组成的横线
        # 支持多种横线格式：---, *** , ___ , - - -, * * *, _ _ _ 等
        horizontal_rule_patterns = [
            r'^\s*[-]{3,}\s*$',       # --- 或更多连字符
            r'^\s*[*]{3,}\s*$',       # *** 或更多星号
            r'^\s*[_]{3,}\s*$',       # ___ 或更多下划线
            r'^\s*[-\s]{3,}\s*$',     # - - - 或带空格的连字符
            r'^\s*[*\s]{3,}\s*$',     # * * * 或带空格的星号
            r'^\s*[_\s]{3,}\s*$',     # _ _ _ 或带空格的下划线
        ]
        
        # 收集需要删除的段落
        paragraphs_to_delete = []
        
        for paragraph in document.paragraphs:
            text = paragraph.text
            
            # 检查是否匹配横线模式
            for pattern in horizontal_rule_patterns:
                if re.match(pattern, text):
                    paragraphs_to_delete.append(paragraph)
                    fixed_count += 1
                    break
        
        # 删除匹配的段落
        for paragraph in paragraphs_to_delete:
            # 获取段落的父元素并删除
            p_element = paragraph._element
            p_element.getparent().remove(p_element)
        
        if fixed_count > 0:
            details.append(f"移除了{fixed_count}个横线段落")
        else:
            details.append("文档中没有需要移除的横线")
        
        return RuleResult(
            rule_id=self.rule_id,
            success=True,
            fixed_count=fixed_count,
            details=details
        )
