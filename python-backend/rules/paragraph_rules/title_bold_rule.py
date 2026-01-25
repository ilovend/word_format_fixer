# -*- coding: utf-8 -*-
"""标题加粗规则"""

from rules.base_rule import BaseRule, RuleResult
from schemas.rule_params import RuleConfigSchema, BoolParam


class TitleBoldRule(BaseRule):
    """标题加粗规则 - 所有标题文本加粗"""

    display_name = "标题加粗"
    category = "段落规则"
    description = "设置是否将所有标题文本加粗显示"
    
    # 参数 Schema 定义
    param_schema = RuleConfigSchema(params=[
        BoolParam(
            name="bold",
            display_name="加粗标题",
            default=True,
            description="开启后所有标题文本将加粗显示"
        ),
    ])

    def apply(self, doc_context) -> RuleResult:
        """
        应用标题加粗规则
        :param doc_context: 文档上下文对象
        """
        document = doc_context.get_document()
        fixed_count = 0
        details = []

        for paragraph in document.paragraphs:
            if paragraph.style.name.startswith('Heading'):
                for run in paragraph.runs:
                    run.font.bold = self.config['bold']
                    fixed_count += 1

        bold_text = "加粗" if self.config['bold'] else "取消加粗"
        details.append(f"标题文字{bold_text}")
        details.append(f"处理了 {fixed_count} 个标题文本运行")

        return RuleResult(
            rule_id=self.rule_id,
            success=True,
            fixed_count=fixed_count,
            details=details
        )
