from rules.base_rule import BaseRule, RuleResult

class TitleBoldRule(BaseRule):
    """标题加粗规则 - 所有标题文本加粗"""

    display_name = "标题加粗"
    category = "段落规则"

    def __init__(self, config=None):
        default_params = {
            'bold': True,  # 是否加粗
        }
        super().__init__({**default_params, **(config or {})})

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
