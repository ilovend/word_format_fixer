from rules.base_rule import BaseRule, RuleResult
from docx.shared import Cm

class PageLayoutRule(BaseRule):
    """页面布局规则 - 设置文档页面布局，包括页面大小和边距"""

    display_name = "页面布局设置"
    category = "页面规则"

    def __init__(self, config=None):
        default_params = {
            'page_width_cm': 21.0,  # A4纸宽度
            'page_height_cm': 29.7,  # A4纸高度
            'page_margin_top_cm': 2.54,
            'page_margin_bottom_cm': 2.54,
            'page_margin_left_cm': 2.54,
            'page_margin_right_cm': 2.54,
        }
        super().__init__({**default_params, **(config or {})})

    def apply(self, doc_context) -> RuleResult:
        """
        核心执行逻辑
        :param doc_context: 文档上下文对象
        """
        document = doc_context.get_document()
        fixed_count = 0
        details = []

        for section in document.sections:
            # 确保页面大小已设置
            if section.page_width is None:
                section.page_width = Cm(self.config['page_width_cm'])
                fixed_count += 1
            if section.page_height is None:
                section.page_height = Cm(self.config['page_height_cm'])
                fixed_count += 1

            # 设置边距
            section.top_margin = Cm(self.config.get('page_margin_top_cm', 2.54))
            section.bottom_margin = Cm(self.config.get('page_margin_bottom_cm', 2.54))
            section.left_margin = Cm(self.config.get('page_margin_left_cm', 2.54))
            section.right_margin = Cm(self.config.get('page_margin_right_cm', 2.54))

            details.append(f"设置页面布局: 宽{section.page_width.cm:.1f}cm × 高{section.page_height.cm:.1f}cm")
            details.append(f"边距: 上{section.top_margin.cm:.1f}cm, 下{section.bottom_margin.cm:.1f}cm, 左{section.left_margin.cm:.1f}cm, 右{section.right_margin.cm:.1f}cm")

        return RuleResult(
            rule_id=self.rule_id,
            success=True,
            fixed_count=fixed_count,
            details=details
        )
