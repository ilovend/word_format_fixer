from rules.base_rule import BaseRule, RuleResult
from docx.enum.text import WD_ALIGN_PARAGRAPH

class TitleAlignmentRule(BaseRule):
    """标题对齐规则 - 一级标题居中，其他标题左对齐"""

    display_name = "标题对齐设置"
    category = "段落规则"

    def __init__(self, config=None):
        default_params = {
            'heading1_align': 'center',  # 一级标题对齐方式
            'other_heading_align': 'left',  # 其他标题对齐方式
        }
        super().__init__({**default_params, **(config or {})})

    def apply(self, doc_context) -> RuleResult:
        """
        应用标题对齐规则
        :param doc_context: 文档上下文对象
        """
        document = doc_context.get_document()
        fixed_count = 0
        details = []

        # 对齐方式映射
        align_map = {
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
        }

        heading1_align = align_map.get(self.config['heading1_align'], WD_ALIGN_PARAGRAPH.CENTER)
        other_heading_align = align_map.get(self.config['other_heading_align'], WD_ALIGN_PARAGRAPH.LEFT)

        for paragraph in document.paragraphs:
            if paragraph.style.name.startswith('Heading'):
                if paragraph.style.name == 'Heading 1':
                    paragraph.alignment = heading1_align
                else:
                    paragraph.alignment = other_heading_align
                fixed_count += 1

        details.append(f"一级标题对齐: {self.config['heading1_align']}")
        details.append(f"其他标题对齐: {self.config['other_heading_align']}")
        details.append(f"调整了 {fixed_count} 个标题的对齐方式")

        return RuleResult(
            rule_id=self.rule_id,
            success=True,
            fixed_count=fixed_count,
            details=details
        )
