# -*- coding: utf-8 -*-
"""标题对齐规则"""

from rules.base_rule import BaseRule, RuleResult
from docx.enum.text import WD_ALIGN_PARAGRAPH
from schemas.rule_params import RuleConfigSchema, EnumParam


class TitleAlignmentRule(BaseRule):
    """标题对齐规则 - 一级标题居中，其他标题左对齐"""

    display_name = "标题对齐设置"
    category = "段落规则"
    description = "分别设置一级标题和其他级别标题的对齐方式"
    
    # 参数 Schema 定义
    param_schema = RuleConfigSchema(params=[
        EnumParam(
            name="heading1_align",
            display_name="一级标题对齐",
            options=[
                {"value": "center", "label": "居中"},
                {"value": "left", "label": "左对齐"},
                {"value": "right", "label": "右对齐"},
                {"value": "justify", "label": "两端对齐"},
            ],
            default="center",
            description="一级标题的对齐方式"
        ),
        EnumParam(
            name="other_heading_align",
            display_name="其他标题对齐",
            options=[
                {"value": "left", "label": "左对齐"},
                {"value": "center", "label": "居中"},
                {"value": "right", "label": "右对齐"},
                {"value": "justify", "label": "两端对齐"},
            ],
            default="left",
            description="二级及以下标题的对齐方式"
        ),
    ])

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
