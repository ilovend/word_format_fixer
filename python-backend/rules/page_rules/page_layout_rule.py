# -*- coding: utf-8 -*-
"""页面布局规则"""

from rules.base_rule import BaseRule, RuleResult
from docx.shared import Cm
from schemas.rule_params import (
    RuleConfigSchema,
    RangeParam,
    EnumParam
)


class PageLayoutRule(BaseRule):
    """页面布局规则 - 设置文档页面布局，包括页面大小和边距"""

    display_name = "页面布局设置"
    category = "页面规则"
    description = "设置文档的页面大小、边距等布局参数"
    
    # 参数 Schema 定义
    param_schema = RuleConfigSchema(params=[
        EnumParam(
            name="page_size",
            display_name="纸张大小",
            options=[
                {"value": "a4", "label": "A4 (21.0 × 29.7 cm)"},
                {"value": "letter", "label": "Letter (21.6 × 27.9 cm)"},
                {"value": "a3", "label": "A3 (29.7 × 42.0 cm)"},
                {"value": "b5", "label": "B5 (17.6 × 25.0 cm)"},
                {"value": "custom", "label": "自定义"},
            ],
            default="a4",
            description="选择标准纸张大小或自定义"
        ),
        RangeParam(
            name="page_width_cm",
            display_name="页面宽度",
            default=21.0,
            min_value=10.0,
            max_value=50.0,
            step=0.1,
            unit="cm",
            description="页面宽度（自定义时使用）"
        ),
        RangeParam(
            name="page_height_cm",
            display_name="页面高度",
            default=29.7,
            min_value=10.0,
            max_value=100.0,
            step=0.1,
            unit="cm",
            description="页面高度（自定义时使用）"
        ),
        RangeParam(
            name="page_margin_top_cm",
            display_name="上边距",
            default=2.54,
            min_value=0.5,
            max_value=5.0,
            step=0.1,
            unit="cm",
            description="页面上边距"
        ),
        RangeParam(
            name="page_margin_bottom_cm",
            display_name="下边距",
            default=2.54,
            min_value=0.5,
            max_value=5.0,
            step=0.1,
            unit="cm",
            description="页面下边距"
        ),
        RangeParam(
            name="page_margin_left_cm",
            display_name="左边距",
            default=2.54,
            min_value=0.5,
            max_value=5.0,
            step=0.1,
            unit="cm",
            description="页面左边距"
        ),
        RangeParam(
            name="page_margin_right_cm",
            display_name="右边距",
            default=2.54,
            min_value=0.5,
            max_value=5.0,
            step=0.1,
            unit="cm",
            description="页面右边距"
        ),
    ])
    
    # 预设纸张大小
    PAGE_SIZES = {
        'a4': (21.0, 29.7),
        'letter': (21.6, 27.9),
        'a3': (29.7, 42.0),
        'b5': (17.6, 25.0),
    }

    def apply(self, doc_context) -> RuleResult:
        """
        核心执行逻辑
        :param doc_context: 文档上下文对象
        """
        document = doc_context.get_document()
        fixed_count = 0
        details = []
        
        # 获取页面尺寸
        page_size = self.config.get('page_size', 'a4')
        if page_size in self.PAGE_SIZES:
            page_width, page_height = self.PAGE_SIZES[page_size]
        else:
            page_width = self.config.get('page_width_cm', 21.0)
            page_height = self.config.get('page_height_cm', 29.7)

        for section in document.sections:
            # 设置页面大小
            section.page_width = Cm(page_width)
            section.page_height = Cm(page_height)
            fixed_count += 1

            # 设置边距
            section.top_margin = Cm(self.config.get('page_margin_top_cm', 2.54))
            section.bottom_margin = Cm(self.config.get('page_margin_bottom_cm', 2.54))
            section.left_margin = Cm(self.config.get('page_margin_left_cm', 2.54))
            section.right_margin = Cm(self.config.get('page_margin_right_cm', 2.54))

            details.append(f"页面大小: {page_width:.1f}cm × {page_height:.1f}cm")
            details.append(f"边距: 上{section.top_margin.cm:.1f}cm, 下{section.bottom_margin.cm:.1f}cm, 左{section.left_margin.cm:.1f}cm, 右{section.right_margin.cm:.1f}cm")

        return RuleResult(
            rule_id=self.rule_id,
            success=True,
            fixed_count=fixed_count,
            details=details
        )
