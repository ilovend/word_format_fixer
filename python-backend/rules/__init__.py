# Rules package initialization
# Note: Rules are dynamically loaded by RuleEngine from the rules/ subdirectories.
# This file only provides convenient imports for direct usage if needed.

# 字体规则
from .font_rules.font_color_rule import FontColorRule
from .font_rules.font_standard_rule import FontNameRule, TitleFontRule, FontSizeRule

# 段落规则
from .paragraph_rules.paragraph_spacing_rule import ParagraphSpacingRule
from .paragraph_rules.title_bold_rule import TitleBoldRule
from .paragraph_rules.title_alignment_rule import TitleAlignmentRule
from .paragraph_rules.list_numbering_rule import ListNumberingRule
from .paragraph_rules.horizontal_rule_removal_rule import HorizontalRuleRemovalRule

# 表格规则
from .table_rules.table_width_rule import TableWidthRule
from .table_rules.table_border_rule import TableBorderRule
from .table_rules.table_borders_rule import TableBordersRule

# 页面规则
from .page_rules.page_layout_rule import PageLayoutRule

__all__ = [
    # 字体规则
    'FontColorRule',
    'FontNameRule',
    'TitleFontRule',
    'FontSizeRule',
    # 段落规则
    'ParagraphSpacingRule',
    'TitleBoldRule',
    'TitleAlignmentRule',
    'ListNumberingRule',
    'HorizontalRuleRemovalRule',
    # 表格规则
    'TableWidthRule',
    'TableBorderRule',
    'TableBordersRule',
    # 页面规则
    'PageLayoutRule',
]
