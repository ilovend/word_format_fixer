# Rules package initialization

# 字体规则
from .font_rules.font_color_rule import FontColorRule
from .font_rules.font_standard_rule import FontNameRule, TitleFontRule, FontSizeRule

# 段落规则
from .paragraph_rules.paragraph_spacing_rule import ParagraphSpacingRule
from .paragraph_rules.title_bold_rule import TitleBoldRule
from .paragraph_rules.title_alignment_rule import TitleAlignmentRule
from .paragraph_rules.list_numbering_rule import ListNumberingRule

# 表格规则
from .table_rules.table_width_rule import TableWidthRule
from .table_rules.table_border_rule import TableBorderRule

# 页面规则
from .page_rules.page_layout_rule import PageLayoutRule

# 规则注册表
RULES_REGISTRY = {
    # 字体规则
    'font_color': FontColorRule(),
    'font_name': FontNameRule(),
    'title_font': TitleFontRule(),
    'font_size': FontSizeRule(),
    
    # 段落规则
    'paragraph_spacing': ParagraphSpacingRule(),
    'title_bold': TitleBoldRule(),
    'title_alignment': TitleAlignmentRule(),
    'list_numbering': ListNumberingRule(),
    
    # 表格规则
    'table_width': TableWidthRule(),
    'table_border': TableBorderRule(),
    
    # 页面规则
    'page_layout': PageLayoutRule(),
}

# 获取所有规则
def get_all_rules():
    """获取所有规则实例"""
    return list(RULES_REGISTRY.values())

# 根据规则ID获取规则
def get_rule(rule_id):
    """根据规则ID获取规则实例"""
    return RULES_REGISTRY.get(rule_id)

# 获取启用的规则
def get_enabled_rules():
    """获取所有启用的规则"""
    return [rule for rule in RULES_REGISTRY.values() if rule.enabled]
