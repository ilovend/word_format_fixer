"""配置管理模块"""

import yaml
import os
from typing import Dict, Optional


def load_config(config_path: Optional[str] = None) -> Dict:
    """
    加载配置文件
    
    Args:
        config_path: 配置文件路径，如果为None则使用默认配置
    
    Returns:
        配置字典
    """
    if config_path and os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        return config
    else:
        # 返回默认配置
        return {
            # 字体设置
            'chinese_font': '宋体',
            'western_font': 'Arial',
            'title_font': '黑体',
            
            # 字号设置
            'font_size_body': 12,
            'font_size_title1': 22,
            'font_size_title2': 18,
            'font_size_title3': 16,
            'font_size_table_header': 14,
            'font_size_table_content': 12,
            
            # 颜色设置
            'text_color': (0, 0, 0),
            'table_header_bg_color': (217, 217, 217),  # 表头背景色
            
            # 页面设置
            'page_width_cm': 21.0,  # A4纸宽度
            'page_height_cm': 29.7,  # A4纸高度
            'page_margin_top_cm': 2.54,
            'page_margin_bottom_cm': 2.54,
            'page_margin_left_cm': 2.54,
            'page_margin_right_cm': 2.54,
            
            # 表格设置
            'table_width_percent': 95,  # 表格宽度占页面宽度的百分比
            'table_alignment': 1,  # CENTER
            'cell_margin_top_cm': 0.1,
            'cell_margin_bottom_cm': 0.1,
            'cell_margin_left_cm': 0.2,
            'cell_margin_right_cm': 0.2,
            'column_widths': [],  # 自定义列宽百分比
            
            # 修复选项
            'fix_table_width': True,
            'adjust_table_layout': True,
            'add_table_header_format': True,
            'center_vertically': False,
            'auto_adjust_columns': True,
        }


def save_config(config: Dict, config_path: str) -> None:
    """
    保存配置文件
    
    Args:
        config: 配置字典
        config_path: 配置文件路径
    """
    with open(config_path, 'w', encoding='utf-8') as f:
        yaml.dump(config, f, default_flow_style=False, allow_unicode=True)


def get_preset_config(preset_name: str) -> Dict:
    """
    获取预设配置
    
    Args:
        preset_name: 预设名称
    
    Returns:
        配置字典
    """
    presets = {
        'default': {
            'chinese_font': '宋体',
            'western_font': 'Arial',
            'title_font': '黑体',
            'font_size_body': 12,
            'font_size_title1': 22,
            'font_size_title2': 18,
            'font_size_title3': 16,
        },
        'bid_document': {
            'chinese_font': '宋体',
            'western_font': 'Arial',
            'title_font': '黑体',
            'font_size_body': 12,
            'font_size_title1': 24,
            'font_size_title2': 20,
            'font_size_title3': 16,
            'page_margin_left_cm': 3.0,
            'page_margin_right_cm': 2.0,
        },
        'compact': {
            'chinese_font': '宋体',
            'western_font': 'Arial',
            'title_font': '黑体',
            'font_size_body': 11,
            'font_size_title1': 18,
            'font_size_title2': 16,
            'font_size_title3': 14,
            'page_margin_top_cm': 2.0,
            'page_margin_bottom_cm': 2.0,
            'page_margin_left_cm': 2.0,
            'page_margin_right_cm': 2.0,
        },
        'print_ready': {
            'chinese_font': '宋体',
            'western_font': 'Times New Roman',
            'title_font': '黑体',
            'font_size_body': 12,
            'font_size_title1': 22,
            'font_size_title2': 18,
            'font_size_title3': 16,
            'table_width_percent': 100,
        },
    }
    
    return presets.get(preset_name, presets['default'])
