"""工具函数模块"""

import re
from typing import Dict, List


def detect_numbering_patterns() -> Dict[str, str]:
    """
    检测文档中的编号模式
    
    Returns:
        编号模式字典
    """
    patterns = {
        # 基本阿拉伯数字编号
        'arabic_with_dot': r'^\s*(\d+)\.\s+',  # 1. 项目
        'arabic_with_paren': r'^\s*(\d+)\)\s+',  # 1) 项目
        'arabic_with_hyphen': r'^\s*(\d+)-\s+',  # 1- 项目
        'arabic_with_colon': r'^\s*(\d+):\s+',  # 1: 项目
        'arabic_with_bracket': r'^\s*\[(\d+)\]\s+',  # [1] 项目
        
        # 中文数字编号
        'chinese_with_dot': r'^\s*([一二三四五六七八九十]+)\.\s+',  # 一. 项目
        'chinese_with_bracket': r'^\s*([一二三四五六七八九十]+)、',  # 一、项目
        'chinese_with_paren': r'^\s*\(([一二三四五六七八九十]+)\)\s+',  # (一) 项目
        
        # 字母编号
        'lower_alpha_with_dot': r'^\s*([a-z])\.\s+',  # a. 项目
        'upper_alpha_with_dot': r'^\s*([A-Z])\.\s+',  # A. 项目
        'lower_alpha_with_paren': r'^\s*([a-z])\)\s+',  # a) 项目
        'upper_alpha_with_paren': r'^\s*([A-Z])\)\s+',  # A) 项目
        
        # 罗马数字编号
        'lower_roman_with_dot': r'^\s*([ivxlcdm]+)\.\s+',  # i. 项目
        'upper_roman_with_dot': r'^\s*([IVXLCDM]+)\.\s+',  # I. 项目
        'lower_roman_with_paren': r'^\s*([ivxlcdm]+)\)\s+',  # i) 项目
        'upper_roman_with_paren': r'^\s*([IVXLCDM]+)\)\s+',  # I) 项目
        
        # 多级编号
        'multi_level_arabic': r'^\s*(\d+\.\d+(?:\.\d+)*)\s+',  # 1.1. 项目、1.1.1. 项目
        'multi_level_chinese': r'^\s*([一二三四五六七八九十]+)、\s*(\d+)\.\s+',  # 一、1. 项目
        'multi_level_mixed': r'^\s*(\d+)\.\s*\(([a-z])\)\s+',  # 1. (a) 项目
        
        # 其他常见格式
        'number_with_slash': r'^\s*(\d+)/(\d+)\s+',  # 1/2 项目
        'number_with_period': r'^\s*(\d+)\.(\d+)\s+',  # 1.2 项目（无空格）
    }
    return patterns


def get_bullet_patterns() -> List[str]:
    """
    获取项目符号模式
    
    Returns:
        项目符号模式列表
    """
    return [
        r'^\s*·\s+',  # 中文点号
        r'^\s*\*\s+',  # 星号
        r'^\s*-\s+',   # 连字符
        r'^\s+•\s+',   # 实心圆点
        r'^\s*◦\s+',   # 空心圆点
        r'^\s*▪\s+',   # 实心方块
        r'^\s*▫\s+',   # 空心方块
        r'^\s*▸\s+',   # 右箭头
        r'^\s*▶\s+',   # 实心右箭头
        r'^\s*→\s+',   # 箭头
        r'^\s*⇒\s+',   # 双箭头
        r'^\s*✓\s+',   # 对勾
        r'^\s*✗\s+',   # 叉号
        r'^\s*☑\s+',   # 复选框
        r'^\s*☒\s+',   # 选中的复选框
    ]


def is_title_paragraph(paragraph_text: str) -> bool:
    """
    判断是否为标题段落
    
    Args:
        paragraph_text: 段落文本
    
    Returns:
        是否为标题
    """
    # 检查是否以#开头（Markdown格式）
    if paragraph_text.strip().startswith('#'):
        return True
    
    # 检查是否为编号标题（如"1. 标题"、"一、标题"等）
    numbering_patterns = detect_numbering_patterns()
    for pattern in numbering_patterns.values():
        if re.match(pattern, paragraph_text):
            return True
    
    return False


def extract_numbering(paragraph_text: str) -> str:
    """
    提取段落中的编号部分
    
    Args:
        paragraph_text: 段落文本
    
    Returns:
        编号部分文本
    """
    numbering_patterns = detect_numbering_patterns()
    for pattern in numbering_patterns.values():
        match = re.match(pattern, paragraph_text)
        if match:
            return match.group()
    return ''


def extract_content(paragraph_text: str) -> str:
    """
    提取段落中的内容部分（去除编号和项目符号）
    
    Args:
        paragraph_text: 段落文本
    
    Returns:
        内容部分文本
    """
    # 去除项目符号
    bullet_patterns = get_bullet_patterns()
    for pattern in bullet_patterns:
        match = re.match(pattern, paragraph_text)
        if match:
            return paragraph_text[len(match.group()):].strip()
    
    # 去除编号
    numbering_patterns = detect_numbering_patterns()
    for pattern in numbering_patterns.values():
        match = re.match(pattern, paragraph_text)
        if match:
            return paragraph_text[len(match.group()):].strip()
    
    return paragraph_text.strip()
