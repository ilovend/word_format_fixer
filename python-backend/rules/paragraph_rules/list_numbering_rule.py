import re
from rules.base_rule import BaseRule, RuleResult
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt, Cm
from docx.oxml.ns import qn

class ListNumberingRule(BaseRule):
    """编号列表规则 - 修复文档中的编号列表和项目符号"""

    display_name = "编号列表标准化"
    category = "段落规则"
    
    def __init__(self, config=None):
        # 搬运原 default_config 中的相关配置
        default_params = {
            'chinese_font': '宋体',
            'western_font': 'Arial',
            'font_size_body': 12,
            'text_color': (0, 0, 0),
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
        
        # 获取编号模式
        patterns = self.detect_numbering_patterns()
        
        for paragraph in document.paragraphs:
            original_text = paragraph.text.strip()
            
            if not original_text or paragraph.style.name.startswith('Heading'):
                continue
            
            # 处理不需要的项目符号（如·）
            bullet_patterns = [
                r'^\s*·\s+',  # 中文点号
                r'^\s*\*\s+',  # 星号
                r'^\s*-\s+',   # 连字符
                r'^\s+•\s+',   # 实心圆点
            ]
            
            bullet_matched = False
            for bullet_pattern in bullet_patterns:
                bullet_match = re.match(bullet_pattern, original_text)
                if bullet_match:
                    bullet_matched = True
                    # 提取内容（去掉项目符号）
                    content = original_text[len(bullet_match.group()):].strip()
                    
                    # 清除原有内容
                    paragraph.clear()
                    
                    # 添加内容
                    run = paragraph.add_run(content)
                    run.font.name = self.config['western_font']
                    run._element.rPr.rFonts.set(
                        qn('w:eastAsia'), 
                        self.config['chinese_font']
                    )
                    run.font.color.rgb = self._get_rgb_color(*self.config['text_color'])
                    run.font.size = Pt(self.config['font_size_body'])
                    
                    # 设置为无序列表样式
                    if 'List Paragraph' in [style.name for style in document.styles]:
                        paragraph.style = document.styles['List Paragraph']
                    
                    # 设置缩进
                    paragraph_format = paragraph.paragraph_format
                    paragraph_format.left_indent = Cm(1.27)
                    paragraph_format.first_line_indent = Cm(-0.64)
                    paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                    paragraph_format.line_spacing = 1.5
                    
                    fixed_count += 1
                    details.append(f"修复项目符号段落: {original_text[:15]}...")
                    break
            
            if bullet_matched:
                continue
            
            # 检查各种编号模式
            for pattern_name, pattern in patterns.items():
                match = re.match(pattern, original_text)
                if match:
                    self._format_numbered_paragraph(paragraph, pattern_name, match, original_text, document)
                    fixed_count += 1
                    details.append(f"修复编号段落: {original_text[:15]}...")
                    break
        
        details.append(f"总共修复了 {fixed_count} 个编号或项目符号段落")
        
        return RuleResult(
            rule_id=self.rule_id,
            success=True,
            fixed_count=fixed_count,
            details=details
        )
    
    def detect_numbering_patterns(self):
        """检测文档中的编号模式"""
        patterns = {
            'arabic_with_dot': r'^\s*(\d+)\.\s+',  # 1. 项目
            'arabic_with_paren': r'^\s*(\d+)\)\s+',  # 1) 项目
            'chinese_with_dot': r'^\s*([一二三四五六七八九十]+)\.\s+',  # 一. 项目
            'chinese_with_bracket': r'^\s*([一二三四五六七八九十]+)、\s+',  # 一、项目
            'chinese_with_paren': r'^\s*\(([一二三四五六七八九十]+)\)\s+',  # (一) 项目
            'lower_alpha_with_dot': r'^\s*([a-z])\.\s+',  # a. 项目
            'upper_alpha_with_dot': r'^\s*([A-Z])\.\s+',  # A. 项目
            'lower_roman_with_dot': r'^\s*([ivxlcdm]+)\.\s+',  # i. 项目
            'upper_roman_with_dot': r'^\s*([IVXLCDM]+)\.\s+',  # I. 项目
        }
        return patterns
    
    def _format_numbered_paragraph(self, paragraph, pattern_type: str, match, original_text: str, document):
        """格式化编号段落"""
        # 提取编号和内容
        number_part = match.group()
        content = original_text[len(number_part):].strip()
        
        # 清除原有内容
        paragraph.clear()
        
        # 添加内容
        run = paragraph.add_run(content)
        run.font.name = self.config['western_font']
        run._element.rPr.rFonts.set(
            qn('w:eastAsia'), 
            self.config['chinese_font']
        )
        run.font.color.rgb = self._get_rgb_color(*self.config['text_color'])
        run.font.size = Pt(self.config['font_size_body'])
        
        # 设置段落格式
        if 'List Paragraph' in [style.name for style in document.styles]:
            paragraph.style = document.styles['List Paragraph']
        else:
            paragraph.style = document.styles['Normal']
        
        # 设置缩进
        paragraph_format = paragraph.paragraph_format
        
        # 根据编号类型设置不同缩进
        if pattern_type.startswith('arabic') or pattern_type.startswith('chinese'):
            # 主要列表项
            paragraph_format.left_indent = Cm(1.27)
            paragraph_format.first_line_indent = Cm(-0.64)
        elif pattern_type.startswith('lower_') or pattern_type.startswith('upper_'):
            # 次级列表项
            paragraph_format.left_indent = Cm(2.54)
            paragraph_format.first_line_indent = Cm(-0.64)
        else:
            # 默认缩进
            paragraph_format.left_indent = Cm(1.27)
            paragraph_format.first_line_indent = Cm(-0.64)
        
        # 设置行距
        paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        paragraph_format.line_spacing = 1.5
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(6)
    
    def _get_rgb_color(self, r, g, b):
        """获取RGB颜色对象"""
        from docx.shared import RGBColor
        return RGBColor(r, g, b)
