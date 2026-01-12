"""测试工具函数模块"""

import unittest
from word_format_fixer.core.utils import (
    detect_numbering_patterns,
    get_bullet_patterns,
    is_title_paragraph,
    extract_numbering,
    extract_content
)


class TestUtils(unittest.TestCase):
    """测试工具函数"""

    def test_detect_numbering_patterns(self):
        """测试检测编号模式"""
        patterns = detect_numbering_patterns()
        self.assertIsInstance(patterns, dict)
        self.assertGreater(len(patterns), 0)
        self.assertIn('arabic_with_dot', patterns)
        self.assertIn('chinese_with_bracket', patterns)

    def test_get_bullet_patterns(self):
        """测试获取项目符号模式"""
        patterns = get_bullet_patterns()
        self.assertIsInstance(patterns, list)
        self.assertEqual(len(patterns), 4)

    def test_is_title_paragraph(self):
        """测试判断是否为标题段落"""
        # 测试Markdown格式标题
        self.assertTrue(is_title_paragraph('# 标题'))
        self.assertTrue(is_title_paragraph('## 副标题'))
        
        # 测试编号标题
        self.assertTrue(is_title_paragraph('1. 标题'))
        self.assertTrue(is_title_paragraph('一、标题'))
        
        # 测试普通段落
        self.assertFalse(is_title_paragraph('普通段落'))

    def test_extract_numbering(self):
        """测试提取编号部分"""
        self.assertEqual(extract_numbering('1. 标题'), '1. ')
        self.assertEqual(extract_numbering('一、标题'), '一、')
        self.assertEqual(extract_numbering('(一) 标题'), '(一) ')
        self.assertEqual(extract_numbering('普通段落'), '')

    def test_extract_content(self):
        """测试提取内容部分"""
        self.assertEqual(extract_content('1. 标题内容'), '标题内容')
        self.assertEqual(extract_content('一、标题内容'), '标题内容')
        self.assertEqual(extract_content('· 项目内容'), '项目内容')
        self.assertEqual(extract_content('普通内容'), '普通内容')


if __name__ == '__main__':
    unittest.main()
