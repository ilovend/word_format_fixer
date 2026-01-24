"""测试UI模块"""

import unittest
import tkinter as tk
from word_format_fixer.ui.main import WordFixerApp


class TestUI(unittest.TestCase):
    """测试UI模块"""

    def setUp(self):
        """设置测试环境"""
        self.root = tk.Tk()
        self.root.withdraw()  # 隐藏窗口

    def tearDown(self):
        """清理测试环境"""
        if self.root:
            self.root.destroy()

    def test_app_initialization(self):
        """测试应用初始化"""
        app = WordFixerApp(self.root)
        self.assertIsInstance(app, WordFixerApp)
        self.assertEqual(self.root.title(), "Word文档格式修复工具")

    def test_config_defaults(self):
        """测试默认配置"""
        app = WordFixerApp(self.root)
        self.assertEqual(app.chinese_font.get(), "宋体")
        self.assertEqual(app.western_font.get(), "Arial")
        self.assertEqual(app.title_font.get(), "黑体")
        self.assertEqual(app.font_size_body.get(), 12)

    def test_preset_configs(self):
        """测试预设配置"""
        app = WordFixerApp(self.root)
        # 测试默认预设
        self.assertEqual(app.preset.get(), "default")
        # 测试切换预设
        app.preset.set("bid_document")
        # 触发预设变更
        app.on_preset_change(None)
        # 验证配置是否更新
        self.assertIsInstance(app.chinese_font.get(), str)


if __name__ == '__main__':
    unittest.main()
