"""测试规则引擎模块"""

import unittest
import os
from unittest.mock import MagicMock, patch
import sys
import os

# 获取项目根目录
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 添加python-backend目录到Python路径
sys.path.insert(0, os.path.join(project_root, 'python-backend'))

# 直接从core导入，因为python-backend是一个目录，不是模块
from core.engine import RuleEngine
from rules.base_rule import BaseRule, RuleResult
from core.config_loader import ConfigLoader


class TestRuleEngine(unittest.TestCase):
    """测试规则引擎"""

    def setUp(self):
        """设置测试环境"""
        self.engine = RuleEngine()

    def test_init(self):
        """测试规则引擎初始化"""
        self.assertIsInstance(self.engine, RuleEngine)
        self.assertIsInstance(self.engine.rules, dict)

    def test_load_rules(self):
        """测试规则加载"""
        # 直接测试规则加载
        self.engine._load_rules()
        self.assertGreater(len(self.engine.rules), 0)

    def test_get_rules_info(self):
        """测试获取规则信息"""
        self.engine._load_rules()
        rules_info = self.engine.get_rules_info()
        self.assertIsInstance(rules_info, list)
        self.assertGreater(len(rules_info), 0)
        
        # 检查规则信息包含必要字段
        for rule_info in rules_info:
            self.assertIn("id", rule_info)
            self.assertIn("name", rule_info)
            self.assertIn("description", rule_info)
            self.assertIn("category", rule_info)
            self.assertIn("params", rule_info)


class TestBaseRule(unittest.TestCase):
    """测试规则基类"""

    def test_base_rule_init(self):
        """测试规则基类初始化"""
        # 创建一个简单的测试规则
        class TestRule(BaseRule):
            display_name = "测试规则"
            category = "测试类别"
            
            def apply(self, doc_context):
                return RuleResult(
                    rule_id=self.rule_id,
                    success=True,
                    fixed_count=0,
                    details=["测试规则执行"]
                )
        
        # 测试规则初始化
        rule = TestRule({"param1": "value1"})
        self.assertEqual(rule.config["param1"], "value1")
        self.assertEqual(rule.display_name, "测试规则")
        self.assertEqual(rule.category, "测试类别")

    def test_get_metadata(self):
        """测试获取规则元数据"""
        # 创建一个简单的测试规则
        class TestRule(BaseRule):
            """这是一个测试规则"""
            display_name = "测试规则"
            category = "测试类别"
            
            def apply(self, doc_context):
                return RuleResult(
                    rule_id=self.rule_id,
                    success=True,
                    fixed_count=0,
                    details=["测试规则执行"]
                )
        
        # 测试获取元数据
        rule = TestRule({"param1": "value1"})
        metadata = rule.get_metadata()
        
        self.assertEqual(metadata["id"], "TestRule")
        self.assertEqual(metadata["name"], "测试规则")
        self.assertEqual(metadata["description"], "这是一个测试规则")
        self.assertEqual(metadata["category"], "测试类别")
        self.assertEqual(metadata["params"]["param1"], "value1")


class TestConfigLoader(unittest.TestCase):
    """测试配置加载器"""

    def setUp(self):
        """设置测试环境"""
        self.config_loader = ConfigLoader()

    def test_get_all_presets(self):
        """测试获取所有预设"""
        presets = self.config_loader.get_all_presets()
        self.assertIsInstance(presets, dict)
        self.assertGreater(len(presets), 0)
        self.assertIn("default", presets)
        self.assertIn("minimal", presets)

    def test_get_preset(self):
        """测试获取指定预设"""
        preset = self.config_loader.get_preset("default")
        self.assertIsInstance(preset, dict)
        self.assertIn("name", preset)
        self.assertIn("description", preset)
        self.assertIn("rules", preset)

    def test_load_preset_config(self):
        """测试加载预设配置"""
        config = self.config_loader.load_preset_config("default")
        self.assertIsInstance(config, dict)

    def test_get_enabled_rules(self):
        """测试获取预设中启用的规则"""
        enabled_rules = self.config_loader.get_enabled_rules("default")
        self.assertIsInstance(enabled_rules, dict)
        self.assertGreater(len(enabled_rules), 0)

    def test_save_preset(self):
        """测试保存预设"""
        # 测试保存一个新的预设
        new_preset = {
            "name": "test_preset",
            "description": "测试预设",
            "rules": {
                "font_standard": {"enabled": True}
            }
        }
        self.config_loader.save_preset("test_preset", new_preset)
        
        # 验证预设是否保存成功
        presets = self.config_loader.get_all_presets()
        self.assertIn("test_preset", presets)

    def test_delete_preset(self):
        """测试删除预设"""
        # 先保存一个测试预设
        new_preset = {
            "name": "test_preset_to_delete",
            "description": "测试删除预设",
            "rules": {
                "font_standard": {"enabled": True}
            }
        }
        self.config_loader.save_preset("test_preset_to_delete", new_preset)
        
        # 验证预设是否存在
        presets = self.config_loader.get_all_presets()
        self.assertIn("test_preset_to_delete", presets)
        
        # 删除预设
        self.config_loader.delete_preset("test_preset_to_delete")
        
        # 验证预设是否已删除
        presets = self.config_loader.get_all_presets()
        self.assertNotIn("test_preset_to_delete", presets)


class TestFontRules(unittest.TestCase):
    """测试字体规则"""

    @patch('rules.font_rules.font_standard_rule')
    def test_font_name_rule(self, mock_font_standard_rule):
        """测试字体名称规则"""
        from rules.font_rules.font_standard_rule import FontNameRule
        
        # 测试规则初始化
        rule = FontNameRule({"chinese_font": "宋体", "western_font": "Arial"})
        self.assertEqual(rule.config["chinese_font"], "宋体")
        self.assertEqual(rule.config["western_font"], "Arial")
        
        # 测试规则元数据
        metadata = rule.get_metadata()
        self.assertEqual(metadata["name"], "字体名称标准化")
        self.assertEqual(metadata["category"], "字体规则")

    @patch('rules.font_rules.font_standard_rule')
    def test_title_font_rule(self, mock_font_standard_rule):
        """测试标题字体规则"""
        from rules.font_rules.font_standard_rule import TitleFontRule
        
        # 测试规则初始化
        rule = TitleFontRule({"title_font": "黑体"})
        self.assertEqual(rule.config["title_font"], "黑体")
        
        # 测试规则元数据
        metadata = rule.get_metadata()
        self.assertEqual(metadata["name"], "标题字体设置")
        self.assertEqual(metadata["category"], "字体规则")

    @patch('rules.font_rules.font_standard_rule')
    def test_font_size_rule(self, mock_font_standard_rule):
        """测试字号规则"""
        from rules.font_rules.font_standard_rule import FontSizeRule
        
        # 测试规则初始化
        rule = FontSizeRule({
            "font_size_body": 12,
            "font_size_title1": 22,
            "font_size_title2": 18,
            "font_size_title3": 16
        })
        self.assertEqual(rule.config["font_size_body"], 12)
        
        # 测试规则元数据
        metadata = rule.get_metadata()
        self.assertEqual(metadata["name"], "字号标准化")
        self.assertEqual(metadata["category"], "字体规则")


class TestParagraphRules(unittest.TestCase):
    """测试段落规则"""

    @patch('rules.paragraph_rules.paragraph_spacing_rule')
    def test_paragraph_spacing_rule(self, mock_paragraph_spacing_rule):
        """测试段落间距规则"""
        from rules.paragraph_rules.paragraph_spacing_rule import ParagraphSpacingRule
        
        # 测试规则初始化
        rule = ParagraphSpacingRule({
            "line_spacing": 1.5,
            "space_before": 0,
            "space_after": 6
        })
        self.assertEqual(rule.config["line_spacing"], 1.5)
        
        # 测试规则元数据
        metadata = rule.get_metadata()
        self.assertEqual(metadata["name"], "段落间距统一")
        self.assertEqual(metadata["category"], "段落规则")

    @patch('rules.paragraph_rules.title_bold_rule')
    def test_title_bold_rule(self, mock_title_bold_rule):
        """测试标题加粗规则"""
        from rules.paragraph_rules.title_bold_rule import TitleBoldRule
        
        # 测试规则初始化
        rule = TitleBoldRule({"enabled": True})
        self.assertEqual(rule.config["enabled"], True)
        
        # 测试规则元数据
        metadata = rule.get_metadata()
        self.assertEqual(metadata["name"], "标题加粗")
        self.assertEqual(metadata["category"], "段落规则")


class TestTableRules(unittest.TestCase):
    """测试表格规则"""

    @patch('rules.table_rules.table_width_rule')
    def test_table_width_rule(self, mock_table_width_rule):
        """测试表格宽度规则"""
        from rules.table_rules.table_width_rule import TableWidthRule
        
        # 测试规则初始化
        rule = TableWidthRule({
            "table_width_percent": 95,
            "auto_adjust_columns": True
        })
        self.assertEqual(rule.config["table_width_percent"], 95)
        
        # 测试规则元数据
        metadata = rule.get_metadata()
        self.assertEqual(metadata["name"], "表格宽度优化")
        self.assertEqual(metadata["category"], "表格规则")


if __name__ == '__main__':
    unittest.main()
