"""规则基类测试"""

import unittest
from rules.base_rule import BaseRule, RuleResult


class RuleResultTestCase(unittest.TestCase):
    """测试RuleResult类"""

    def test_init_with_all_params(self):
        """测试使用所有参数初始化RuleResult"""
        result = RuleResult(
            rule_id="test_rule",
            success=True,
            fixed_count=5,
            details=["detail1", "detail2"]
        )

        self.assertEqual(result.rule_id, "test_rule")
        self.assertTrue(result.success)
        self.assertEqual(result.fixed_count, 5)
        self.assertEqual(len(result.details), 2)

    def test_init_with_empty_details(self):
        """测试使用空详情列表初始化RuleResult"""
        result = RuleResult(
            rule_id="test_rule",
            success=True,
            fixed_count=0,
            details=[]
        )

        self.assertEqual(result.details, [])

    def test_dict_method(self):
        """测试RuleResult的dict方法"""
        result = RuleResult(
            rule_id="test_rule",
            success=True,
            fixed_count=3,
            details=["fixed item 1", "fixed item 2", "fixed item 3"]
        )

        result_dict = result.dict()

        self.assertIsInstance(result_dict, dict)
        self.assertEqual(result_dict["rule_id"], "test_rule")
        self.assertTrue(result_dict["success"])
        self.assertEqual(result_dict["fixed_count"], 3)
        self.assertEqual(len(result_dict["details"]), 3)

    def test_dict_returns_correct_structure(self):
        """测试dict方法返回正确的数据结构"""
        result = RuleResult("test", True, 1, ["test detail"])
        result_dict = result.dict()

        required_keys = ["rule_id", "success", "fixed_count", "details"]
        for key in required_keys:
            self.assertIn(key, result_dict)


class BaseRuleTestCase(unittest.TestCase):
    """测试BaseRule类"""

    def setUp(self):
        """设置测试环境"""
        self.test_rule_class = self.create_test_rule()

    def create_test_rule(self):
        """创建测试规则类"""
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

        return TestRule

    def test_init_without_config(self):
        """测试不使用配置初始化规则"""
        rule = self.test_rule_class()
        self.assertEqual(rule.config, {})
        self.assertTrue(rule.enabled)

    def test_init_with_config(self):
        """测试使用配置初始化规则"""
        config = {
            "param1": "value1",
            "param2": 42
        }
        rule = self.test_rule_class(config)
        self.assertEqual(rule.config["param1"], "value1")
        self.assertEqual(rule.config["param2"], 42)

    def test_rule_id_generation(self):
        """测试规则ID自动生成"""
        rule = self.test_rule_class()
        self.assertEqual(rule.rule_id, "TestRule")

    def test_display_name_property(self):
        """测试显示名称属性"""
        rule = self.test_rule_class()
        self.assertEqual(rule.display_name, "测试规则")

    def test_enabled_property(self):
        """测试启用属性"""
        rule = self.test_rule_class()
        self.assertTrue(rule.enabled)

    def test_get_metadata(self):
        """测试获取规则元数据"""
        rule = self.test_rule_class({
            "param1": "value1",
            "param2": 42
        })

        metadata = rule.get_metadata()

        # 检查元数据结构
        self.assertIsInstance(metadata, dict)

        # 检查必需字段
        self.assertIn("id", metadata)
        self.assertIn("name", metadata)
        self.assertIn("description", metadata)
        self.assertIn("category", metadata)
        self.assertIn("params", metadata)

        # 检查字段值
        self.assertEqual(metadata["id"], "TestRule")
        self.assertEqual(metadata["name"], "测试规则")
        self.assertEqual(metadata["category"], "测试类别")
        self.assertEqual(metadata["params"]["param1"], "value1")
        self.assertEqual(metadata["params"]["param2"], 42)

    def test_get_metadata_without_category(self):
        """测试没有category属性的规则的元数据"""
        class TestRuleWithoutCategory(BaseRule):
            display_name = "无类别规则"

            def apply(self, doc_context):
                return RuleResult(
                    rule_id=self.rule_id,
                    success=True,
                    fixed_count=0,
                    details=["测试"]
                )

        rule = TestRuleWithoutCategory()
        metadata = rule.get_metadata()

        # category应该有默认值"其他规则"
        self.assertEqual(metadata["category"], "其他规则")

    def test_explain_method(self):
        """测试explain方法"""
        rule = self.test_rule_class()
        explanation = rule.explain()

        self.assertIsInstance(explanation, str)
        # explain方法返回__doc__的值，或者"无描述"
        self.assertTrue(explanation in [None, "无描述", ""])

    def test_explain_with_docstring(self):
        """测试有文档字符串的规则的explain方法"""
        class TestRuleWithDoc(BaseRule):
            display_name = "带文档的规则"
            category = "测试"
            """这是规则的详细说明文档"""

            def apply(self, doc_context):
                return RuleResult(
                    rule_id=self.rule_id,
                    success=True,
                    fixed_count=0,
                    details=["测试"]
                )

        rule = TestRuleWithDoc()
        explanation = rule.explain()

        self.assertIn("详细说明", explanation)

    def test_config_update(self):
        """测试更新规则配置"""
        rule = self.test_rule_class({"param1": "value1"})

        # 更新配置
        rule.config["param1"] = "new_value"
        rule.config["param2"] = "value2"

        self.assertEqual(rule.config["param1"], "new_value")
        self.assertEqual(rule.config["param2"], "value2")

    def test_enabled_toggle(self):
        """测试切换规则启用状态"""
        rule = self.test_rule_class()

        # 禁用规则
        rule.enabled = False
        self.assertFalse(rule.enabled)

        # 启用规则
        rule.enabled = True
        self.assertTrue(rule.enabled)

    def test_abstract_apply_method(self):
        """测试apply方法是抽象的"""
        from abc import ABC

        # BaseRule应该继承自ABC
        self.assertTrue(issubclass(BaseRule, ABC))


if __name__ == '__main__':
    unittest.main()
