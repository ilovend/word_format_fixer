"""配置仓库测试"""

import unittest
import tempfile
import os
from pathlib import Path
from core.config_repository import IConfigRepository
from core.yaml_config_repository import YamlConfigRepository


class YamlConfigRepositoryTestCase(unittest.TestCase):
    """测试YAML配置仓库"""

    def setUp(self):
        """设置测试环境"""
        self.temp_dir = tempfile.TemporaryDirectory()
        self.temp_path = Path(self.temp_dir.name)

        # 创建临时配置文件
        self.config_path = self.temp_path / "presets.yaml"

        # 创建初始配置
        initial_config = {
            "presets": {
                "default": {
                    "name": "默认预设",
                    "description": "这是默认预设",
                    "rules": {
                        "font_standard": {"enabled": True}
                    }
                },
                "minimal": {
                    "name": "最小预设",
                    "description": "这是最小预设",
                    "rules": {}
                }
            }
        }

        # 写入初始配置
        import yaml
        with open(self.config_path, 'w', encoding='utf-8') as f:
            yaml.dump(initial_config, f, allow_unicode=True, default_flow_style=False)

        # 创建仓库实例
        self.repository = YamlConfigRepository(str(self.config_path))

    def tearDown(self):
        """清理测试环境"""
        self.temp_dir.cleanup()

    def test_init(self):
        """测试仓库初始化"""
        repo = YamlConfigRepository(str(self.config_path))
        self.assertEqual(repo.config_path, str(self.config_path))

    def test_init_without_path(self):
        """测试不提供路径初始化"""
        repo = YamlConfigRepository()
        self.assertIsNotNone(repo.config_path)

    def test_load_config(self):
        """测试加载配置"""
        config = self.repository.load_config()

        self.assertIsInstance(config, dict)
        self.assertIn("presets", config)

    def test_load_config_contains_presets(self):
        """测试加载的配置包含预设"""
        config = self.repository.load_config()

        presets = config.get("presets", {})
        self.assertIn("default", presets)
        self.assertIn("minimal", presets)

    def test_save_config(self):
        """测试保存配置"""
        new_config = {
            "presets": {
                "test": {
                    "name": "测试预设",
                    "description": "测试",
                    "rules": {}
                }
            }
        }

        # 保存配置
        self.repository.save_config(new_config)

        # 重新加载并验证
        loaded_config = self.repository.load_config()
        self.assertEqual(loaded_config["presets"]["test"]["name"], "测试预设")

    def test_get_preset_exists(self):
        """测试获取存在的预设"""
        preset = self.repository.get_preset("default")

        self.assertIsNotNone(preset)
        self.assertEqual(preset["name"], "默认预设")
        self.assertEqual(preset["description"], "这是默认预设")

    def test_get_preset_not_exists(self):
        """测试获取不存在的预设"""
        preset = self.repository.get_preset("nonexistent")

        self.assertIsNone(preset)

    def test_save_preset(self):
        """测试保存预设"""
        preset_data = {
            "name": "新预设",
            "description": "这是一个新预设",
            "rules": {
                "test_rule": {"enabled": True}
            }
        }

        # 保存预设
        self.repository.save_preset("new_preset", preset_data)

        # 验证预设已保存
        saved_preset = self.repository.get_preset("new_preset")
        self.assertIsNotNone(saved_preset)
        self.assertEqual(saved_preset["name"], "新预设")

    def test_save_preset_overwrites(self):
        """测试保存预设覆盖已有预设"""
        preset_data = {
            "name": "更新的预设",
            "description": "这是更新的描述",
            "rules": {}
        }

        # 覆盖默认预设
        self.repository.save_preset("default", preset_data)

        # 验证预设已更新
        updated_preset = self.repository.get_preset("default")
        self.assertEqual(updated_preset["name"], "更新的预设")
        self.assertEqual(updated_preset["description"], "这是更新的描述")

    def test_delete_preset_exists(self):
        """测试删除存在的预设"""
        # 验证预设存在
        preset = self.repository.get_preset("minimal")
        self.assertIsNotNone(preset)

        # 删除预设
        self.repository.delete_preset("minimal")

        # 验证预设已删除
        deleted_preset = self.repository.get_preset("minimal")
        self.assertIsNone(deleted_preset)

    def test_delete_preset_not_exists(self):
        """测试删除不存在的预设"""
        # 删除不存在的预设不应该抛出异常
        self.repository.delete_preset("nonexistent")

        # 配置应该仍然有效
        config = self.repository.load_config()
        self.assertIsNotNone(config)

    def test_implements_interface(self):
        """测试实现了接口"""
        self.assertTrue(isinstance(self.repository, IConfigRepository))

    def test_load_and_save_roundtrip(self):
        """测试加载和保存的往返一致性"""
        # 加载配置
        original_config = self.repository.load_config()

        # 保存配置
        self.repository.save_config(original_config)

        # 重新加载
        reloaded_config = self.repository.load_config()

        # 验证一致性
        self.assertEqual(original_config, reloaded_config)

    def test_multiple_save_operations(self):
        """测试多次保存操作"""
        # 第一次保存
        preset1 = {
            "name": "预设1",
            "description": "描述1",
            "rules": {}
        }
        self.repository.save_preset("preset1", preset1)

        # 第二次保存
        preset2 = {
            "name": "预设2",
            "description": "描述2",
            "rules": {}
        }
        self.repository.save_preset("preset2", preset2)

        # 验证两个预设都存在
        self.assertIsNotNone(self.repository.get_preset("preset1"))
        self.assertIsNotNone(self.repository.get_preset("preset2"))

    def test_yaml_encoding(self):
        """测试YAML编码正确处理中文"""
        preset_data = {
            "name": "中文预设名称",
            "description": "这是中文描述",
            "rules": {}
        }

        # 保存中文内容
        self.repository.save_preset("chinese", preset_data)

        # 重新加载并验证
        loaded_preset = self.repository.get_preset("chinese")
        self.assertEqual(loaded_preset["name"], "中文预设名称")
        self.assertEqual(loaded_preset["description"], "这是中文描述")

    def test_empty_config_file(self):
        """测试处理空配置文件"""
        # 创建空配置文件
        empty_config_path = self.temp_path / "empty.yaml"
        empty_config_path.write_text("", encoding='utf-8')

        repo = YamlConfigRepository(str(empty_config_path))

        # 加载空配置
        config = repo.load_config()
        # yaml.safe_load对空文件返回None，但我们处理了这种情况
        self.assertIsNotNone(config)


if __name__ == '__main__':
    unittest.main()
