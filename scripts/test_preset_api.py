#!/usr/bin/env python3
"""
测试预设管理API功能
"""

import requests
import json

# 后端服务器地址
BASE_URL = "http://127.0.0.1:7777"


def test_get_presets():
    """测试获取预设列表"""
    print("=== 测试获取预设列表 ===")
    try:
        response = requests.get(f"{BASE_URL}/api/presets")
        print(f"状态码: {response.status_code}")
        presets = response.json()
        print(f"获取到 {len(presets)} 个预设")
        for preset_id, preset in presets.items():
            print(f"- {preset_id}: {preset['name']} - {preset['description']}")
        print("✓ 获取预设列表成功")
        return presets
    except Exception as e:
        print(f"✗ 获取预设列表失败: {e}")
        return None


def test_save_new_preset():
    """测试保存新预设"""
    print("\n=== 测试保存新预设 ===")
    try:
        new_preset = {
            "preset_id": "test_preset",
            "preset_data": {
                "name": "测试预设",
                "description": "这是一个测试预设",
                "rules": {
                    "FontNameRule": {"enabled": True},
                    "TitleFontRule": {"enabled": True}
                }
            }
        }
        response = requests.post(
            f"{BASE_URL}/api/presets/save",
            data=json.dumps(new_preset),
            headers={"Content-Type": "application/json"}
        )
        print(f"状态码: {response.status_code}")
        print(f"响应: {response.json()}")
        print("✓ 保存新预设成功")
        return True
    except Exception as e:
        print(f"✗ 保存新预设失败: {e}")
        return False


def test_delete_preset():
    """测试删除预设"""
    print("\n=== 测试删除预设 ===")
    try:
        delete_data = {
            "preset_id": "test_preset"
        }
        response = requests.delete(
            f"{BASE_URL}/api/presets/delete",
            data=json.dumps(delete_data),
            headers={"Content-Type": "application/json"}
        )
        print(f"状态码: {response.status_code}")
        print(f"响应: {response.json()}")
        print("✓ 删除预设成功")
        return True
    except Exception as e:
        print(f"✗ 删除预设失败: {e}")
        return False


def test_get_rules():
    """测试获取规则列表"""
    print("\n=== 测试获取规则列表 ===")
    try:
        response = requests.get(f"{BASE_URL}/api/rules")
        print(f"状态码: {response.status_code}")
        rules = response.json()
        print(f"获取到 {len(rules)} 个规则")
        for rule in rules:
            print(f"- {rule['id']}: {rule['name']} ({rule['category']})")
        print("✓ 获取规则列表成功")
        return rules
    except Exception as e:
        print(f"✗ 获取规则列表失败: {e}")
        return None


if __name__ == "__main__":
    print("开始测试预设管理API功能...")
    print(f"后端地址: {BASE_URL}")
    print("=" * 50)
    
    # 测试获取规则列表
    rules = test_get_rules()
    
    # 测试获取预设列表
    presets = test_get_presets()
    
    # 测试保存新预设
    save_success = test_save_new_preset()
    
    # 再次获取预设列表，验证新预设是否保存成功
    if save_success:
        print("\n=== 验证新预设是否保存成功 ===")
        presets_after_save = test_get_presets()
        if presets_after_save and "test_preset" in presets_after_save:
            print("✓ 新预设已成功保存")
        else:
            print("✗ 新预设保存失败")
    
    # 测试删除预设
    delete_success = test_delete_preset()
    
    # 再次获取预设列表，验证预设是否删除成功
    if delete_success:
        print("\n=== 验证预设是否删除成功 ===")
        presets_after_delete = test_get_presets()
        if presets_after_delete and "test_preset" not in presets_after_delete:
            print("✓ 预设已成功删除")
        else:
            print("✗ 预设删除失败")
    
    print("\n" + "=" * 50)
    print("测试完成！")
