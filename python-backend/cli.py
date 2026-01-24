#!/usr/bin/env python3
"""
Word Format Fixer 命令行工具
直接处理命令，输出JSON结果
"""

import sys
import json
from typing import Dict, Any, List
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 导入服务层
from services import (
    DocumentProcessingService,
    ConfigManagementService,
    RuleManagementService
)

# 全局服务实例
doc_service = DocumentProcessingService()
config_service = ConfigManagementService()
rule_service = RuleManagementService()


def get_version() -> str:
    """
    读取并返回当前版本号
    
    Returns:
        当前版本号
    """
    version_file = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "VERSION")
    if os.path.exists(version_file):
        with open(version_file, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "1.0.0"


def process_command(command: str, data: Dict[str, Any]) -> Dict[str, Any]:
    """
    处理命令并返回结果
    
    Args:
        command: 命令名称
        data: 命令参数
        
    Returns:
        处理结果，JSON格式
    """
    try:
        if command == "get-version":
            # 获取版本号
            return {"version": get_version()}
        
        elif command == "get-presets":
            # 获取所有预设
            return config_service.get_all_presets()
        
        elif command == "get-rules":
            # 获取所有规则
            return rule_service.get_all_rules()
        
        elif command == "save-preset":
            # 保存预设
            preset_id = data.get('preset_id')
            preset_data = data.get('preset_data')
            return config_service.save_preset(preset_id, preset_data)
        
        elif command == "delete-preset":
            # 删除预设
            preset_id = data.get('preset_id')
            return config_service.delete_preset(preset_id)
        
        elif command == "process-document":
            # 处理文档
            file_path = data.get('file_path')
            active_rules = data.get('active_rules', [])
            result = doc_service.process_document(file_path, active_rules)
            
            # 构建响应
            return {
                "status": result.get("status", "success"),
                "summary": result.get("summary", {}),
                "results": result.get("results", []),
                "save_success": result.get("save_success", False),
                "saved_to": result.get("saved_to", file_path)
            }
        
        else:
            return {"error": f"Unknown command: {command}"}
            
    except Exception as e:
        return {"error": str(e)}


def main():
    """命令行入口"""
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Missing command"}))
        sys.exit(1)
    
    # 获取命令
    command = sys.argv[1]
    
    # 获取数据（如果有）
    data = {}
    if len(sys.argv) > 2:
        try:
            data = json.loads(sys.argv[2])
        except json.JSONDecodeError as e:
            print(json.dumps({"error": f"Invalid JSON data: {str(e)}"}))
            sys.exit(1)
    
    # 处理命令
    result = process_command(command, data)
    
    # 输出结果
    print(json.dumps(result))
    
    # 设置退出码
    if "error" in result:
        sys.exit(1)
    else:
        sys.exit(0)


if __name__ == "__main__":
    main()
