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
    RuleManagementService,
    DiffService
)

# 全局服务实例
doc_service = DocumentProcessingService()
config_service = ConfigManagementService()
rule_service = RuleManagementService()
diff_service = DiffService()


def get_version() -> str:
    """
    读取并返回当前版本号
    
    Returns:
        当前版本号
    """
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
    version_file = os.path.join(base_dir, "VERSION")
    
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
        
        elif command == "configure-rules":
            # 配置规则参数
            configs = data.get('configs', [])
            results = []
            for config in configs:
                rule_id = config.get('rule_id')
                params = config.get('params', {})
                result = rule_service.update_rule_config(rule_id, params)
                results.append(result)
            return {"status": "success", "results": results}
        
        elif command == "prepare-diff":
            # 准备对比：缓存原始文档
            file_path = data.get('file_path')
            return diff_service.prepare_diff(file_path)
        
        elif command == "generate-diff":
            # 生成对比：对比修改后的文档
            file_path = data.get('file_path')
            return diff_service.generate_diff(file_path)
        
        elif command == "get-preview":
            # 获取文档HTML预览
            file_path = data.get('file_path')
            return diff_service.get_document_preview(file_path)
        
        else:
            return {"error": f"Unknown command: {command}"}
            
    except Exception as e:
        return {"error": str(e)}


def run_interactive_mode():
    """
    交互模式：持续读取标准输入，输出 JSON 结果
    每行输入格式：{"id": "req-1", "command": "cmd_name", "data": {...}}
    输出格式：{"id": "req-1", "result": {...}}
    """
    # 强制标准输出使用 UTF-8
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stdin.reconfigure(encoding='utf-8')
    
    # 打印就绪信号
    print(json.dumps({"status": "ready", "pid": os.getpid()}))
    sys.stdout.flush()
    
    while True:
        try:
            line = sys.stdin.readline()
            if not line:
                break
                
            line = line.strip()
            if not line:
                continue
                
            try:
                request = json.loads(line)
                req_id = request.get('id')
                command = request.get('command')
                data = request.get('data', {})
                
                # 处理命令
                result = process_command(command, data)
                
                # 输出响应
                response = {
                    "id": req_id,
                    "success": "error" not in result,
                    "result": result
                }
                print(json.dumps(response))
                sys.stdout.flush()
                
            except json.JSONDecodeError:
                print(json.dumps({
                    "id": None, 
                    "success": False, 
                    "error": "Invalid JSON format"
                }))
                sys.stdout.flush()
                
        except Exception as e:
            # 捕获循环中的未处理异常，防止进程退出
            print(json.dumps({
                "id": None,
                "success": False, 
                "error": f"Internal error: {str(e)}"
            }))
            sys.stdout.flush()


def main():
    """命令行入口"""
    # 检查是否进入交互模式
    if len(sys.argv) > 1 and sys.argv[1] == "--interactive":
        run_interactive_mode()
        return

    # --- 兼容原有单命令模式 ---
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
