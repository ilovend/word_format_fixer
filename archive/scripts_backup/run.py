#!/usr/bin/env python3
"""Word文档格式修复工具启动脚本"""

import sys
import os

# 获取项目根目录
project_root = os.path.dirname(os.path.abspath(__file__))

# 添加python-backend目录到Python路径
python_backend_path = os.path.join(project_root, "python-backend")
sys.path.insert(0, python_backend_path)

# 直接从python-backend目录导入模块
from ipc.adapter import start_server
from core.engine import RuleEngine
from rules import get_all_rules

if __name__ == "__main__":
    # 初始化规则引擎
    engine = RuleEngine()
    
    # 注册所有规则
    rules = get_all_rules()
    for rule in rules:
        engine.register_rule(rule)
    
    # 启动Python后端服务器
    print("Starting Word Format Fixer IPC Server...")
    print("Server will run on http://127.0.0.1:7777")
    print("Press Ctrl+C to stop")
    
    try:
        # 传递命令行参数给start_server函数
        start_server(host="127.0.0.1", port=7777)
    except KeyboardInterrupt:
        print("Server stopped by user")
    except Exception as e:
        print(f"Error starting server: {e}")
        sys.exit(1)
