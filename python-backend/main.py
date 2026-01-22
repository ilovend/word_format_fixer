#!/usr/bin/env python3
"""Word Format Fixer 启动入口"""

import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ipc.adapter import start_server
from core.engine import RuleEngine

# 导入规则注册
from rules import get_all_rules

if __name__ == "__main__":
    # 初始化规则引擎
    engine = RuleEngine()
    
    # 注册所有规则
    rules = get_all_rules()
    for rule in rules:
        engine.register_rule(rule)
    
    # 启动IPC服务器
    print("Starting Word Format Fixer IPC Server...")
    print("Server will run on http://127.0.0.1:7777")
    print("Press Ctrl+C to stop")
    
    try:
        start_server(host="127.0.0.1", port=7777)
    except KeyboardInterrupt:
        print("Server stopped by user")
    except Exception as e:
        print(f"Error starting server: {e}")
        sys.exit(1)