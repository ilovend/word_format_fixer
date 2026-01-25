#!/usr/bin/env python3
"""
Word Format Fixer 版本管理脚本
用于更新项目的版本号
"""

import os
import sys
import argparse


def get_current_version() -> str:
    """
    获取当前版本号
    
    Returns:
        当前版本号
    """
    version_file = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "VERSION")
    if os.path.exists(version_file):
        with open(version_file, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "1.0.0"


def update_version(new_version: str) -> None:
    """
    更新项目的版本号
    
    Args:
        new_version: 新的版本号
    """
    # 更新VERSION文件
    version_file = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "VERSION")
    with open(version_file, "w", encoding="utf-8") as f:
        f.write(new_version)
    
    print(f"版本号已更新为: {new_version}")
    print("\n使用以下命令来应用版本号:")
    print("- 前端打包: npm run prepack && npm run pack")
    print("- 后端构建: 重新构建Python后端即可")
    print("\n下次打包时，版本号将自动应用到所有组件中")


def main():
    """
    主函数
    """
    parser = argparse.ArgumentParser(description="Word Format Fixer 版本管理脚本")
    parser.add_argument("--show", action="store_true", help="显示当前版本号")
    parser.add_argument("--set", type=str, help="设置新的版本号")
    
    args = parser.parse_args()
    
    if args.show:
        current_version = get_current_version()
        print(f"当前版本号: {current_version}")
    elif args.set:
        update_version(args.set)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()