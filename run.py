"""Word文档格式修复工具启动脚本"""

import sys
from word_format_fixer.ui.main import run_app
from word_format_fixer.cli.main import main


def print_help():
    """打印帮助信息"""
    print("Word文档格式修复工具")
    print("使用方法:")
    print("  python run.py          # 启动GUI界面")
    print("  python run.py --cli    # 使用命令行模式")
    print("  python run.py --help   # 显示此帮助信息")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == "--cli":
            # 启动命令行模式
            sys.argv = sys.argv[1:]
            main()
        elif sys.argv[1] == "--help":
            # 显示帮助信息
            print_help()
        else:
            # 其他参数传递给命令行模式
            sys.argv = [sys.argv[0]] + sys.argv[1:]
            main()
    else:
        # 启动GUI界面
        run_app()
