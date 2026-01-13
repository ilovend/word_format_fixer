"""Word文档格式修复工具启动脚本"""

import sys
from word_format_fixer.ui.main import run_app as run_tk_app
from word_format_fixer.ui.qt.main_window import run_app as run_qt_app
from word_format_fixer.cli.main import main


def print_help():
    """打印帮助信息"""
    print("Word文档格式修复工具")
    print("使用方法:")
    print("  python run.py          # 启动默认GUI界面")
    print("  python run.py --qt     # 启动PyQt-Fluent-Widgets界面")
    print("  python run.py --tk     # 启动tkinter界面")
    print("  python run.py --cli    # 使用命令行模式")
    print("  python run.py --help   # 显示此帮助信息")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == "--cli":
            # 启动命令行模式
            sys.argv = sys.argv[1:]
            main()
        elif sys.argv[1] == "--qt":
            # 启动PyQt-Fluent-Widgets界面
            run_qt_app()
        elif sys.argv[1] == "--tk":
            # 启动tkinter界面
            run_tk_app()
        elif sys.argv[1] == "--help":
            # 显示帮助信息
            print_help()
        else:
            # 其他参数传递给命令行模式
            sys.argv = [sys.argv[0]] + sys.argv[1:]
            main()
    else:
        # 启动默认GUI界面（使用PyQt-Fluent-Widgets）
        run_qt_app()
