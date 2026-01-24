#!/usr/bin/env python3
"""测试运行脚本"""

import sys
import subprocess
import argparse
from pathlib import Path


def run_pytest(args):
    """使用pytest运行测试"""
    cmd = ["pytest"]

    # 添加pytest参数
    if args.verbose:
        cmd.append("-v")
    if args.coverage:
        cmd.extend(["--cov=python-backend", "--cov-report=html", "--cov-report=term"])
    if args.marks:
        cmd.extend(["-m", args.marks])
    if args.pattern:
        cmd.append(args.pattern)
    else:
        cmd.append("tests/")

    print(f"运行命令: {' '.join(cmd)}")
    subprocess.run(cmd, check=True)


def run_unittest(args):
    """使用unittest运行测试"""
    cmd = [sys.executable, "-m", "unittest", "discover", "tests/", "-v"]

    print(f"运行命令: {' '.join(cmd)}")
    subprocess.run(cmd, check=True)


def main():
    parser = argparse.ArgumentParser(description="运行Word Format Fixer测试")
    parser.add_argument(
        "--framework",
        choices=["pytest", "unittest", "both"],
        default="pytest",
        help="选择测试框架"
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="详细输出"
    )
    parser.add_argument(
        "-c", "--coverage",
        action="store_true",
        help="生成覆盖率报告"
    )
    parser.add_argument(
        "-m", "--marks",
        help="运行特定标记的测试 (pytest only)"
    )
    parser.add_argument(
        "-p", "--pattern",
        help="测试文件模式"
    )

    args = parser.parse_args()

    if args.framework == "pytest":
        run_pytest(args)
    elif args.framework == "unittest":
        run_unittest(args)
    else:  # both
        print("=" * 60)
        print("使用pytest运行测试")
        print("=" * 60)
        run_pytest(args)

        print("\n" + "=" * 60)
        print("使用unittest运行测试")
        print("=" * 60)
        run_unittest(args)


if __name__ == "__main__":
    main()
