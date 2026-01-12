@echo off
echo 开始构建Word文档格式修复工具...

:: 清理之前的构建结果
echo 清理之前的构建结果...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "WordFormatFixer.spec" del "WordFormatFixer.spec"
echo 清理完成!

:: 检查并安装PyInstaller
echo 检查PyInstaller是否已安装...
pip list | findstr pyinstaller > nul
if %errorlevel% neq 0 (
    echo 正在安装PyInstaller...
    pip install pyinstaller
)

:: 安装项目依赖
echo 安装项目依赖...
pip install -r requirements.txt

:: 构建可执行文件
echo 构建可执行文件...
pyinstaller --onefile --windowed --name WordFormatFixer --add-data "word_format_fixer;word_format_fixer" run.py

echo 构建完成！
echo 可执行文件位置: dist\WordFormatFixer.exe
echo.
echo 使用方法:
echo 1. 双击 dist\WordFormatFixer.exe 运行GUI界面
echo 2. 命令行模式: dist\WordFormatFixer.exe --cli 输入文件.docx -o 输出文件.docx
echo.
pause
