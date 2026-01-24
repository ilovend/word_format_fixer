@echo off

rem 构建Word格式修复工具

rem 切换到项目根目录
cd /d %~dp0

echo 正在构建Word格式修复工具...
echo =============================

rem 检查Python环境
echo 检查Python环境...
python --version
if %errorlevel% neq 0 (
    echo 错误: 未找到Python环境，请先安装Python 3.7+
    pause
    exit /b 1
)

rem 检查Node.js环境
echo 检查Node.js环境...
npm --version
if %errorlevel% neq 0 (
    echo 错误: 未找到Node.js环境，请先安装Node.js 16+
    pause
    exit /b 1
)

rem 安装Python依赖
echo 安装Python依赖...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo 错误: 安装Python依赖失败
    pause
    exit /b 1
)

rem 安装Electron依赖
echo 安装Electron依赖...
cd electron
npm install
if %errorlevel% neq 0 (
    echo 错误: 安装Electron依赖失败
    pause
    exit /b 1
)

rem 构建Electron应用
echo 构建Electron应用...
npm run build
if %errorlevel% neq 0 (
    echo 错误: 构建Electron应用失败
    pause
    exit /b 1
)

cd /d %~dp0

rem 复制必要文件到构建目录
echo 复制必要文件到构建目录...
if not exist build mkdir build
if not exist build\python-backend mkdir build\python-backend
if not exist build\python-backend\core mkdir build\python-backend\core
if not exist build\python-backend\rules mkdir build\python-backend\rules
if not exist build\python-backend\ipc mkdir build\python-backend\ipc
if not exist build\python-backend\config mkdir build\python-backend\config

xcopy python-backend\core build\python-backend\core /s /e
xcopy python-backend\rules build\python-backend\rules /s /e
xcopy python-backend\ipc build\python-backend\ipc /s /e
xcopy python-backend\config build\python-backend\config /s /e
copy python-backend\main.py build\python-backend\main.py
copy requirements.txt build\requirements.txt

rem 创建构建完成的启动脚本
echo 创建构建完成的启动脚本...
(>
@echo off
cd /d %%~dp0
start "Word Format Fixer" python python-backend/main.py
timeout /t 3
start "Word Format Fixer UI" electron\dist\word-format-fixer-electron.exe
) > build\run.bat

echo =============================
echo 构建完成！
echo 构建结果位于 build 目录
pause
