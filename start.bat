@echo off

rem 启动Word格式修复工具

rem 切换到项目根目录
cd /d %~dp0

echo 正在启动Word格式修复工具...
echo =============================

rem 检查Python环境
echo 检查Python环境...
python --version
if %errorlevel% neq 0 (
    echo 错误: 未找到Python环境，请先安装Python 3.7+
    pause
    exit /b 1
)

rem 安装依赖
echo 安装Python依赖...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo 错误: 安装依赖失败
    pause
    exit /b 1
)

rem 启动Python后端服务
echo 启动Python后端服务...
start "Python Backend" python python-backend/main.py

rem 等待后端服务启动
echo 等待后端服务启动...
timeout /t 3 /nobreak > nul

rem 启动Electron前端
echo 启动Electron前端...
cd electron
if exist "node_modules" (
    npm start
) else (
    echo 安装Electron依赖...
    npm install
    if %errorlevel% neq 0 (
        echo 错误: 安装Electron依赖失败
        pause
        exit /b 1
    )
    npm start
)

cd /d %~dp0
pause
