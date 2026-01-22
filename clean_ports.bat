@echo off
echo 🔥 清理端口占用进程...

:: 杀死Python相关进程
taskkill /f /im python.exe 2>nul
taskkill /f /im pythonw.exe 2>nul

:: 杀死Node.js相关进程  
taskkill /f /im node.exe 2>nul
taskkill /f /im electron.exe 2>nul

:: 强制释放指定端口范围
for /l %%p in (7777,1,7790) do (
    echo 检查端口 %%p...
    for /f "tokens=5" %%i in ('netstat -ano ^| findstr ":%%p"') do (
        if not "%%i"=="" (
            echo 终止占用端口 %%p 的进程 PID: %%i
            taskkill /f /pid %%i 2>nul
        )
    )
)

echo ✅ 端口清理完成
timeout 2
