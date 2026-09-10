@echo off
chcp 65001 >nul 2>&1
cd /d "<PROJECT_ROOT>"

set "PY=python.exe"
if not exist "%PY%" (
    echo [错误] 找不到 Python: %PY%
    pause
    exit /b 1
)

if not exist logs mkdir logs

rem stdout/stderr 全部重定向到日志文件: 即使控制台被选中/关闭也不阻塞 worker 线程
"%PY%" main.py >> logs\console.log 2>&1

echo.
echo 程序已退出, 运行日志见 logs\console.log
pause
