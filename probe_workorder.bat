@echo off
rem 双击启动 probe_workorder.py
rem 支持传参: probe_workorder.bat --url https://xxx

setlocal
cd /d "%~dp0"

set "PY=C:\Users\35144\AppData\Local\Programs\Python\Python312\python.exe"
if not exist "%PY%" (
    echo [错误] 找不到 Python: %PY%
    pause
    exit /b 1
)

"%PY%" "%~dp0probe_workorder.py" %*

echo.
echo ============================================================
echo  跑完看 output\workorder_probe.json 和 probe_*.png 截图
echo ============================================================
pause