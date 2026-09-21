@echo off
rem 双击启动 probe_one_query.py —— 单条工单查询验证
rem 可选传参: probe_one_query.bat "公司名" "代理" "国家" "服务项目"

setlocal
cd /d "%~dp0"

set "PY=python.exe"
if not exist "%PY%" (
    echo [错误] 找不到 Python: %PY%
    pause
    exit /b 1
)

"%PY%" "%~dp0probe_one_query.py" %*

echo.
echo ============================================================
echo  看上面 "结果表解析状态" 一行:
echo    ok        = 解析正常
echo    failed    = 仍然读不到结果表(把 output\debug\ 最新快照发我)
echo ============================================================
pause
