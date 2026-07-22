@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
title 采购自动化工作台

if not exist "requirements.txt" (
    echo [错误] 未找到 requirements.txt。请保持启动文件与项目文件在同一目录。
    goto :failed
)
if not exist "装潢透视表工具.py" (
    echo [错误] 未找到主程序「装潢透视表工具.py」。
    goto :failed
)

call :find_python
if not defined PYTHON_CMD (
    echo [错误] 未找到可用的 Python 3。请联系维护人员安装 Python 后再重试。
    goto :failed
)

set "VENV_PY=.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo 正在创建本机运行环境，首次运行请稍候...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo [错误] 无法创建 .venv 运行环境。
        goto :failed
    )
    set "NEEDS_DEPENDENCIES=1"
)

if not exist ".venv\.requirements-installed.txt" (
    set "NEEDS_DEPENDENCIES=1"
) else (
    fc /b "requirements.txt" ".venv\.requirements-installed.txt" >nul
    if errorlevel 1 set "NEEDS_DEPENDENCIES=1"
)

if defined NEEDS_DEPENDENCIES (
    echo 正在安装或更新运行依赖，首次运行请保持联网...
    "%VENV_PY%" -m pip install --upgrade pip
    if errorlevel 1 goto :failed
    "%VENV_PY%" -m pip install -r requirements.txt
    if errorlevel 1 goto :failed
    copy /y "requirements.txt" ".venv\.requirements-installed.txt" >nul
)

if not exist ".venv\.playwright-msedge-ready.txt" (
    echo 正在准备 Edge 自动化组件，首次运行请稍候...
    "%VENV_PY%" -m playwright install msedge
    if errorlevel 1 goto :failed
    > ".venv\.playwright-msedge-ready.txt" echo Edge automation component is ready.
)

echo 正在启动采购自动化工作台...
"%VENV_PY%" "装潢透视表工具.py"
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
    echo.
    echo [错误] 工作台异常退出，错误信息已保留在本窗口中。
    pause
)
exit /b %EXIT_CODE%

:find_python
py -3 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
    exit /b 0
)
python --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    exit /b 0
)
exit /b 1

:failed
echo.
echo 请将本窗口截图发送给维护人员。
pause
exit /b 1
