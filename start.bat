@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

cd /d "%~dp0copilot_consolidator"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py --lang en %*
    exit /b %ERRORLEVEL%
)

if exist "start.bat" (
    call "%~dp0copilot_consolidator\start.bat" %*
    exit /b %ERRORLEVEL%
)

echo [ERROR] Application directory or venv not found!
exit /b 1
