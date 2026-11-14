@echo off
chcp 65001 >nul
title Local Document Consolidator for Microsoft Copilot (Windows 11)

cd /d "%~dp0copilot_consolidator"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py --lang en %*
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [ERROR] Process finished with error code %ERRORLEVEL%.
        pause
    )
    exit /b %ERRORLEVEL%
)

if exist "start.bat" (
    call "%~dp0copilot_consolidator\start.bat" %*
    exit /b %ERRORLEVEL%
)

echo [ERROR] Application directory or venv not found!
echo Please ensure Python 3.10+ is installed and run start.bat from the project folder.
pause
exit /b 1
