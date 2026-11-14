@echo off
setlocal
chcp 65001 >nul
title Local Document Consolidator for Microsoft Copilot (Windows 11) v6.0

cd /d "%~dp0copilot_consolidator"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py --lang en %*
    if errorlevel 1 (
        echo.
        echo [ERROR] Process finished with exit code %ERRORLEVEL%.
        pause
    )
    exit /b %ERRORLEVEL%
)

if exist "run.bat" (
    call "%~dp0copilot_consolidator\run.bat" %*
    exit /b %ERRORLEVEL%
)

echo [ERROR] Application directory or virtual environment not found!
echo Please ensure Python 3.10+ is installed and run run.bat or start.bat.
pause
exit /b 1
