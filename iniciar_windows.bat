@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

cd /d "%~dp0copilot_consolidator"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py --lang pt %*
    exit /b %ERRORLEVEL%
)

if exist "iniciar_windows.bat" (
    call "%~dp0copilot_consolidator\iniciar_windows.bat" %*
    exit /b %ERRORLEVEL%
)

echo [ERRO] Diretorio do aplicativo ou venv nao encontrado!
exit /b 1
