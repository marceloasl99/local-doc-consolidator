@echo off
chcp 65001 >nul
title Consolidador Local de Documentos para Microsoft Copilot (Windows 11)

cd /d "%~dp0copilot_consolidator"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py --lang pt %*
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [ERRO] O processo finalizou com codigo de erro %ERRORLEVEL%.
        pause
    )
    exit /b %ERRORLEVEL%
)

if exist "iniciar_windows.bat" (
    call "%~dp0copilot_consolidator\iniciar_windows.bat" %*
    exit /b %ERRORLEVEL%
)

echo [ERRO] Diretorio do aplicativo ou venv nao encontrado!
pause
exit /b 1
