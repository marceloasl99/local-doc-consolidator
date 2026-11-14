@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
title Local Document Consolidator for Microsoft Copilot (Windows 11)

cd /d "%~dp0"

set PYTHON_CMD=
where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set PYTHON_CMD=python
) else (
    where py >nul 2>nul
    if !ERRORLEVEL! EQU 0 (
        set PYTHON_CMD=py -3
    )
)

if "%PYTHON_CMD%"=="" (
    echo ============================================================================
    echo [ERROR] Python not found on Windows!
    echo Please install Python 3.10+ at https://www.python.org/downloads/
    echo IMPORTANT: Check the "Add Python to PATH" box in the installer.
    echo ============================================================================
    pause
    exit /b 1
)

if exist ".venv\Scripts\python.exe" goto check_packages

echo [1/2] Creating local virtual environment (.venv)...
%PYTHON_CMD% -m venv .venv
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to create .venv virtual environment
    pause
    exit /b 1
)

:check_packages
if exist ".venv\.installed_ok" goto run_app

echo [2/2] Installing local PDF, Office and OCR libraries (100%% offline)...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install pymupdf pypdf openpyxl python-docx python-pptx pandas xlrd pillow rapidocr-onnxruntime
if %ERRORLEVEL% EQU 0 (
    echo ok > ".venv\.installed_ok"
) else (
    echo [WARNING] Dependency installation finished with notices.
)

:run_app
".venv\Scripts\python.exe" app.py --lang en %*
pause
