@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
title Consolidador Local de Documentos para Microsoft Copilot (Windows 11)

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
    echo [ERRO] Python nao encontrado no Windows!
    echo Instale o Python 3.10+ em https://www.python.org/downloads/
    echo IMPORTANTE: Marque a caixa "Add Python to PATH" no instalador.
    echo ============================================================================
    pause
    exit /b 1
)

if exist ".venv\Scripts\python.exe" goto check_packages

echo [1/2] Criando ambiente virtual local venv...
%PYTHON_CMD% -m venv .venv
if %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Falha ao criar o ambiente virtual .venv
    pause
    exit /b 1
)

:check_packages
if exist ".venv\.installed_ok" goto run_app

echo [2/2] Instalando bibliotecas locais de PDF, Office e OCR (100%% offline)...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install pymupdf pypdf openpyxl python-docx python-pptx pandas xlrd pillow rapidocr-onnxruntime
if %ERRORLEVEL% EQU 0 (
    echo ok > ".venv\.installed_ok"
) else (
    echo [AVISO] Ocorreu um aviso na instalacao de dependencias.
)

:run_app
".venv\Scripts\python.exe" app.py --lang pt %*
pause
