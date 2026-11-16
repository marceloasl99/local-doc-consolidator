@echo off
setlocal EnableExtensions
REM Launcher unico para Windows. Uso: _launcher_windows.bat en^|pt [opcoes]
set "BASE=%~dp0"
set "LANG_APP=%~1"
if "%LANG_APP%"=="" set "LANG_APP=pt"
set "ARGS="
:coletar
shift
if "%~1"=="" goto fim_args
set "ARGS=%ARGS% %1"
goto coletar
:fim_args
chcp 65001 >nul
cd /d "%BASE%"
title Document Consolidator for Copilot v7.0

set "VENV_PY=%BASE%.venv\Scripts\python.exe"
if exist "%VENV_PY%" if exist "%BASE%.venv\.installed_ok" goto run_app

REM --- Localiza um Python real (ignora o atalho falso da Microsoft Store) ---
set "PYTHON_CMD="
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul && set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD python -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul && set "PYTHON_CMD=python"
if not defined PYTHON_CMD goto no_python

if /i "%LANG_APP%"=="en" (echo [1/2] Preparing the program - first run only, takes a few minutes...) else (echo [1/2] Preparando o programa - so na primeira vez, leva alguns minutos...)
if not exist "%VENV_PY%" %PYTHON_CMD% -m venv "%BASE%.venv"
if not exist "%VENV_PY%" goto setup_failed
echo [2/2] pip install...
"%VENV_PY%" -m pip install --upgrade pip --disable-pip-version-check -q
"%VENV_PY%" -m pip install -r "%BASE%requirements.txt" --disable-pip-version-check
if errorlevel 1 goto setup_failed
echo ok> "%BASE%.venv\.installed_ok"

:run_app
"%VENV_PY%" "%BASE%app.py" --lang %LANG_APP%%ARGS%
if errorlevel 1 (
    echo.
    if /i "%LANG_APP%"=="en" (echo [ERROR] The program ended with an error. See the messages above.) else (echo [ERRO] O programa terminou com erro. Veja as mensagens acima.)
    pause
)
exit /b 0

:no_python
echo ============================================================================
if /i "%LANG_APP%"=="en" (
    echo  Python 3.10+ was not found on this computer.
    echo  1. Open https://www.python.org/downloads/ and install Python.
    echo  2. On the first installer screen, CHECK "Add python.exe to PATH".
    echo  3. Close this window and open start.bat again.
) else (
    echo  O Python 3.10+ nao foi encontrado neste computador.
    echo  1. Acesse https://www.python.org/downloads/ e instale o Python.
    echo  2. Na primeira tela do instalador, MARQUE "Add python.exe to PATH".
    echo  3. Feche esta janela e abra o iniciar_windows.bat novamente.
)
echo ============================================================================
pause
exit /b 1

:setup_failed
if /i "%LANG_APP%"=="en" (echo [ERROR] Could not install the libraries. Check your internet connection and try again.) else (echo [ERRO] Nao foi possivel instalar as bibliotecas. Verifique a internet e tente de novo.)
if exist "%BASE%.venv\.installed_ok" del "%BASE%.venv\.installed_ok"
pause
exit /b 1
