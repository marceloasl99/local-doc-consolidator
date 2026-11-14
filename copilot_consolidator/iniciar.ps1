<#
.SYNOPSIS
    Consolidador Local de Documentos para Microsoft Copilot (Inicializador em Português)
.DESCRIPTION
    Executa diretamente na janela do terminal ativo (PowerShell / Windows Terminal).
#>
[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$AppArgs
)

$Host.UI.RawUI.WindowTitle = "Consolidador Local de Documentos para Microsoft Copilot (Windows 11)"
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$VenvPython = Join-Path $AppDir ".venv\Scripts\python.exe"
$AppPy = Join-Path $AppDir "app.py"

if (-not (Test-Path $VenvPython)) {
    Write-Host "[1/2] Inicializando ambiente virtual Python local..." -ForegroundColor Yellow
    Push-Location $AppDir
    python -m venv .venv
    & $VenvPython -m pip install --upgrade pip
    & $VenvPython -m pip install pymupdf pypdf openpyxl python-docx python-pptx pandas xlrd pillow rapidocr-onnxruntime
    Pop-Location
}

Push-Location $AppDir
try {
    & $VenvPython $AppPy --lang pt @AppArgs
} finally {
    Pop-Location
}
