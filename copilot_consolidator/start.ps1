<#
.SYNOPSIS
    Local Document Consolidator for Microsoft Copilot (English Launcher)
.DESCRIPTION
    Runs natively inside the active console/terminal window (PowerShell / Windows Terminal).
#>
[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$AppArgs
)

$Host.UI.RawUI.WindowTitle = "Local Document Consolidator for Microsoft Copilot (Windows 11)"
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$VenvPython = Join-Path $AppDir ".venv\Scripts\python.exe"
$AppPy = Join-Path $AppDir "app.py"

if (-not (Test-Path $VenvPython)) {
    Write-Host "[1/2] Local Python environment not initialized. Running setup..." -ForegroundColor Yellow
    Push-Location $AppDir
    python -m venv .venv
    & $VenvPython -m pip install --upgrade pip
    & $VenvPython -m pip install pymupdf pypdf openpyxl python-docx python-pptx pandas xlrd pillow rapidocr-onnxruntime
    Pop-Location
}

Push-Location $AppDir
try {
    & $VenvPython $AppPy --lang en @AppArgs
} finally {
    Pop-Location
}
