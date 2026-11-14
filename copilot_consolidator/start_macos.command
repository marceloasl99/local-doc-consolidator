#!/bin/bash
cd "$(dirname "$0")"

clear
echo "============================================================================"
echo "  LOCAL DOCUMENT CONSOLIDATOR FOR MICROSOFT COPILOT (Multicore - macOS)"
echo "============================================================================"

if ! command -v python3 &> /dev/null; then
    echo "[ERROR] python3 not found on macOS!"
    echo "Install via Homebrew (brew install python) or from https://www.python.org/"
    read -p "Press ENTER to exit..."
    exit 1
fi

if [ ! -f ".venv/bin/python3" ]; then
    echo "[1/2] Creating local virtual environment (.venv)..."
    python3 -m venv .venv
fi

if [ ! -f ".venv/.installed_ok" ]; then
    echo "[2/2] Installing local libraries (PDF, Office, offline RapidOCR and Multiprocessing)..."
    ./.venv/bin/python3 -m pip install --upgrade pip
    ./.venv/bin/python3 -m pip install pymupdf pypdf openpyxl python-docx python-pptx pandas xlrd pillow rapidocr-onnxruntime
    touch ".venv/.installed_ok"
fi

./.venv/bin/python3 app.py --lang en "$@"
