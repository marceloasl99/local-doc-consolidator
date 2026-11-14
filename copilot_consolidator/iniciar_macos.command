#!/bin/bash
cd "$(dirname "$0")"

clear
echo "============================================================================"
echo "  CONSOLIDADOR LOCAL PARA MICROSOFT COPILOT (Multicore - macOS)"
echo "============================================================================"

if ! command -v python3 &> /dev/null; then
    echo "[ERRO] python3 nao encontrado no macOS!"
    echo "Instale via Homebrew (brew install python) ou em https://www.python.org/"
    read -p "Pressione ENTER para sair..."
    exit 1
fi

if [ ! -f ".venv/bin/python3" ]; then
    echo "[1/2] Criando ambiente virtual isolado (.venv)..."
    python3 -m venv .venv
fi

if [ ! -f ".venv/.installed_ok" ]; then
    echo "[2/2] Instalando bibliotecas locais (PDF, Office, OCR offline e Multiprocessing)..."
    ./.venv/bin/python3 -m pip install --upgrade pip
    ./.venv/bin/python3 -m pip install pymupdf pypdf openpyxl python-docx python-pptx pandas xlrd pillow rapidocr-onnxruntime
    touch ".venv/.installed_ok"
fi

./.venv/bin/python3 app.py --lang pt "$@"
