#!/bin/bash
# Launcher unico para macOS / Linux. Uso: _launcher_macos.sh en|pt [opcoes]
LANG_APP="${1:-pt}"
shift
BASE="$(cd "$(dirname "$0")" && pwd)"
cd "$BASE" || exit 1

msg() { if [ "$LANG_APP" = "en" ]; then echo "$2"; else echo "$1"; fi; }
pausar() { echo; read -r -p "$(msg 'Pressione ENTER para fechar...' 'Press ENTER to close...')" _; }

VENV_PY="$BASE/.venv/bin/python3"

if [ ! -x "$VENV_PY" ] || [ ! -f "$BASE/.venv/.installed_ok" ]; then
    # Procura um Python 3.10+ de verdade (o python3 do sistema costuma ser 3.9)
    PYTHON_CMD=""
    for c in python3.14 python3.13 python3.12 python3.11 python3.10 \
             /opt/homebrew/bin/python3 /usr/local/bin/python3 \
             /Library/Frameworks/Python.framework/Versions/Current/bin/python3 python3; do
        if command -v "$c" >/dev/null 2>&1 && "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' >/dev/null 2>&1; then
            PYTHON_CMD="$c"; break
        fi
    done
    if [ -z "$PYTHON_CMD" ]; then
        echo "============================================================================"
        msg " O Python 3.10+ nao foi encontrado neste Mac." " Python 3.10+ was not found on this Mac."
        msg " 1. Acesse https://www.python.org/downloads/ e instale o Python (botao amarelo)." \
            " 1. Open https://www.python.org/downloads/ and install Python (yellow button)."
        msg " 2. Feche esta janela e abra o programa novamente." " 2. Close this window and open the program again."
        echo "============================================================================"
        pausar; exit 1
    fi
    msg "[1/2] Preparando o programa - so na primeira vez, leva alguns minutos..." \
        "[1/2] Preparing the program - first run only, takes a few minutes..."
    [ -x "$VENV_PY" ] || "$PYTHON_CMD" -m venv "$BASE/.venv"
    echo "[2/2] pip install..."
    "$VENV_PY" -m pip install --upgrade pip --disable-pip-version-check -q
    if ! "$VENV_PY" -m pip install -r "$BASE/requirements.txt" --disable-pip-version-check; then
        msg "[ERRO] Nao foi possivel instalar as bibliotecas. Verifique a internet e tente de novo." \
            "[ERROR] Could not install the libraries. Check your internet connection and try again."
        pausar; exit 1
    fi
    touch "$BASE/.venv/.installed_ok"
fi

"$VENV_PY" "$BASE/app.py" --lang "$LANG_APP" "$@"
if [ $? -ne 0 ]; then
    msg "[ERRO] O programa terminou com erro. Veja as mensagens acima." "[ERROR] The program ended with an error. See the messages above."
    pausar
fi
