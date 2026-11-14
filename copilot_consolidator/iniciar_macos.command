#!/bin/bash
# Duplo clique no Finder (ou: bash "iniciar_macos.command")
exec /bin/bash "$(cd "$(dirname "$0")" && pwd)/_launcher_macos.sh" pt "$@"
