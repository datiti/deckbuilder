#!/usr/bin/env bash
# deck.sh : lance build_deck.py avec le bon Python.
# Priorité au venv du projet (.venv), sinon python3 du système.
# Usage : bash deck.sh <arguments de build_deck.py>
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
if [ -x ".venv/bin/python" ]; then
  PY=".venv/bin/python"
elif command -v python3 >/dev/null 2>&1; then
  PY="python3"
else
  echo "ERREUR : python3 introuvable. Lancer verifier-environnement.sh à la racine du kit." >&2
  exit 2
fi
exec "$PY" "$SCRIPT_DIR/build_deck.py" "$@"
