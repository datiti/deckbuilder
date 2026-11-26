#!/usr/bin/env bash
# verifier-environnement.sh : à lancer AUJOURD'HUI sur le Mac de démo, depuis la racine du kit.
# Usage : bash verifier-environnement.sh
cd "$(dirname "$0")"
OK=1
ok()  { printf "\033[32m[OK]\033[0m %s\n" "$1"; }
ko()  { printf "\033[31m[KO]\033[0m %s\n     -> %s\n" "$1" "$2"; OK=0; }
SK=".claude/skills/deck-builder/scripts"

# 1. Claude Code
if command -v claude >/dev/null 2>&1; then ok "Claude Code : $(claude --version 2>/dev/null | head -1)"
else ko "Claude Code introuvable dans le PATH" "Vérifier l'installation de Claude Code."; fi

# 2. Python 3
if command -v python3 >/dev/null 2>&1 && python3 --version >/dev/null 2>&1; then ok "Python : $(python3 --version 2>&1)"
else ko "Python 3 absent" "xcode-select --install (outils en ligne de commande) ou demander au support."; fi

# 3. python-pptx : venv du projet, sinon python3 système
if [ -x .venv/bin/python ] && .venv/bin/python -c "import pptx" 2>/dev/null; then
  ok "python-pptx (venv .venv) : $(.venv/bin/python -c 'import pptx;print(pptx.__version__)')"
elif python3 -c "import pptx" 2>/dev/null; then
  ok "python-pptx (système) : $(python3 -c 'import pptx;print(pptx.__version__)')"
else
  echo "     python-pptx absent : tentative d'installation dans un venv local (.venv)..."
  python3 -m venv .venv && .venv/bin/pip install -q python-pptx \
    && ok "python-pptx installé dans .venv" \
    || ko "Installation de python-pptx impossible" \
          "Voir le miroir PyPI interne : 'pip3 config list', puis .venv/bin/pip install --index-url <URL miroir> python-pptx"
fi

# 4. Génération du deck exemple
if OUT=$(bash "$SK/deck.sh" exemples/deck-exemple.json -o sortie/deck-exemple.pptx 2>&1); then ok "Génération : $(echo "$OUT" | tail -1)"
else ko "Génération du deck exemple" "$OUT"; fi

# 5. Export PDF via PowerPoint (déclenche les demandes d'autorisation macOS la première fois)
if [ -f sortie/deck-exemple.pptx ] && OUT=$(bash "$SK/rendu-pdf.sh" sortie/deck-exemple.pptx 2>&1); then ok "Export PDF : $OUT"
else ko "Export PDF via PowerPoint" "Autorisations refusées ou bloquées par la gestion du Mac : la démo reste possible, vérification visuelle à la main. Détail : $OUT"; fi

echo
[ $OK -eq 1 ] && printf "\033[32mEnvironnement prêt pour la démo.\033[0m\n" \
              || printf "\033[31mAu moins un point KO : corriger aujourd'hui ou basculer sur le plan B (DEMO.md).\033[0m\n"
