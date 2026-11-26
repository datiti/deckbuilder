#!/usr/bin/env bash
# rendu-pdf.sh : exporte un .pptx en PDF via PowerPoint pour Mac (AppleScript).
# Claude Code lit ensuite le PDF avec l'outil Read pour vérifier le rendu.
# Usage : bash rendu-pdf.sh sortie/deck.pptx [sortie/deck.pdf]
#
# Premier lancement : macOS demande d'autoriser le Terminal à contrôler PowerPoint
# (Réglages Système > Confidentialité et sécurité > Automatisation), puis PowerPoint
# peut demander l'accès au dossier. Accepter les deux AVANT la démo.
set -euo pipefail
[ $# -ge 1 ] || { echo "Usage : bash rendu-pdf.sh fichier.pptx [fichier.pdf]" >&2; exit 1; }
[ -f "$1" ] || { echo "ERREUR : fichier introuvable : $1" >&2; exit 1; }

IN="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
OUT_REL="${2:-${1%.*}.pdf}"
mkdir -p "$(dirname "$OUT_REL")"
OUT="$(cd "$(dirname "$OUT_REL")" && pwd)/$(basename "$OUT_REL")"
rm -f "$OUT"

if ! osascript -e 'id of application "Microsoft PowerPoint"' >/dev/null 2>&1; then
  echo "ERREUR : Microsoft PowerPoint n'est pas installé. Ouvrir le .pptx à la main pour vérifier." >&2
  exit 3
fi

osascript <<APPLESCRIPT
set inFile to POSIX file "$IN"
set outPath to (POSIX file "$OUT") as text
set wasRunning to application "Microsoft PowerPoint" is running
tell application "Microsoft PowerPoint"
    open inFile
    delay 3
    set p to active presentation
    save p in outPath as save as PDF
    close p saving no
    if not wasRunning then quit
end tell
APPLESCRIPT

# l'écriture du PDF peut être asynchrone
for _ in 1 2 3 4 5 6 7 8 9 10; do
  [ -s "$OUT" ] && { echo "OK : $OUT"; exit 0; }
  sleep 1
done
echo "ERREUR : PDF non produit. Vérifier les autorisations Automatisation et accès fichiers de PowerPoint." >&2
exit 4
