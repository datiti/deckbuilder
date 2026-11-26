# Kit Claude Code : decks PowerPoint (macOS)

1. Dézipper ce dossier, ouvrir un Terminal dedans.
2. `bash verifier-environnement.sh` (installe python-pptx dans un venv local si besoin, teste la génération et l'export PDF).
3. Déposer le modèle PowerPoint de l'entreprise dans `modeles/`.
4. Lancer `claude` dans ce dossier et demander un deck en langage naturel.

Contenu :
- `CLAUDE.md` : règles de l'espace de travail
- `.claude/skills/deck-builder/` : le skill (méthode, format de spec, script de rendu python-pptx, export PDF via PowerPoint)
- `exemples/` : un brief et une spec de démonstration (données fictives)
- `DEMO.md` : déroulé de la démo, prompts, questions attendues, plan B

Seule dépendance : `python-pptx`.
