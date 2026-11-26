---
name: deck-builder
description: Fabrique un deck PowerPoint (.pptx) à partir d'un brief, d'un document ou d'une spec JSON. À utiliser dès que l'utilisateur demande un deck, des slides, une présentation, un support COMEX ou CODIR, ou veut modifier un .pptx généré.
allowed-tools: Bash(bash ${CLAUDE_SKILL_DIR}/scripts/*)
---

# Fabriquer un deck PowerPoint

Tu produis des decks de niveau direction : un message par slide, des titres d'action, aucun chiffre inventé.
Le rendu est fait par un script déterministe ; ton travail est la storyline et la spec JSON, pas le code de mise en page.

## Étapes

1. **Cadrer** (une seule question groupée si l'info manque) : public, décision attendue, nombre de slides (défaut : 8 maximum), modèle PowerPoint à utiliser (fichier `.potx` ou `.pptx` dans `modeles/`).
2. **Storyline d'abord** : affiche dans le chat un tableau « n° | type | titre d'action ». Le titre d'action est une phrase qui énonce la conclusion (« Le parc applicatif concentre l'essentiel du potentiel »), jamais un thème (« Parc applicatif »). Continue sans attendre sauf si l'utilisateur a demandé à valider.
3. **Écrire la spec** dans `decks/<nom-court>.json` au format décrit dans [reference/spec-format.md](reference/spec-format.md).
4. **Contrôler** :
   `bash "${CLAUDE_SKILL_DIR}/scripts/deck.sh" decks/<nom>.json --lint-only`
   Corrige chaque avertissement (raccourcir, scinder) et relance jusqu'à zéro avertissement.
5. **Générer** :
   `bash "${CLAUDE_SKILL_DIR}/scripts/deck.sh" decks/<nom>.json -o sortie/<nom>.pptx [--template modeles/<modele>.potx]`
   Avant le premier usage d'un modèle, lance `bash "${CLAUDE_SKILL_DIR}/scripts/deck.sh" --list-layouts modeles/<modele>.potx` pour vérifier qu'il contient un layout « Titre seul » ou « Title Only ».
6. **Vérifier visuellement** (Mac avec PowerPoint installé) :
   `bash "${CLAUDE_SKILL_DIR}/scripts/rendu-pdf.sh" sortie/<nom>.pptx`
   Puis lis le PDF produit avec l'outil Read et corrige : texte qui déborde, titre sur trois lignes, slide vide ou surchargée.
   Si l'export échoue (PowerPoint absent, autorisation refusée), dis-le et demande à l'utilisateur d'ouvrir le fichier.
7. **Restituer** : chemin du fichier, nombre de slides, liste des `[À COMPLÉTER]`, hypothèses prises.

`deck.sh` utilise le venv `.venv` du projet s'il existe, sinon `python3`. Si `python-pptx` manque, arrête-toi et renvoie vers `bash verifier-environnement.sh` au lieu de bricoler une alternative.

## Règles de contenu

- Aucun chiffre, nom, date ou citation qui ne vient pas des documents fournis ou de l'utilisateur. Une donnée manquante s'écrit `[À COMPLÉTER : description]`.
- Indique la source de chaque chiffre dans le champ `source` de la slide.
- 6 puces de premier niveau au maximum, 12 mots par puce visés.
- Le champ `message` porte la phrase que le lecteur doit retenir ; utilise-le sur les slides qui portent une décision.
- Mets ce que l'orateur doit dire dans `notes`, pas sur la slide.
- Pas de tiret cadratin : utilise virgule ou deux-points.

## Modifier un deck existant

Modifie la spec JSON puis régénère. Ne modifie pas le `.pptx` directement : la spec est la source de vérité, c'est ce qui rend l'itération rapide et traçable.
Pour reprendre un deck PowerPoint fait à la main, extrais d'abord son texte (python-pptx), propose une spec équivalente, puis régénère.
