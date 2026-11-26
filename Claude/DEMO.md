# Démo : fabriquer des decks PowerPoint avec Claude Code

Durée : 15 minutes. Message central à faire passer :
**Claude Code ne « dessine » pas des slides : il écrit la storyline, la traduit en spec, exécute un script qui applique notre charte, puis relit le résultat.** La qualité vient du skill et du modèle PowerPoint, pas de la chance du prompt.

---

## Ce soir (obligatoire)

1. Dézipper le kit dans un dossier de travail du Mac pro, par exemple `~/decks-demo/`.
2. Dans le Terminal, depuis ce dossier : `bash verifier-environnement.sh`. Tout point KO se règle ce soir.
3. Pendant l'étape d'export PDF, macOS demande d'autoriser le Terminal à contrôler PowerPoint, puis PowerPoint peut demander l'accès au dossier : accepter les deux. Si la gestion du Mac (MDM) bloque l'Automatisation, tu le sauras ce soir et pas devant ta cheffe.
4. Déposer le modèle PowerPoint officiel dans `modeles/` et générer une fois le deck exemple avec `--template` : c'est ce rendu que ta cheffe jugera, pas le style par défaut.
5. Lancer `claude` une première fois et accepter la permission sur les scripts du skill (« Yes, don't ask again ») : sinon chaque commande demandera une validation pendant la démo.
6. Répéter les 4 prompts ci-dessous une fois, en entier. Garder les fichiers produits : ce sont ton plan B.

## Déroulé

Ouvrir le Terminal dans le dossier du kit, lancer `claude`.

### 1. Le cadre (2 min, sans clavier)
Montrer l'arborescence : `CLAUDE.md` (règles de l'espace), `.claude/skills/deck-builder/` (le savoir-faire), `decks/` (les specs), `sortie/` (les pptx).
Point clé : le skill est un fichier texte versionnable, partageable avec toute l'équipe.

### 2. Du brief au deck (5 min)
```
Lis exemples/brief-exemple.md et fais-en un deck de 6 slides pour le comité d'architecture,
avec le modèle modeles/<ton-modele>.potx. Montre-moi d'abord la storyline.
```
À commenter pendant que ça tourne : il affiche la storyline en titres d'action, écrit la spec JSON, lance le contrôle, corrige ses avertissements, génère, exporte en PDF via PowerPoint et relit le PDF.
Ouvrir le pptx produit dans PowerPoint.

### 3. L'itération (3 min)
```
Slide 3 : remplace la liste par une matrice valeur / faisabilité.
Ajoute des notes orateur sur chaque slide. Raccourcis tous les titres à moins de 70 caractères.
```
Point clé : on modifie la spec, pas le pptx. Chaque version est traçable et régénérable en 2 secondes.

### 4. Le garde-fou (2 min)
```
Ajoute une slide avec le budget du pilote et le ROI attendu.
```
Résultat attendu : des `[À COMPLÉTER]` au lieu de chiffres inventés, car le brief n'en contient pas. C'est la réponse à la question « et s'il invente ? ».

### 5. Ouverture (3 min)
Autres usages du même mécanisme : fiche de décision d'architecture, synthèse d'un document long en 5 slides COMEX, mise au gabarit d'un deck existant.

## Questions probables de ta cheffe

| Question | Réponse |
|---|---|
| Les données saisies sortent-elles de la banque ? | Dépend du contrat entreprise et de la configuration validée par la sécurité ; à confirmer auprès d'eux, ne pas improviser. Pour la démo : données fictives uniquement. |
| C'est mieux qu'un consultant ou que Copilot dans PowerPoint ? | Ce n'est pas plus beau, c'est plus rapide sur la structure et reproductible : la storyline et la charte sont figées dans le skill. La finition graphique reste humaine. |
| Combien de temps gagné ? | Ne pas donner de chiffre non mesuré. Proposer de le mesurer sur un pilote. |
| Qui maintient le skill ? | Un fichier texte et un script de 500 lignes, versionnés dans un dépôt d'équipe. |

## Plan B

| Incident | Parade |
|---|---|
| `python-pptx` non installable (pas d'accès au miroir PyPI) | Demander à Claude Code d'écrire lui-même le pptx en OpenXML pur (zipfile, bibliothèque standard). Plus lent et plus fragile : à n'utiliser qu'en secours, et ne le montrer en démo que si tu l'as répété. |
| Export PDF bloqué (Automatisation refusée) | Ouvrir le pptx à la main ; la démo reste valable. |
| Claude Code lent ou indisponible | Montrer les fichiers produits la veille (spec JSON + pptx) et dérouler l'explication. |
| Le kit ne peut pas être copié sur le poste pro | Coller le contenu de `SKILL.md` et de `spec-format.md` dans Claude Code et lui demander d'écrire le script de rendu. Démo plus longue mais encore plus parlante. |
