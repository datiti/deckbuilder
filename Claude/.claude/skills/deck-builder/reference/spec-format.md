# Format de la spec JSON

```json
{
  "title": "Titre du document",
  "template": "modeles/modele-entreprise.potx",
  "theme": { "primary": "1F3A5F", "accent": "00915A", "font": "Arial", "confidentiality": "Interne" },
  "slides": [ ... ]
}
```

`template` et `theme` sont optionnels. Avec un modèle, les polices et couleurs de fond viennent du modèle ; `primary` et `accent` colorent les formes ajoutées.

## Champs communs à toutes les slides

| Champ | Rôle |
|---|---|
| `type` | obligatoire, voir ci-dessous |
| `title` | titre d'action, moins de 90 caractères |
| `message` | bandeau « à retenir » sous le titre (optionnel) |
| `source` | source des données, en pied de slide (optionnel) |
| `notes` | notes de l'orateur (optionnel) |

## Types

**title** : `title`, `subtitle`, `date`

**section** : `title`, `subtitle`

**bullets** : `bullets` = liste de chaînes ; une liste imbriquée crée des sous-puces.
```json
{"type": "bullets", "title": "...", "bullets": ["Point 1", "Point 2", ["Sous-point a", "Sous-point b"], "Point 3"]}
```

**two_columns** : `left` et `right`, chacun `{ "heading": "...", "bullets": [...] }`

**kpis** : `kpis` = jusqu'à 4 objets `{ "value": "35 %", "label": "..." }`, plus `comment` optionnel.

**table** : `columns` (en-têtes), `rows` (liste de lignes), `col_widths` optionnel (fractions dont la somme vaut 1). 8 lignes au maximum par slide.

**roadmap** : `phases` = 3 à 5 objets `{ "label": "...", "period": "...", "items": [...] }`

**matrix** : 2x2. `x_label`, `y_label`, `quadrants` = 4 objets `{ "title": "...", "items": [...] }` dans l'ordre haut-gauche, haut-droite, bas-gauche, bas-droite.
