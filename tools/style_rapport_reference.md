# Rapport de style — Référence TCO 01 - DESAMIANTAGE - CURAGE - GO.xlsx

Généré lors de la session du 2026-02-25.
Outil : openpyxl `load_workbook(data_only=True)`.

---

## 1. Feuilles

| N° | Nom | État |
|----|-----|------|
| 0 | Feuille 1 | visible |

Dimensions relevées : **A1:X138**

---

## 2. Freeze panes et filtres

| Propriété | Valeur |
|-----------|--------|
| freeze_panes | `A115` (lignes 1–114 figées, aucune colonne figée) |
| auto_filter | Aucun |
| print_title_rows | Aucun |

---

## 3. Mise en page

| Propriété | Valeur |
|-----------|--------|
| Orientation | portrait |
| Taille papier | non définie |
| Marges (gauche/droite/haut/bas/entête/pied) | 0.0 / 0.0 / 0.0 / 0.0 / 0.0 / 0.0 |

---

## 4. Largeurs de colonnes

| Colonne | Largeur | Masquée |
|---------|---------|---------|
| A (Code) | 9.5 | Non |
| B (Désignation) | 56.75 | Non |
| C (Qté) | 9.5 | Non |
| D (Unité) | 7.125 | Non |
| E (Px U. HT) | 14.125 | Non |
| F (Px Tot HT) | 16.5 | Non |
| G (UUID interne) | 25.0 | **Oui** |
| K (Qté entreprise 1) | 9.5 | Non |
| L (Px U. HT entreprise 1) | 14.125 | Non |
| M (Px Tot HT entreprise 1) | 16.5 | Non |
| N (Commentaire entreprise 1) | 14.125 | Non |
| O (UUID interne 2) | 25.0 | **Oui** |
| P (Qté entreprise 2) | 9.5 | Non |
| Q (Px U. HT entreprise 2) | 14.125 | Non |
| R (Px Tot HT entreprise 2) | 16.5 | Non |
| S (Commentaire entreprise 2) | 14.125 | Non |
| T (UUID interne 3) | 25.0 | **Oui** |
| U (Qté entreprise 3) | 9.5 | Non |
| V (Px U. HT entreprise 3) | 14.125 | Non |
| W (Px Tot HT entreprise 3) | 16.5 | Non |
| X (Commentaire entreprise 3) | 14.125 | Non |

_H, I, J : non définis (largeur par défaut ~8.43, colonnes visibles/vides dans le template)._

---

## 5. Hauteurs de lignes

| Plage | Hauteur (pt) | Note |
|-------|-------------|------|
| Ligne 1 | 14.25 | En-tête groupe (labels phase) |
| Ligne 2 | 14.25 | En-tête colonnes |
| Lignes 3–129 | **28.5** | Corps de données (hauteur double) |
| Lignes 130–131 | 14.25 | Zone résumé |
| Lignes 132–133 | 15.0 | Zone résumé |
| Ligne 134 | **56.25** | Ligne signature (très haute) |
| Lignes 135–136 | 15.0 | Bas de page |
| Ligne 137 | 28.5 | |
| Ligne 138 | 15.0 | |

---

## 6. Cellules fusionnées (34 plages)

### Fusions en-têtes (ligne 1 — groupes de colonnes)
- `C1:F1` — groupe "Estimation" (Qté/U/Px U/Px Tot)
- `K1:N1` — groupe entreprise 1
- `P1:S1` — groupe entreprise 2
- `U1:X1` — groupe entreprise 3

### Fusions lignes totaux (colonne A:B fusionnées)
- Environ 27 fusions A:B sur les lignes de totaux/sous-totaux (ex : `A15:B15`, `A25:B25`, …)
- Ces fusions permettent au label "Total XYZ :" de s'étendre sur Code + Désignation.

### Fusions zone résumé (bas de feuille)
- `K134:N134`, `P134:S134`, `U134:X134` — ligne signature fusionnée par groupe

---

## 7. Styles de cellules

### Police (UNIQUE : Tahoma partout)

| Usage | Taille | Gras | Couleur hex |
|-------|--------|------|-------------|
| Sections niveau 1 | **11 pt** | Oui | `AC2C18` (rouge foncé) |
| Sous-sections niv. 2–3 | 9 pt | Oui | `314E85` (bleu foncé) |
| Articles (lignes détail) | 9 pt | Non | `000000` (noir) |
| Variants offre (sous-articles) | 9 pt | Non | `DC9329` (ambre/orange) |
| Lignes totaux / récaps | **11 pt** | Oui | `000000` (noir) |
| Commentaires (col N) | 9 pt | Oui | `1800FC` (bleu vif) |
| En-tête phase (ligne 1) | 10 pt | Oui | `75A4F8` (bleu clair) / `A7AAA9` (gris) |
| En-tête colonnes (ligne 2) | 10 pt | Oui | identique ligne 1 |

### Couleur de fond (fill)

| Zone | Couleur | Pattern |
|------|---------|---------|
| **Toutes les lignes de données** | `FFFFFF` (blanc pur) | solid |
| En-têtes ligne 1–2 | `00000000` (transparent / aucun) | — |

> **Règle clé : le fichier de référence n'utilise AUCUN fond coloré sur les données.
> La hiérarchie visuelle est portée UNIQUEMENT par la couleur et la taille de police.**

### Bordures

- Style unique : **`thin`** sur tous les côtés (aucun medium/thick).
- ~3 920 côtés avec bordure thin.
- Toutes les cellules A–N des lignes 3–100 ont des bordures.

### Alignement

- Général : horizontal `None` (left par défaut), vertical `None`
- Pas de `wrap_text` explicite détecté dans les premiers 100 lignes.
- Indentation : non utilisée dans la référence.

---

## 8. Formats numériques

| Format | Nb cellules | Usage |
|--------|-------------|-------|
| `###,###,###,##0.00\ \€;\-###,###,###,##0.00\ \€;` | 332 | Montants en euros |
| `###,###,###,##0.00;\-###,###,###,##0.00;` | 139 | Quantités décimales |
| `General` | — | Codes, désignations, unités |

---

## 9. Formatage conditionnel

Aucun formatage conditionnel détecté.

---

## 10. Colonnes/lignes masquées

| Colonne | Statut |
|---------|--------|
| G | Masquée (colonne UUID interne) |
| O | Masquée (UUID interne entreprise 2) |
| T | Masquée (UUID interne entreprise 3) |

Aucune ligne masquée détectée.

---

## Récapitulatif des règles visuelles à reproduire

1. **Police Tahoma** — utilisée partout sans exception.
2. **Fond blanc pur** (`FFFFFF`) pour toutes les lignes de données.
3. **Hiérarchie par couleur de police** :
   - Sections → rouge foncé `AC2C18`, bold, 11 pt
   - Sous-sections → bleu foncé `314E85`, bold, 9 pt
   - Articles → noir `000000`, normal, 9 pt
   - Totaux/récaps → noir `000000`, bold, 11 pt
4. **Hauteur des lignes de données** : **28.5 pt** uniformément.
5. **Largeurs de colonnes** : voir tableau §4.
6. **Bordures uniquement thin** — aucune bordure medium ou thick.
7. **Format euro** : `###,###,###,##0.00\ \€;\-###,###,###,##0.00\ \€;`
8. **Format quantité** : `###,###,###,##0.00;\-###,###,###,##0.00;`
9. **Freeze panes** : le template fige les lignes 1–114 (pas de colonne figée).
   Notre export utilise `C3` (lignes 1–2 + colonnes A–B figées) — adapté à notre structure 2 lignes d'en-tête.
