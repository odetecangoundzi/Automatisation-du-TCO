# 📘 README.md — **TCO Automator**

TCO Automator est une application Web conçue pour automatiser l’intégration des fichiers **DPGF** d’entreprises dans un **TCO** modèle. Elle applique des règles strictes de clean code, sécurité et bonnes pratiques, et s’appuie exclusivement sur **Python, Pandas, OpenPyXL et Streamlit**.

---

# 🧭 1. Objectifs du projet
- Importer un **TCO modèle** (Code, Désignation, Qu., U, Px U. HT, Px Tot HT).
- Importer un **DPGF entreprise** (Code, Désignation, Cc, U, Px U, Px Total).
- Normaliser automatiquement les données (Cc→Qu., extraction des annotations, nettoyage).
- Détecter les incohérences (totaux faux, codes absents, texte dans valeurs numériques).
- Fusionner le DPGF dans le TCO.
- Générer un **TCO Final Excel** propre, consolidé et coloré.

---

# 🗂️ 2. Formats des fichiers attendus
## 📌 TCO (modèle)
```
Code | Désignation | Qu. | U | Px U. HT | Px Tot HT
```

## 📌 DPGF (entrée entreprise)
```
Code | Désignation | Cc | U | Px U | Px Total
```
Les champs numériques peuvent contenir du texte : "SANS OBJET", "COMPRIS", "nc", etc.

## 📌 Export final
```
Code | Désignation | Qu_estimation | U_estimation | PU_estimation | Total_estimation |
EntrepriseX_Qte | EntrepriseX_PU_HT | EntrepriseX_Total | EntrepriseX_Com | …
```

---

# 🧱 3. Architecture du projet
```
project/
│
├── app.py                 # Interface Streamlit
│
├── core/
│   ├── parser_tco.py      # Lecture + validation TCO
│   ├── parser_dpgf.py     # Normalisation + extraction annotations
│   ├── merger.py          # Fusion TCO + Entreprise
│   └── exporter.py        # Export Excel + couleurs
│
├── uploads/               # Fichiers importés
├── outputs/               # Exports générés
└── README.md
```

---

# 🚀 4. Installation
## Environnement
```bash
python -m venv venv
source venv/bin/activate      # Linux/Mac
venv\Scriptsctivate         # Windows
```

## Installation dépendances
```bash
pip install -r requirements.txt
```

## Lancer l’application
```bash
streamlit run app.py
```

---

# 🧠 5. Règles métier
## 5.1 Normalisation DPGF
- Renommer `Cc` → `Qu.`
- Nettoyage des champs (FR→float, suppression espaces parasites)
- Extraction automatique **nombre + texte**
- Annotation envoyée dans `Commentaire`

## 5.2 Détection des erreurs
- **Erreur rouge** : Qu×PU ≠ Total (tolérance 0.01)
- **Avertissement orange** : Code du DPGF non présent dans TCO
- **Note jaune** : Texte dans une cellule numérique
- **Info bleu** : Mot‑clé (SANS OBJET, COMPRIS, nc, P‑M…)

## 5.3 Fusion
- Fusion strictement par colonne `Code`.
- Aucune ligne supprimée.
- Colonnes Entreprise ajoutées dynamiquement.

---

# 🔒 6. Sécurité
## Upload sécurisé
- Accepter **.xlsx** uniquement
- Refuser : `.xlsm`, `.csv`, `.pdf`, `.zip`, `.exe`
- Taille maximale configurable

## Backend sécurisé
- Aucune exécution de code externe
- Jamais de `eval()` ou `exec()`
- Aucun accès réseau

## Données
- Pas de stockage permanent
- Nettoyage périodique de `/uploads` et `/outputs`

---

# 🧼 7. Clean Code — Règles obligatoires
## Structure
- Tout le traitement métier est dans `core/`
- `app.py` contient uniquement l’interface

## Nommage
- Variables et fonctions en `snake_case`
- Noms explicites : `normalize_dpgf()`, `merge_company_into_tco()`, etc.

## Fonctions courtes
- 1 fonction = 1 responsabilité
- Max ~25 lignes

## Commentaires & docstrings
```python
def merge_company_into_tco(tco_df, company_df):
    """
    Fusionne un DPGF normalisé dans le TCO.
    - tco_df : DataFrame du TCO modèle
    - company_df : DataFrame normalisé du DPGF
    Retourne : DataFrame fusionné
    """
```

## Non‑duplication
- Utiliser des helpers pour conversions, regex, couleurs, etc.

## Conformité PEP8
- Largeur max 88–100 chars
- Imports triés

---

# 🔧 8. Bonnes pratiques
- Valider systématiquement les colonnes TCO & DPGF avant traitement.
- Ne jamais modifier la structure du DataFrame TCO.
- Ne jamais supprimer une ligne même si elle est invalide.
- Log léger : actions, pas de contenu sensible.
- Tests unitaires pour cas extrêmes :
  - DPGF sans nombre (ex : "SANS OBJET")
  - Totaux incohérents
  - Codes manquants

---

# 🧪 9. Workflow interne du programme
1. **Parser TCO** → structure valide
2. **Parser DPGF** → normalisation + flags
3. **Fusion** TCO + entreprise
4. **Export Excel** avec couleurs & colonnes masquées

---

# 🤖 10. Règles IA (Antigravity / Gemini)
À coller pour cadrer l’IA :
```
Tu dois respecter strictement la structure du projet.
Tu ne dois utiliser que Streamlit pour l’interface.
Tu ne dois utiliser que Python + Pandas + OpenPyXL pour la logique.
Tu ne dois jamais proposer React, Vue, Angular, Node.js ou Docker.
Tu dois suivre la structure du dossier core/.
Tu dois respecter les colonnes exactes des fichiers TCO et DPGF.
Tes fonctions doivent être courtes, documentées, sans duplication.
Tu dois appliquer toutes les règles de sécurité.
Tu ne dois jamais exécuter de code provenant d’un fichier.
Tu dois aider un développeur débutant à produire un code propre et maintenable.
```

---

# 📌 11. Roadmap
## v1.0
✔ Import TCO
✔ Import DPGF
✔ Normalisation
✔ Fusion
✔ Export coloré

## v1.1
⬜ Coloration cellule par cellule
⬜ Édition UI

## v1.2
⬜ Export PDF
⬜ Comparateur d’entreprises

---

# 📄 12. Licence
MIT

---

# 🎉 Conclusion
TCO Automator est conçu pour être :
- propre
- sécurisé
- simple
- maintenable
- compatible avec le développement assisté par IA

Il sert de base fiable pour automatiser l’intégration des DPGF dans un TCO.
