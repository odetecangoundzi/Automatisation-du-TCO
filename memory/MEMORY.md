# TCO Automator — Mémoire persistante

## Architecture clés
- `app/__init__.py` — CSS Streamlit + logique rendu UI (step 0 landing, step 1 upload, step 2 consolidation)
- `core/parser_tco.py` — Lit le template TCO (DPGF LOT modèle) → DataFrame normalisé
- `core/parser_dpgf.py` — Lit le DPGF entreprise → DataFrame normalisé + alertes
- `core/merger.py` — Fusionne DPGF dans TCO, calcule totaux sections, TVA/TTC
- `core/exporter.py` — Export Excel avec formules dynamiques SUM/multiplication
- `core/utils.py` — `classify_row`, `find_header_row`, `find_column_index`
- `config.py` — TVA_DEFAULT=0.20, MAX_COMPANIES=100

## Patterns de classification `classify_row`
Priorité: Entete > Désignation > Structure code
- Entete `_Niv1`/`_Niv2` → `sub_section`
- Entete `_Art` → `article`
- Entete `Bord_xxx_Recap` → `recap`
- Entete `Bd_xxx_Bord` → `section_header`
- Entete `RecapBord_xxx` → `recap_summary`
- Entete `LignesTot_xxx` → `total_line`
- Heuristique code: 2 parties → section_header, 3 → sub_section, ≥4 → article

## Bugs corrigés (sessions précédentes + session courante)

### [CORRIGÉ] Lignes DPGF absentes du TCO (ex: 03.5.2.5/03.5.2.6)
- Cause: codes absents du template TCO, recap parent introuvable
- Fix: fallback hiérarchique dans merger.py — remonte 03.5.2 → 03.5 → ...
- Fichier: `core/merger.py` (~ligne 190)

### [CORRIGÉ] sub_sections exclues des SUM recap dans l'export Excel (CAS 2,3,4)
- Cause: `exporter.py` ne trackait que les `article` dans `section_articles`, pas les `sub_section`
- Résultat: formule `=SUM(...)` du recap omettait les lignes `_Niv1`/`_Niv2` avec prix
- Exemple: 06.5.3.2 (4130€), 06.5.3.3 (602€), 06.5.3.4 (4648€) manquaient dans "Total BATIMENT G"
- Fix: `if row_type in ("article", "sub_section") and current_section_code:` dans exporter.py
- Fichier: `core/exporter.py` (~ligne 244)

### [CORRIGÉ] Warning silencieux lignes sans code dans merger (CAS 1)
- Cause: merger skippait sans alerter les article/sub_section sans code mais avec montant
- Fix: warning émis si Px_Tot_HT > 0 et code vide
- Fichier: `core/merger.py` (~ligne 147)

## Comportements importants

### _compute_section_totals (merger.py)
- Passe 1: section_headers reçoivent somme de leurs articles + sub_sections enfants
- Passe 2: recap reçoit valeur du section_header parent
- Passe 3: recap_summary reçoit valeur du section_header correspondant
- Passe 4: Montant HT = somme de TOUTES les section_headers (pas imbriquées pour LOT 06)

### exporter.py — Formules Excel
- Articles → `=C*E` (colonne F TCO) et `={qu}*{px}` (colonnes entreprise)
- Sub_sections avec Qu.+Px_U non nuls → formule `=C*E` / `={qu}*{px}` DYNAMIQUE
- Sub_sections sans prix (titres = BATIMENT F) → statique (correct, pas de décomposition)
- Sub_sections sans prix → style FONT_MAIN_TITLE + FILL_MAIN_TITLE (titre principal)
- Section_headers → `=F{recap_row}` via injection DIFFÉRÉE post-boucle (section_header_rows dict)
- Recap → `=SUM(section_articles)` — inclut articles ET sub_sections
- Montant HT → `=SUM(recap_summary)` si recap_summary présents, sinon statique
- TVA → `=F_ht * tva_rate`, TTC → `=F_ht + F_tva`
- Freeze pane : `ws.freeze_panes = "C3"` → correct (lignes 1+2 figées + col A+B)

### data_only=True (openpyxl)
Utilisé dans parse_tco et parse_dpgf pour lire valeurs cachées des formules Excel.

## Session 2025-02 : 7 améliorations appliquées

### (1) Nom fichier export — app.py step 3
- Extrait `tco_meta["project_info"]["lot"]` ou fallback `current_project`
- Normalise : majuscules, espaces/spéciaux → `_` → `TCO_FINAL_LOT06_ELECTRICITE_<ts>.xlsx`

### (2) Bouton Enregistrer — app.py sidebar
- Juste sous le div nom-du-projet : `st.button("💾 Enregistrer")` → `save_project(curr_name, st.session_state)`

### (3) Freeze panes — exporter.py
- `ws.freeze_panes = "C3"` déjà correct : rows 1+2 figés, cols A+B figées. Pas de cellule fusionnée dans la zone figée.

### (4) Match rate — merger.py
- Après la boucle articles DPGF : `match_rate = matched_count / total_dpgf * 100`
- <50% → alerte error rouge; 50-90% → warning orange; ≥90% → OK

### (5) Doublons de code — parser_dpgf.py
- Post-DataFrame : `Code_source` column ajoutée si doublons détectés
- 1ère occurrence garde le code original; suivantes → `CODE_DUP02`
- Warning généré par code dupliqué

### (6) Formules plages — exporter.py
- `_rows_to_sum_formula(col, rows)` : génère `=SUM(F3:F5,F7:F8)` — jamais d'énumération
- 6 points remplacés dans l'export (recap, montant HT, section_header fallback, colonnes entreprise)

### (7) Rouge anomalies — exporter.py + parser_dpgf.py + app.py
- exporter : alerte `type=error` → `FILL_ERROR` sur toute la ligne (1..max_col); warning → seulement colonnes entreprise
- parser_dpgf : détection unité manquante (article avec montant, U vide) → warning
- app.py step 3 : résumé `🔴 N erreur(s) — 🟡 N warn — 🔵 N info` sous les métriques

## Session 2026-02-25 : Conformité style référence (ÉTAPE 3 mission visuelle)

### Objectif
Rendre l'export Excel visuellement identique au fichier de référence `TCO 01 - DESAMIANTAGE - CURAGE - GO.xlsx`.

### Livrables créés
- `tools/style_rapport_reference.md` — inventaire complet du style référence
- `tools/compare_excel_style.py` — outil de validation automatique (checks : largeurs, hauteurs, freeze, formats, polices, fonds, bordures, fusions)

### Modifications exporter.py (style uniquement, logique inchangée)
- Toutes les polices → `name="Tahoma"` (police unique référence)
- FONT_SECTION : `AC2C18` (rouge foncé), bold, 11pt
- FONT_SUB_SECTION : `314E85` (bleu foncé), bold, 9pt — **NOUVEAU**
- FONT_RECAP / FONT_TOTAL : `000000` (noir), bold, 11pt
- FONT_DATA : `000000`, 9pt
- FILL_WHITE = `FFFFFF` solid — toutes les lignes données (FILL_SECTION/RECAP/SUB_SECTION/TOTAL_LINE = FILL_WHITE)
- THICK_TOP_BORDER = THIN_BORDER (plus de medium)
- MONEY_FORMAT = `r'###,###,###,##0.00\ \€;\-###,###,###,##0.00\ \€;'`
- QTY_FORMAT   = `r'###,###,###,##0.00;\-###,###,###,##0.00;'`
- Largeurs colonnes : A=9.5, B=56.75, C=9.5, D=7.125, E=14.125, F=16.5 (exact référence)
- Entreprises : Qu=9.5, PU=14.125, Tot=16.5, Commentaire=25.0
- Hauteur en-têtes : 14.25pt (lignes 1+2)
- Hauteur données : 28.5pt (boucle principale)
- `_get_row_style` : sub_section → FONT_SUB_SECTION

### Résultat validation (compare_excel_style.py)
- Avant : 12 écarts (2 critiques, 9 importants, 1 mineur)
- Après : **2 écarts, 0 critique** → `=> Rendu visuel correct`
- Écarts restants intentionnels : fond coloré HT/TVA/TTC + alertes (fonctionnel), pas de fusion A:B (cosmétique)

## Fichiers de test disponibles
- `test_discrepancy.py`, `test_discrepancy_lot06.py` — diagnostics LOT 01 et LOT 06
- `check_lot.py` — inspection rapide d'un fichier DPGF
- `verify_fixes.py`, `compare_templates.py` — scripts de vérification

## Dossier entreprise (fichiers testés)
- `DEVIS DPGF GUSTAVE EIFFEL BDX - Exemplaire Client.xlsx` → LOT 06 ELECTRICITE
- `DPGF B2R - BAT EXTERNATS LYCEE EIFFEL BDX - LOT 03 - PLATRERIE.xlsx`
- `14-DE-20251282 - DPGF LOT 02 - MENUISERIES - SERRURERIE - APPS MUSCULATION.xlsx`
- `DPGF LOT 04 - REVETEMENT INTERIEUR ET EXTERIEUR - PEINTURE - SIGNALETIQUE.xlsx`
