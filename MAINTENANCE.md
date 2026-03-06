# TCO Automator — Guide complet (version locale .exe)

> **Pour qui ?** Ce guide s'adresse a toute personne qui utilise ou maintient l'application TCO Automator installee en local sur un poste Windows, meme sans connaissance en informatique ou en developpement logiciel.

---

## Partie 1 — Guide d'utilisation

### Qu'est-ce que TCO Automator ?

TCO Automator est un logiciel installe sur votre poste (fichier `TCO_Automator.exe`). Il automatise la construction du **Tableau Comparatif des Offres (TCO)**. A partir du bordereau de prix de reference (le "TCO modele") et des DPGF remis par les entreprises soumissionnaires, il produit automatiquement un fichier Excel de comparaison pret a l'emploi, avec mise en evidence des anomalies.

Le logiciel fonctionne entierement en local : **aucune connexion internet n'est necessaire**, vos donnees ne quittent jamais votre poste.

---

### 1.1 Lancer l'application

1. Double-cliquez sur `TCO_Automator.exe`.
2. Une fenetre noire (console) s'ouvre — **c'est normal**, ne la fermez pas, elle fait tourner le logiciel en arriere-plan.
3. Quelques secondes plus tard, votre navigateur internet s'ouvre automatiquement sur l'application.
4. Si le navigateur ne s'ouvre pas tout seul, ouvrez-le manuellement et tapez l'adresse : `http://localhost:8501`

> **Important :** Ne fermez jamais la fenetre noire pendant que vous travaillez. La fermer revient a eteindre le logiciel. Fermez-la uniquement quand vous avez termine votre session.

**Structure des dossiers crees automatiquement a cote de l'exe :**

```
TCO_Automator.exe
projects/       <- vos projets sauvegardes
logs/           <- journaux d'erreur
uploads/        <- fichiers temporaires (vides apres chaque session)
```

---

### 1.2 Creer un nouveau projet

Lors du premier lancement, la barre laterale (a gauche) propose de **creer un nouveau projet** ou d'en **charger un existant**.

1. Dans le champ "Nom du projet", tapez un nom descriptif (ex : `Bergerac_LOT03_2025`).
   - Evitez les accents, les espaces et les caracteres speciaux (`/`, `\`, `:`).
2. Cliquez sur **Creer le projet**.
3. Le projet est immediatement cree et pret a recevoir des donnees.

Pour **sauvegarder** votre travail en cours, cliquez sur le bouton **Sauvegarder le projet** dans la barre laterale. Le fichier `.tco.json.gz` est conserve dans le dossier `projects/` a cote de l'exe.

> **Partager un projet entre postes :** Copiez simplement le fichier `.tco.json.gz` du dossier `projects/` d'un poste a l'autre.

---

### 1.3 Etape 1 — Importer le TCO modele

Le TCO modele est votre bordereau de reference, celui qui liste tous les postes du marche.

1. Cliquez sur la section **Etape 1**.
2. Cliquez sur **Parcourir** et selectionnez votre fichier TCO (format `.xlsx` ou `.xlsm`).
3. L'application analyse le fichier et affiche un apercu du tableau avec le nombre de postes detectes.
4. Si le fichier est valide, un message vert confirme l'import.

> **En cas d'erreur a cette etape :** Verifiez que votre fichier TCO respecte bien le format attendu (colonnes "Code", "Designation", "U.", "Quantite", "PU HT", "Px_Tot_HT"). Si vous avez un doute, contactez le responsable du projet.

---

### 1.4 Etape 2 — Importer les DPGF des entreprises

Ici, vous ajoutez un a un les DPGF remis par chaque entreprise soumissionnaire.

1. Dans le champ **Nom de l'entreprise**, saisissez le nom (ex : `GECIP`, `CMR Batiment`).
2. Cliquez sur **Parcourir** et selectionnez le fichier DPGF de cette entreprise (`.xlsx`, `.xlsm`, `.xls`, `.xlsb`, ou `.pdf`).
3. Cliquez sur **Ajouter l'entreprise**.
4. L'application analyse le DPGF et l'integre dans le tableau comparatif.
5. Repetez l'operation pour chaque entreprise.

**Choisir le taux de TVA :**
Un menu deroulant permet de choisir le taux applicable (5,5 %, 10 % ou 20 %). Selectionnez-le avant d'exporter.

**Supprimer une entreprise :**
Chaque entreprise ajoutee apparait dans la liste avec un bouton **Supprimer**. Un message de confirmation vous sera demande avant la suppression definitive.

> **Nombre maximum d'entreprises :** L'application accepte jusqu'a 100 entreprises simultanement.

---

### 1.5 Comprendre les alertes

Apres l'import de chaque DPGF, l'application analyse automatiquement les donnees et signale les anomalies. Il existe trois niveaux d'alerte :

| Couleur | Niveau | Ce que ca signifie |
|---------|--------|---------------------|
| **Rouge** | Erreur | Probleme serieux : erreur de calcul, code DPGF non reconnu, poste manquant. A corriger imperativement. |
| **Jaune/Orange** | Avertissement | Anomalie probable : unite manquante, montant suspect, texte dans un champ numerique. A verifier. |
| **Bleu** | Information | Observation neutre : mot-cle detecte (SANS OBJET, COMPRIS), code supplementaire ajoute par l'entreprise. Aucune action requise. |

Ces alertes sont visibles :
- **Dans l'interface** (panneau d'alertes detaille apres chaque import)
- **Dans le fichier Excel exporte** (cases colorees + message explicatif dans la colonne "Commentaire")

---

### 1.6 Comprendre les montants en vert dans l'Excel

Dans le fichier Excel exporte, certains montants de Prix Unitaire HT (PU HT) apparaissent en **vert gras**. Cela signifie que cette entreprise propose le **prix le plus bas** sur ce poste precis, parmi toutes les entreprises importees.

**Exemple :**

| Poste | GECIP PU HT | CMR PU HT | Passerelles PU HT |
|-------|------------|-----------|------------------|
| 01.1.1 Terrassement | 12,00 € | **8,50 €** *(vert)* | 11,00 € |

Points importants :
- La comparaison se fait **poste par poste**, pas sur le total. Une entreprise peut etre la moins chere sur beaucoup de postes mais ne pas etre la moins disante globalement.
- Si deux entreprises ont exactement le meme prix, les deux s'affichent en vert.
- Le recapitulatif en bas du fichier Excel indique aussi l'entreprise globalement la moins disante avec la mention **"Meilleur Prix"**.

---

### 1.7 Etape 3 — Visualiser et exporter

1. Une fois tous les DPGF importes, cliquez sur la section **Etape 3**.
2. Le tableau comparatif complet s'affiche a l'ecran.
3. Verifiez les alertes et les montants.
4. Cliquez sur **Telecharger le TCO Excel** pour obtenir le fichier final.

Le fichier Excel genere contient :
- Le tableau de comparaison avec toutes les entreprises cote a cote
- Les PU HT en vert pour les meilleurs prix poste par poste
- Les totaux par section et le recapitulatif general HT et TTC
- Les cellules colorees en rouge ou jaune aux endroits problematiques
- Un message explicatif dans la colonne "Commentaire" de chaque anomalie

---

### 1.8 Gerer plusieurs lots

Si votre marche comporte plusieurs lots, chaque lot est traite comme un **projet independant**. Creez un projet par lot (ex : `Bergerac_LOT01`, `Bergerac_LOT02`) et exportez un fichier Excel par lot.

---

## Partie 2 — Gerer les problemes et les bugs

### 2.1 Problemes courants et solutions rapides

#### L'application ne demarre pas (rien ne s'ouvre)

- Attendez 15 a 30 secondes apres avoir double-clique sur l'exe — le demarrage peut etre lent.
- Verifiez que la fenetre noire (console) est bien ouverte. Si elle a disparu immediatement, il y a une erreur au demarrage (voir section 2.2).
- Essayez de relancer l'exe en **clic droit > "Executer en tant qu'administrateur"**.
- Verifiez qu'aucun antivirus ne bloque l'exe (ajoutez une exception si necessaire).

#### Le navigateur s'ouvre mais affiche une erreur blanche ou rouge

- Fermez l'application (fermez la fenetre noire), attendez 10 secondes, relancez.
- Si le probleme persiste, ouvrez un navigateur different (Chrome, Edge, Firefox) et tapez : `http://localhost:8501`

#### "Le fichier de projet n'existe plus"

- Les projets sont sauvegardes dans le dossier `projects/` a cote de l'exe. Verifiez que ce dossier existe et n'a pas ete deplace ou supprime.
- Si le projet a ete cree sur un autre poste, copiez le fichier `.tco.json.gz` dans le dossier `projects/` de votre machine.

#### Un DPGF est refuse au chargement ("Format non reconnu")

- Verifiez que le fichier n'est pas protege par un mot de passe (Excel demande un mot de passe a l'ouverture).
- Verifiez que le fichier fait moins de 20 Mo.
- Assurez-vous que le fichier est bien un vrai fichier Excel (extension `.xlsx`, `.xlsm`, etc.) et non un fichier renomme.
- Si c'est un fichier `.pdf`, verifiez qu'il s'agit bien d'un vrai PDF (pas d'une image scannee).

#### L'export Excel est vide ou incomplet

- Revenez a l'etape 2 et verifiez qu'au moins un DPGF est bien importe (la liste des entreprises doit etre non vide).
- Sauvegardez le projet, fermez et relancez l'application, rechargez le projet.

#### Le port 8501 est deja utilise (erreur "address already in use")

Cela arrive si une ancienne instance de l'application tourne encore en arriere-plan.

1. Appuyez sur `Ctrl + Alt + Suppr` > **Gestionnaire des taches**.
2. Dans l'onglet **Processus**, cherchez `TCO_Automator.exe` ou `python.exe`.
3. Faites un clic droit > **Fin de tache**.
4. Relancez l'application.

---

### 2.2 Recuperer les logs d'erreur

Les logs sont votre meilleur outil pour comprendre ce qui s'est passe. Ils enregistrent tout ce que l'application fait, y compris les erreurs detaillees.

**Ou les trouver :**

```
[dossier de l'exe]/
    logs/
        tco_automator.log     <- fichier principal
        tco_automator.log.1   <- archive precedente (rotation automatique)
```

**Comment les lire :**
Double-cliquez sur `tco_automator.log` (s'ouvre dans le Bloc-notes) ou faites clic droit > "Ouvrir avec" > "Bloc-notes".

Cherchez les lignes contenant les mots `ERROR` ou `Traceback` — ce sont elles qui decrivent le probleme.

**Exemple de ligne d'erreur :**
```
2025-03-06 14:32:11 ERROR core.merger — KeyError: 'Px_Tot_HT' (merger.py, line 847)
```

> **Astuce :** Dans le Bloc-notes, utilisez `Ctrl + F` et cherchez le mot `ERROR` pour aller directement aux erreurs.

---

### 2.3 Corriger un bug avec Antigravity IA

Antigravity est un agent de developpement base sur l'intelligence artificielle. Il peut lire le code source de l'application, comprendre l'erreur et appliquer une correction directement dans les fichiers, meme si vous n'avez aucune connaissance en programmation.

**Prerequis :** Antigravity doit etre ouvert dans le **dossier du projet** (le dossier qui contient `app.py`, `core/`, etc. — pas le dossier de l'exe).

#### Quand l'utiliser ?

- Un message d'erreur rouge apparait dans l'application et ne disparait pas apres relancement.
- L'export Excel plante ou produit un fichier corrompu ou incomplet.
- Un DPGF que vous savez correct est systematiquement rejete.
- Une anomalie apparait apres remplacement du fichier TCO modele.

#### Procedure pas a pas

**1. Copiez le message d'erreur**

Dans l'interface, notez exactement le message d'erreur affiche. Depuis les logs (`logs/tco_automator.log`), copiez les 20 a 30 lignes autour du mot `ERROR` ou `Traceback`.

**2. Ouvrez Antigravity**

Lancez Antigravity en vous assurant qu'il est ouvert sur le dossier source du projet (celui qui contient `app.py`).

**3. Formulez votre demande**

Copiez-collez ce modele en remplissant les parties entre crochets :

---

> Bonjour, le logiciel TCO Automator a rencontre une erreur.
>
> Voici le message d'erreur copie depuis les logs :
> ```
> [COLLEZ ICI LE CONTENU DU LOG]
> ```
>
> Contexte : l'erreur se produit quand [ex : "j'importais le DPGF de l'entreprise GECIP" / "je cliquais sur Telecharger le TCO Excel"].
>
> Peux-tu :
> 1. Analyser la cause de cette erreur
> 2. Corriger le code dans les fichiers concernes
> 3. M'expliquer en termes simples ce qui s'est passe
> 4. Ecrire un test automatique pour eviter que ca recommence

---

**4. Laissez Antigravity travailler**

L'IA va :
- Lire les fichiers du projet concernes (generalement dans `core/`, `app/` ou `services/`)
- Identifier la ligne de code responsable de l'erreur
- Appliquer une correction directement dans les fichiers
- Lancer les tests automatiques (`pytest`) pour verifier que la correction ne casse rien d'autre

**5. Regenerez l'exe apres correction**

Une fois la correction appliquee par Antigravity dans les fichiers sources, vous devez regenerer l'exe pour que vos utilisateurs beneficient du correctif. Dans le terminal Antigravity (ou un terminal Windows ouvert dans le dossier du projet) :

```
pyinstaller TCO_Automator.spec --clean
```

Le nouvel exe se trouve dans `dist/TCO_Automator/`. Remplacez l'ancien exe par le nouveau.

---

### 2.4 Informations pour Antigravity : architecture du projet

Si l'IA vous demande des informations sur le projet, voici un resume a lui fournir :

```
Projet : TCO Automator v2.2 — version locale .exe (PyInstaller onedir)
Technologie : Python 3.11, Streamlit, openpyxl, pandas, rapidfuzz
Point d'entree exe : run_app.py (lance stcli.main() en mode frozen)
Structure :
  - app.py              : interface principale Streamlit (3 etapes)
  - core/
      parser_tco.py     : lecture du TCO modele
      parser_dpgf.py    : lecture des DPGF entreprises (Excel)
      parser_dpgf_pdf.py: lecture des DPGF au format PDF
      merger.py         : fusion TCO + DPGF, generation des alertes
      exporter.py       : production du fichier Excel final avec colorisation
  - services/
      persistence.py    : sauvegarde/chargement des projets (.tco.json.gz)
      file_validator.py : validation des fichiers uploades (magic bytes)
  - config.py           : constantes metier (TVA, tolerances, chemins)
  - logger.py           : configuration des journaux rotatifs
  - run_app.py          : point d'entree PyInstaller
  - TCO_Automator.spec  : fichier de build PyInstaller
Tests : dossier tests/ (pytest)
Donnees utilisateur stockees a cote de l'exe : projects/, logs/, uploads/
```

---

### 2.5 Maintenance reguliere

| Tache | Frequence | Comment faire |
|-------|-----------|---------------|
| Sauvegarder les projets | Apres chaque session | Bouton "Sauvegarder" dans la barre laterale |
| Archiver les projets sur le reseau | Apres chaque marche | Copier le dossier `projects/` sur votre serveur ou GED |
| Archiver les TCO Excel exportes | Apres chaque marche | Copier les fichiers `.xlsx` dans votre GED ou reseau |
| Verifier les logs | En cas de comportement anormal | Ouvrir `logs/tco_automator.log` |
| Mettre a jour l'application | Quand une nouvelle version est disponible | Remplacer le dossier `dist/TCO_Automator/` par la nouvelle version |

> **Sauvegarde importante :** Le dossier `projects/` contient tout votre travail. Copiez-le regulierement sur un disque reseau ou un stockage partage.

---

### 2.6 Contacts et escalade

Si Antigravity ne parvient pas a resoudre le probleme, ou si l'erreur touche une partie critique (perte de donnees, corruption de projets), transmettez au developpeur en charge :

1. Le fichier `logs/tco_automator.log` complet
2. Une description precise de ce que vous faisiez au moment de l'erreur
3. Le fichier DPGF ou TCO qui pose probleme (si applicable et non confidentiel)
4. La version de l'application (visible en bas de la barre laterale de l'interface)

---

*TCO Automator v2.2 — Odetec — Version locale .exe*
