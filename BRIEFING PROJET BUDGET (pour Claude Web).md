# Contexte projet — Budget SNRT 2026

Je travaille sur le suivi budgétaire 2026 de la SNRT. Voici le contexte à connaître avant de m'aider sur les fichiers Excel de ce projet.

## Fichiers principaux

- **`PPM 2026 AVEC ESTIMATION.xlsx`** — Programme Prévisionnel des Marchés. Feuille "PPM", en-têtes ligne 9, données lignes 10 à 169 (160 marchés). Colonnes : A=DIRECTION, B=OBJET, C=NATURE DE LA PRESTATION, D=MODE DE PASSATION, E=PERIODE PREVUE POUR LE LANCEMENT (T1..T4), F=LIEU D'EXECUTION, G=AOO RESERVES AUX PME, H=ESTIMATION, **I=NATURE BUDGET** (Investissement/Fonctionnement/Exploitation/Honoraires — ajoutée en croisant avec `PPA2026Global...xlsx` sur l'Objet).
- **`Numéros décision 2026.xlsx`** — feuille "AO 2026", en-têtes ligne 3, données à partir de ligne 4. Colonnes : N° PROJET, Date de Réception, Direction, Objet, **intitulé PPM** (colonne de lien manuel vers l'objet du PPM — parfois vide, parfois plusieurs intitulés séparés par un retour à la ligne), Estimation, N° décision, Date Statut, STATUT (Lancé / Annulé / Complément MO / Traitement DA / 8 jours...).
- **`ETAT D'ENGAGEMENTS ....xlsx`** (le plus récent, nom variable) — feuille "ENGAGEMENTS 2026". Colonnes clés : B=DIRECTION, C=REF, D=BCI (en-tête mal nommé, = TYPE D'ACHAT : CONTRAT/AOO-M/AOO-CT/BCL/BCI), F=OBJET, H=STATUT SUPPORT D'ENGAGEMENT, N=MONTANT PREENGAGE INITIAL, AD=TOTAL MARCHE TTC/TOTAL ENGAGE TTC, AE=STATUT D'ENGAGEMENT.
- **`BC 2026 AU 30-06-2026.xlsx`** — feuille "BC APPROUVE", en-têtes ligne 6. Colonnes : BC, N° JDE, FOURNISSEURS, NATURE DE LA DEPENSE, OBJET, MONTANT HT/TTC, DIRECTION, BUDGET, APPROBATION, APPROUVE (date).
- **`COMPARATIF PPM 2026 vs NUMEROS DECISION 2026.xlsx`** — le tableau de bord principal généré (voir plus bas), reconstruit à chaque mise à jour du PPM ou du fichier Numéros décision.
- **`PPM 2026 - Projets non recus (relance directions).xlsx`** — livrable simplifié destiné au directeur, listant les projets PPM sans décision.

## Règle métier — montant à utiliser selon STATUT D'ENGAGEMENT (fichier État d'engagements)

- **ENGAGE**, **EN COURS D'ENGAGEMENT**, **NON ENGAGE** → colonne AD (Total Marché TTC)
- **EN COURS DE PREENGAGEMENT** → colonne N (Montant Préengagé Initial), inconditionnellement
- **PREENGAGE** → si STATUT SUPPORT D'ENGAGEMENT (col H) = `ATTRIBUE` → AD, sinon → N

## Méthode de croisement PPM ↔ Numéros décision

1. Lien principal = la colonne **"intitulé PPM"** du fichier décision (texte copié à la main, à comparer normalisé — accents/casse/espaces) avec l'Objet du PPM.
2. Une décision peut contenir plusieurs lots dans son Objet (`Lot n°1: ... Lot n°2: ...`) — éclater et comparer chaque lot séparément à l'estimation (qui a aussi parfois un format `Lot 1: X Lot 2: Y`, avec des numéros parfois non alignés type "01" vs "1" → normaliser).
3. **Détection "infructueux"** : quand un même texte (segment ou objet entier) apparaît plusieurs fois pour un même intitulé PPM, la/les occurrence(s) la/les plus ancienne(s) (par date de réception) sont l'appel d'offres infructueux relancé ensuite — les marquer INFRUCTUEUX et exclure leur montant de la somme. Parfois le texte diffère légèrement (abréviation "SNRT" vs nom complet, faute de frappe) et il faut valider avec moi au cas par cas plutôt que deviner.
4. Statut global d'un projet PPM = "Lancé" si au moins une décision active a le statut Lancé, sinon "En cours" si des décisions existent mais aucune Lancée, sinon "Projet non reçu" si aucune décision liée.

## Structure du fichier COMPARATIF PPM 2026 vs NUMEROS DECISION 2026.xlsx

- **Onglet "Accès rapide"** (premier onglet) : menu de boutons (cellules avec hyperliens internes) par Direction / Période / Statut, qui sautent vers l'onglet "Vues filtrées".
- **Onglet "PPM vs Décisions"** : tableau principal, converti en vrai Tableau Excel (filtres sur les en-têtes). Une ligne par projet PPM (en gras) suivie de sous-lignes indentées par décision/lot liée (groupées avec l'outline Excel, +/- cliquable). Colonnes : N°, Direction, Objet, Nature, Nature Budget, Mode de passation, Période prévue, Estimation PPM, Estimation Décision, Écart, Écart %, Statut global, Statut (catégorie), N° décision, Date statut, STATUT décision, Lot, N° Projet (décision).
- **Onglet "Vues filtrées"** : mêmes données, sectionnées par Direction puis Période puis Statut, avec bouton "Retour au menu".
- **Onglet "Synthèse"**, **"PPM - Projet non reçu"**, **"Anomalies fichier décision"** (intitulés PPM vides ou introuvables à corriger).

## Conventions visuelles

- Bleu marine `#1F4E78` pour les bandeaux/en-têtes, blanc pour le texte.
- Vert clair = Lancé/OK, orange = En cours, rouge clair = Non reçu/manquant, gris = Infructueux (texte barré).
- Toujours proposer un fichier de test séparé avant d'écraser un fichier de travail si la manipulation est risquée (ex. édition XML bas niveau).
- Toujours vérifier qu'un fichier n'est pas ouvert dans Excel (fichier `~$...`) avant d'essayer de l'enregistrer.

## Outillage

Sur ma machine personnelle, Python réel est ici : `C:\Users\HP\AppData\Local\Programs\Python\Python312\python.exe` avec openpyxl 3.1.5 (le `python`/`py` du PATH ne marche pas, c'est le stub Microsoft Store). Sur le PC professionnel, l'environnement sera sans doute différent — à vérifier.
