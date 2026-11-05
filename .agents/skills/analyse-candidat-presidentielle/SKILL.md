---
name: analyse-candidat-presidentielle
description: >-
  Réalise une recherche encyclopédique, un classement thématique en 11 dossiers
  (28 fiches Markdown), un chiffrage budgétaire ligne à ligne sous Excel (5 onglets
  avec formules et 5 graphiques natifs) et un tableau de bord HTML interactif
  comparant le programme d'un candidat à l'élection présidentielle 2027 avec le
  Budget actuel de l'État français (modèle complet David Lisnard). Utiliser dès
  que l'utilisateur demande d'analyser, approfondir ou chiffrer un candidat.
---

# Skill : Analyse & Chiffrage Complet d'un Candidat à la Présidentielle 2027

Ce skill standardise la méthode d'investigation, de structuration documentaire, de chiffrage budgétaire et de comparaison graphique avec les comptes de l'État (au niveau d'exigence réalisé pour **David Lisnard**).

## Emplacement des Dossiers et Référentiels

- **Racine du projet Présidentielle :** `C:\Users\user\Desktop\projets_IA\projets\presidentielle`
- **Dossier du candidat :** `C:\Users\user\Desktop\projets_IA\projets\presidentielle\candidats\<Prénom Nom>\`
- **Budget de référence de l'État français :** `C:\Users\user\Desktop\projets_IA\projets\presidentielle\Budget_Detaille_Etat_Francais.xlsx`
- **Chiffres de référence (Baseline État & APU) :** Voir [baseline_budget_etat.md](./references/baseline_budget_etat.md)
- **Modèle de référence complet (Gold Standard) :** `C:\Users\user\Desktop\projets_IA\projets\presidentielle\candidats\David Lisnard\`

---

## Procédure Complète en 5 Étapes Obligatoires

Lorsqu'on te demande d'analyser un candidat (ou de « faire comme pour David Lisnard »), exécute **intégralement** les 5 étapes ci-dessous sans en omettre aucune :

### Étape 1 : Recherche Approfondie Multi-Sources (Web, Réseaux, Vidéos, Chiffrages)

Effectue entre **5 et 10 recherches web ciblées** (`search_web`) couvrant :
1. **Programme officiel, livre-programme et parti** : propositions économiques, sociales, régaliennes, institutionnelles, écologiques, européennes.
2. **Interventions médias, YouTube et Réseaux Sociaux (X/Twitter, Instagram, TikTok, Facebook)** : discours fondateurs, meetings, grandes interviews (TV, radio, presse écrite), punchlines et angles d'attaque.
3. **Évaluations financières et chiffrages** : chiffrage officiel du candidat, évaluations des think tanks (Institut Montaigne, IFRAP, OFCE, Terra Nova, Fondapol, Institut Rousseau, Cour des comptes) pour documenter le coût ou l'économie de chaque mesure en **Année 1** et en **Année 5 (plein régime)**.

> 💡 **Astuce parallélisation :** Si l'utilisateur demande d'analyser plusieurs candidats ou d'aller très vite, délègue la rédaction des dossiers thématiques ou la collecte web à des sous-agents (`invoke_subagent`).

---

### Étape 2 : Création des 11 Sous-Dossiers Thématiques (28 Fiches Détaillées)

Dans `C:\Users\user\Desktop\projets_IA\projets\presidentielle\candidats\<Prénom Nom>\`, crée systématiquement l'arborescence suivante avec des fiches Markdown **exhaustives, sourcées et chiffrées** (minimum 80 à 150 lignes par fiche, pas de résumés superficiels) :

```text
<Prénom Nom>/
├── 00_Introduction_et_Candidature/
│   ├── 01_Biographie_Parcours_et_Positionnement.md
│   ├── 02_Parti_Mouvement_Equipe_et_Reseaux.md
│   └── 03_Doctrine_Philosophie_et_Vision_2027.md
├── 01_Economie_et_Finances_Publiques/
│   ├── 01_Fiscalite_Impots_Menages_et_Patrimoine.md
│   ├── 02_Entreprises_Production_et_Competitivite.md
│   ├── 03_Travail_Salaires_et_Pouvoir_d_Achat.md
│   └── 04_Dette_Depense_Publique_et_Trajectoire.md
├── 02_Securite_Justice_et_Immigration/
│   ├── 01_Securite_Police_Gendarmerie_et_Ordre_Public.md
│   ├── 02_Justice_Prisons_et_Politique_Penale.md
│   └── 03_Immigration_Frontieres_Assimilation_et_Laicite.md
├── 03_Sante_et_Affaires_Sociales/
│   ├── 01_Hopital_Medecine_de_Ville_et_Deserts_Medicaux.md
│   ├── 02_Retraites_et_Financement_Protection_Sociale.md
│   └── 03_Minima_Sociaux_RSA_Chomage_et_Lutte_Anti_Fraude.md
├── 04_Education_Culture_et_Jeunesse/
│   ├── 01_Ecole_Savoirs_Fondamentaux_et_Enseignants.md
│   ├── 02_Universites_Recherche_et_Innovation.md
│   └── 03_Culture_Patrimoine_Audiovisuel_et_Jeunesse.md
├── 05_Ecologie_Energie_et_Agriculture/
│   ├── 01_Politique_Energetique_Nucleaire_et_Renouvelables.md
│   ├── 02_Ecologie_Transports_et_Industrie_Verte.md
│   └── 03_Agriculture_Peche_et_Souverainete_Alimentaire.md
├── 06_Institutions_et_Decentralisation/
│   ├── 01_Institutions_Constitution_Referendum_et_Democratie.md
│   └── 02_Reforme_de_l_Etat_Fonction_Publique_et_Collectivites.md
├── 07_Europe_et_Geopolitique/
│   ├── 01_Union_Europeenne_Traites_et_Commerce.md
│   └── 02_Defense_Armees_Diplomatie_et_Francophonie.md
├── 08_Veille_Numerique_et_Sources/
│   ├── 01_Presence_Numerique_YouTube_Reseaux_Sociaux.md
│   └── 02_Bibliographie_Discours_Clés_et_Sources_Verifiees.md
├── 09_Logement_Famille_et_Demographie/
│   ├── 01_Logement_Construction_et_Fiscalite_Immobiliere.md
│   └── 02_Politique_Familiale_Natalite_et_Transmission.md
└── 10_Chiffrage_et_Simulateurs_Budgetaires/
    ├── 01_Chiffrage_Ligne_a_Ligne_par_Famille.md
    └── 02_Impact_Macroeconomique_et_Cas_Types_Menages.md
```

Chaque fiche thématique doit détailler pour chaque proposition :
- **Le constat du candidat** (Pourquoi cette mesure ?)
- **Le mécanisme précis** (Comment ça marche, quel texte de loi / réforme constitutionnelle / décret ?)
- **Qui est concerné** (Bénéficiaires et perdants / payeurs)
- **L'impact budgétaire chiffré** (en Mds €/an en An 1 et An 5)
- **Les citations et sources exactes** (discours, livre, interview, rapport).

---

### Étape 3 : Création du Fichier Maître `INDEX.md`

À la racine de `C:\Users\user\Desktop\projets_IA\projets\presidentielle\candidats\<Prénom Nom>\INDEX.md`, crée un index général contenant :
1. **Carte d'identité politique et stratégique** du candidat.
2. **Tableau récapitulatif de l'équilibre budgétaire** de son programme (Économies, Dépenses nouvelles, Hausses d'impôts, Baisses d'impôts, Impact sur le déficit et la dette).
3. **Sommaire cliquable** vers les 11 sous-dossiers, les 28 fiches Markdown, le fichier Excel et le tableau de bord HTML.
4. **Synthèse des 15 mesures phares** du candidat.

---

### Étape 4 : Génération du Classeur Excel à 5 Onglets (`Chiffrage_Programme_<Prenom_Nom>_2027.xlsx`)

Utilise et adapte le script modèle [template_candidate_generator.py](./scripts/template_candidate_generator.py) pour générer le classeur Excel dans le dossier du candidat avec **exactement ces 5 onglets** (noms <= 31 caractères) :

1. **`Chiffrage_Detaille_Familles`** :
   - Entre **25 et 40 lignes de mesures détaillées** regroupées en **8 Grandes Familles** (ex. Économies État/Agences, Sphère Sociale/Retraites, Immigration/Fraude, Collectivités/Normes, Cotisations/Travail, Fiscalité Entreprises, Fiscalité Ménages/Patrimoine, Dépenses Nouvelles/Réinvestissements).
   - Colonnes obligatoires : `Code`, `Famille`, `Ministère / Sphère`, `Intitulé Détaillé de la Mesure`, `Nature (Économie / Dépense / Baisse d'impôt / Hausse d'impôt)`, `Impact An 1 (Mds €)`, `Impact An 5 Plein Régime (Mds €/an)`, `Cumul 5 Ans (Mds €)`, `QUI EST CONCERNÉ ?`, `COMMENT & QUAND ? (Mécanisme)`, `POURQUOI ? (Objectif)`, `Source / Référence`.
   - **Sous-totaux dynamiques (`=SUM(...)`)** à la fin de chaque famille + **Récapitulatif général en bas de tableau** (Total Économies, Total Dépenses nouvelles, Total Baisses d'impôts, Total Hausses d'impôts, Solde primaire statique, Effet retour macroéconomique, Solde net sur le déficit public).
2. **`Synthese_Familles_Equilibre`** :
   - Tableau de synthèse lié par formules au 1er onglet (`=Chiffrage_Detaille_Familles!F12`, etc.) + Analyse de la doctrine budgétaire propre au candidat.
3. **`Simulateur_Cas_Types_Micro`** :
   - Impact concret en **€/an** et **€/mois** sur **8 profils types** (Salarié SMIC, Salarié médian 2 000 € net, Couple classe moyenne avec 2 enfants, Cadre supérieur / Chef d'entreprise, Retraité pension 1 500 € vs 2 800 €, Propriétaire / Héritier, PME industrielle 50 salariés, Agent public / Enseignant / Soignant).
4. **`04_Comparatif_Etat_vs_Candidat`** :
   - Comparaison ligne à ligne avec les chiffres officiels de [baseline_budget_etat.md](./references/baseline_budget_etat.md) :
     - **Section 1 :** Tableau de bord macroéconomique (Dépenses APU `1 514,0 Mds €`, Recettes APU `1 368,0 Mds €`, Déficit Maastricht `-146,0 Mds €`, Budget Général État `503,6 Mds €`, Recettes Nettes État `308,3 Mds €` vs Budget Candidat An 1 et An 5).
     - **Section 2 :** Les **10 Grands Pôles de Dépenses Publiques** (`Retraites 382 Mds`, `Santé 254 Mds`, `Solidarité/Famille/APL/RSA 212 Mds`, `Collectivités 156 Mds`, `Éducation/Recherche 115 Mds`, `Régalien Défense/Police/Justice 107 Mds`, `Économie/Écologie/Agences 105 Mds`, `Chômage 68 Mds`, `Charge Dette 58 Mds`, `Admin État & UE 57 Mds` = `1 514,0 Mds €`) comparés aux montants prévus par le candidat en An 1 et An 5.
     - **Section 3 :** Les **6 Grandes Familles de Recettes Publiques** (`Cotisations & CSG 678 Mds`, `TVA & Consommation 319 Mds`, `IS & Impôts de production 137 Mds`, `IR & Épargne 95 Mds`, `Patrimoine & Successions 92 Mds`, `Recettes non fiscales + Effet retour 47 Mds` = `1 368,0 Mds €`) comparées aux recettes prévues par le candidat.
5. **`05_Graphiques_Comparatifs`** :
   - 5 tableaux de données sources (colonnes A à D) + **5 graphiques Excel natifs (`openpyxl.chart`)** intégrés en colonnes F à P :
     1. `BarChart` : Impact du programme par famille (An 1 vs An 5).
     2. `DoughnutChart` ou `PieChart` : Répartition des mesures du programme.
     3. `BarChart` groupé : Budget Actuel vs Budget Candidat An 1 vs Budget Candidat An 5 (Dépenses, Recettes, Déficit, État seul).
     4. `BarChart` horizontal : Où va l'argent — Comparaison des 10 pôles de dépenses (Actuel vs Candidat An 5).
     5. `BarChart` groupé : D'où vient l'argent — Comparaison des 6 familles de recettes (Actuel vs Candidat An 5).

---

### Étape 5 : Génération du Tableau de Bord Interactif HTML (`Visualisation_Programme_et_Comparatif_Budget_<Prenom_Nom>.html`)

Crée dans le dossier du candidat un fichier HTML autonome utilisant **Chart.js** (thème sombre exécutif identique à celui de David Lisnard) comprenant :
- **5 cartes KPI** en haut de page résumant les grands agrégats du programme du candidat.
- **6 graphiques interactifs Chart.js** :
  1. Le Programme du candidat par Famille (An 1 vs An 5).
  2. La Répartition / Clé budgétaire du candidat (Doughnut chart).
  3. Comparatif Macroéconomique Global (Budget Actuel vs Candidat An 1 vs Candidat An 5).
  4. Trajectoire Quinquennale sur 5 ans (Courbes Dépenses vs Recettes vs Déficit de 2026 à An 5).
  5. Où va l'argent : Comparaison horizontale des 10 Grands Pôles de Dépenses (Budget Actuel `1 514 Mds €` vs Budget Candidat An 5).
  6. D'où vient l'argent : Comparaison des 6 Grandes Familles de Recettes (Budget Actuel `1 368 Mds €` vs Budget Candidat An 5).
- **Le grand tableau récapitulatif HTML** ligne à ligne en bas de page.

---

## Vérification Finale Avant Réponse à l'Utilisateur

1. Vérifie par un script Python que le fichier `.xlsx` s'ouvre sans erreur, que les 5 onglets existent et que **les sommes des 10 pôles de dépenses et des 6 familles de recettes correspondent exactement aux totaux macroéconomiques** (aucun écart d'arrondi).
2. Mets à jour `C:\Users\user\Desktop\projets_IA\projets\presidentielle\candidats\INDEX_CANDIDATS.md` pour indiquer que le dossier du candidat est désormais **Complet (Programme + Excel + Graphiques)**.
3. Dans ta réponse à l'utilisateur, affiche directement :
   - Les liens cliquables `file:///` vers `INDEX.md`, le fichier `.xlsx` et le fichier `.html`.
   - Les diagrammes visuels ASCII/Unicode et les tableaux comparatifs (Dépenses vs Recettes actuelles vs Budget du candidat).
