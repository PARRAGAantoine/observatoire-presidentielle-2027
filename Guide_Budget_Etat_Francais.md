# Guide Explicatif : Le Budget Complet et Ultra-Détaillé de l'État Français

Ce document accompagne le fichier Excel **[Budget_Detaille_Etat_Francais.xlsx](file:///c:/Users/user/Desktop/projets_IA/projets/presidentielle/Budget_Detaille_Etat_Francais.xlsx)** créé dans le répertoire `C:\Users\user\Desktop\projets_IA\projets\presidentielle`.

Il présente la nomenclature officielle et le chiffrement exhaustif des **recettes fiscales et non fiscales**, des **dépenses ministérielles par mission et programme (LOLF)**, des **comptes spéciaux** et des **finances publiques consolidées (APU)** de la République française.

---

## 🏛️ 1. Architecture du Fichier Excel (`Budget_Detaille_Etat_Francais.xlsx`)

Le classeur est structuré en **4 onglets de référence** entièrement dynamiques (formules de calculs automatiques `=SUM(...)`, ratios de pourcentages et liaisons inter-onglets) :

1. **`00_Tableau_Bord_Synthese` :**
   * **Équilibre Général du Budget de l'État :** Recettes fiscales brutes, dégrèvements, recettes fiscales nettes, recettes non fiscales, prélèvements sur recettes (UE et Collectivités), total des dépenses du budget général et solde budgétaire net (déficit de l'État).
   * **Périmètre Consolidé des Administrations Publiques (APU selon l'INSEE et Maastricht) :** État et ODAC (APUC), Sécurité Sociale (ASSO), Collectivités locales (APUL) avec dépenses, recettes et solde de chaque sous-secteur.
   * **Indicateurs Clés Macroéconomiques :** PIB nominal (2 875 Mds €), Dette publique (3 228,4 Mds € soit 112,3 % du PIB), Dépenses publiques totales (52,7 % du PIB hors transferts), Taux de prélèvements obligatoires (44,7 % du PIB), Charge annuelle de la dette (54,8 Mds €).

2. **`01_Recettes_Detaillees` :**
   * **Recettes Fiscales Ligne par Ligne :**
     * **Impôt sur le Revenu (IR) :** Prélèvement à la source (78,5 Mds €), acomptes contemporains des indépendants (16,2 Mds €), PFU part IR (8,4 Mds €), soldes liquidés et rôles (7,5 Mds €), dégrèvements et crédits d'impôt (-14,2 Mds €).
     * **Impôt sur les Sociétés (IS) :** Acomptes et solde au taux normal 25 % (82,4 Mds €), taux réduit PME à 15 % (4,8 Mds €), remboursements de créances de Crédit Impôt Recherche - CIR et CICE (-18,6 Mds €).
     * **Taxe sur la Valeur Ajoutée (TVA) :** TVA brute totale collectée (214,5 Mds €), remboursements de crédits de TVA (-68,4 Mds €), part de TVA transférée à la Sécurité Sociale (-62,8 Mds €), part de TVA transférée aux Collectivités territoriales (-42,5 Mds €) et TVA nette conservée par le Budget Général de l'État (40,8 Mds €).
     * **Accises sur l'Énergie (ex-TICPE) :** Carburants (26,2 Mds €), électricité et gaz (7,4 Mds €), part transférée aux collectivités et à l'AFITF (-16,8 Mds €), accises nettes État (16,8 Mds €).
     * **Enregistrement, Mutations et Successions :** Droits de succession / héritage (17,6 Mds €), donations entre vifs (3,4 Mds €), droits de mutation immobilière DMTO part État (3,8 Mds €), Impôt sur la Fortune Immobilière / IFI (2,4 Mds €), timbres fiscaux et droits d'actes (2,8 Mds €).
     * **Autres Impôts et Taxes Intérieures :** Accises tabacs (3,6 Mds €), alcools (1,2 Mds €), Taxe sur les Transactions Financières - TTF (2,1 Mds €), Taxe sur les services numériques GAFAM (0,8 Mds €), malus automobile CO2 et au poids (2,2 Mds €), jeux et paris FDJ/PMU (1,1 Mds €), logements vacants (0,9 Mds €), frais d'assiette et recouvrement perçus pour compte de tiers (5,2 Mds €).
   * **Recettes Non Fiscales Ligne par Ligne :**
     * Dividendes des participations publiques (Bpifrance, ADP, Safran, Renault, Thales, EDF : 4,5 Mds €), reversements et dividendes de la Banque de France (1,8 Mds €), redevances hertziennes 4G/5G (2,6 Mds €), redevances concessions autoroutes et barrages (3,1 Mds €), cessions immobilières (0,9 Mds €), ventes de biens et services (3,2 Mds €), amendes radars et circulation (1,7 Mds €), sanctions autorités de régulation AMF/Concurrence (1,4 Mds €), intérêts des prêts à des États étrangers (1,5 Mds €), rémunération des garanties de l'État (2,2 Mds €).
   * **Prélèvements sur Recettes (Déductions) :**
     * Prélèvement au profit de l'Union Européenne (PSR-UE : -25,6 Mds €).
     * Dotation Globale de Fonctionnement aux collectivités (DGF : -27,4 Mds €).
     * Dotations de compensation fiscale locale (-14,2 Mds €) et fonds de péréquation (-4,6 Mds €).

3. **`02_Depenses_Missions` :**
   * **Les 32 Missions Ministérielles Officielles du Budget Général (Crédits de Paiement & Autorisations d'Engagement) :**
     * *Enseignement scolaire* (64,2 Mds € - 1er poste de dépense ministériel)
     * *Engagements financiers de l'État / Intérêts de la dette* (54,8 Mds € - 2ème poste)
     * *Défense* (49,7 Mds € en CP / 56,4 Mds € en AE)
     * *Recherche et enseignement supérieur* (32,8 Mds €)
     * *Solidarité, insertion et égalité des chances / AAH & Prime d'activité* (31,4 Mds €)
     * *Sécurités / Police et Gendarmerie* (25,6 Mds €)
     * *Travail et emploi / Apprentissage & France Travail* (22,8 Mds €)
     * *Écologie, développement et mobilité durables* (19,4 Mds €)
     * *Cohésion des territoires / APL part État & hébergement d'urgence* (18,6 Mds €)
     * *Gestion des finances publiques / DGFiP & Douanes* (11,2 Mds €)
     * *Justice / Tribunaux et Prisons* (10,8 Mds €)
     * *Régimes sociaux et de retraite / Subventions équilibre SNCF/RATP* (6,8 Mds €)
     * *France 2030* (5,8 Mds €)
     * *Aide publique au développement* (5,2 Mds €)
     * *Administration générale et territoriale de l'État / Préfectures* (4,8 Mds €)
     * *Culture* (4,4 Mds €), *Agriculture et forêt* (4,2 Mds €), *Action extérieure / Diplomatie* (3,6 Mds €), *Économie* (3,4 Mds €), *Outre-mer* (3,1 Mds €), *Immigration et asile* (2,3 Mds €), *Santé / AME* (1,9 Mds €), *Direction action du Gouvernement* (1,8 Mds €), *Sports* (1,7 Mds €), *Anciens combattants* (1,6 Mds €), *Fonction publique* (1,2 Mds €), *Pouvoirs publics / Élysée, Assemblée, Sénat* (1,1 Mds €), *Médias* (0,9 Mds €), *Conseil et contrôle* (0,8 Mds €), *Relance résiduelle* (0,6 Mds €), *Crédits non répartis* (0,9 Mds €), *Solidarités Outre-mer* (0,5 Mds €).

4. **`03_Comptes_Speciaux` :**
   * **Comptes d'Affectation Spéciale (CAS) :** CAS Pensions (retraites fonctionnaires civils et militaires : 64,8 Mds €), CAS Contrôle de la circulation routière / Radars (1,8 Mds €), CAS Participations financières de l'État (4,2 Mds €), CAS Transition énergétique (7,8 Mds €).
   * **Budgets Annexes :** Contrôle et exploitation aériens (BACEA / DGAC : 2,3 Mds €), Publications officielles et information administrative (DILA : 0,2 Mds €).
   * **Comptes de Concours Financiers :** Avances aux collectivités territoriales (128,4 Mds € - mécanisme de trésorerie d'avance des impôts locaux par l'État), Prêts à des États étrangers (1,9 Mds €).

---

## 📊 2. Synthèse Chiffrée Consolidée

| Grand Agrégat du Budget de l'État | Montant Annuel (Mds €) | Règle / Mécanisme Associé |
|:---|---:|:---|
| **Recettes Fiscales Brutes** | **497,2** | Totalité des impôts recouvrés sur le territoire |
| Remboursements et Dégrèvements d'Impôts | **-155,4** | Restitutions de TVA, CIR, dégrèvements divers |
| **Recettes Fiscales Nettes** | **341,8** | Ressources fiscales directes nettes |
| **Recettes Non Fiscales** | **32,1** | Dividendes, domaine hertzien/autoroutier, amendes |
| Fonds de Concours et Divers | **+6,2** | Participations de tiers affectées |
| **Prélèvements sur Recettes (Déductions)** | **-71,8** | Dont -46,2 Mds € Collectivités et -25,6 Mds € Union Européenne |
| **TOTAL DES RECETTES NETTES DU BUDGET GÉNÉRAL** | **308,3** | Ressources réelles de fonctionnement de l'État |
| **TOTAL DES DÉPENSES DU BUDGET GÉNÉRAL (CP)** | **503,6** | Somme des 32 missions ministérielles (hors CAS) |
| **SOLDE DU BUDGET GÉNÉRAL (DÉFICIT BUDGÉTAIRE ÉTAT)** | **-195,3** | Besoin de financement couvert par émission de dette (OAT) |
| Solde net des Comptes Spéciaux | **+1,8** | Excédent net technique des CAS |
| **BESOIN DE FINANCEMENT TOTAL DE L'ÉTAT** | **-193,5** | Déficit net annuel de l'État |

---

## 📁 3. Emplacement des Fichiers

* Fichier Excel : **`C:\Users\user\Desktop\projets_IA\projets\presidentielle\Budget_Detaille_Etat_Francais.xlsx`**
* Guide documentaire : **`C:\Users\user\Desktop\projets_IA\projets\presidentielle\Guide_Budget_Etat_Francais.md`**
* Dossier des candidats présidentiels : **`C:\Users\user\Desktop\projets_IA\projets\presidentielle\candidats\David Lisnard\`**
