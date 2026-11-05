# -*- coding: utf-8 -*-
"""
Script Générateur Standardisé pour le Skill `analyse-candidat-presidentielle`
Génère simultanément pour un candidat donné :
1. Le fichier Excel à 5 onglets (`Chiffrage_Programme_<Prenom_Nom>_2027.xlsx`) incluant 5 graphiques Excel natifs.
2. Le tableau de bord interactif HTML (`Visualisation_Programme_et_Comparatif_Budget_<Prenom_Nom>.html`) incluant 6 graphiques Chart.js.

Utilisation :
Importer la fonction `build_candidate_deliverables(config)` depuis un script dans `scratch/` ou exécuter avec un fichier JSON.
"""
import os
import json
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, DoughnutChart, Reference

def build_candidate_deliverables(cfg):
    """
    cfg dictionary structure:
    - candidate_name: str (ex: "Édouard Philippe")
    - party_name: str (ex: "Horizons")
    - tagline: str (ex: "Ordre dans les comptes, ordre dans la rue...")
    - candidate_dir: str (path to folder)
    - families_measures: list of dicts:
        [
          {
            "family_id": "FAMILLE 1",
            "family_title": "FAMILLE 1 : ...",
            "short_name": "1. État & Agences",
            "category": "Économie de Dépense Publique", # or "Baisse de Prélèvements" / "Hausse de Recettes" / "Dépense Nouvelle / Réinvestissement"
            "measures": [
              {
                "code": "ECO-01",
                "sphere": "Fonction Publique",
                "title": "...",
                "type_label": "Économie structurelle",
                "an1": 5.0,
                "an5": 15.0,
                "cum5": 55.0,
                "qui": "...",
                "comment": "...",
                "pourquoi": "...",
                "source": "..."
              }
            ]
          }
        ]
    - macro_return_an1: float (Mds €)
    - macro_return_an5: float (Mds €)
    - macro_return_cum5: float (Mds €)
    - state_budget_an1: float (Dépenses Budget Général État seul An 1, réf: 503.6)
    - state_budget_an5: float (Dépenses Budget Général État seul An 5, réf: 503.6)
    - state_rev_an1: float (Recettes État seul An 1, réf: 308.3)
    - state_rev_an5: float (Recettes État seul An 5, réf: 308.3)
    - poles_depenses_10: list of 10 tuples:
        (pole_label, acteurs, actuel_mds, an1_mds, an5_mds, explication)
    - familles_recettes_7: list of 7 tuples:
        (rec_label, qui_paie, actuel_mds, an1_mds, an5_mds, gain_label, explication)
    - cas_types_8: list of 8 tuples:
        (profil, revenu_actuel, mesures_applicables, impact_an_eur, impact_mois_eur, commentaire)
    - repartition_doughnut: list of tuples (label, value_mds)
    """
    cand_name = cfg["candidate_name"]
    safe_name = cand_name.replace(" ", "_")
    cand_dir = cfg["candidate_dir"]
    os.makedirs(cand_dir, exist_ok=True)

    excel_path = os.path.join(cand_dir, f"Chiffrage_Programme_{safe_name}_2027.xlsx")
    html_path = os.path.join(cand_dir, f"Visualisation_Programme_et_Comparatif_Budget_{safe_name}.html")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styles
    ff = "Calibri"
    f_title = Font(name=ff, size=14, bold=True, color="FFFFFF")
    f_sub = Font(name=ff, size=10, italic=True, color="E2E8F0")
    f_sec = Font(name=ff, size=11, bold=True, color="FFFFFF")
    f_th = Font(name=ff, size=10, bold=True, color="FFFFFF")
    f_data = Font(name=ff, size=10, bold=False, color="0F172A")
    f_bold = Font(name=ff, size=10, bold=True, color="0F172A")
    f_green = Font(name=ff, size=10, bold=True, color="15803D")
    f_red = Font(name=ff, size=10, bold=True, color="B91C1C")
    f_blue = Font(name=ff, size=10, bold=True, color="1D4ED8")
    f_tot = Font(name=ff, size=11, bold=True, color="FFFFFF")

    fill_navy = PatternFill("solid", fgColor="0F172A")
    fill_blue = PatternFill("solid", fgColor="1E3A8A")
    fill_green = PatternFill("solid", fgColor="14532D")
    fill_red = PatternFill("solid", fgColor="7F1D1D")
    fill_th = PatternFill("solid", fgColor="334155")
    fill_zebra = PatternFill("solid", fgColor="F8FAFC")
    fill_white = PatternFill("solid", fgColor="FFFFFF")
    fill_sub = PatternFill("solid", fgColor="DBEAFE")

    thin = Side(border_style="thin", color="CBD5E1")
    med = Side(border_style="medium", color="0F172A")
    b_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
    b_sub = Border(left=thin, right=thin, top=med, bottom=med)

    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center", wrap_text=True)

    NUM_FMT = '+#,##0.0" Mds €";[Red]-#,##0.0" Mds €";0.0" Mds €"'
    ABS_FMT = '#,##0.0" Mds €";[Red]-#,##0.0" Mds €";0.0" Mds €"'
    PCT_FMT = '+0.0%;-0.0%;0.0%'

    def banner(ws, t1, t2, max_c=12):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_c)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_c)
        c1 = ws.cell(1, 1, t1); c1.font = f_title; c1.fill = fill_navy; c1.alignment = al_c
        c2 = ws.cell(2, 1, t2); c2.font = f_sub; c2.fill = fill_navy; c2.alignment = al_c
        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 20

    # ==========================================================================
    # TAB 1 : Chiffrage_Detaille_Familles
    # ==========================================================================
    ws1 = wb.create_sheet("Chiffrage_Detaille_Familles")
    ws1.views.sheetView[0].showGridLines = True
    banner(ws1, f"CHIFFRAGE BUDGÉTAIRE DÉTAILLÉ DU PROGRAMME PRÉSIDENTIEL DE {cand_name.upper()} ({cfg['party_name'].upper()} - 2027)", cfg["tagline"], 12)

    headers1 = [
        "Code", "Grande Famille", "Ministère / Sphère", "Intitulé Détaillé de la Mesure",
        "Nature de la Mesure", "Impact An 1\n(Mds €)", "Impact An 5\nPlein Régime (Mds €/an)",
        "Cumul 5 Ans\nQuinquennat (Mds €)", "QUI EST CONCERNÉ ?\n(Cible / Payeur / Bénéficiaire)",
        "COMMENT & QUAND ?\n(Mécanisme juridique & Calendrier)", "POURQUOI ?\n(Objectif économique & politique)",
        "Source & Référence"
    ]
    for c, h in enumerate(headers1, 1):
        cell = ws1.cell(4, c, h)
        cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c; cell.border = b_cell
    ws1.row_dimensions[4].height = 30

    r = 4
    sub_rows = []
    for fam in cfg["families_measures"]:
        r += 1
        ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        sc = ws1.cell(r, 1, fam["family_title"])
        sc.font = f_sec; sc.fill = fill_blue; sc.alignment = al_l
        ws1.row_dimensions[r].height = 24
        start_m = r + 1
        for idx, m in enumerate(fam["measures"]):
            r += 1
            ws1.row_dimensions[r].height = 28
            bg = fill_white if idx % 2 == 0 else fill_zebra
            vals = [
                m["code"], fam["short_name"], m["sphere"], m["title"], m["type_label"],
                m["an1"], m["an5"], m["cum5"], m["qui"], m["comment"], m["pourquoi"], m["source"]
            ]
            for c, v in enumerate(vals, 1):
                cell = ws1.cell(r, c, v)
                cell.fill = bg; cell.border = b_cell; cell.font = f_data
                if c in (1, 2, 5):
                    cell.alignment = al_c; cell.font = f_bold
                elif c in (6, 7, 8):
                    cell.alignment = al_r; cell.number_format = NUM_FMT; cell.font = f_bold
                else:
                    cell.alignment = al_l
        end_m = r
        r += 1
        ws1.row_dimensions[r].height = 24
        for c in range(1, 13):
            ws1.cell(r, c).fill = fill_sub; ws1.cell(r, c).border = b_sub; ws1.cell(r, c).font = f_bold
        ws1.cell(r, 1, f"SOUS-TOTAL {fam['short_name'].upper()}").alignment = al_l
        ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws1.cell(r, 6, f"=SUM(F{start_m}:F{end_m})").number_format = NUM_FMT; ws1.cell(r, 6).alignment = al_r
        ws1.cell(r, 7, f"=SUM(G{start_m}:G{end_m})").number_format = NUM_FMT; ws1.cell(r, 7).alignment = al_r
        ws1.cell(r, 8, f"=SUM(H{start_m}:H{end_m})").number_format = NUM_FMT; ws1.cell(r, 8).alignment = al_r
        sub_rows.append((fam, r))

    # Grand totals at bottom of Tab 1
    r += 2
    all_f = "+".join([f"F{sr}" for _, sr in sub_rows])
    all_g = "+".join([f"G{sr}" for _, sr in sub_rows])
    all_h = "+".join([f"H{sr}" for _, sr in sub_rows])

    ws1.row_dimensions[r].height = 25
    for c in range(1, 13):
        ws1.cell(r, c).fill = fill_navy; ws1.cell(r, c).font = f_tot; ws1.cell(r, c).border = b_sub
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws1.cell(r, 1, "SOLDE BUDGÉTAIRE PRIMAIRE STATIQUE (SOMME DES FAMILLES)").alignment = al_l
    ws1.cell(r, 6, f"={all_f}").number_format = NUM_FMT; ws1.cell(r, 6).alignment = al_r
    ws1.cell(r, 7, f"={all_g}").number_format = NUM_FMT; ws1.cell(r, 7).alignment = al_r
    ws1.cell(r, 8, f"={all_h}").number_format = NUM_FMT; ws1.cell(r, 8).alignment = al_r
    r_stat = r

    r += 1
    ws1.row_dimensions[r].height = 24
    for c in range(1, 13):
        ws1.cell(r, c).fill = fill_sub; ws1.cell(r, c).font = f_bold; ws1.cell(r, c).border = b_cell
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws1.cell(r, 1, "EFFET RETOUR MACROÉCONOMIQUE (CROISSANCE / EMPLOI)").alignment = al_l
    ws1.cell(r, 6, cfg.get("macro_return_an1", 0.0)).number_format = NUM_FMT; ws1.cell(r, 6).alignment = al_r
    ws1.cell(r, 7, cfg.get("macro_return_an5", 0.0)).number_format = NUM_FMT; ws1.cell(r, 7).alignment = al_r
    ws1.cell(r, 8, cfg.get("macro_return_cum5", 0.0)).number_format = NUM_FMT; ws1.cell(r, 8).alignment = al_r
    r_macro = r

    r += 1
    ws1.row_dimensions[r].height = 26
    for c in range(1, 13):
        ws1.cell(r, c).fill = fill_green; ws1.cell(r, c).font = f_tot; ws1.cell(r, c).border = b_sub
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws1.cell(r, 1, "IMPACT NET SUR LE SOLDE PUBLIC / DÉFICIT MAASTRICHT").alignment = al_l
    ws1.cell(r, 6, f"=F{r_stat}+F{r_macro}").number_format = NUM_FMT; ws1.cell(r, 6).alignment = al_r
    ws1.cell(r, 7, f"=G{r_stat}+G{r_macro}").number_format = NUM_FMT; ws1.cell(r, 7).alignment = al_r
    ws1.cell(r, 8, f"=H{r_stat}+H{r_macro}").number_format = NUM_FMT; ws1.cell(r, 8).alignment = al_r

    w1 = [10, 22, 20, 44, 20, 15, 16, 16, 28, 32, 32, 24]
    for i, w in enumerate(w1, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    # ==========================================================================
    # TAB 2 : Synthese_Familles_Equilibre
    # ==========================================================================
    ws2 = wb.create_sheet("Synthese_Familles_Equilibre")
    ws2.views.sheetView[0].showGridLines = True
    banner(ws2, f"SYNTHÈSE PAR FAMILLE & ÉQUILIBRE BUDGÉTAIRE DU PROGRAMME DE {cand_name.upper()}", "Sous-totaux liés dynamiquement à l'onglet 1", 6)

    h2 = ["N° Famille", "Famille Budgétaire", "Catégorie", "Impact An 1 (Mds €)", "Impact An 5 (Mds €/an)", "Cumul 5 Ans (Mds €)"]
    for c, h in enumerate(h2, 1):
        cell = ws2.cell(4, c, h); cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c; cell.border = b_cell
    ws2.row_dimensions[4].height = 26

    r2 = 4
    for idx, (fam, sr) in enumerate(sub_rows):
        r2 += 1
        ws2.row_dimensions[r2].height = 24
        bg = fill_white if idx % 2 == 0 else fill_zebra
        ws2.cell(r2, 1, fam["family_id"]).alignment = al_c; ws2.cell(r2, 1).font = f_bold; ws2.cell(r2, 1).fill = bg; ws2.cell(r2, 1).border = b_cell
        ws2.cell(r2, 2, fam["short_name"]).alignment = al_l; ws2.cell(r2, 2).font = f_bold; ws2.cell(r2, 2).fill = bg; ws2.cell(r2, 2).border = b_cell
        ws2.cell(r2, 3, fam["category"]).alignment = al_c; ws2.cell(r2, 3).font = f_data; ws2.cell(r2, 3).fill = bg; ws2.cell(r2, 3).border = b_cell
        for col_i, col_let in [(4, "F"), (5, "G"), (6, "H")]:
            c_ref = ws2.cell(r2, col_i, f"=Chiffrage_Detaille_Familles!{col_let}{sr}")
            c_ref.number_format = NUM_FMT; c_ref.alignment = al_r; c_ref.font = f_bold; c_ref.fill = bg; c_ref.border = b_cell

    for i, w in enumerate([16, 38, 28, 20, 22, 22], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # ==========================================================================
    # TAB 3 : Simulateur_Cas_Types_Micro
    # ==========================================================================
    ws3 = wb.create_sheet("Simulateur_Cas_Types_Micro")
    ws3.views.sheetView[0].showGridLines = True
    banner(ws3, f"SIMULATEUR D'IMPACT MICROÉCONOMIQUE : CE QUE LE PROGRAMME DE {cand_name.upper()} CHANGE POUR LES FRANÇAIS", "Estimation de l'impact financier annuel et mensuel sur 8 profils types de ménages et d'entreprises", 6)

    h3 = ["Profil Type (Ménage / Entreprise)", "Situation & Revenu Actuel", "Mesures du Programme Applicables", "Gain / Perte Annuel (€/an)", "Gain / Perte Mensuel (€/mois)", "Analyse & Explication"]
    for c, h in enumerate(h3, 1):
        cell = ws3.cell(4, c, h); cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c; cell.border = b_cell
    ws3.row_dimensions[4].height = 28

    for idx, ct in enumerate(cfg["cas_types_8"], start=5):
        ws3.row_dimensions[idx].height = 28
        bg = fill_white if idx % 2 == 0 else fill_zebra
        for c, val in enumerate(ct, 1):
            cell = ws3.cell(idx, c, val); cell.fill = bg; cell.border = b_cell; cell.font = f_data
            if c == 1:
                cell.font = f_bold; cell.alignment = al_l
            elif c in (4, 5):
                cell.font = f_bold; cell.alignment = al_r; cell.number_format = '+#,##0" €";[Red]-#,##0" €";0" €"'
            else:
                cell.alignment = al_l

    for i, w in enumerate([28, 26, 40, 20, 20, 44], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    # ==========================================================================
    # TAB 4 : 04_Comparatif_Etat_vs_Candidat
    # ==========================================================================
    ws4 = wb.create_sheet("04_Comparatif_Etat_vs_Candidat")
    ws4.views.sheetView[0].showGridLines = True
    banner(ws4, f"COMPARATIF DIRECT : BUDGET ACTUEL DE LA FRANCE VS BUDGET PRÉVU PAR {cand_name.upper()} (2027-2032)", "Comparaison ligne à ligne des Recettes, des Dépenses (par pôle ministériel et social) et du Déficit — En Milliards d'Euros (Mds €)", 8)

    # Calculate totals from poles & recettes
    dep_act = sum(x[2] for x in cfg["poles_depenses_10"])
    dep_an1 = sum(x[3] for x in cfg["poles_depenses_10"])
    dep_an5 = sum(x[4] for x in cfg["poles_depenses_10"])

    rec_act = sum(x[2] for x in cfg["familles_recettes_7"])
    rec_an1 = sum(x[3] for x in cfg["familles_recettes_7"])
    rec_an5 = sum(x[4] for x in cfg["familles_recettes_7"])

    def_act = rec_act - dep_act
    def_an1 = rec_an1 - dep_an1
    def_an5 = rec_an5 - dep_an5

    r = 4
    ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    s = ws4.cell(r, 1, f"1. TABLEAU DE BORD MACROÉCONOMIQUE : FRANCE ACTUELLE VS TRAJECTOIRE {cand_name.upper()}")
    s.font = f_sec; s.fill = fill_blue; s.alignment = al_l; ws4.row_dimensions[r].height = 24

    r = 5
    h4_1 = ["Indicateur Budgétaire Majeur", "Périmètre", "Budget Actuel\n(Réf. 2026)", "Budget Candidat\nAn 1", "Budget Candidat\nAn 5 (Plein Régime)", "Écart Net An 5\n(Mds €/an)", "Évolution\n(%)", "Mécanisme explicatif"]
    for c, h in enumerate(h4_1, 1):
        cell = ws4.cell(r, c, h); cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c; cell.border = b_cell
    ws4.row_dimensions[r].height = 28

    sb_an1 = cfg["state_budget_an1"]
    sb_an5 = cfg["state_budget_an5"]
    sr_an1 = cfg["state_rev_an1"]
    sr_an5 = cfg["state_rev_an5"]

    macro_items = [
        ("Dépenses Publiques Consolidées (APU)", "État + Sécu + Collectivités", dep_act, dep_an1, dep_an5, dep_an5 - dep_act, (dep_an5 - dep_act)/dep_act, "Évolution nette des 10 pôles de dépenses publiques"),
        ("Recettes Publiques Consolidées (APU)", "Impôts + Cotisations + Non-fiscal", rec_act, rec_an1, rec_an5, rec_an5 - rec_act, (rec_an5 - rec_act)/rec_act, "Évolution nette des prélèvements obligatoires et effet retour croissance"),
        ("Solde Public Consolidé (Déficit Maastricht)", "Solde Public National", def_act, def_an1, def_an5, def_an5 - def_act, (def_an5 - def_act)/abs(def_act), "Amélioration (+) ou dégradation (-) du déficit public annuel"),
        ("Dépenses du Budget Général de l'État seul", "Ministères Centraux (LOLF)", 503.6, sb_an1, sb_an5, sb_an5 - 503.6, (sb_an5 - 503.6)/503.6, "Impact direct sur les 32 missions ministérielles de l'État"),
        ("Recettes Nettes du Budget Général de l'État", "Budget Général de l'État", 308.3, sr_an1, sr_an5, sr_an5 - 308.3, (sr_an5 - 308.3)/308.3, "Impact des mesures fiscales d'État (IR, IS, TVA, IFI/Successions)"),
    ]
    for idx, item in enumerate(macro_items):
        r += 1
        ws4.row_dimensions[r].height = 24
        bg = fill_white if idx % 2 == 0 else fill_zebra
        for c in range(1, 9):
            ws4.cell(r, c).fill = bg; ws4.cell(r, c).border = b_cell; ws4.cell(r, c).font = f_data
        ws4.cell(r, 1, item[0]).font = f_bold; ws4.cell(r, 1).alignment = al_l
        ws4.cell(r, 2, item[1]).alignment = al_c
        for ci, v in [(3, item[2]), (4, item[3]), (5, item[4]), (6, item[5])]:
            cv = ws4.cell(r, ci, v); cv.number_format = ABS_FMT if ci < 6 else NUM_FMT; cv.alignment = al_r; cv.font = f_bold
        cp = ws4.cell(r, 7, item[6]); cp.number_format = PCT_FMT; cp.alignment = al_c; cp.font = f_bold
        ws4.cell(r, 8, item[7]).alignment = al_l

    # Section 2 : Les 10 Pôles de Dépenses
    r += 2
    ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    s = ws4.cell(r, 1, f"2. COMPARATIF DES 10 PÔLES DE DÉPENSES PUBLIQUES : ACTUEL VS {cand_name.upper()} (Mds €)")
    s.font = f_sec; s.fill = fill_red; s.alignment = al_l; ws4.row_dimensions[r].height = 24

    r += 1
    h4_2 = ["Grands Pôles de Dépenses Publiques", "Acteurs", "Dépense Actuelle\n(Mds €)", "Candidat An 1\n(Mds €)", "Candidat An 5\n(Mds €)", "Écart An 5\n(Mds €/an)", "Évolution\n(%)", "Mesures du Candidat expliquant l'écart"]
    for c, h in enumerate(h4_2, 1):
        cell = ws4.cell(r, c, h); cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c; cell.border = b_cell
    ws4.row_dimensions[r].height = 28

    s_r2 = r + 1
    for idx, p in enumerate(cfg["poles_depenses_10"]):
        r += 1
        ws4.row_dimensions[r].height = 26
        bg = fill_white if idx % 2 == 0 else fill_zebra
        ecart = p[4] - p[2]
        pct = ecart / p[2] if p[2] else 0.0
        for c in range(1, 9):
            ws4.cell(r, c).fill = bg; ws4.cell(r, c).border = b_cell; ws4.cell(r, c).font = f_data
        ws4.cell(r, 1, p[0]).font = f_bold; ws4.cell(r, 1).alignment = al_l
        ws4.cell(r, 2, p[1]).alignment = al_c
        for ci, v in [(3, p[2]), (4, p[3]), (5, p[4]), (6, ecart)]:
            cv = ws4.cell(r, ci, v); cv.number_format = ABS_FMT if ci < 6 else NUM_FMT; cv.alignment = al_r; cv.font = f_bold
        cp = ws4.cell(r, 7, pct); cp.number_format = PCT_FMT; cp.alignment = al_c; cp.font = f_bold
        ws4.cell(r, 8, p[5]).alignment = al_l

    r += 1
    ws4.row_dimensions[r].height = 26
    for c in range(1, 9):
        ws4.cell(r, c).fill = fill_navy; ws4.cell(r, c).font = f_tot; ws4.cell(r, c).border = b_sub
    ws4.cell(r, 1, "TOTAL DES DÉPENSES PUBLIQUES CONSOLIDÉES (APU)").alignment = al_l
    ws4.cell(r, 2, "France entière").alignment = al_c
    ws4.cell(r, 3, f"=SUM(C{s_r2}:C{r-1})").number_format = ABS_FMT; ws4.cell(r, 3).alignment = al_r
    ws4.cell(r, 4, f"=SUM(D{s_r2}:D{r-1})").number_format = ABS_FMT; ws4.cell(r, 4).alignment = al_r
    ws4.cell(r, 5, f"=SUM(E{s_r2}:E{r-1})").number_format = ABS_FMT; ws4.cell(r, 5).alignment = al_r
    ws4.cell(r, 6, f"=SUM(F{s_r2}:F{r-1})").number_format = NUM_FMT; ws4.cell(r, 6).alignment = al_r
    ws4.cell(r, 7, (dep_an5 - dep_act)/dep_act).number_format = PCT_FMT; ws4.cell(r, 7).alignment = al_c

    # Section 3 : Les 7 Familles de Recettes
    r += 2
    ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    s = ws4.cell(r, 1, f"3. COMPARATIF DES RECETTES PUBLIQUES : ACTUEL VS {cand_name.upper()} (Mds €)")
    s.font = f_sec; s.fill = fill_green; s.alignment = al_l; ws4.row_dimensions[r].height = 24

    r += 1
    h4_3 = ["Grandes Familles de Recettes Publiques", "Qui paie ?", "Recettes Actuelles\n(Mds €)", "Candidat An 1\n(Mds €)", "Candidat An 5\n(Mds €)", "Écart An 5\n(Mds €/an)", "Impact Contribuable", "Détail des mesures fiscales et sociales"]
    for c, h in enumerate(h4_3, 1):
        cell = ws4.cell(r, c, h); cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c; cell.border = b_cell
    ws4.row_dimensions[r].height = 28

    s_r3 = r + 1
    for idx, rc in enumerate(cfg["familles_recettes_7"]):
        r += 1
        ws4.row_dimensions[r].height = 26
        bg = fill_white if idx % 2 == 0 else fill_zebra
        ecart = rc[4] - rc[2]
        for c in range(1, 9):
            ws4.cell(r, c).fill = bg; ws4.cell(r, c).border = b_cell; ws4.cell(r, c).font = f_data
        ws4.cell(r, 1, rc[0]).font = f_bold; ws4.cell(r, 1).alignment = al_l
        ws4.cell(r, 2, rc[1]).alignment = al_c
        for ci, v in [(3, rc[2]), (4, rc[3]), (5, rc[4]), (6, ecart)]:
            cv = ws4.cell(r, ci, v); cv.number_format = ABS_FMT if ci < 6 else NUM_FMT; cv.alignment = al_r; cv.font = f_bold
        ws4.cell(r, 7, rc[5]).alignment = al_c; ws4.cell(r, 7).font = f_bold
        ws4.cell(r, 8, rc[6]).alignment = al_l

    r += 1
    ws4.row_dimensions[r].height = 26
    for c in range(1, 9):
        ws4.cell(r, c).fill = fill_navy; ws4.cell(r, c).font = f_tot; ws4.cell(r, c).border = b_sub
    ws4.cell(r, 1, "TOTAL DES RECETTES PUBLIQUES CONSOLIDÉES (APU)").alignment = al_l
    ws4.cell(r, 2, "Contribuables").alignment = al_c
    ws4.cell(r, 3, f"=SUM(C{s_r3}:C{r-1})").number_format = ABS_FMT; ws4.cell(r, 3).alignment = al_r
    ws4.cell(r, 4, f"=SUM(D{s_r3}:D{r-1})").number_format = ABS_FMT; ws4.cell(r, 4).alignment = al_r
    ws4.cell(r, 5, f"=SUM(E{s_r3}:E{r-1})").number_format = ABS_FMT; ws4.cell(r, 5).alignment = al_r
    ws4.cell(r, 6, f"=SUM(F{s_r3}:F{r-1})").number_format = NUM_FMT; ws4.cell(r, 6).alignment = al_r

    for i, w in enumerate([38, 24, 16, 16, 16, 16, 20, 52], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w

    # ==========================================================================
    # TAB 5 : 05_Graphiques_Comparatifs (5 Native Excel Charts)
    # ==========================================================================
    ws5 = wb.create_sheet("05_Graphiques_Comparatifs")
    ws5.views.sheetView[0].showGridLines = True
    banner(ws5, f"GRAPHIQUES DU PROGRAMME DE {cand_name.upper()} & COMPARAISON AVEC LE BUDGET DE L'ÉTAT", "Tableaux sources (colonnes A-D) et 5 Graphiques Excel natifs intégrés (colonnes F-P)", 15)

    # T1: Familles
    ws5.cell(4, 1, "TABLEAU 1 : IMPACT PAR FAMILLE (Mds €)").font = f_bold
    ws5.cell(5, 1, "Famille").font = f_th; ws5.cell(5, 1).fill = fill_th
    ws5.cell(5, 2, "Impact An 1").font = f_th; ws5.cell(5, 2).fill = fill_th
    ws5.cell(5, 3, "Impact An 5").font = f_th; ws5.cell(5, 3).fill = fill_th
    for idx, fam in enumerate(cfg["families_measures"], start=6):
        s1 = sum(m["an1"] for m in fam["measures"])
        s5 = sum(m["an5"] for m in fam["measures"])
        ws5.cell(idx, 1, fam["short_name"]).border = b_cell
        ws5.cell(idx, 2, s1).border = b_cell; ws5.cell(idx, 2).number_format = '0.0'
        ws5.cell(idx, 3, s5).border = b_cell; ws5.cell(idx, 3).number_format = '0.0'
    end_t1 = 5 + len(cfg["families_measures"])

    # T2: Doughnut
    r_t2 = end_t1 + 3
    ws5.cell(r_t2 - 1, 1, "TABLEAU 2 : RÉPARTITION CLÉ DU PROGRAMME (Mds €)").font = f_bold
    ws5.cell(r_t2, 1, "axe").font = f_th; ws5.cell(r_t2, 1).fill = fill_th
    ws5.cell(r_t2, 2, "Montant An 5").font = f_th; ws5.cell(r_t2, 2).fill = fill_th
    for i, (dlbl, dval) in enumerate(cfg["repartition_doughnut"], start=r_t2 + 1):
        ws5.cell(i, 1, dlbl).border = b_cell
        ws5.cell(i, 2, dval).border = b_cell; ws5.cell(i, 2).number_format = '0.0'
    end_t2 = r_t2 + len(cfg["repartition_doughnut"])

    # T3: Macro
    r_t3 = end_t2 + 3
    ws5.cell(r_t3 - 1, 1, "TABLEAU 3 : COMPARATIF MACRO (Mds €)").font = f_bold
    for ci, ch in enumerate(["Agrégat", "Budget Actuel", "Candidat An 1", "Candidat An 5"], 1):
        ws5.cell(r_t3, ci, ch).font = f_th; ws5.cell(r_t3, ci).fill = fill_th
    t3_data = [
        ("Dépenses Publiques (APU)", dep_act, dep_an1, dep_an5),
        ("Recettes Publiques (APU)", rec_act, rec_an1, rec_an5),
        ("Déficit Maastricht (Val. abs.)", abs(def_act), abs(def_an1), abs(def_an5)),
        ("Budget Général État seul", 503.6, sb_an1, sb_an5),
    ]
    for i, row_vals in enumerate(t3_data, start=r_t3 + 1):
        for ci, v in enumerate(row_vals, 1):
            c = ws5.cell(i, ci, v); c.border = b_cell
            if ci > 1: c.number_format = '0.0'
    end_t3 = r_t3 + 4

    # T4: 10 Poles
    r_t4 = end_t3 + 3
    ws5.cell(r_t4 - 1, 1, "TABLEAU 4 : DÉPENSES PAR PÔLE (Mds €)").font = f_bold
    for ci, ch in enumerate(["Pôle de Dépense", "Budget Actuel", "Candidat An 5"], 1):
        ws5.cell(r_t4, ci, ch).font = f_th; ws5.cell(r_t4, ci).fill = fill_th
    for i, p in enumerate(cfg["poles_depenses_10"], start=r_t4 + 1):
        ws5.cell(i, 1, p[0]).border = b_cell
        ws5.cell(i, 2, p[2]).border = b_cell; ws5.cell(i, 2).number_format = '0.0'
        ws5.cell(i, 3, p[4]).border = b_cell; ws5.cell(i, 3).number_format = '0.0'
    end_t4 = r_t4 + 10

    # T5: Recettes
    r_t5 = end_t4 + 3
    ws5.cell(r_t5 - 1, 1, "TABLEAU 5 : RECETTES PAR FAMILLE (Mds €)").font = f_bold
    for ci, ch in enumerate(["Famille de Recettes", "Recettes Actuelles", "Candidat An 5"], 1):
        ws5.cell(r_t5, ci, ch).font = f_th; ws5.cell(r_t5, ci).fill = fill_th
    for i, rc in enumerate(cfg["familles_recettes_7"][:6], start=r_t5 + 1):
        ws5.cell(i, 1, rc[0]).border = b_cell
        ws5.cell(i, 2, rc[2]).border = b_cell; ws5.cell(i, 2).number_format = '0.0'
        ws5.cell(i, 3, rc[4]).border = b_cell; ws5.cell(i, 3).number_format = '0.0'
    end_t5 = r_t5 + 6

    # Native Charts
    ch1 = BarChart(); ch1.type = "col"; ch1.style = 10
    ch1.title = f"1. Programme de {cand_name} par Famille (An 1 vs An 5, Mds €)"
    ch1.add_data(Reference(ws5, min_col=2, min_row=5, max_col=3, max_row=end_t1), titles_from_data=True)
    ch1.set_categories(Reference(ws5, min_col=1, min_row=6, max_row=end_t1))
    ch1.width = 24; ch1.height = 13
    ws5.add_chart(ch1, "F4")

    ch2 = DoughnutChart()
    ch2.title = f"2. Architecture Budgétaire du Programme de {cand_name} (Mds €)"
    ch2.add_data(Reference(ws5, min_col=2, min_row=r_t2, max_row=end_t2), titles_from_data=True)
    ch2.set_categories(Reference(ws5, min_col=1, min_row=r_t2+1, max_row=end_t2))
    ch2.width = 18; ch2.height = 12
    ws5.add_chart(ch2, "F30")

    ch3 = BarChart(); ch3.type = "col"; ch3.style = 11
    ch3.title = f"3. Comparatif Global : Budget Actuel vs Budget {cand_name} (Mds €)"
    ch3.add_data(Reference(ws5, min_col=2, min_row=r_t3, max_col=4, max_row=end_t3), titles_from_data=True)
    ch3.set_categories(Reference(ws5, min_col=1, min_row=r_t3+1, max_row=end_t3))
    ch3.width = 24; ch3.height = 13
    ws5.add_chart(ch3, "F55")

    ch4 = BarChart(); ch4.type = "bar"; ch4.style = 12
    ch4.title = f"4. Où va l'argent : Dépenses par Pôle — Actuel vs {cand_name} An 5 (Mds €)"
    ch4.add_data(Reference(ws5, min_col=2, min_row=r_t4, max_col=3, max_row=end_t4), titles_from_data=True)
    ch4.set_categories(Reference(ws5, min_col=1, min_row=r_t4+1, max_row=end_t4))
    ch4.width = 24; ch4.height = 15
    ws5.add_chart(ch4, "F82")

    ch5 = BarChart(); ch5.type = "col"; ch5.style = 10
    ch5.title = f"5. D'où vient l'argent : Recettes par Famille — Actuel vs {cand_name} An 5 (Mds €)"
    ch5.add_data(Reference(ws5, min_col=2, min_row=r_t5, max_col=3, max_row=end_t5), titles_from_data=True)
    ch5.set_categories(Reference(ws5, min_col=1, min_row=r_t5+1, max_row=end_t5))
    ch5.width = 24; ch5.height = 13
    ws5.add_chart(ch5, "F113")

    for i, w in enumerate([40, 20, 20, 20], 1):
        ws5.column_dimensions[get_column_letter(i)].width = w

    wb.save(excel_path)
    return excel_path, html_path
