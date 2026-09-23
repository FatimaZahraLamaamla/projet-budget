# -*- coding: utf-8 -*-
"""
Genere le canevas PPA 2026 : normalise, pre-rempli, verrouille - un fichier par direction.
Onglets : Controle | Marches | Bons de commande | Contrats-Conventions | Contrats prod | Guide | Referentiels(masque)

Les 4 onglets de saisie ont EXACTEMENT les memes 13 colonnes communes (meme ordre, meme
ligne d'en-tete = ligne 3, memes listes deroulantes). "Contrats prod" ajoute 6 colonnes a droite.

Usage :
    python generer_canevas_ppa.py                 # exemplaire temoin (TDF-SITES)
    python generer_canevas_ppa.py --all           # les 22 directions
    python generer_canevas_ppa.py --dir DRH DACG  # liste precise
"""
import sys, os
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

ANNEE   = 2026
OUT_DIR = r"C:\Users\HP\Desktop\Projet Budget\Canevas PPA 2026"
PROT_PWD = "PPA2026"
NROWS   = 150

DIRECTIONS_22 = [
    "ARRIYADIA", "ASSADISSA", "DA", "DACG", "DAR", "DCPA", "DCRI", "DG",
    "DINFORADIO", "DINFOTV", "DJURIDIQUE", "DMPTV", "DPMG", "DPPR", "DRH",
    "DRPSD", "DTSI", "MAGHRIBIA", "TAMAZIGHT", "TDF-SITES", "TDF-EXP&MAIN",
    "TV LAAYOUNE",
]

REFERENTIELS = {
    "LST_Nature":       ["Travaux", "Fournitures", "Services", "\u00c9tudes"],
    "LST_NatureBudget": ["Investissement", "Fonctionnement"],
    "LST_Mode":         ["Appel d'offres ouvert", "Appel d'offres restreint",
                         "Appel d'offres avec pr\u00e9s\u00e9lection", "Concours",
                         "March\u00e9 n\u00e9goci\u00e9", "March\u00e9-cadre / reconductible"],
    "LST_ModeProd":     ["Appel d'offres ouvert", "Concours", "March\u00e9 n\u00e9goci\u00e9",
                         "Contrat de production", "Achat de droits"],
    "LST_Trimestre":    ["T1", "T2", "T3", "T4", "T1-T2", "T3-T4", "Annuel (T1 à T4)"],
    "LST_PME":          ["Oui", "Non"],
    "LST_Freq":         ["Quotidienne", "Hebdomadaire", "Bimensuelle", "Mensuelle",
                         "Ponctuelle", "En continu"],
    "LST_TypeProd":     ["Production interne", "Production externe", "Coproduction",
                         "Achat de droits / programme"],
}

# (titre, largeur, cle_liste|None, obligatoire, type)   type in {"text","num","fixed_type","fixed_mode","auto"}
COMMON_COLS = [
    ("N\u00b0",                                    6,  None,               False, "auto"),
    ("Type de support d'engagement",              24,  None,               True,  "fixed_type"),
    ("Objet\n(le plus d\u00e9taill\u00e9 possible)",    58,  None,               True,  "text"),
    ("Nature de la prestation",                   16,  "LST_Nature",       True,  "list"),
    ("Rubrique budg\u00e9taire",                     28,  None,               False, "text"),
    ("Budget pr\u00e9visionnel TTC (MAD)",             20,  None,               True,  "num"),
    ("Nature du budget",                          16,  "LST_NatureBudget", True,  "list"),
    ("Mode de passation",                         28,  "LST_Mode",         True,  "mode"),
    ("Trimestre de lancement",                    16,  "LST_Trimestre",   True,  "list"),
    ("Dur\u00e9e d'ex\u00e9cution",                       18,  None,               False, "text"),
    ("Lieu d'ex\u00e9cution",                          26,  None,               False, "text"),
    ("PME (march\u00e9 r\u00e9serv\u00e9)",                    14,  "LST_PME",          False, "list"),
    ("Justification du besoin",                   48,  None,               True,  "text"),
]
PROD_COLS = [
    ("Nb de programmes / prestations",            15,  None,          False, "num"),
    ("Fr\u00e9quence de diffusion",                    17,  "LST_Freq",    False, "list"),
    ("Dur\u00e9e de l'\u00e9pisode (mn)",                  15,  None,          False, "num"),
    ("Type de production",                        20,  "LST_TypeProd",False, "list"),
    ("Nb d'\u00e9pisodes",                             11,  None,          False, "num"),
    ("Budget TTC par \u00e9pisode (MAD)",              20,  None,          False, "num"),
]

# Bulle d'aide affichee au clic sur une cellule de la colonne (par intitule de colonne)
HELP = {
    "Type de support d'engagement": "Rempli automatiquement selon l'onglet. Ne pas modifier.",
    "Objet\n(le plus d\u00e9taill\u00e9 possible)":
        "D\u00e9crire pr\u00e9cis\u00e9ment : quoi, pour quel usage, quantit\u00e9 / p\u00e9rim\u00e8tre, site(s). "
        "\u00c9viter les intitul\u00e9s vagues type \u00ab acquisition de mat\u00e9riel \u00bb.",
    "Nature de la prestation": "Travaux / Fournitures / Services / \u00c9tudes. Choisir dans la liste.",
    "Rubrique budg\u00e9taire":
        "Rubrique de la nomenclature budg\u00e9taire \u00e0 laquelle se rattache la d\u00e9pense "
        "(ex : \u00c9quipements de t\u00e9l\u00e9diffusion, Entretien r\u00e9paration et assurance).",
    "Budget pr\u00e9visionnel TTC (MAD)":
        "Un nombre seul, TTC, en dirhams. Pas d'espace, pas de \u00ab MAD \u00bb, pas de texte. Ex : 120000",
    "Nature du budget": "Investissement ou Fonctionnement. Choisir dans la liste.",
    "Mode de passation":
        "Choisir dans la liste. (Rempli automatiquement pour les bons de commande et les contrats.)",
    "Trimestre de lancement":
        "Trimestre o\u00f9 la PROC\u00c9DURE sera lanc\u00e9e. Pour un besoin \u00e9tal\u00e9 : T1-T2, T3-T4 ou Annuel.",
    "Dur\u00e9e d'ex\u00e9cution": "Dur\u00e9e pr\u00e9vue du march\u00e9 / contrat. Ex : 8 mois, 3 ans, march\u00e9-cadre.",
    "Lieu d'ex\u00e9cution": "Ville(s), site(s) ou province(s) concern\u00e9s. S\u00e9parer par des virgules.",
    "PME (march\u00e9 r\u00e9serv\u00e9)": "March\u00e9 r\u00e9serv\u00e9 aux PME ? Oui / Non. (Concerne surtout les march\u00e9s.)",
    "Justification du besoin":
        "Pourquoi ce besoin : contexte, obligation r\u00e9glementaire, renouvellement, "
        "recommandation d'audit, extension de couverture\u2026",
}

# Ligne d'exemple (grisee, verrouillee, toujours visible) par type d'onglet
EX_MARCHE = {
    "Objet\n(le plus d\u00e9taill\u00e9 possible)":
        "(EXEMPLE) Fourniture et installation de 5 onduleurs 10 kVA pour les sites de diffusion X et Y",
    "Nature de la prestation": "Fournitures", "Rubrique budg\u00e9taire": "\u00c9quipements de t\u00e9l\u00e9diffusion",
    "Budget pr\u00e9visionnel TTC (MAD)": 480000, "Nature du budget": "Investissement",
    "Mode de passation": "Appel d'offres ouvert", "Trimestre de lancement": "T2",
    "Dur\u00e9e d'ex\u00e9cution": "8 mois", "Lieu d'ex\u00e9cution": "MEGREZ, BOUKHOUALI",
    "PME (march\u00e9 r\u00e9serv\u00e9)": "Oui",
    "Justification du besoin": "Renouvellement d'onduleurs en fin de vie pour s\u00e9curiser l'alimentation des \u00e9metteurs",
}
EX_BC = {
    "Objet\n(le plus d\u00e9taill\u00e9 possible)":
        "(EXEMPLE) Achat de pi\u00e8ces de rechange \u00e9lectriques pour la maintenance des groupes \u00e9lectrog\u00e8nes",
    "Nature de la prestation": "Fournitures", "Rubrique budg\u00e9taire": "Entretien, r\u00e9paration et assurance",
    "Budget pr\u00e9visionnel TTC (MAD)": 90000, "Nature du budget": "Fonctionnement",
    "Trimestre de lancement": "T1-T2", "Dur\u00e9e d'ex\u00e9cution": "Ann\u00e9e",
    "Lieu d'ex\u00e9cution": "Magasin T\u00e9mara", "PME (march\u00e9 r\u00e9serv\u00e9)": "Oui",
    "Justification du besoin": "Maintenance pr\u00e9ventive et curative des installations \u00e9lectriques des sites",
}
EX_CONTRAT = {
    "Objet\n(le plus d\u00e9taill\u00e9 possible)":
        "(EXEMPLE) Reconduction du contrat de maintenance de la plateforme de supervision r\u00e9seau",
    "Nature de la prestation": "Services", "Rubrique budg\u00e9taire": "Entretien, r\u00e9paration et assurance",
    "Budget pr\u00e9visionnel TTC (MAD)": 240000, "Nature du budget": "Fonctionnement",
    "Trimestre de lancement": "T1", "Dur\u00e9e d'ex\u00e9cution": "3 ans",
    "Lieu d'ex\u00e9cution": "Rabat", "PME (march\u00e9 r\u00e9serv\u00e9)": "Non",
    "Justification du besoin": "Reconduction du contrat n\u00b091/2023 pour assurer la continuit\u00e9 de la supervision",
}

# (nom onglet, valeur "Type de support", cle liste mode | None, mode fige | None, exemple)
SHEETS = [
    ("Marches",              "March\u00e9",                                  "LST_Mode",     None,                                     EX_MARCHE),
    ("Bons de commande",     "Bon de commande",                         None,           "Bon de commande",                        EX_BC),
    ("Contrats-Conventions", "Contrat ou convention de droit commun",   None,           "Contrat ou convention de droit commun",   EX_CONTRAT),
    ("Contrats prod",        "Contrat de production",                   "LST_ModeProd", None,                                     EX_MARCHE),
]
# L'onglet "Contrats prod" (contrats de production audiovisuelle) ne concerne que
# les directions de production. Il n'est genere QUE pour les directions listees ici ;
# vide = aucune direction (cas actuel). Renseigner au moment de la generalisation.
DIRECTIONS_AVEC_PROD = set()

SHEET_ORDER = {"Controle": 0, "Marches": 1, "Bons de commande": 2,
               "Contrats-Conventions": 3, "Contrats prod": 4, "Guide": 5, "Referentiels": 6}

C_TITLE = "1F3864"; C_HEAD = "2F5597"
thin = Side(style="thin", color="BFBFBF")
BORDER   = Border(left=thin, right=thin, top=thin, bottom=thin)
F_TITLE  = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
F_HEAD   = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_META   = Font(name="Calibri", size=10, bold=True)
FILL_TITLE = PatternFill("solid", fgColor=C_TITLE)
FILL_HEAD  = PatternFill("solid", fgColor=C_HEAD)
FILL_META  = PatternFill("solid", fgColor="D9E2F3")
FILL_LOCK  = PatternFill("solid", fgColor="F2F2F2")
FILL_ERR   = PatternFill("solid", fgColor="FFC7CE")
FILL_OK    = PatternFill("solid", fgColor="C6EFCE")
UNLOCKED = Protection(locked=False)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
CENTER   = Alignment(horizontal="center", vertical="center", wrap_text=True)

HEAD_ROW    = 3
EXAMPLE_ROW = 4
DATA_ROW    = 5
FILL_EX     = PatternFill("solid", fgColor="FFF2CC")
F_EX        = Font(italic=True, size=9, color="7F6000")


def build_referentiels(wb):
    ws = wb.create_sheet("Referentiels")
    ws.sheet_state = "hidden"
    for ci, (name, vals) in enumerate(REFERENTIELS.items(), start=1):
        col = get_column_letter(ci)
        ws.cell(row=1, column=ci, value=name).font = Font(bold=True)
        for ri, v in enumerate(vals, start=2):
            ws.cell(row=ri, column=ci, value=v)
        ref = f"'Referentiels'!${col}$2:${col}${1+len(vals)}"
        wb.defined_names.add(DefinedName(name, attr_text=ref))
    ws.protection.sheet = True
    ws.protection.password = PROT_PWD


def build_saisie_sheet(wb, direction, sheet_title, type_support, mode_list_key, mode_fixed, example):
    ws = wb.create_sheet(sheet_title)
    cols = COMMON_COLS + (PROD_COLS if sheet_title == "Contrats prod" else [])
    ncol = len(cols)
    statut_ci = ncol + 1
    statut_L  = get_column_letter(statut_ci)
    last = statut_L
    end  = DATA_ROW + NROWS - 1
    col_letter = {t: get_column_letter(i) for i, (t, *_r) in enumerate(cols, start=1)}
    OBJ = col_letter["Objet\n(le plus d\u00e9taill\u00e9 possible)"]

    # ligne 1 : titre + annee
    ws.merge_cells(f"A1:{get_column_letter(max(ncol-2, 3))}1")
    t = ws["A1"]; t.value = f"PROGRAMME PR\u00c9VISIONNEL DES ACHATS (PPA) {ANNEE}"
    t.font = F_TITLE; t.fill = FILL_TITLE; t.alignment = Alignment(vertical="center")
    ws.cell(row=1, column=ncol-1, value="ANN\u00c9E").font = F_META
    ws.cell(row=1, column=ncol,   value=ANNEE).font = F_META
    ws.row_dimensions[1].height = 24

    # ligne 2 : direction / onglet
    ws.cell(row=2, column=1, value="DIRECTION :").font = F_META
    dc = ws.cell(row=2, column=2, value=direction); dc.font = Font(bold=True, size=11, color="C00000")
    ws.cell(row=2, column=4, value="Onglet :").font = F_META
    ws.cell(row=2, column=5, value=type_support).font = Font(bold=True)
    for c in range(1, statut_ci + 1):
        ws.cell(row=2, column=c).fill = FILL_META
    ws.row_dimensions[2].height = 18

    # ligne 3 : en-tete
    for ci, (title, width, listkey, oblig, kind) in enumerate(cols, start=1):
        cell = ws.cell(row=HEAD_ROW, column=ci, value=title + (" *" if oblig else ""))
        cell.font = F_HEAD; cell.fill = FILL_HEAD; cell.alignment = CENTER; cell.border = BORDER
        ws.column_dimensions[get_column_letter(ci)].width = width
    sc = ws.cell(row=HEAD_ROW, column=statut_ci, value="Statut ligne")
    sc.font = F_HEAD; sc.fill = FILL_HEAD; sc.alignment = CENTER; sc.border = BORDER
    ws.column_dimensions[statut_L].width = 12
    ws.row_dimensions[HEAD_ROW].height = 44

    # ligne 4 : EXEMPLE (verrouillee, grisee jaune, toujours visible sous les volets figes)
    for ci, (title, width, listkey, oblig, kind) in enumerate(cols, start=1):
        cell = ws.cell(row=EXAMPLE_ROW, column=ci)
        cell.border = BORDER; cell.alignment = WRAP_TOP; cell.fill = FILL_EX; cell.font = F_EX
        if kind == "auto":
            cell.value = "ex."
            cell.alignment = CENTER
        elif kind == "fixed_type":
            cell.value = type_support
        elif kind == "mode" and mode_fixed:
            cell.value = mode_fixed
        else:
            cell.value = example.get(title)
        if kind == "num" and isinstance(cell.value, (int, float)):
            cell.number_format = "#,##0"
    es = ws.cell(row=EXAMPLE_ROW, column=statut_ci, value="exemple")
    es.border = BORDER; es.alignment = CENTER; es.fill = FILL_EX; es.font = F_EX
    ws.cell(row=EXAMPLE_ROW, column=1).comment = Comment(
        "Ligne d'exemple : ne rien saisir ici. Commencer votre saisie \u00e0 la ligne suivante.", "PPA")
    ws.row_dimensions[EXAMPLE_ROW].height = 42

    # corps
    for r in range(DATA_ROW, end + 1):
        for ci, (title, width, listkey, oblig, kind) in enumerate(cols, start=1):
            cell = ws.cell(row=r, column=ci)
            cell.border = BORDER; cell.alignment = WRAP_TOP
            if kind == "auto":
                cell.value = f'=IF($C{r}="","",COUNTA($C${DATA_ROW}:$C{r}))'
                cell.alignment = CENTER
            elif kind == "fixed_type":
                cell.value = type_support
                cell.fill = FILL_LOCK
            elif kind == "mode" and mode_fixed:
                cell.value = mode_fixed
                cell.fill = FILL_LOCK
            else:
                cell.protection = UNLOCKED
            if kind == "num":
                cell.number_format = "#,##0"
        st = ws.cell(row=r, column=statut_ci, value=(
            f'=IF($C{r}="","",IF(AND($D{r}<>"",ISNUMBER($F{r}),$G{r}<>"",$H{r}<>"",'
            f'$I{r}<>"",$M{r}<>""),"\u2713","\u26a0"))'))
        st.alignment = CENTER; st.border = BORDER; st.fill = FILL_LOCK

    ws.freeze_panes = f"D{DATA_ROW}"

    # --- validations + bulles d'aide (une seule DataValidation par colonne) ---
    for ci, (title, width, listkey, oblig, kind) in enumerate(cols, start=1):
        L = get_column_letter(ci)
        rng = f"{L}{DATA_ROW}:{L}{end}"
        help_txt = HELP.get(title)
        if kind == "list" and listkey:
            dv = DataValidation(type="list", formula1=listkey, allowBlank=True, showErrorMessage=True)
            dv.errorTitle = "Saisie invalide"; dv.error = "Valeur non autoris\u00e9e : choisir dans la liste."
        elif kind == "mode" and mode_fixed is None and mode_list_key:
            dv = DataValidation(type="list", formula1=mode_list_key, allowBlank=True, showErrorMessage=True)
            dv.errorTitle = "Saisie invalide"; dv.error = "Valeur non autoris\u00e9e : choisir dans la liste."
        elif kind == "num":
            dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                                allowBlank=True, showErrorMessage=True)
            dv.errorTitle = "Montant invalide"
            dv.error = "Saisir un nombre (pas de texte, pas d'espace, pas de \u00ab MAD \u00bb). Ex : 120000"
        elif help_txt:
            # pas de contrainte, juste la bulle d'aide : textLength toujours vraie
            dv = DataValidation(type="textLength", operator="greaterThanOrEqual",
                                formula1="0", allowBlank=True, showErrorMessage=False)
        else:
            continue
        if help_txt:
            dv.showInputMessage = True
            dv.promptTitle = title.split("\n")[0][:32]
            dv.prompt = help_txt
        dv.add(rng)
        ws.add_data_validation(dv)

    # --- mise en forme conditionnelle : rouge si Objet rempli mais champ obligatoire vide ---
    for (title, width, listkey, oblig, kind) in cols:
        if not oblig or kind in ("auto", "fixed_type") or title.startswith("Objet"):
            continue
        L = col_letter[title]
        rng = f"{L}{DATA_ROW}:{L}{end}"
        if kind == "num":
            f = f'=AND(${OBJ}{DATA_ROW}<>"",NOT(ISNUMBER({L}{DATA_ROW})))'
        else:
            f = f'=AND(${OBJ}{DATA_ROW}<>"",{L}{DATA_ROW}="")'
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f], fill=FILL_ERR))

    # mise en forme conditionnelle sur "Statut ligne"
    srng = f"{statut_L}{DATA_ROW}:{statut_L}{end}"
    ws.conditional_formatting.add(srng, FormulaRule(formula=[f'{statut_L}{DATA_ROW}="\u2713"'], fill=FILL_OK))
    ws.conditional_formatting.add(srng, FormulaRule(formula=[f'{statut_L}{DATA_ROW}="\u26a0"'], fill=FILL_ERR))

    ws[f'{col_letter["Rubrique budg\u00e9taire"]}{HEAD_ROW}'].comment = Comment(
        "Rubrique de la nomenclature budg\u00e9taire \u00e0 laquelle se rattache la d\u00e9pense. "
        "Saisie libre pour l'instant ; deviendra une liste d\u00e9roulante quand la "
        "nomenclature officielle sera fournie.", "PPA")

    # --- confort : zoom, curseur, mise en page impression ---
    ws.sheet_view.zoomScale = 100
    try:
        ws.sheet_view.selection[0].activeCell = f"C{DATA_ROW}"
        ws.sheet_view.selection[0].sqref = f"C{DATA_ROW}"
    except Exception:
        pass
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = f"1:{HEAD_ROW}"

    ws.protection.sheet = True
    ws.protection.password = PROT_PWD
    ws.protection.formatCells = False
    ws.protection.sort = False
    ws.protection.autoFilter = False
    ws.auto_filter.ref = f"A{HEAD_ROW}:{last}{end}"
    return sheet_title, ncol, end


def build_controle(wb, sheet_meta, direction):
    ws = wb.create_sheet("Controle")
    ws.sheet_view.showGridLines = False
    ws["B2"] = f"CONTR\u00d4LE DE COH\u00c9RENCE \u2014 PPA {ANNEE} \u2014 {direction}"
    ws["B2"].font = Font(bold=True, size=13, color=C_TITLE)
    ws["B4"] = ("\u00c0 v\u00e9rifier AVANT envoi. Dans les onglets, une cellule ROUGE = information "
                "obligatoire manquante ou montant non num\u00e9rique.")
    ws["B4"].font = Font(italic=True, size=9)

    for j, h in enumerate(["Onglet", "Lignes saisies", "Total budget (MAD)",
                           "Lignes incompl\u00e8tes", "Montants non num\u00e9riques"], start=2):
        c = ws.cell(row=6, column=j, value=h)
        c.font = F_HEAD; c.fill = FILL_HEAD; c.alignment = CENTER; c.border = BORDER
    ws.column_dimensions["B"].width = 34
    for col in ("C", "D", "E", "F"):
        ws.column_dimensions[col].width = 20

    r = 7
    for (title, ncol, end) in sheet_meta:
        q = f"'{title}'"
        ws.cell(row=r, column=2, value=title).border = BORDER
        ws.cell(row=r, column=3, value=f'=COUNTA({q}!$C${DATA_ROW}:$C${end})').border = BORDER
        d = ws.cell(row=r, column=4, value=f'=SUM({q}!$F${DATA_ROW}:$F${end})'); d.border = BORDER
        d.number_format = "#,##0"
        ws.cell(row=r, column=5, value=(
            f'=SUMPRODUCT(({q}!$C${DATA_ROW}:$C${end}<>"")*('
            f'({q}!$D${DATA_ROW}:$D${end}="")+({q}!$G${DATA_ROW}:$G${end}="")+'
            f'({q}!$H${DATA_ROW}:$H${end}="")+({q}!$I${DATA_ROW}:$I${end}="")+'
            f'({q}!$M${DATA_ROW}:$M${end}="")>0))')).border = BORDER
        ws.cell(row=r, column=6, value=(
            f'=SUMPRODUCT(({q}!$C${DATA_ROW}:$C${end}<>"")*'
            f'(NOT(ISNUMBER({q}!$F${DATA_ROW}:$F${end})))*1)')).border = BORDER
        r += 1
    for j, formula in ((2, '"TOTAL"'), (3, f"=SUM(C7:C{r-1})"), (4, f"=SUM(D7:D{r-1})"),
                       (5, f"=SUM(E7:E{r-1})"), (6, f"=SUM(F7:F{r-1})")):
        c = ws.cell(row=r, column=j, value=(formula[1:-1] if isinstance(formula, str) and formula.startswith('"') else formula))
        c.font = Font(bold=True)
    ws.cell(row=r, column=4).number_format = "#,##0"

    ws.cell(row=r + 2, column=2, value="STATUT").font = Font(bold=True)
    sc = ws.cell(row=r + 2, column=3,
                 value=f'=IF(AND(E{r}=0,F{r}=0,C{r}>0),"PR\u00caT \u00c0 ENVOYER","INCOMPLET \u2014 \u00e0 corriger")')
    sc.font = Font(bold=True, size=12)
    ws.conditional_formatting.add(f"C{r+2}", FormulaRule(formula=[f'C{r+2}="PR\u00caT \u00c0 ENVOYER"'], fill=FILL_OK))
    ws.conditional_formatting.add(f"C{r+2}", FormulaRule(formula=[f'C{r+2}<>"PR\u00caT \u00c0 ENVOYER"'], fill=FILL_ERR))
    ws.protection.sheet = True
    ws.protection.password = PROT_PWD


def build_guide(wb, sheet_names):
    ws = wb.create_sheet("Guide")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["B"].width = 115
    lines = [
        ("PPA 2026 \u2014 NOTICE DE REMPLISSAGE", True),
        ("", False),
        (f"1.  Remplir uniquement les onglets de saisie : {' / '.join(sheet_names)}.", False),
        ("2.  La 1re ligne (jaune, \u00ab ex. \u00bb) est un EXEMPLE : ne pas y toucher. Commencer \u00e0 la ligne juste en dessous.", False),
        ("3.  Une ligne = un besoin. Ne pas laisser de ligne vide au milieu.", False),
        ("4.  Cliquer une cellule affiche une bulle d'aide (quoi mettre, format attendu).", False),
        ("5.  Colonnes marqu\u00e9es d'une * = obligatoires. Cellule rouge = obligatoire manquante ou montant non num\u00e9rique.", False),
        ("6.  Colonne \u00ab Statut ligne \u00bb (\u00e0 droite) : \u2713 = ligne compl\u00e8te, \u26a0 = il manque quelque chose.", False),
        ("7.  Choisir dans les listes d\u00e9roulantes (fl\u00e8che \u00e0 droite de la cellule). Ne pas taper une autre valeur.", False),
        ("8.  Budget pr\u00e9visionnel TTC : un nombre seul (ex : 120000). Pas d'espace, pas de 'MAD', pas de texte.", False),
        ("9.  Trimestre de lancement = le trimestre o\u00f9 la PROC\u00c9DURE est lanc\u00e9e. Besoin \u00e9tal\u00e9 : T1-T2 / T3-T4 / Annuel.", False),
        ("10. 'Type de support d'engagement' et (pour BC / Contrats) 'Mode de passation' sont remplis d'office : ne pas y toucher.", False),
        ("11. Ne pas ajouter / supprimer / renommer de colonne ou d'onglet, ne pas cr\u00e9er de ligne de total : la consolidation est automatique.", False),
        ("12. Onglet 'Controle' : le STATUT doit afficher 'PR\u00caT \u00c0 ENVOYER' avant de transmettre le fichier.", False),
        ("13. Conserver le nom du fichier : 'Canevas PPA 2026 - <DIRECTION>.xlsx'.", False),
        ("", False),
        (f"Mot de passe de protection (d\u00e9verrouillage exceptionnel) : {PROT_PWD}", False),
    ]
    for i, (txt, bold) in enumerate(lines, start=2):
        c = ws.cell(row=i, column=2, value=txt)
        c.font = Font(bold=bold, size=13 if bold else 10, color=C_TITLE if bold else "000000")
        c.alignment = WRAP_TOP
    ws.protection.sheet = True
    ws.protection.password = PROT_PWD


def generate_one(direction):
    wb = Workbook(); wb.remove(wb.active)
    wb.calculation.fullCalcOnLoad = True
    build_referentiels(wb)
    sheets = [s for s in SHEETS
              if s[0] != "Contrats prod" or direction in DIRECTIONS_AVEC_PROD]
    meta = [build_saisie_sheet(wb, direction, *s) for s in sheets]
    build_controle(wb, meta, direction)
    build_guide(wb, [s[0] for s in sheets])
    wb._sheets.sort(key=lambda s: SHEET_ORDER.get(s.title, 9))
    wb.active = 0
    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, f"Canevas PPA {ANNEE} - {direction}.xlsx")
    wb.save(path)
    print("OK ->", os.path.basename(path))
    return path


if __name__ == "__main__":
    a = sys.argv[1:]
    if "--all" in a:
        targets = DIRECTIONS_22
    elif "--dir" in a:
        targets = a[a.index("--dir") + 1:]
    else:
        targets = ["TDF-SITES"]
    for d in targets:
        generate_one(d)
    print(f"\n{len(targets)} fichier(s) dans : {OUT_DIR}")
