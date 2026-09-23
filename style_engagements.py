"""
Mise en forme unifiée de l'onglet "ENGAGEMENTS 2026".

- une seule police (Calibri 10), gras réservé aux colonnes clés ;
- en-tête bleu marine #1F4E78, groupes de colonnes séparés ;
- alignements et formats homogènes par type de colonne (texte, code, date, montant) ;
- couleur de ligne selon le STATUT D'ENGAGEMENT (règles recréées proprement,
  couvrant aussi les lignes futures) ;
- suppression des surlignages manuels et des règles cassées (#REF!).
"""
import math

from openpyxl.formatting.rule import FormulaRule, Rule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.table import TableStyleInfo

MARINE = "1F4E78"
TEXTE = "262626"
GRILLE = "D9D9D9"
DERNIERE_COL = "AM"
LIGNES_STYLE = 3000  # les règles de couleur couvrent aussi les lignes à venir

# Type de chaque colonne : c = code/centré, t = texte long, m = montant,
# d = date, p = pourcentage, n = nombre court
COLONNES = {
    "A": ("c", 11), "B": ("c", 14), "C": ("c", 23), "D": ("c", 10), "E": ("c", 7),
    "F": ("t", 60), "G": ("c", 14), "H": ("c", 17), "I": ("c", 18),
    "J": ("c", 14), "K": ("c", 18), "L": ("t", 24), "M": ("t", 34),
    "N": ("m", 16), "O": ("m", 16), "P": ("m", 16), "Q": ("c", 17),
    "R": ("d", 11), "S": ("d", 11), "T": ("t", 26), "U": ("c", 13),
    "V": ("d", 11), "W": ("d", 11), "X": ("d", 11),
    "Y": ("m", 16), "Z": ("m", 15), "AA": ("m", 16), "AB": ("m", 15),
    "AC": ("m", 16), "AD": ("m", 17), "AE": ("c", 20),
    "AF": ("t", 20), "AG": ("m", 14), "AH": ("m", 15), "AI": ("m", 15),
    "AJ": ("m", 15), "AK": ("p", 11), "AL": ("p", 11), "AM": ("t", 50),
}

# Colonnes en gras : identification du dossier, montant retenu, statut
GRAS = {"B", "C", "AD", "AE"}

# Premières colonnes de chaque groupe : un trait plus marqué les sépare
DEBUT_GROUPE = {"G", "J", "N", "R", "T", "Y", "AE", "AF", "AM"}

FORMATS = {
    "m": "#,##0.00",
    "d": "dd/mm/yyyy",
    "p": "0.0%",
}

# Couleurs par statut : fond de ligne pastel + texte du statut
STATUTS = {
    "ENGAGE": ("E2EFDA", "006100"),
    "PREENGAGE": ("FFF2CC", "9C6500"),
    "EN COURS D'ENGAGEMENT": ("DDEBF7", "1F4E79"),
    "EN COURS DE PREENGAGEMENT": ("E4DFEC", "5F497A"),
    "NON ENGAGE": ("FCE4E4", "9C0006"),
}


def _hauteur(ws, r):
    """Hauteur de ligne suffisante pour les colonnes de texte long."""
    lignes = 1
    for col, (typ, larg) in COLONNES.items():
        if typ != "t":
            continue
        v = ws[f"{col}{r}"].value
        if v:
            txt = str(v)
            n = sum(math.ceil(max(len(p), 1) / (larg * 1.15)) for p in txt.split("\n"))
            lignes = max(lignes, n)
    return max(30, min(lignes, 5) * 13.5 + 4)


def appliquer_style(ws, fin):
    """Met en forme l'onglet des engagements, lignes 1 à fin."""
    fin_grille = Side(style="thin", color=GRILLE)
    sep = Side(style="medium", color="8EA9C1")
    blanc = Side(style="thin", color="FFFFFF")
    sep_entete = Side(style="medium", color="FFFFFF")
    aucun = PatternFill(fill_type=None)

    # En-tête
    ws.row_dimensions[1].height = 45
    for col, (typ, larg) in COLONNES.items():
        c = ws[f"{col}1"]
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=MARINE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=sep_entete if col in DEBUT_GROUPE else blanc,
                          right=blanc, top=blanc, bottom=blanc)
        dim = ws.column_dimensions[col]
        dim.width = larg

    # Données
    for r in range(2, fin + 1):
        ws.row_dimensions[r].height = _hauteur(ws, r)
        for col, (typ, _) in COLONNES.items():
            c = ws[f"{col}{r}"]
            c.font = Font(name="Calibri", size=10, bold=col in GRAS, color=TEXTE)
            c.fill = aucun
            c.border = Border(left=sep if col in DEBUT_GROUPE else fin_grille,
                              right=fin_grille, top=fin_grille, bottom=fin_grille)
            if typ == "t":
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif typ in ("m", "p"):
                c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if typ in FORMATS:
                c.number_format = FORMATS[typ]
            elif c.number_format != "General" and not isinstance(c.value, (int, float)):
                c.number_format = "General"

    # Mise en forme conditionnelle : on repart de zéro
    ws.conditional_formatting._cf_rules.clear()
    plage = f"A2:{DERNIERE_COL}{LIGNES_STYLE}"
    for statut, (fond, texte) in STATUTS.items():
        s = statut.replace('"', '""')
        ws.conditional_formatting.add(plage, FormulaRule(
            formula=[f'$AE2="{s}"'],
            fill=PatternFill("solid", fgColor=fond, bgColor=fond)))
        dxf = DifferentialStyle(font=Font(bold=True, color=texte))
        regle = Rule(type="expression", dxf=dxf, formula=[f'$AE2="{s}"'])
        ws.conditional_formatting.add(f"AE2:AE{LIGNES_STYLE}", regle)

    # Tableau Excel : style neutre, sans bandes (la couleur vient du statut)
    for t in ws.tables.values():
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False,
                                          showColumnStripes=False,
                                          showFirstColumn=False, showLastColumn=False)

    ws.sheet_view.zoomScale = 90
    ws.sheet_view.showGridLines = False
