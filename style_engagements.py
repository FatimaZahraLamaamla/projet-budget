"""
Mise en forme unifiée de l'onglet "ENGAGEMENTS 2026".

- une seule police (Calibri 11, colonnes en gras en 10), gras pour les colonnes clés
  (direction, référence, objet, lignes budgétaires, montants, statut) ;
- en-tête bleu marine #1F4E78, groupes de colonnes séparés ;
- alignements et formats homogènes par type de colonne (texte, code, date, montant) ;
- colonne ÉTAT en tête : icône d'avancement calculée depuis le STATUT
  D'ENGAGEMENT ;
- couleur de ligne selon le statut (règles recréées proprement, couvrant
  aussi les lignes futures) ;
- suppression des surlignages manuels et des règles cassées (#REF!).
"""
import math
from datetime import datetime

from openpyxl.formatting.rule import FormulaRule, Rule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.table import TableFormula, TableStyleInfo

MARINE = "1F4E78"
TEXTE = "262626"
GRILLE = "D9D9D9"
LIGNES_STYLE = 3000  # les règles de couleur couvrent aussi les lignes à venir

# Colonne d'icônes insérée en tête de l'onglet (les colonnes du fichier
# source sont décalées d'un cran : A -> B, ..., AM -> AN)
ENTETE_ETAT = "ÉTAT"


def _dec(col):
    """Lettre de colonne après insertion de la colonne ÉTAT en A."""
    return get_column_letter(column_index_from_string(col) + 1)


# Type de chaque colonne du fichier source : c = code/centré, t = texte long,
# m = montant, d = date, p = pourcentage
COLONNES_SOURCE = {
    "A": ("c", 13), "B": ("c", 14), "C": ("c", 23), "D": ("c", 10), "E": ("c", 7),
    "F": ("t", 66), "G": ("c", 14), "H": ("c", 17), "I": ("c", 18),
    "J": ("c", 14), "K": ("c", 18), "L": ("t", 24), "M": ("t", 34),
    "N": ("m", 17), "O": ("m", 17), "P": ("m", 17), "Q": ("c", 17),
    "R": ("d", 11), "S": ("d", 11), "T": ("t", 26), "U": ("c", 15),
    "V": ("d", 11), "W": ("d", 11), "X": ("d", 11),
    "Y": ("m", 17), "Z": ("m", 17), "AA": ("m", 17), "AB": ("m", 17),
    "AC": ("m", 17), "AD": ("m", 18), "AE": ("c", 24),
    "AF": ("t", 20), "AG": ("m", 17), "AH": ("m", 17), "AI": ("m", 17),
    "AJ": ("m", 17), "AK": ("p", 11), "AL": ("p", 11), "AM": ("t", 50),
}
# Colonnes en gras (lettres du fichier source) : direction, référence, objet,
# lignes budgétaires, statut et tous les montants
GRAS_SOURCE = {"B", "C", "F", "M", "AE"} | {
    col for col, (typ, _) in COLONNES_SOURCE.items() if typ == "m"}
# Premières colonnes de chaque groupe : un trait plus marqué les sépare
DEBUT_GROUPE_SOURCE = {"G", "J", "N", "R", "T", "Y", "AE", "AF", "AM"}

# Mise en page finale, colonne ÉTAT (i = icône) comprise
COLONNES = {"A": ("i", 7)} | {_dec(k): v for k, v in COLONNES_SOURCE.items()}
GRAS = {_dec(c) for c in GRAS_SOURCE}
DEBUT_GROUPE = {"B"} | {_dec(c) for c in DEBUT_GROUPE_SOURCE}
COL_OBJET = _dec("F")
COL_STATUT = _dec("AE")
DERNIERE_COL = _dec("AM")

# Tailles de police : 11 pour le texte normal ; les colonnes en gras sont
# un cran plus petites pour ne pas écraser la lecture (OBJET reste à 11)
TAILLE = 11
TAILLE_GRAS = 10
TAILLE_OBJET = 11
TAILLE_ICONE = 14

FORMATS = {
    "m": "#,##0.00",
    "d": "dd/mm/yyyy",
    "p": "0.0%",
}

# Par statut : fond de ligne pastel, couleur du statut et de l'icône, icône.
# L'icône suit l'avancement du dossier :
# ◔ préengagement en cours → ◑ préengagé → ◕ engagement en cours → ✔ engagé.
STATUTS = {
    "ENGAGE": ("E2EFDA", "006100", "✔"),
    "EN COURS D'ENGAGEMENT": ("DDEBF7", "1F4E79", "◕"),
    "PREENGAGE": ("FFF2CC", "9C6500", "◑"),
    "EN COURS DE PREENGAGEMENT": ("E4DFEC", "5F497A", "◔"),
    "NON ENGAGE": ("FCE4E4", "9C0006", "✖"),
}


def formule_icone(r):
    """Formule de la colonne ÉTAT : icône selon le STATUT D'ENGAGEMENT."""
    f = '""'
    for statut, (_, _, icone) in reversed(STATUTS.items()):
        s = statut.replace('"', '""')
        f = f'IF({COL_STATUT}{r}="{s}","{icone}",{f})'
    return "=" + f


def _taille(col):
    if col == "A":
        return TAILLE_ICONE
    if col == COL_OBJET:
        return TAILLE_OBJET
    return TAILLE_GRAS if col in GRAS else TAILLE


def _hauteur(ws, r):
    """Hauteur de ligne suffisante pour les colonnes de texte long."""
    hauteur = 32
    for col, (typ, larg) in COLONNES.items():
        if typ != "t":
            continue
        v = ws[f"{col}{r}"].value
        if v:
            taille = _taille(col)
            # caractères par ligne : ~1 par unité de largeur à 11 pt, moins
            # en gras ou en plus grand
            par_ligne = larg * 11 / taille * (0.9 if col in GRAS else 1.0)
            n = sum(math.ceil(max(len(p), 1) / par_ligne) for p in str(v).split("\n"))
            hauteur = max(hauteur, min(n, 6) * taille * 1.35 + 6)
    return hauteur


def appliquer_style(ws, fin):
    """Met en forme l'onglet des engagements, lignes 1 à fin."""
    fin_grille = Side(style="thin", color=GRILLE)
    sep = Side(style="medium", color="8EA9C1")
    blanc = Side(style="thin", color="FFFFFF")
    sep_entete = Side(style="medium", color="FFFFFF")
    aucun = PatternFill(fill_type=None)

    # En-tête
    ws.row_dimensions[1].height = 48
    for col, (typ, larg) in COLONNES.items():
        c = ws[f"{col}1"]
        c.font = Font(name="Calibri", size=TAILLE, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=MARINE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=sep_entete if col in DEBUT_GROUPE else blanc,
                          right=blanc, top=blanc, bottom=blanc)
        dim = ws.column_dimensions[col]
        dim.width = larg

    ws["A1"] = ENTETE_ETAT

    # Données
    for r in range(2, fin + 1):
        ws[f"A{r}"] = formule_icone(r)
        ws.row_dimensions[r].height = _hauteur(ws, r)
        for col, (typ, _) in COLONNES.items():
            c = ws[f"{col}{r}"]
            police = "Segoe UI Symbol" if typ == "i" else "Calibri"
            c.font = Font(name=police, size=_taille(col), bold=col in GRAS, color=TEXTE)
            c.fill = aucun
            c.border = Border(left=sep if col in DEBUT_GROUPE else fin_grille,
                              right=fin_grille, top=fin_grille, bottom=fin_grille)
            if typ == "i":
                c.alignment = Alignment(horizontal="center", vertical="center")
                continue
            if typ == "t":
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif typ in ("m", "p"):
                c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if typ in FORMATS:
                c.number_format = FORMATS[typ]
            elif c.is_date and isinstance(c.value, datetime):
                # Nombre saisi dans une colonne de code : une vraie date reste
                # une date, un petit nombre (année, n°) mal formaté en date
                # redevient un nombre (ex. 2022 affiché « 14/07/1905 »).
                if c.value.year < 1910:
                    c.number_format = "General"
                    c.value = (c.value - datetime(1899, 12, 30)).days
                else:
                    c.number_format = FORMATS["d"]
            elif (c.number_format != "General" and not c.is_date
                  and not isinstance(c.value, (int, float))):
                c.number_format = "General"

    # Mise en forme conditionnelle : on repart de zéro
    ws.conditional_formatting._cf_rules.clear()
    plage = f"A2:{DERNIERE_COL}{LIGNES_STYLE}"
    for statut, (fond, texte, _) in STATUTS.items():
        s = statut.replace('"', '""')
        condition = f'${COL_STATUT}2="{s}"'
        ws.conditional_formatting.add(plage, FormulaRule(
            formula=[condition],
            fill=PatternFill("solid", fgColor=fond, bgColor=fond)))
        dxf = DifferentialStyle(font=Font(bold=True, color=texte))
        for cible in (f"A2:A{LIGNES_STYLE}", f"{COL_STATUT}2:{COL_STATUT}{LIGNES_STYLE}"):
            ws.conditional_formatting.add(
                cible, Rule(type="expression", dxf=dxf, formula=[condition]))

    # Tableau Excel : style neutre, sans bandes (la couleur vient du statut).
    # La colonne ÉTAT est déclarée « colonne calculée » : Excel recopie
    # automatiquement la formule de l'icône sur chaque nouvelle ligne.
    for t in ws.tables.values():
        t.tableColumns[0].calculatedColumnFormula = TableFormula(
            attr_text=formule_icone(2)[1:])
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False,
                                          showColumnStripes=False,
                                          showFirstColumn=False, showLastColumn=False)

    ws.sheet_view.zoomScale = 100
    ws.sheet_view.showGridLines = False
