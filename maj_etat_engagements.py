"""
Mise à jour du classeur de tableaux de bord ETAT_D_ENGAGEMENTS.xlsx
à partir d'un nouvel état d'engagements source.

Les onglets sources (ENGAGEMENTS 2026, Parametres, lignes) sont remplacés
par ceux du fichier source ; les onglets de tableaux de bord (Tableau de Bord,
Filtre Interactif, Tableau de Bord Suivi, Données Suivi, PPM 2026) sont
conservés tels quels et se recalculent à l'ouverture dans Excel.

Usage :
    python maj_etat_engagements.py <fichier_source.xlsx> [modele.xlsx] [sortie.xlsx]

Par défaut : modele = ETAT_D_ENGAGEMENTS.xlsx,
             sortie = ETAT_D_ENGAGEMENTS - MAJ <date>.xlsx
"""
import sys
import warnings
from copy import copy
from datetime import date

import openpyxl
from openpyxl.worksheet.table import Table

warnings.filterwarnings("ignore", category=UserWarning)

ONGLETS_SOURCES = ["Parametres", "ENGAGEMENTS 2026", "lignes"]
ONGLET_ENGAGEMENTS = "ENGAGEMENTS 2026"
LIGNES_FORMULES_SUIVI = 3000  # plage couverte par l'onglet "Données Suivi"


def derniere_ligne(ws, col=2):
    """Dernière ligne dont la colonne DIRECTION (B) est renseignée."""
    fin = 1
    for r, (v,) in enumerate(
            ws.iter_rows(min_row=2, min_col=col, max_col=col, values_only=True), start=2):
        if v not in (None, ""):
            fin = r
    return fin


def copier_onglet(src, dst):
    for row in src.iter_rows():
        for c in row:
            d = dst.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                d.font = copy(c.font)
                d.fill = copy(c.fill)
                d.border = copy(c.border)
                d.alignment = copy(c.alignment)
                d.number_format = c.number_format
                d.protection = copy(c.protection)
            if c.hyperlink:
                d.hyperlink = copy(c.hyperlink)
            if c.comment:
                d.comment = copy(c.comment)

    for key, dim in src.column_dimensions.items():
        dd = dst.column_dimensions[key]
        dd.width, dd.hidden = dim.width, dim.hidden
        dd.min, dd.max = dim.min, dim.max
        dd.outlineLevel = dim.outlineLevel
    for key, dim in src.row_dimensions.items():
        dd = dst.row_dimensions[key]
        dd.height, dd.hidden = dim.height, dim.hidden
        dd.outlineLevel = dim.outlineLevel

    for rng in src.merged_cells.ranges:
        dst.merge_cells(str(rng))
    for cf in src.conditional_formatting:
        for rule in cf.rules:
            dst.conditional_formatting.add(str(cf.sqref), rule)
    for dv in src.data_validations.dataValidation:
        dst.add_data_validation(copy(dv))
    for img in getattr(src, "_images", []):
        dst.add_image(img)

    dst.freeze_panes = src.freeze_panes
    dst.auto_filter.ref = src.auto_filter.ref
    dst.sheet_properties.tabColor = src.sheet_properties.tabColor
    dst.sheet_view.zoomScale = src.sheet_view.zoomScale
    dst.sheet_view.showGridLines = src.sheet_view.showGridLines
    dst.page_setup.orientation = src.page_setup.orientation
    dst.page_setup.paperSize = src.page_setup.paperSize

    for t in src.tables.values():
        nt = Table(displayName=t.displayName, ref=t.ref)
        nt.tableStyleInfo = copy(t.tableStyleInfo)
        nt.autoFilter = copy(t.autoFilter)
        nt._initialise_columns()
        for col_src, col_dst in zip(t.tableColumns, nt.tableColumns):
            col_dst.name = col_src.name
        dst.add_table(nt)


def etendre_table(ws, fin):
    """Étend le tableau Excel de l'onglet jusqu'à la dernière ligne de données."""
    for t in ws.tables.values():
        debut, _ = t.ref.split(":")
        col_fin = "".join(ch for ch in t.ref.split(":")[1] if ch.isalpha())
        ancien = t.ref
        t.ref = f"{debut}:{col_fin}{fin}"
        if t.autoFilter is not None:
            t.autoFilter.ref = t.ref
        return ancien, t.ref
    return None, None


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    source = sys.argv[1]
    modele = sys.argv[2] if len(sys.argv) > 2 else "ETAT_D_ENGAGEMENTS.xlsx"
    sortie = sys.argv[3] if len(sys.argv) > 3 else (
        f"ETAT_D_ENGAGEMENTS - MAJ {date.today():%d-%m-%Y}.xlsx")

    wb_src = openpyxl.load_workbook(source)
    wb = openpyxl.load_workbook(modele)

    manquants = [n for n in ONGLETS_SOURCES if n not in wb_src.sheetnames]
    if manquants:
        sys.exit(f"Onglets absents du fichier source : {manquants}")

    # Contrôle des en-têtes : la structure doit être identique au modèle
    h_mod = [c.value for c in wb[ONGLET_ENGAGEMENTS][1]]
    h_src = [c.value for c in wb_src[ONGLET_ENGAGEMENTS][1]]
    if h_mod[:39] != h_src[:39]:
        diff = [(i + 1, a, b) for i, (a, b) in enumerate(zip(h_mod, h_src)) if a != b]
        sys.exit(f"Colonnes différentes entre modèle et source (A à AM) : {diff}")

    for nom in ONGLETS_SOURCES:
        idx = wb.sheetnames.index(nom)
        wb.remove(wb[nom])
        dst = wb.create_sheet(nom, idx)
        copier_onglet(wb_src[nom], dst)

    ws = wb[ONGLET_ENGAGEMENTS]
    fin = derniere_ligne(ws)
    if fin > LIGNES_FORMULES_SUIVI:
        sys.exit(f"{fin} lignes : dépasse la plage de 'Données Suivi' "
                 f"({LIGNES_FORMULES_SUIVI}), à étendre.")
    ancien, nouveau = etendre_table(ws, fin)

    wb.active = 0
    wb.calculation.fullCalcOnLoad = True  # recalcul complet à l'ouverture
    wb.save(sortie)

    print(f"Source      : {source}")
    print(f"Modèle      : {modele}")
    print(f"Sortie      : {sortie}")
    print(f"Lignes      : {fin - 1} dossiers (jusqu'à la ligne {fin})")
    if ancien:
        print(f"Tableau     : {ancien} -> {nouveau}")


if __name__ == "__main__":
    main()
