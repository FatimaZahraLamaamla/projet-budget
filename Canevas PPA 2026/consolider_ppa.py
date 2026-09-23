# -*- coding: utf-8 -*-
r"""
Consolide tous les canevas PPA 2026 renvoyes par les directions en un seul classeur.

Lit :  <dossier>\Canevas PPA 2026 - *.xlsx   (4 onglets de saisie chacun)
Ecrit : <Projet Budget>\Consolidation PPA 2026.xlsx
  - "PPA Consolide"      : toutes les lignes a plat (1 ligne = 1 besoin) + Direction + Onglet
  - "Synthese"          : budget par Direction / Nature / Nature budget / Trimestre / Type / Mode
  - "Anomalies"         : lignes avec champ obligatoire manquant ou montant non numerique
  - "Suivi des retours" : qui a renvoye, nb lignes, total, statut

Usage : python consolider_ppa.py  [dossier_des_retours]
"""
import sys, os, glob, datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

_BASE    = r"C:\Users\HP\Desktop\Projet Budget\Canevas PPA 2026"
SRC_DIR  = sys.argv[1] if len(sys.argv) > 1 else os.path.join(_BASE, "Retours")
OUT_PATH = os.path.join(_BASE, "Consolidation PPA 2026.xlsx")
DATA_ROW = 5   # ligne 4 = exemple, saisie a partir de la ligne 5

SHEET_TYPE = {
    "Marches":              "Marché",
    "Bons de commande":     "Bon de commande",
    "Contrats-Conventions": "Contrat ou convention de droit commun",
    "Contrats prod":        "Contrat de production",
}
# Onglets ou "Mode de passation" est pre-rempli/verrouille (formule) -> valeur imposee cote conso
MODE_FIXE = {
    "Bons de commande":     "Bon de commande",
    "Contrats-Conventions": "Contrat ou convention de droit commun",
}
COMMON = ["Objet", "Nature de la prestation", "Rubrique budgétaire",
          "Budget previsionnel TTC (MAD)", "Nature du budget", "Mode de passation",
          "Trimestre de lancement", "Duree d'execution", "Lieu d'execution",
          "PME (marche reserve)", "Justification du besoin"]
PROD  = ["Nb de programmes / prestations", "Frequence de diffusion",
         "Duree de l'episode (mn)", "Type de production", "Nb d'episodes",
         "Budget TTC par episode (MAD)"]
OUT_COLS = ["Direction", "Onglet", "Type de support d'engagement"] + COMMON + PROD + \
           ["Fichier source", "Ligne source"]
# indices (1-based) dans les onglets de saisie
C_OBJET, C_NATURE, C_RUBR, C_BUDGET, C_NATBUD, C_MODE, C_TRIM, C_DUREE, C_LIEU, C_PME, C_JUST = range(3, 14)
PROD_IDX = list(range(14, 20))
OBLIG = {"Nature de la prestation": C_NATURE, "Budget previsionnel TTC (MAD)": C_BUDGET,
         "Nature du budget": C_NATBUD, "Mode de passation": C_MODE,
         "Trimestre de lancement": C_TRIM, "Justification du besoin": C_JUST}

C_TITLE = "1F3864"; C_HEAD = "2F5597"
thin = Side(style="thin", color="BFBFBF")
BORD = Border(left=thin, right=thin, top=thin, bottom=thin)
FH = Font(bold=True, color="FFFFFF"); FILLH = PatternFill("solid", fgColor=C_HEAD)
WRAP = Alignment(wrap_text=True, vertical="top")


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace("\u202f", "").replace(" ", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None
    return None


def main():
    os.makedirs(SRC_DIR, exist_ok=True)
    files = sorted(glob.glob(os.path.join(SRC_DIR, "Canevas PPA 2026 - *.xlsx")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    # directions attendues = fichiers vierges dans le dossier parent
    expected = sorted(os.path.basename(p).replace("Canevas PPA 2026 - ", "").replace(".xlsx", "")
                      for p in glob.glob(os.path.join(_BASE, "Canevas PPA 2026 - *.xlsx"))
                      if not os.path.basename(p).startswith("~$"))
    if not files:
        print("Aucun retour dans", SRC_DIR, "- generation du seul suivi des retours.")

    rows, anomalies, suivi = [], [], []
    for path in files:
        fname = os.path.basename(path)
        dir_from_name = fname.replace("Canevas PPA 2026 - ", "").replace(".xlsx", "")
        try:
            wb = load_workbook(path, data_only=True)
        except Exception as e:
            print("  ! illisible :", fname, e); continue
        direction = dir_from_name
        for sh in SHEET_TYPE:
            if sh in wb.sheetnames:
                v = wb[sh]["B2"].value
                if isinstance(v, str) and v.strip():
                    direction = v.strip()
                break
        n_lignes = 0; total_dir = 0.0; n_anom = 0
        for sh, type_support in SHEET_TYPE.items():
            if sh not in wb.sheetnames:
                continue
            ws = wb[sh]
            for r in range(DATA_ROW, ws.max_row + 1):
                objet = ws.cell(row=r, column=C_OBJET).value
                if objet is None or str(objet).strip() == "":
                    continue
                n_lignes += 1
                rec = {"Direction": direction, "Onglet": sh,
                       "Type de support d'engagement": type_support,
                       "Fichier source": fname, "Ligne source": r}
                rec["Objet"] = objet
                rec["Nature de la prestation"]            = ws.cell(row=r, column=C_NATURE).value
                rec["Rubrique budgétaire"]     = ws.cell(row=r, column=C_RUBR).value
                braw = ws.cell(row=r, column=C_BUDGET).value
                bval = num(braw)
                rec["Budget previsionnel TTC (MAD)"]      = bval if bval is not None else braw
                rec["Nature du budget"]                   = ws.cell(row=r, column=C_NATBUD).value
                rec["Mode de passation"]                  = MODE_FIXE.get(sh) or ws.cell(row=r, column=C_MODE).value
                rec["Trimestre de lancement"]             = ws.cell(row=r, column=C_TRIM).value
                rec["Duree d'execution"]                  = ws.cell(row=r, column=C_DUREE).value
                rec["Lieu d'execution"]                   = ws.cell(row=r, column=C_LIEU).value
                rec["PME (marche reserve)"]               = ws.cell(row=r, column=C_PME).value
                rec["Justification du besoin"]            = ws.cell(row=r, column=C_JUST).value
                for label, idx in zip(PROD, PROD_IDX):
                    rec[label] = ws.cell(row=r, column=idx).value if sh == "Contrats prod" else None
                rows.append(rec)
                if bval is not None:
                    total_dir += bval
                # anomalies
                missing = [lab for lab, idx in OBLIG.items()
                           if idx != C_BUDGET
                           and not (idx == C_MODE and sh in MODE_FIXE)
                           and (ws.cell(row=r, column=idx).value in (None, "")
                                or str(ws.cell(row=r, column=idx).value).strip() == "")]
                if bval is None:
                    missing.append("Budget non numerique" if braw not in (None, "") else "Budget previsionnel TTC (MAD)")
                if missing:
                    n_anom += 1
                    anomalies.append({"Direction": direction, "Onglet": sh, "Ligne source": r,
                                      "Objet": objet, "Champs en anomalie": ", ".join(missing),
                                      "Valeur budget brute": braw})
        suivi.append({"Direction": direction, "Fichier": fname,
                      "Lignes saisies": n_lignes, "Total budget (MAD)": total_dir,
                      "Lignes en anomalie": n_anom,
                      "Statut": "OK" if (n_lignes > 0 and n_anom == 0) else
                                ("INCOMPLET" if n_lignes > 0 else "VIDE")})
        wb.close()
        print(f"  {direction:<16} {n_lignes:4d} lignes  {total_dir:15,.0f} MAD  anomalies={n_anom}")

    recu = {s["Direction"] for s in suivi}
    for d in expected:
        if d not in recu:
            suivi.append({"Direction": d, "Fichier": "", "Lignes saisies": 0,
                          "Total budget (MAD)": 0, "Lignes en anomalie": 0, "Statut": "NON RECU"})

    out = Workbook(); out.remove(out.active)
    _sheet_consolide(out, rows)
    _sheet_synthese(out, rows)
    _sheet_anomalies(out, anomalies)
    _sheet_suivi(out, suivi, files, expected)
    out.save(OUT_PATH)
    print("\n->", OUT_PATH)
    print(f"   {len(rows)} lignes consolidees | {len(anomalies)} anomalies | {len(files)} fichiers lus")


def _write_table(ws, headers, records, start=1):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=start, column=j, value=h)
        c.font = FH; c.fill = FILLH; c.alignment = Alignment(wrap_text=True, vertical="center",
                                                             horizontal="center"); c.border = BORD
    for i, rec in enumerate(records, start=start + 1):
        for j, h in enumerate(headers, start=1):
            c = ws.cell(row=i, column=j, value=rec.get(h))
            c.border = BORD; c.alignment = WRAP
            if "Budget" in h or "Total" in h:
                if isinstance(c.value, (int, float)):
                    c.number_format = "#,##0"
    ws.freeze_panes = ws.cell(row=start + 1, column=1).coordinate
    if records:
        ws.auto_filter.ref = f"A{start}:{get_column_letter(len(headers))}{start + len(records)}"


def _sheet_consolide(out, rows):
    ws = out.create_sheet("PPA Consolide")
    _write_table(ws, OUT_COLS, rows)
    widths = {"Objet": 55, "Justification du besoin": 45, "Lieu d'execution": 25,
              "Rubrique budgétaire": 28, "Direction": 16, "Onglet": 20,
              "Type de support d'engagement": 24, "Mode de passation": 26}
    for j, h in enumerate(OUT_COLS, start=1):
        ws.column_dimensions[get_column_letter(j)].width = widths.get(h, 15)


def _agg(rows, key):
    d = {}
    for r in rows:
        k = r.get(key) or "(non renseigne)"
        b = r.get("Budget previsionnel TTC (MAD)")
        d.setdefault(k, [0, 0.0])
        d[k][0] += 1
        if isinstance(b, (int, float)):
            d[k][1] += b
    return d


def _sheet_synthese(out, rows):
    ws = out.create_sheet("Synthese")
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"SYNTHESE PPA 2026 - {len(rows)} lignes"
    ws["A1"].font = Font(bold=True, size=14, color=C_TITLE)
    ws["A2"] = "Genere le " + datetime.date.today().strftime("%d/%m/%Y")
    ws["A2"].font = Font(italic=True, size=9)
    row = 4
    blocks = [("Par Direction", "Direction"), ("Par Nature de la prestation", "Nature de la prestation"),
              ("Par Nature du budget", "Nature du budget"), ("Par Trimestre de lancement", "Trimestre de lancement"),
              ("Par Type de support", "Type de support d'engagement"), ("Par Mode de passation", "Mode de passation")]
    first_chart_ref = None
    for titre, key in blocks:
        ws.cell(row=row, column=1, value=titre).font = Font(bold=True, size=12, color=C_TITLE)
        row += 1
        for j, h in enumerate(["Valeur", "Nb lignes", "Budget total (MAD)"], start=1):
            c = ws.cell(row=row, column=j, value=h); c.font = FH; c.fill = FILLH; c.border = BORD
        hdr_row = row
        row += 1
        data = sorted(_agg(rows, key).items(), key=lambda kv: -kv[1][1])
        start_data = row
        for k, (n, b) in data:
            ws.cell(row=row, column=1, value=k).border = BORD
            ws.cell(row=row, column=2, value=n).border = BORD
            cc = ws.cell(row=row, column=3, value=round(b)); cc.border = BORD; cc.number_format = "#,##0"
            row += 1
        tot = ws.cell(row=row, column=1, value="TOTAL"); tot.font = Font(bold=True)
        ws.cell(row=row, column=2, value=sum(v[0] for v in (x[1] for x in data))).font = Font(bold=True)
        tc = ws.cell(row=row, column=3, value=round(sum(v[1] for v in (x[1] for x in data))))
        tc.font = Font(bold=True); tc.number_format = "#,##0"
        if key in ("Direction", "Nature de la prestation") and row > start_data:
            ch = BarChart(); ch.type = "bar"; ch.title = titre; ch.height = 7; ch.width = 16
            ch.add_data(Reference(ws, min_col=3, min_row=hdr_row, max_row=row - 1), titles_from_data=True)
            ch.set_categories(Reference(ws, min_col=1, min_row=start_data, max_row=row - 1))
            ws.add_chart(ch, f"F{hdr_row}")
        row += 2
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 20


def _sheet_anomalies(out, anomalies):
    ws = out.create_sheet("Anomalies")
    heads = ["Direction", "Onglet", "Ligne source", "Objet", "Champs en anomalie", "Valeur budget brute"]
    _write_table(ws, heads, anomalies)
    for j, h in enumerate(heads, start=1):
        ws.column_dimensions[get_column_letter(j)].width = {"Objet": 50, "Champs en anomalie": 40}.get(h, 16)
    if not anomalies:
        ws["A3"] = "Aucune anomalie."; ws["A3"].font = Font(bold=True, color="1E7B34")


def _sheet_suivi(out, suivi, files, expected):
    ws = out.create_sheet("Suivi des retours", 0)
    ws.sheet_view.showGridLines = False
    ws["A1"] = "SUIVI DES RETOURS PPA 2026"
    ws["A1"].font = Font(bold=True, size=14, color=C_TITLE)
    ws["A2"] = (f"{len(files)} / {len(expected) or len(suivi)} retour(s) recu(s) - "
                f"genere le {datetime.date.today().strftime('%d/%m/%Y')}")
    ws["A2"].font = Font(italic=True, size=9)
    heads = ["Direction", "Fichier", "Lignes saisies", "Total budget (MAD)", "Lignes en anomalie", "Statut"]
    _write_table(ws, heads, sorted(suivi, key=lambda s: s["Direction"]), start=4)
    for j, h in enumerate(heads, start=1):
        ws.column_dimensions[get_column_letter(j)].width = {"Fichier": 38, "Direction": 16,
                                                            "Total budget (MAD)": 20}.get(h, 15)
    last = 4 + len(suivi)
    ws.cell(row=last + 1, column=1, value="TOTAL").font = Font(bold=True)
    tc = ws.cell(row=last + 1, column=4, value=sum(s["Total budget (MAD)"] for s in suivi))
    tc.font = Font(bold=True); tc.number_format = "#,##0"
    ws.cell(row=last + 1, column=3, value=sum(s["Lignes saisies"] for s in suivi)).font = Font(bold=True)


if __name__ == "__main__":
    main()
