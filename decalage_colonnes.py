"""
Décalage des références de colonnes dans les formules Excel, pour insérer
une colonne en tête d'un onglet sans casser les formules qui le lisent.

Exemple : insertion d'une colonne A dans 'ENGAGEMENTS 2026'
    ='ENGAGEMENTS 2026'!B2          -> ='ENGAGEMENTS 2026'!C2
    ='ENGAGEMENTS 2026'!$L:$L       -> ='ENGAGEMENTS 2026'!$M:$M
    =+Z3 (dans l'onglet lui-même)   -> =+AA3
Les références vers d'autres onglets, les noms définis et les références
structurées (Tableau[[#This Row],[Colonne]]) ne changent pas.
"""
import re

from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils import column_index_from_string, get_column_letter

_PARTIE_CELLULE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d*)$")
_PARTIE_LIGNE = re.compile(r"^\$?\d+$")


def _nom_onglet(prefixe):
    p = prefixe
    if p.startswith("'") and p.endswith("'"):
        p = p[1:-1].replace("''", "'")
    return p


def _decaler_plage(plage, n, a_partir_de):
    """Décale de n colonnes les colonnes >= a_partir_de d'une plage sans onglet
    (A1, $B$2:$C9, $L:$L). Renvoie None si ce n'est pas une référence."""
    parties = plage.split(":")
    sortie = []
    for p in parties:
        if _PARTIE_LIGNE.match(p):
            sortie.append(p)
            continue
        m = _PARTIE_CELLULE.match(p)
        if not m:
            return None
        d1, col, d2, ligne = m.groups()
        # une référence colonne seule (sans ligne) n'est valable qu'en plage A:B
        if not ligne and len(parties) == 1:
            return None
        idx = column_index_from_string(col.upper())
        if idx >= a_partir_de:
            idx += n
        sortie.append(f"{d1}{get_column_letter(idx)}{d2}{ligne}")
    return ":".join(sortie)


def decaler_formule(formule, onglet_cible, onglet_formule, n=1, a_partir_de=1):
    """Décale les références vers onglet_cible dans une formule située dans
    onglet_formule. Renvoie la formule inchangée si rien n'est concerné."""
    if not isinstance(formule, str) or not formule.startswith("="):
        return formule
    tok = Tokenizer(formule)
    modifie = False
    for t in tok.items:
        if t.type != Token.OPERAND or t.subtype != Token.RANGE or "[" in t.value:
            continue
        valeur = t.value
        if "!" in valeur:
            prefixe, plage = valeur.rsplit("!", 1)
            if _nom_onglet(prefixe) != onglet_cible:
                continue
        else:
            if onglet_formule != onglet_cible:
                continue
            prefixe, plage = None, valeur
        nouvelle = _decaler_plage(plage, n, a_partir_de)
        if nouvelle is None or nouvelle == plage:
            continue
        t.value = f"{prefixe}!{nouvelle}" if prefixe is not None else nouvelle
        modifie = True
    if not modifie:
        return formule
    return "=" + "".join(t.value for t in tok.items)


def decaler_reference(ref, n=1, a_partir_de=1):
    """Décale une plage ou une liste de plages séparées par des espaces
    (sqref de validation), sans préfixe d'onglet."""
    return " ".join(_decaler_plage(p, n, a_partir_de) or p for p in str(ref).split())
