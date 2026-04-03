"""
calculs.py — Logique métier SOMME.SI pour le suivi d'allocation.
Reproduit fidèlement les formules Excel de l'onglet Alloc.
"""

import numpy as np
import pandas as pd

# ══════════════════════════════════════════════════════════════════════════════
#  STRUCTURE FIXE : ordre des lignes, labels, type et règles de calcul
#  type    : "blue"  → ligne groupe  (somme des enfants)
#            "white" → ligne détail  (formules SOMME.SI)
#            "total" → Total Général (somme des bleus)
#  detail  : "normal"            → D, G, J calculés
#            "nanties"           → D vide, J vide, G calculé
#            "nanties_no_formula"→ D vide, G=0, J vide
# ══════════════════════════════════════════════════════════════════════════════
ROW_DEFS = [
    (3,  "Obligations classiques",                  "blue",  None),
    (4,  "Obligations souveraines",                 "white", "normal"),
    (5,  "Obligations privées",                     "white", "normal"),
    (6,  "Obligations nanties",                     "blue",  None),
    (7,  "Obligations souveraines",                 "white", "nanties"),
    (8,  "Obligations privées",                     "white", "nanties_no_formula"),
    (9,  "Autres produits de taux",                 "blue",  None),
    (10, "Dettes privées",                          "white", "normal"),
    (11, "Alternatifs",                             "white", "normal"),
    (12, "Actions",                                 "blue",  None),
    (13, "Actions internationales",                 "white", "normal"),
    (14, "Actions Zone Euro",                       "white", "normal"),
    (15, "Autres actions (capital investissement)", "white", "normal"),
    (16, "Actifs réels",                            "blue",  None),
    (17, "Immobilier placement",                    "white", "normal"),
    (18, "Infrastructures",                         "white", "normal"),
    (19, "Stratégique",                             "blue",  None),
    (20, "Prêts stratégiques",                      "white", "normal"),
    (21, "Immobilier stratégique",                  "white", "normal"),
    (22, "Actions stratégiques",                    "white", "normal"),
    (23, "Trésorerie",                              "blue",  None),
    (24, "Trésorerie",                              "white", "normal"),
    (25, "Total Général",                           "total", None),
]

# Enfants de chaque ligne bleue
BLUE_CHILDREN = {
    3:  [4, 5],
    6:  [7, 8],
    9:  [10, 11],
    12: [13, 14, 15],
    16: [17, 18],
    19: [20, 21, 22],
    23: [24],
}
BLUE_ROWS_LIST = [3, 6, 9, 12, 16, 19, 23]
TOTAL_ROW = 25

# Onglets requis pour que l'appli fonctionne
REQUIRED_SHEETS = ["Portefeuille", "Retraitements", "Alloc"]
ALL_EXPECTED_SHEETS = [
    "Inputs>>", "TPT", "ListeActifs", "Portefeuille",
    "Calculs>>", "Retraitements", "Alloc", "A exlure>>",
    "Liste actif v.31.07.23", "Feuil1", "Definitions",
]


# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def somme_si(df: pd.DataFrame, criteria_col: str, criteria_val, sum_col: str) -> float:
    """Équivalent pandas de SOMME.SI — case & whitespace insensitive."""
    if criteria_val is None or (isinstance(criteria_val, float) and np.isnan(criteria_val)):
        return 0.0
    mask = (
        df[criteria_col]
        .astype(str)
        .str.strip()
        .str.upper()
        == str(criteria_val).strip().upper()
    )
    return float(df.loc[mask, sum_col].sum())


def validate_sheets(wb) -> tuple[bool, list[str], list[str]]:
    """
    Vérifie la présence des onglets.
    Retourne (ok_critical, missing_critical, missing_optional).
    """
    present = wb.sheetnames
    missing_critical = [s for s in REQUIRED_SHEETS if s not in present]
    missing_optional = [s for s in ALL_EXPECTED_SHEETS if s not in present and s not in REQUIRED_SHEETS]
    return len(missing_critical) == 0, missing_critical, missing_optional


# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT DES DATAFRAMES
# ══════════════════════════════════════════════════════════════════════════════
def load_portefeuille(wb) -> pd.DataFrame:
    ws = wb["Portefeuille"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) <= 36:
            continue
        rows.append({
            "F":  row[5],    # Catégorie d'instrument
            "W":  row[22],   # Valeur actuelle comptable (VNC)
            "Y":  row[24],   # Valeur de marché hors cc (Allocation)
            "AK": row[36],   # Classification
        })
    df = pd.DataFrame(rows)
    df["W"] = pd.to_numeric(df["W"], errors="coerce").fillna(0)
    df["Y"] = pd.to_numeric(df["Y"], errors="coerce").fillna(0)
    return df


def load_retraitements(wb) -> pd.DataFrame:
    ws = wb["Retraitements"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 3 or row[1] is None:
            continue
        rows.append({"B": row[1], "C": row[2] or 0})
    if not rows:
        return pd.DataFrame(columns=["B", "C"])
    df = pd.DataFrame(rows)
    df["C"] = pd.to_numeric(df["C"], errors="coerce").fillna(0)
    return df


def load_alloc_criteria(wb) -> dict:
    """
    Lit depuis l'onglet Alloc :
      - col B (index 2) : critère AK / Classification
      - col C (index 3) : critère F  / Catégorie instrument
      - col E (index 5) : Allocation cible (%)
      - col F (index 6) : Marge de manoeuvre (texte)
    pour chaque ligne blanche/bleue/total.
    """
    ws = wb["Alloc"]
    data = {}
    for row_num, _, _, _ in ROW_DEFS:
        data[row_num] = {
            "B": ws.cell(row=row_num, column=2).value,
            "C": ws.cell(row=row_num, column=3).value,
            "E": ws.cell(row=row_num, column=5).value,   # Allocation cible
            "F": ws.cell(row=row_num, column=6).value,   # Marge de manoeuvre
        }
    return data


# ══════════════════════════════════════════════════════════════════════════════
#  CALCUL PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def compute_allocation(wb) -> dict:
    df_port = load_portefeuille(wb)
    df_ret  = load_retraitements(wb)
    criteria = load_alloc_criteria(wb)

    results = {}

    # ── Lignes blanches ──────────────────────────────────────────────────────
    for row_num, label, row_type, detail in ROW_DEFS:
        if row_type != "white":
            continue

        b_val = criteria[row_num]["B"]
        c_val = criteria[row_num]["C"]
        e_val = criteria[row_num]["E"]   # Alloc cible
        f_val = criteria[row_num]["F"]   # Marge de manoeuvre

        # Col D — Retraitements (uniquement lignes "normal")
        if detail == "normal" and not df_ret.empty:
            d = somme_si(df_ret, "B", b_val, "C") / 1e6
        else:
            d = 0.0

        # Col G — Allocation M€
        if detail == "nanties_no_formula":
            g = 0.0
        else:
            g = (somme_si(df_port, "F", c_val, "Y") + somme_si(df_port, "AK", b_val, "Y")) / 1e6 + d

        # Col J — VNC M€ (vide pour obligations nanties)
        if detail in ("nanties", "nanties_no_formula"):
            j = None
        else:
            j = (somme_si(df_port, "F", c_val, "W") + somme_si(df_port, "AK", b_val, "W")) / 1e6

        results[row_num] = {
            "label": label, "type": row_type, "detail": detail,
            "D": d, "E": e_val, "F": f_val, "G": g, "J": j,
        }

    # ── Lignes bleues (somme des enfants) ────────────────────────────────────
    for blue_row, children in BLUE_CHILDREN.items():
        g_sum = sum(results[c]["G"] for c in children)
        j_vals = [results[c]["J"] for c in children if results[c]["J"] is not None]
        j_sum  = sum(j_vals) if j_vals else 0.0
        lbl    = next(l for r, l, t, _ in ROW_DEFS if r == blue_row)
        e_val  = criteria[blue_row]["E"]
        f_val  = criteria[blue_row]["F"]
        results[blue_row] = {
            "label": lbl, "type": "blue", "detail": None,
            "D": 0.0, "E": e_val, "F": f_val, "G": g_sum, "J": j_sum,
        }

    # ── Total Général ─────────────────────────────────────────────────────────
    g_total = sum(results[r]["G"] for r in BLUE_ROWS_LIST)
    j_total = sum(results[r]["J"] for r in BLUE_ROWS_LIST if isinstance(results[r]["J"], (int, float)))
    e_total = criteria[TOTAL_ROW]["E"]
    f_total = criteria[TOTAL_ROW]["F"]
    results[TOTAL_ROW] = {
        "label": "Total Général", "type": "total", "detail": None,
        "D": 0.0, "E": e_total, "F": f_total, "G": g_total, "J": j_total,
    }

    # ── Pourcentages H et K ───────────────────────────────────────────────────
    for r in results.values():
        r["H"] = r["G"] / g_total if g_total else 0.0
        r["K"] = (r["J"] / j_total) if (j_total and r["J"] is not None) else None

    return results
