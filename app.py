import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="Suivi d'Allocation", layout="wide")

# ══════════════════════════════════════════════════════════════════
#  STRUCTURE FIXE — ROW_DEFS
#  (label, type, detail, B=critère AK, C=critère F, alloc_cible_fixe, C_knl)
#
#  detail :
#    normal            → calcul depuis Portefeuille (AK + F)
#    nantissement_souv → calcul depuis Nantissement (F + AK) + retraitement
#    nantissement_priv → calcul depuis Nantissement (F + OPCVM MONETAIRES)
#
#  alloc_cible_fixe :
#    float  → hardcodé (ex: 0.08 pour 8%)
#    None   → calculé dynamiquement (lignes 0-5)
# ══════════════════════════════════════════════════════════════════
ROW_DEFS = [
    # idx  label                                    type              detail                  B (AK)                                     C (F)                              alloc_fixe  C_knl
    # 0
    ("Obligations classiques",                  "blue",           None,                   None,                                      None,                              None,       None),
    # 1
    ("Obligations souveraines",                 "white",          "normal",               "Obligations souveraines",                 "EMPRUNTS ETATS & OBLIG GARANTIES", None,       None),
    # 2
    ("Obligations privées",                     "white",          "normal",               "Obligations privées",                     "OBLIGATIONS COTEES",               None,       None),
    # 3
    ("Obligations nanties",                     "nanties_parent", None,                   None,                                      None,                              None,       None),
    # 4
    ("Obligations souveraines",                 "white",          "nantissement_souv",    "Obligations souveraines",                 "EMPRUNTS ETATS & OBLIG GARANTIES", None,       None),
    # 5
    ("Obligations privées",                     "white",          "nantissement_priv",    "Obligations privées",                     "OBLIGATIONS COTEES",               None,       None),
    # 6
    ("Autres produits de taux",                 "blue",           None,                   None,                                      None,                              0.13,       None),
    # 7
    ("Dettes privées",                          "white",          "normal",               "Dettes privées",                          None,                              0.08,       "DETTE PRIVEE"),
    # 8
    ("Alternatifs",                             "white",          "normal",               "Alternatifs",                             None,                              0.05,       "ALTERNATIF"),
    # 9
    ("Actions",                                 "blue",           None,                   None,                                      None,                              0.11,       None),
    # 10
    ("Actions internationales",                 "white",          "normal",               "Actions internationales",                 None,                              0.00,       None),
    # 11
    ("Actions Zone Euro",                       "white",          "normal",               "Actions Zone Euro",                       None,                              0.06,       "ACTIONS"),
    # 12
    ("Autres actions (capital investissement)", "white",          "normal",               "Autres actions (capital investissement)", None,                              0.05,       "Private Equity"),
    # 13
    ("Actifs réels",                            "blue",           None,                   None,                                      None,                              0.19,       None),
    # 14
    ("Immobilier placement",                    "white",          "normal",               "Immobilier placement",                    None,                              0.13,       "IMMOBILIER"),
    # 15
    ("Infrastructures",                         "white",          "normal",               "Infrastructures",                         None,                              0.06,       "INFRASTRUCTURE"),
    # 16
    ("Stratégique",                             "blue",           None,                   None,                                      None,                              0.12,       None),
    # 17
    ("Prêts stratégiques",                      "white",          "normal",               "Prêts stratégiques",                      None,                              0.02,       None),
    # 18
    ("Immobilier stratégique",                  "white",          "normal",               "Immobilier stratégique",                  None,                              0.01,       None),
    # 19
    ("Actions stratégiques",                    "white",          "normal",               "Actions stratégiques",                    None,                              0.09,       None),
    # 20
    ("Trésorerie",                              "blue",           None,                   None,                                      None,                              0.01,       None),
    # 21
    ("Trésorerie",                              "white",          "normal",               "Trésorerie",                              None,                              0.01,       None),
    # 22
    ("Total Général",                           "total",          None,                   None,                                      None,                              1.00,       None),
]

BLUE_CHILDREN = {
    0:  [1, 2],
    3:  [4, 5],
    6:  [7, 8],
    9:  [10, 11, 12],
    13: [14, 15],
    16: [17, 18, 19],
    20: [21],
}
BLUE_IDX  = [0, 3, 6, 9, 13, 16, 20]
TOTAL_IDX = 22

# Marge de manoeuvre hardcodée (None = vide)
MARGE = {
    0: "-20% / +5%", 1: "-20% / +5%", 2: "-20% / +5%",
    6: "-13% / +3%", 7: "-8% / +3%",  8: "-5% / +3%",
    9: "-11% / +3%", 10: "-0% / +3%", 11: "-6% / +3%", 12: "-5% / +3%",
    13: "-19% / +3%", 14: "-13% / +3%", 15: "-6% / +3%",
}

REQUIRED_SHEETS = ["Portefeuille", "Retraitements", "Nantissement"]

# ══════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════
def normalize(s):
    import unicodedata
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.split()).upper()


def somme_si(df, crit_col, crit_val, sum_col):
    if crit_val is None or df.empty:
        return 0.0
    if isinstance(crit_val, float) and np.isnan(crit_val):
        return 0.0
    target = normalize(crit_val)
    mask = df[crit_col].astype(str).map(normalize) == target
    return float(df.loc[mask, sum_col].sum())


def fmt_m(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    return f"{val:,.0f}"


def fmt_pct(val, decimals=0):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    return f"{val * 100:.{decimals}f}%"


# ══════════════════════════════════════════════════════════════════
#  CALCUL
# ══════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def compute(file_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)

    def read_portefeuille(ws):
        """Portefeuille : F=5, W=22, Y=24, AK=36."""
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) <= 36:
                continue
            rows.append({"F": row[5], "W": row[22], "Y": row[24], "AK": row[36]})
        df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["F","W","Y","AK"])
        df["W"] = pd.to_numeric(df["W"], errors="coerce").fillna(0)
        df["Y"] = pd.to_numeric(df["Y"], errors="coerce").fillna(0)
        return df

    def read_nantissement(ws):
        """
        Nantissement — indices confirmés depuis le CSV debug :
          F  = index 5  (Catégorie d Instrument)
          W  = index 22 (Valeur actuelle comptable hors cc)
          Y  = index 24 (Valeur de marché hors cc)
          AK = absent   → None
        Nantissement a 36 colonnes (indices 0-35), on skip si < 25.
        """
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 25:
                continue
            rows.append({
                "F":  row[5],
                "W":  row[22],
                "Y":  row[24],
                "AK": None,
            })
        df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["F","W","Y","AK"])
        df["W"] = pd.to_numeric(df["W"], errors="coerce").fillna(0)
        df["Y"] = pd.to_numeric(df["Y"], errors="coerce").fillna(0)
        return df

    # ── Portefeuille & Nantissement
    df_p    = read_portefeuille(wb["Portefeuille"])
    df_nant = read_nantissement(wb["Nantissement"]) if "Nantissement" in wb.sheetnames else pd.DataFrame(columns=["F","W","Y","AK"])

    # ── Retraitements : C=2 classe, D=3 montant
    ws_r, rows_r = wb["Retraitements"], []
    for row in ws_r.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 4:
            continue
        classe, montant = row[2], row[3]
        if classe is None or str(classe).strip().lower() in ("classe d'actifs", "montant", ""):
            continue
        try:
            rows_r.append({"classe": str(classe).strip(), "D": float(montant)})
        except (TypeError, ValueError):
            continue
    df_r = pd.DataFrame(rows_r) if rows_r else pd.DataFrame(columns=["classe", "D"])

    # ── KNL : B=1, K=10, M=12, N=13 + date E1
    has_knl = "KNL" in wb.sheetnames
    df_knl  = pd.DataFrame(columns=["B","K","M","N"])
    knl_year = None
    if has_knl:
        ws_knl = wb["KNL"]
        # Lecture date KNL!E1 (col index 4)
        knl_e1 = ws_knl.cell(row=1, column=5).value
        if hasattr(knl_e1, "year"):
            knl_year = knl_e1.year
        rows_knl = []
        for row in ws_knl.iter_rows(min_row=3, values_only=True):
            if not row or len(row) <= 13 or row[1] is None:
                continue
            rows_knl.append({"B": row[1], "K": row[10], "M": row[12], "N": row[13]})
        if rows_knl:
            df_knl = pd.DataFrame(rows_knl)
            for col in ["K","M","N"]:
                df_knl[col] = pd.to_numeric(df_knl[col], errors="coerce").fillna(0)

    # ── Calcul L8 (engagements nanties privées)
    # L8 = -35 * (ANNEE(KNL!E1) - ANNEE(AUJOURDHUI()))
    l8 = 0.0
    if knl_year is not None:
        l8 = -35.0 * (knl_year - date.today().year)

    # ══════════════════════════════════════════════
    #  ETAPE 1 : calcul nanties EN PREMIER (nécessaires pour classiques)
    # ══════════════════════════════════════════════
    res = {}

    # idx 4 — Oblig souveraines nanties
    # G7 = SOMME.SI(Nant!F, C7, Y) + SOMME.SI(Nant!AK, B7, Y)
    C4 = ROW_DEFS[4][4]  # = "EMPRUNTS ETATS & OBLIG GARANTIES"
    B4 = ROW_DEFS[4][3]  # = "Obligations souveraines"
    g4 = (somme_si(df_nant, "F", C4, "Y") + somme_si(df_nant, "AK", B4, "Y")) / 1e6
    j4 = (somme_si(df_nant, "F", C4, "W") + somme_si(df_nant, "AK", B4, "W")) / 1e6
    res[4] = {"label": ROW_DEFS[4][0], "type": "white", "detail": "nantissement_souv",
              "D": 0.0, "G": g4, "J": j4, "M": 0.0, "N": 0.0, "O": 0.0}

    # idx 5 — Oblig privées nanties
    # G8 = SOMME.SI(Nant!F, C8, Y) + SOMME.SI(Nant!F, "OPCVM MONETAIRES", Y)
    C5 = ROW_DEFS[5][4]  # = "OBLIGATIONS COTEES"
    B5 = ROW_DEFS[5][3]  # = "Obligations privées"
    g5 = (somme_si(df_nant, "F", C5, "Y") + somme_si(df_nant, "F", "OPCVM MONETAIRES", "Y")) / 1e6
    j5 = (somme_si(df_nant, "F", C5, "W") + somme_si(df_nant, "AK", B5, "W")) / 1e6
    res[5] = {"label": ROW_DEFS[5][0], "type": "white", "detail": "nantissement_priv",
              "D": 0.0, "G": g5, "J": j5, "M": l8, "N": 0.0, "O": 0.0}

    # ══════════════════════════════════════════════
    #  ETAPE 2 : toutes les autres lignes blanches
    # ══════════════════════════════════════════════
    for idx, (label, rtype, detail, B, C, _, C_knl) in enumerate(ROW_DEFS):
        if rtype != "white" or idx in (4, 5):
            continue

        d = 0.0; g = 0.0; j = None; m_val = 0.0; n_val = 0.0; o_val = 0.0

        if detail == "normal":
            d     = somme_si(df_r, "classe", B, "D") / 1e6
            g_raw = (somme_si(df_p, "AK", B, "Y") + (somme_si(df_p, "F", C, "Y") if C else 0.0)) / 1e6
            j_raw = (somme_si(df_p, "AK", B, "W") + (somme_si(df_p, "F", C, "W") if C else 0.0)) / 1e6

            # Obligations souveraines classiques (idx=1) : soustraire nanties souv
            if idx == 1:
                g = g_raw + d - res[4]["G"]
                j = j_raw - res[4]["J"]
            # Obligations privées classiques (idx=2) : soustraire nanties priv
            elif idx == 2:
                g = g_raw + d - res[5]["G"]
                j = j_raw - res[5]["J"]
            else:
                g = g_raw + d
                j = j_raw

            m_val = somme_si(df_knl, "B", C_knl, "K") / 1e6 if C_knl else 0.0
            n_val = somme_si(df_knl, "B", C_knl, "M") / 1e6 if C_knl else 0.0
            o_val = somme_si(df_knl, "B", C_knl, "N") / 1e6 if C_knl else 0.0

        res[idx] = {"label": label, "type": rtype, "detail": detail,
                    "D": d, "G": g, "J": j, "M": m_val, "N": n_val, "O": o_val}

    # ══════════════════════════════════════════════
    #  ETAPE 3 : lignes bleues = somme enfants
    # ══════════════════════════════════════════════
    for br, children in BLUE_CHILDREN.items():
        rtype_br = ROW_DEFS[br][1]
        g_sum = sum(res[c]["G"] for c in children if c in res)
        j_vals = [res[c]["J"] for c in children if c in res and res[c]["J"] is not None]
        j_sum  = sum(j_vals) if j_vals else None
        m_sum  = sum(res[c]["M"] for c in children if c in res)
        n_sum  = sum(res[c]["N"] for c in children if c in res)
        o_sum  = sum(res[c]["O"] for c in children if c in res)
        res[br] = {"label": ROW_DEFS[br][0], "type": rtype_br, "detail": None,
                   "D": 0.0, "G": g_sum, "J": j_sum,
                   "M": m_sum, "N": n_sum, "O": o_sum}

    # ══════════════════════════════════════════════
    #  ETAPE 4 : cas spéciaux N + Total
    # ══════════════════════════════════════════════

    # N[Trésorerie] EN PREMIER : N24 = G24 - G_total * alloc_cible_tres
    g_total_tmp = sum(res[i]["G"] for i in BLUE_IDX)
    res[21]["N"] = res[21]["G"] - g_total_tmp * 0.01
    res[20]["N"] = res[21]["N"]

    # N[Oblig privées classiques] : N5 = -N9-N12-N16-N19-N23 + M25 - M6(nanties)
    m_total  = sum(res[i]["M"] for i in BLUE_IDX)  # = M25
    n_autres = sum(res[i]["N"] for i in [6, 9, 13, 16, 20])
    m_nanties = res[3]["M"]  # = M6
    res[2]["N"] = -n_autres + m_total - m_nanties
    res[0]["N"] = res[2]["N"]  # N3 = N5

    # Total général
    g_tot = sum(res[i]["G"] for i in BLUE_IDX)
    j_vals_tot = [res[i]["J"] for i in BLUE_IDX if res[i].get("J") is not None]
    j_tot = sum(j_vals_tot) if j_vals_tot else 0.0
    m_tot = sum(res[i]["M"] for i in BLUE_IDX)
    n_tot = sum(res[i]["N"] for i in BLUE_IDX)
    o_tot = sum(res[i]["O"] for i in BLUE_IDX)
    res[TOTAL_IDX] = {"label": "Total Général", "type": "total", "detail": None,
                      "D": 0.0, "G": g_tot, "J": j_tot,
                      "M": m_tot, "N": n_tot, "O": o_tot}

    # ══════════════════════════════════════════════
    #  ETAPE 5 : alloc cible dynamique
    # ══════════════════════════════════════════════
    # E6=G6/G25, E7=G7/G25, E8=G8/G25
    # E3=44%-E6, E4=20%-E7, E5=24%-E8
    e6 = res[3]["G"] / g_tot if g_tot else 0.0
    e7 = res[4]["G"] / g_tot if g_tot else 0.0
    e8 = res[5]["G"] / g_tot if g_tot else 0.0
    res[3]["alloc_cible_f"] = e6
    res[4]["alloc_cible_f"] = e7
    res[5]["alloc_cible_f"] = e8
    res[0]["alloc_cible_f"] = 0.44 - e6
    res[1]["alloc_cible_f"] = 0.20 - e7
    res[2]["alloc_cible_f"] = 0.24 - e8
    for idx, (_, _, _, _, _, alloc_fixe, _) in enumerate(ROW_DEFS):
        if idx in res and "alloc_cible_f" not in res[idx]:
            res[idx]["alloc_cible_f"] = alloc_fixe

    # ══════════════════════════════════════════════
    #  ETAPE 6 : pourcentages et dérivés
    # ══════════════════════════════════════════════

    # H = G% , K = J%
    for r in res.values():
        r["H"] = r["G"] / g_tot if g_tot else 0.0
        r["K"] = (r["J"] / j_tot) if (j_tot and r.get("J") is not None) else None

    # +/- values = G - J, arrondi pour eviter erreurs floating point
    for r in res.values():
        j_val = r.get("J")
        if j_val is not None:
            raw = r["G"] - j_val
            r["PLUSMOINS"] = round(raw, 6) if abs(raw) > 1e-4 else 0.0
        else:
            r["PLUSMOINS"] = None

    # Allocation projetée P = G + M - N + O
    for idx, r in res.items():
        r["P"] = r["G"] + r["M"] - r["N"] + r["O"]
    # Nanties parent = somme enfants
    res[3]["P"] = sum(res[c]["P"] for c in BLUE_CHILDREN[3] if c in res)

    # Q = P%
    # P3 = O3/O25 - P6, P4 = O4/O25 - P7, P5 = O5/O25 - P8
    p_tot = res[TOTAL_IDX]["P"]
    for idx, r in res.items():
        if p_tot:
            q_raw = r["P"] / p_tot
            if idx == 0:  r["Q"] = q_raw - res[3]["P"] / p_tot
            elif idx == 1: r["Q"] = q_raw - res[4]["P"] / p_tot
            elif idx == 2: r["Q"] = q_raw - res[5]["P"] / p_tot
            else:          r["Q"] = q_raw
        else:
            r["Q"] = None

    return res


# ══════════════════════════════════════════════════════════════════
#  RENDU HTML
# ══════════════════════════════════════════════════════════════════
STYLE = {
    "blue":          {"bg": "#2E75B6", "fg": "#FFFFFF", "fw": "700"},
    "nanties_parent":{"bg": "#8FAADC", "fg": "#FFFFFF", "fw": "700"},
    "white":         {"bg": "#FFFFFF", "fg": "#1a1a1a", "fw": "400"},
    "total":         {"bg": "#1F3864", "fg": "#FFFFFF", "fw": "700"},
}


def render_table(res: dict, show_retraitement: bool = True) -> str:
    headers = [
        ("Catégorie", "l", True),
        ("Cible", "c", True),
        ("Marge", "c", True),
        ("Retrait. M€", "r", show_retraitement),
        ("Alloc. M€", "r", True),
        ("Alloc. %", "r", True),
        ("VNC M€", "r", True),
        ("VNC %", "r", True),
        ("+/- val. M€", "r", True),
        ("Engag. M€", "r", True),
        ("Financement M€", "r", True),
        ("Gains cap. M€", "r", True),
        ("Proj. M€", "r", True),
        ("Proj. %", "r", True),
    ]
    th = "".join(f'<th class="{a}">{h}</th>' for h, a, show in headers if show)
    html = f"""
    <style>
      .at{{border-collapse:collapse;width:100%;font-family:Calibri,sans-serif;font-size:11px}}
      .at th{{padding:5px 6px;text-align:center;border:1px solid #888;
             background:#1F3864;color:#fff;font-weight:700;white-space:nowrap}}
      .at td{{padding:3px 6px;border:1px solid #ccc;white-space:nowrap}}
      .r{{text-align:right}}.l{{text-align:left}}.c{{text-align:center}}
    </style>
    <table class="at"><thead><tr>{th}</tr></thead><tbody>
    """
    for idx, (label, rtype, detail, _, _, _, _) in enumerate(ROW_DEFS):
        r = res.get(idx)
        if not r:
            continue
        s  = STYLE[rtype]
        cs = f'background:{s["bg"]};color:{s["fg"]};font-weight:{s["fw"]};' 

        ac_f  = r.get("alloc_cible_f")
        cells = [
            (label,                                                              "l", True),
            (fmt_pct(ac_f) if ac_f is not None else "",                         "c", True),
            (MARGE.get(idx, ""),                                                 "c", True),
            (fmt_m(r["D"]) if (rtype=="white" and detail in ("normal","nantissement_souv")) else "", "r", show_retraitement),
            (fmt_m(r["G"]),                                                      "r", True),
            (fmt_pct(r["H"]),                                                    "r", True),
            (fmt_m(r["J"]) if r.get("J") is not None else "",                   "r", True),
            (fmt_pct(r["K"]) if r.get("K") is not None else "",                 "r", True),
            (fmt_m(r["PLUSMOINS"]) if r.get("PLUSMOINS") is not None else "",   "r", True),
            (fmt_m(r["M"]) if r.get("M") else "",                               "r", True),
            (fmt_m(r["N"]) if r.get("N") else "",                               "r", True),
            (fmt_m(r["O"]) if r.get("O") else "",                               "r", True),
            (fmt_m(r["P"]) if r.get("P") is not None else "",                   "r", True),
            (fmt_pct(r["Q"]) if r.get("Q") is not None else "",                 "r", True),
        ]
        tds = "".join(f'<td class="{a}" style="{cs}">{v}</td>' for v, a, show in cells if show)
        html += f'<tr>{tds}</tr>\n'
    html += "</tbody></table>"
    return html


# ══════════════════════════════════════════════════════════════════
#  EXPORT EXCEL
# ══════════════════════════════════════════════════════════════════
def export_excel(res: dict) -> BytesIO:
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "Alloc"

    thin   = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    FILLS = {
        "blue":          PatternFill(fgColor="2E75B6", fill_type="solid"),
        "nanties_parent":PatternFill(fgColor="8FAADC", fill_type="solid"),
        "white":         PatternFill(fgColor="FFFFFF", fill_type="solid"),
        "total":         PatternFill(fgColor="1F3864", fill_type="solid"),
        "header":        PatternFill(fgColor="1F3864", fill_type="solid"),
    }
    FONTS = {
        "blue":          Font(bold=True, color="FFFFFF", name="Calibri", size=11),
        "nanties_parent":Font(bold=True, color="FFFFFF", name="Calibri", size=11),
        "white":         Font(color="000000",            name="Calibri", size=11),
        "total":         Font(bold=True, color="FFFFFF", name="Calibri", size=11),
        "header":        Font(bold=True, color="FFFFFF", name="Calibri", size=11),
    }

    headers = [
        "Catégorie d'investissement", "Alloc. cible", "Marge",
        "Retraitement (M€)", "Allocation (M€)", "Allocation (%)",
        "VNC (M€)", "VNC (%)", "+/- values (M€)",
        "Engagements (M€)", "Financement appels (M€)", "Gains capital (M€)",
        "Alloc. projetée (M€)", "Alloc. projetée (%)",
    ]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = FILLS["header"]; c.font = FONTS["header"]; c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center")

    for erow, (idx, (label, rtype, detail, _, _, _, _)) in enumerate(enumerate(ROW_DEFS), 2):
        r = res.get(idx)
        if not r:
            continue
        ac_f  = r.get("alloc_cible_f")
        d_val = r["D"] if (rtype == "white" and detail in ("normal", "nantissement_souv")) else None
        vals  = [
            label,
            ac_f,
            MARGE.get(idx, ""),
            d_val,
            r["G"], r["H"],
            r["J"] if r.get("J") is not None else None,
            r["K"] if r.get("K") is not None else None,
            r["PLUSMOINS"] if r.get("PLUSMOINS") is not None else None,
            r["M"] if r.get("M") else None,
            r["N"] if r.get("N") else None,
            r["O"] if r.get("O") else None,
            r["P"] if r.get("P") is not None else None,
            r["Q"] if r.get("Q") is not None else None,
        ]
        fmts   = [None, "0%", "@", "#,##0", "#,##0", "0.0%",
                  "#,##0", "0.0%", "#,##0",
                  "#,##0", "#,##0", "#,##0", "#,##0", "0.0%"]
        aligns = ["left"] + ["right"] * 13

        for col, (val, fmt, align) in enumerate(zip(vals, fmts, aligns), 1):
            c = ws.cell(row=erow, column=col, value=val)
            c.fill = FILLS[rtype]; c.font = FONTS[rtype]; c.border = border
            c.alignment = Alignment(horizontal=align, vertical="center")
            if fmt and val is not None and val != "":
                c.number_format = fmt

    ws.column_dimensions["A"].width = 40
    for col_letter in ["B","C","D","E","F","G","H","I","J","K","L","M","N"]:
        ws.column_dimensions[col_letter].width = 15
    ws.row_dimensions[1].height = 24

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<h1 style='color:#1F3864;font-family:Calibri,sans-serif;'>Suivi d'Allocation</h1>
<p style='color:#666;font-size:14px;'>
  Dépose le fichier Excel → tous les calculs se font automatiquement.
</p><hr style='border:1px solid #ddd;'>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        file_bytes = uploaded.read()

        wb_chk = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True)
        sheet_names = wb_chk.sheetnames
        wb_chk.close()

        missing = [s for s in REQUIRED_SHEETS if s not in sheet_names]
        if missing:
            st.error(f"Onglets manquants : `{'`, `'.join(missing)}`")
            st.stop()

        if "KNL" not in sheet_names:
            st.warning("Onglet KNL absent — colonnes Engagements/Financement/Gains/Projection à 0.")

        with st.spinner("Calcul en cours..."):
            res = compute(file_bytes)

        # ── Controle classes Retraitements non matchées
        categories_connues = {
            normalize(B)
            for _, rtype, detail, B, _, _, _ in ROW_DEFS
            if rtype == "white" and detail == "normal" and B is not None
        }
        wb_chk2 = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
        ws_rchk = wb_chk2["Retraitements"]
        classes_fichier = set()
        for row in ws_rchk.iter_rows(min_row=2, values_only=True):
            if row and len(row) >= 3 and row[2] is not None:
                val = str(row[2]).strip()
                if val.lower() not in ("classe d'actifs", "montant", ""):
                    classes_fichier.add(normalize(val))
        non_matches = classes_fichier - categories_connues
        if non_matches:
            st.warning(f"Classes Retraitements non reconnues : **{', '.join(sorted(non_matches))}**")

        g_tot = res[TOTAL_IDX]["G"]
        j_tot = res[TOTAL_IDX]["J"]
        p_tot = res[TOTAL_IDX]["P"] or 0.0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Allocation",  f"{g_tot:,.0f} M€")
        c2.metric("Total VNC",         f"{j_tot:,.0f} M€")
        c3.metric("Ecart Alloc - VNC", f"{g_tot - j_tot:,.0f} M€")
        c4.metric("Alloc. projetée",   f"{p_tot:,.0f} M€")

        show_ret = st.checkbox("Afficher colonne Retraitement", value=False)
        components.html(render_table(res, show_retraitement=show_ret), height=800, scrolling=True)
        st.markdown("<br>", unsafe_allow_html=True)

        st.download_button(
            label="Exporter en Excel",
            data=export_excel(res),
            file_name="suivi_allocation_calcule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except KeyError as e:
        st.error(f"Onglet introuvable : {e}")
    except Exception as e:
        st.error(f"Erreur : {e}")
        with st.expander("Details"):
            st.exception(e)
else:
    st.info("En attente du fichier Excel...")
    with st.expander("Onglets requis"):
        st.markdown("""
| Onglet | Colonnes utilisées |
|---|---|
| **Portefeuille** | F (col 6), W (col 23), Y (col 25), AK (col 37) |
| **Nantissement** | même structure que Portefeuille |
| **Retraitements** | A (ISIN), B (nom), C (classe), D (montant) |
| **KNL** *(optionnel)* | B (classe), K (engagements), M (retour capital), N (capital gain), E1 (date échéance) |
        """)