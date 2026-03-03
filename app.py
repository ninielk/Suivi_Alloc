import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="Suivi d'Allocation", layout="wide")

# ══════════════════════════════════════════════════════════════════
#  STRUCTURE FIXE
#  (label, type, detail, B=critère AK, C=critère F,
#   alloc_cible, marge, C_knl=critère KNL!B)
#
#  C_knl : critère pour les SOMME.SI.ENS sur KNL
#          None = ligne non concernée par KNL (M/N/O = 0)
# ══════════════════════════════════════════════════════════════════
ROW_DEFS = [
    # label                                    type    detail                B (AK)                                      C (F)                           alloc   marge           C_knl
    ("Obligations classiques",                 "blue",  None,                None,                                       None,                           "44%",  "-20% / +5%",   None),
    ("Obligations souveraines",                "white", "normal",            "Obligations souveraines",                  "EMPRUNTS ETATS & OBLIG GARANTIES","20%", "-20% / +5%", None),
    ("Obligations privées",                    "white", "normal",            "Obligations privées",                      "OBLIGATIONS COTEES",            "24%",  "-20% / +5%",   None),
    ("Obligations nanties",                    "nanties_parent", None,       None,                                       None,                           "",     "",             None),
    ("Obligations souveraines",                "white", "nanties",           "Obligations souveraines",                  None,                           "",     "",             None),
    ("Obligations privées",                    "white", "nanties_no_formula","Obligations privées",                      None,                           "",     "",             None),
    ("Autres produits de taux",                "blue",  None,                None,                                       None,                           "13%",  "-13% / +3%",   None),
    ("Dettes privées",                         "white", "normal",            "Dettes privées",                           None,                           "8%",   "-8% / +3%",    "DETTE PRIVEE"),
    ("Alternatifs",                            "white", "normal",            "Alternatifs",                              None,                           "5%",   "-5% / +3%",    "ALTERNATIF"),
    ("Actions",                                "blue",  None,                None,                                       None,                           "11%",  "-11% / +3%",   None),
    ("Actions internationales",                "white", "normal",            "Actions internationales",                  None,                           "0%",   "-0% / +3%",    None),
    ("Actions Zone Euro",                      "white", "normal",            "Actions Zone Euro",                        None,                           "6%",   "-6% / +3%",    "ACTIONS"),
    ("Autres actions (capital investissement)","white", "normal",            "Autres actions (capital investissement)",  None,                           "5%",   "-5% / +3%",    "Private Equity"),
    ("Actifs réels",                           "blue",  None,                None,                                       None,                           "19%",  "-19% / +3%",   None),
    ("Immobilier placement",                   "white", "normal",            "Immobilier placement",                     None,                           "13%",  "-13% / +3%",   "IMMOBILIER"),
    ("Infrastructures",                        "white", "normal",            "Infrastructures",                          None,                           "6%",   "-6% / +3%",    "INFRASTRUCTURE"),
    ("Stratégique",                            "blue",  None,                None,                                       None,                           "12%",  "",             None),
    ("Prêts stratégiques",                     "white", "normal",            "Prêts stratégiques",                       None,                           "2%",   "",             None),
    ("Immobilier stratégique",                 "white", "normal",            "Immobilier stratégique",                   None,                           "1%",   "",             None),
    ("Actions stratégiques",                   "white", "normal",            "Actions stratégiques",                     None,                           "9%",   "",             None),
    ("Trésorerie",                             "blue",  None,                None,                                       None,                           "1%",   "",             None),
    ("Trésorerie",                             "white", "normal",            "Trésorerie",                               None,                           "1%",   "",             None),
    ("Total Général",                          "total", None,                None,                                       None,                           "100%", "",             None),
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

# Alloc cible en décimal pour calcul N[Trésorerie]
ALLOC_CIBLE_PCT = {
    1: 0.20, 2: 0.24, 7: 0.08, 8: 0.05,
    10: 0.00, 11: 0.06, 12: 0.05,
    14: 0.13, 15: 0.06,
    17: 0.02, 18: 0.01, 19: 0.09,
    21: 0.01,
}

REQUIRED_SHEETS = ["Portefeuille", "Retraitements"]

# ══════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════
def normalize(s):
    import unicodedata
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.split()).upper()


def somme_si(df, crit_col, crit_val, sum_col):
    if crit_val is None:
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


def fmt_pct(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    return f"{val * 100:.1f}%"


# ══════════════════════════════════════════════════════════════════
#  CALCUL
# ══════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def compute(file_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)

    # ── Portefeuille
    #    F=5  Catégorie instrument | W=22 VNC | Y=24 Valeur marché | AK=36 Classification
    ws_p = wb["Portefeuille"]
    rows_p = []
    for row in ws_p.iter_rows(min_row=2, values_only=True):
        if not row or len(row) <= 36:
            continue
        rows_p.append({"F": row[5], "W": row[22], "Y": row[24], "AK": row[36]})
    df_p = pd.DataFrame(rows_p)
    df_p["W"] = pd.to_numeric(df_p["W"], errors="coerce").fillna(0)
    df_p["Y"] = pd.to_numeric(df_p["Y"], errors="coerce").fillna(0)

    # ── Retraitements
    #    A=0 ISIN | B=1 Nom fonds | C=2 Classe d'actifs | D=3 Montant
    ws_r = wb["Retraitements"]
    rows_r = []
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

    # ── KNL (optionnel)
    #    B=1  Classe d'actif (critère)
    #    K=10 Engagements à financer  → col M Alloc
    #    M=12 Retour en capital       → col N Alloc
    #    N=13 Capital Gain            → col O Alloc
    #    données à partir de la ligne 3 (2 lignes de header)
    has_knl = "KNL" in wb.sheetnames
    df_knl = pd.DataFrame(columns=["B", "K", "M", "N"])
    if has_knl:
        ws_knl = wb["KNL"]
        rows_knl = []
        for row in ws_knl.iter_rows(min_row=3, values_only=True):
            if not row or len(row) <= 13:
                continue
            if row[1] is None:
                continue
            rows_knl.append({
                "B": row[1],
                "K": row[10],
                "M": row[12],
                "N": row[13],
            })
        df_knl = pd.DataFrame(rows_knl) if rows_knl else df_knl
        for col in ["K", "M", "N"]:
            df_knl[col] = pd.to_numeric(df_knl[col], errors="coerce").fillna(0)

    # ── Calcul lignes blanches
    res = {}
    for idx, (label, rtype, detail, B, C, _, _, C_knl) in enumerate(ROW_DEFS):
        if rtype != "white":
            continue

        # Col D – Retraitement
        d = somme_si(df_r, "classe", B, "D") / 1e6 if (detail == "normal" and not df_r.empty) else 0.0

        # Col G – Allocation M€
        if detail == "nanties_no_formula":
            g = 0.0
        else:
            g = (somme_si(df_p, "AK", B, "Y") + (somme_si(df_p, "F", C, "Y") if C else 0.0)) / 1e6 + d

        # Col J – VNC M€
        if detail in ("nanties", "nanties_no_formula"):
            j = None
        else:
            j = (somme_si(df_p, "AK", B, "W") + (somme_si(df_p, "F", C, "W") if C else 0.0)) / 1e6

        # Col M – Engagements à financer (KNL!K)
        m_val = somme_si(df_knl, "B", C_knl, "K") / 1e6 if (C_knl and not df_knl.empty) else 0.0

        # Col N – Financement appels (KNL!M)
        n_val = somme_si(df_knl, "B", C_knl, "M") / 1e6 if (C_knl and not df_knl.empty) else 0.0

        # Col O – Gains en capital (KNL!N)
        # Vide pour Oblig classiques/nanties/Stratégique/Trésorerie → déjà 0 si C_knl=None
        o_val = somme_si(df_knl, "B", C_knl, "N") / 1e6 if (C_knl and not df_knl.empty) else 0.0

        res[idx] = {
            "label": label, "type": rtype, "detail": detail,
            "D": d, "G": g, "J": j,
            "M": m_val, "N": n_val, "O": o_val,
        }

    # ── Lignes bleues = somme enfants
    for br, children in BLUE_CHILDREN.items():
        g_sum = sum(res[c]["G"] for c in children if c in res)
        j_sum = sum(res[c]["J"] for c in children if c in res and res[c]["J"] is not None)
        m_sum = sum(res[c]["M"] for c in children if c in res)
        n_sum = sum(res[c]["N"] for c in children if c in res)
        o_sum = sum(res[c]["O"] for c in children if c in res)
        rtype_br = ROW_DEFS[br][1]
        res[br] = {
            "label": ROW_DEFS[br][0], "type": rtype_br, "detail": None,
            "D": 0.0, "G": g_sum, "J": j_sum,
            "M": m_sum, "N": n_sum, "O": o_sum,
        }

    # ── Cas spécial N[Tresorerie] EN PREMIER (idx=21) car utilise dans N5
    #    N24 = G24 - G_total * alloc_cible_tresorerie
    g_total_tmp = sum(res[i]["G"] for i in BLUE_IDX)
    alloc_tres = ALLOC_CIBLE_PCT.get(21, 0.01)
    res[21]["N"] = res[21]["G"] - g_total_tmp * alloc_tres
    res[20]["N"] = res[21]["N"]  # parent Tresorerie = enfant unique

    # ── Cas spécial N[Obligations privees] (idx=2) :
    #    N5 = -N9-N12-N16-N19-N23 + M25
    m_total = sum(res[i]["M"] for i in BLUE_IDX)  # = M25
    n_autres = sum(res[i]["N"] for i in [6, 9, 13, 16, 20])  # exactement -N9-N12-N16-N19-N23
    res[2]["N"] = -n_autres + m_total
    res[0]["N"] = res[2]["N"]  # N3 = SOMME(N4:N5) = 0 + N5
    res[20]["N"] = res[21]["N"]

    # ── Cas spécial P[Obligations nanties] = 155 hardcodé (idx=3)
    # Géré dans P directement

    # ── Total
    g_tot = sum(res[i]["G"] for i in BLUE_IDX)
    j_tot = sum(res[i]["J"] for i in BLUE_IDX if isinstance(res[i].get("J"), (int, float)))
    m_tot = sum(res[i]["M"] for i in BLUE_IDX)
    n_tot = sum(res[i]["N"] for i in BLUE_IDX)
    o_tot = sum(res[i]["O"] for i in BLUE_IDX)
    res[TOTAL_IDX] = {
        "label": "Total Général", "type": "total", "detail": None,
        "D": 0.0, "G": g_tot, "J": j_tot,
        "M": m_tot, "N": n_tot, "O": o_tot,
    }

    # ── Cas special : Obligations nanties G hardcode a 288.12
    res[3]["G"] = 288.12

    # ── Pourcentages G et J
    for r in res.values():
        r["H"] = r["G"] / g_tot if g_tot else 0.0
        r["K"] = (r["J"] / j_tot) if (j_tot and r.get("J") is not None) else None

    # ── Col P – Allocation projetée M€
    #    P = G + M - N + O  (par ligne)
    #    Sauf Obligations nanties (idx=3) → 155 hardcodé
    for idx, r in res.items():
        if idx == 3:  # Obligations nanties hardcodé
            r["P"] = 155.0
        elif idx in (4, 5):  # enfants obligations nanties → vide
            r["P"] = None
        else:
            r["P"] = r["G"] + r["M"] - r["N"] + r["O"]

    # ── Col Q – Allocation projetée %
    p_tot = res[TOTAL_IDX]["P"] or 0.0
    for r in res.values():
        r["Q"] = (r["P"] / p_tot) if (p_tot and r.get("P") is not None) else None

    return res


# ══════════════════════════════════════════════════════════════════
#  RENDU HTML
# ══════════════════════════════════════════════════════════════════
STYLE = {
    "blue":         {"bg": "#2E75B6", "fg": "#FFFFFF", "fw": "700"},
    "nanties_parent":{"bg": "#8FAADC", "fg": "#FFFFFF", "fw": "700"},
    "white":        {"bg": "#FFFFFF", "fg": "#1a1a1a", "fw": "400"},
    "total":        {"bg": "#1F3864", "fg": "#FFFFFF", "fw": "700"},
}


def render_table(res: dict) -> str:
    html = """
    <style>
      .at{border-collapse:collapse;width:100%;font-family:Calibri,sans-serif;font-size:13px}
      .at th{padding:7px 10px;text-align:center;border:1px solid #888;
             background:#1F3864;color:#fff;font-weight:700;white-space:nowrap}
      .at td{padding:5px 10px;border:1px solid #ccc;white-space:nowrap}
      .r{text-align:right}.l{text-align:left}.c{text-align:center}
    </style>
    <table class="at"><thead><tr>
      <th class="l">Catégorie d'investissement</th>
      <th>Alloc. cible</th>
      <th>Marge</th>
      <th>Retraitement (M€)</th>
      <th>Allocation (M€)</th>
      <th>Allocation (%)</th>
      <th>VNC (M€)</th>
      <th>VNC (%)</th>
      <th>Engagements (M€)</th>
      <th>Financement appels (M€)</th>
      <th>Gains capital (M€)</th>
      <th>Alloc. projetée (M€)</th>
      <th>Alloc. projetée (%)</th>
    </tr></thead><tbody>
    """
    for idx, (label, rtype, detail, _, _, alloc_cible, marge, _) in enumerate(ROW_DEFS):
        r = res.get(idx)
        if not r:
            continue
        s  = STYLE[rtype]
        cs = f'background:{s["bg"]};color:{s["fg"]};font-weight:{s["fw"]};'

        d_s = fmt_m(r["D"])  if (rtype == "white" and detail == "normal") else ""
        g_s = fmt_m(r["G"])
        h_s = fmt_pct(r["H"])
        j_s = fmt_m(r["J"])  if r.get("J") is not None else ""
        k_s = fmt_pct(r["K"]) if r.get("K") is not None else ""
        m_s = fmt_m(r["M"])  if r.get("M") else ""
        n_s = fmt_m(r["N"])  if r.get("N") else ""
        o_s = fmt_m(r["O"])  if r.get("O") else ""
        p_s = fmt_m(r["P"])  if r.get("P") is not None else ""
        q_s = fmt_pct(r["Q"]) if r.get("Q") is not None else ""

        html += (
            f'<tr>'
            f'<td class="l" style="{cs}">{label}</td>'
            f'<td class="c" style="{cs}">{alloc_cible}</td>'
            f'<td class="c" style="{cs}">{marge}</td>'
            f'<td class="r" style="{cs}">{d_s}</td>'
            f'<td class="r" style="{cs}">{g_s}</td>'
            f'<td class="r" style="{cs}">{h_s}</td>'
            f'<td class="r" style="{cs}">{j_s}</td>'
            f'<td class="r" style="{cs}">{k_s}</td>'
            f'<td class="r" style="{cs}">{m_s}</td>'
            f'<td class="r" style="{cs}">{n_s}</td>'
            f'<td class="r" style="{cs}">{o_s}</td>'
            f'<td class="r" style="{cs}">{p_s}</td>'
            f'<td class="r" style="{cs}">{q_s}</td>'
            f'</tr>\n'
        )
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
        "Retraitement (M€)", "Allocation (M€)", "Allocation (%)", "VNC (M€)", "VNC (%)",
        "Engagements (M€)", "Financement appels (M€)", "Gains capital (M€)",
        "Alloc. projetée (M€)", "Alloc. projetée (%)",
    ]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = FILLS["header"]; c.font = FONTS["header"]; c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center")

    for erow, (idx, (label, rtype, detail, _, _, alloc_cible, marge, _)) in enumerate(
        enumerate(ROW_DEFS), 2
    ):
        r = res.get(idx)
        if not r:
            continue
        d_val = r["D"] if (rtype == "white" and detail == "normal") else None
        vals = [
            label, alloc_cible, marge,
            d_val, r["G"], r["H"],
            r["J"] if r.get("J") is not None else None,
            r["K"] if r.get("K") is not None else None,
            r["M"] if r.get("M") else None,
            r["N"] if r.get("N") else None,
            r["O"] if r.get("O") else None,
            r["P"] if r.get("P") is not None else None,
            r["Q"] if r.get("Q") is not None else None,
        ]
        fmts   = [None, "@", "@", "#,##0", "#,##0", "0.0%", "#,##0", "0.0%",
                  "#,##0", "#,##0", "#,##0", "#,##0", "0.0%"]
        aligns = ["left"] + ["right"] * 12

        for col, (val, fmt, align) in enumerate(zip(vals, fmts, aligns), 1):
            c = ws.cell(row=erow, column=col, value=val)
            c.fill = FILLS[rtype]; c.font = FONTS[rtype]; c.border = border
            c.alignment = Alignment(horizontal=align, vertical="center")
            if fmt and val is not None and val != "":
                c.number_format = fmt

    ws.column_dimensions["A"].width = 40
    for col_letter in ["B","C","D","E","F","G","H","I","J","K","L","M"]:
        ws.column_dimensions[col_letter].width = 16
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
  Déposez le fichier Excel → Allocation, VNC, Retraitements et projections recalculés automatiquement.
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
            st.warning("Onglet KNL absent — colonnes Engagements/Financement/Gains/Projection affichées à 0.")

        with st.spinner("Calcul en cours..."):
            res = compute(file_bytes)

        # ── Controle classes Retraitements non matchées
        categories_connues = {
            normalize(B)
            for _, rtype, detail, B, _, _, _, _ in ROW_DEFS
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
            st.warning(
                f"Classes Retraitements non reconnues (ignorées) : "
                f"**{', '.join(sorted(non_matches))}**"
            )

        g_tot = res[TOTAL_IDX]["G"]
        j_tot = res[TOTAL_IDX]["J"]
        p_tot = res[TOTAL_IDX]["P"] or 0.0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Allocation", f"{g_tot:,.0f} M€")
        c2.metric("Total VNC",        f"{j_tot:,.0f} M€")
        c3.metric("Ecart Alloc - VNC", f"{g_tot - j_tot:,.0f} M€")
        c4.metric("Alloc. projetée",   f"{p_tot:,.0f} M€")

        st.markdown("<br>", unsafe_allow_html=True)
        components.html(render_table(res), height=750, scrolling=True)
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
| **Portefeuille** | F, W, Y, AK |
| **Retraitements** | A (ISIN), B (nom), C (classe), D (montant) |
| **KNL**  | B (classe), K (engagements), M (retour capital), N (capital gain) |
        """)