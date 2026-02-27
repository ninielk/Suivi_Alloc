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
#  (label, type, detail, B=critère AK, C=critère F, alloc_cible, marge)
# ══════════════════════════════════════════════════════════════════
ROW_DEFS = [
    ("Obligations classiques",                  "blue",  None,                   None,                                       None,                          "44%",  "-20% / +5%"),
    ("Obligations souveraines",                 "white", "normal",               "Obligations souveraines",                  "EMPRUNTS ETATS & OBLIG GARANTIES", "20%",  "-20% / +5%"),
    ("Obligations privées",                     "white", "normal",               "Obligations privées",                      "OBLIGATIONS COTEES",          "24%",  "-20% / +5%"),
    ("Obligations nanties",                     "blue",  None,                   None,                                       None,                          "",     ""),
    ("Obligations souveraines",                 "white", "nanties",              "Obligations souveraines",                  None,                          "",     ""),
    ("Obligations privées",                     "white", "nanties_no_formula",   "Obligations privées",                      None,                          "",     ""),
    ("Autres produits de taux",                 "blue",  None,                   None,                                       None,                          "13%",  "-13% / +3%"),
    ("Dettes privées",                          "white", "normal",               "Dettes privées",                           None,                          "8%",   "-8% / +3%"),
    ("Alternatifs",                             "white", "normal",               "Alternatifs",                              None,                          "5%",   "-5% / +3%"),
    ("Actions",                                 "blue",  None,                   None,                                       None,                          "11%",  "-11% / +3%"),
    ("Actions internationales",                 "white", "normal",               "Actions internationales",                  None,                          "0%",   "-0% / +3%"),
    ("Actions Zone Euro",                       "white", "normal",               "Actions Zone Euro",                        None,                          "6%",   "-6% / +3%"),
    ("Autres actions (capital investissement)", "white", "normal",               "Autres actions (capital investissement)",  None,                          "5%",   "-5% / +3%"),
    ("Actifs réels",                            "blue",  None,                   None,                                       None,                          "19%",  "-19% / +3%"),
    ("Immobilier placement",                    "white", "normal",               "Immobilier placement",                     None,                          "13%",  "-13% / +3%"),
    ("Infrastructures",                         "white", "normal",               "Infrastructures",                          None,                          "6%",   "-6% / +3%"),
    ("Stratégique",                             "blue",  None,                   None,                                       None,                          "12%",  ""),
    ("Prêts stratégiques",                      "white", "normal",               "Prêts stratégiques",                       None,                          "2%",   ""),
    ("Immobilier stratégique",                  "white", "normal",               "Immobilier stratégique",                   None,                          "1%",   ""),
    ("Actions stratégiques",                    "white", "normal",               "Actions stratégiques",                     None,                          "9%",   ""),
    ("Trésorerie",                              "blue",  None,                   None,                                       None,                          "1%",   ""),
    ("Trésorerie",                              "white", "normal",               "Trésorerie",                               None,                          "1%",   ""),
    ("Total Général",                           "total", None,                   None,                                       None,                          "100%", ""),
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

REQUIRED_SHEETS = ["Portefeuille", "Retraitements"]

# ══════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════
def normalize(s):
    """Retire accents, majuscules, espaces superflus — matching robuste."""
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
    #    F=5  Catégorie instrument
    #    W=22 VNC comptable
    #    Y=24 Valeur de marché
    #    AK=36 Classification
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
        classe  = row[2]
        montant = row[3]
        if classe is None or str(classe).strip().lower() in ("classe d'actifs", "montant", ""):
            continue
        try:
            rows_r.append({"classe": str(classe).strip(), "D": float(montant)})
        except (TypeError, ValueError):
            continue
    df_r = pd.DataFrame(rows_r) if rows_r else pd.DataFrame(columns=["classe", "D"])

    # ── Calcul lignes blanches
    res = {}
    for idx, (label, rtype, detail, B, C, _, _) in enumerate(ROW_DEFS):
        if rtype != "white":
            continue

        # Retraitement (col D)
        if detail == "normal" and not df_r.empty:
            d = somme_si(df_r, "classe", B, "D") / 1e6
        else:
            d = 0.0

        # Allocation M€ (col G)
        if detail == "nanties_no_formula":
            g = 0.0
        else:
            g_ak = somme_si(df_p, "AK", B, "Y")
            g_f  = somme_si(df_p, "F",  C, "Y") if C else 0.0
            g    = (g_ak + g_f) / 1e6 + d

        # VNC M€ (col J)
        if detail in ("nanties", "nanties_no_formula"):
            j = None
        else:
            j_ak = somme_si(df_p, "AK", B, "W")
            j_f  = somme_si(df_p, "F",  C, "W") if C else 0.0
            j    = (j_ak + j_f) / 1e6

        res[idx] = {"label": label, "type": rtype, "detail": detail, "D": d, "G": g, "J": j}

    # ── Lignes bleues = somme enfants
    for br, children in BLUE_CHILDREN.items():
        g_sum = sum(res[c]["G"] for c in children if c in res)
        j_sum = sum(res[c]["J"] for c in children if c in res and res[c]["J"] is not None)
        res[br] = {"label": ROW_DEFS[br][0], "type": "blue", "detail": None,
                   "D": 0.0, "G": g_sum, "J": j_sum}

    # ── Total
    g_tot = sum(res[i]["G"] for i in BLUE_IDX)
    j_tot = sum(res[i]["J"] for i in BLUE_IDX if isinstance(res[i].get("J"), (int, float)))
    res[TOTAL_IDX] = {"label": "Total Général", "type": "total", "detail": None,
                      "D": 0.0, "G": g_tot, "J": j_tot}

    # ── Pourcentages
    for r in res.values():
        r["H"] = r["G"] / g_tot if g_tot else 0.0
        r["K"] = (r["J"] / j_tot) if (j_tot and r.get("J") is not None) else None

    return res


# ══════════════════════════════════════════════════════════════════
#  RENDU HTML
# ══════════════════════════════════════════════════════════════════
STYLE = {
    "blue":  {"bg": "#2E75B6", "fg": "#FFFFFF", "fw": "700"},
    "white": {"bg": "#FFFFFF", "fg": "#1a1a1a", "fw": "400"},
    "total": {"bg": "#1F3864", "fg": "#FFFFFF", "fw": "700"},
}


def render_table(res: dict) -> str:
    html = """
    <style>
      .at{border-collapse:collapse;width:100%;font-family:Calibri,sans-serif;font-size:13px}
      .at th{padding:7px 12px;text-align:center;border:1px solid #888;
             background:#1F3864;color:#fff;font-weight:700;white-space:nowrap}
      .at td{padding:5px 12px;border:1px solid #ccc;white-space:nowrap}
      .r{text-align:right}.l{text-align:left}.c{text-align:center}
    </style>
    <table class="at"><thead><tr>
      <th class="l">Catégorie d'investissement</th>
      <th>Alloc. cible</th>
      <th>Marge de manoeuvre</th>
      <th>Retraitement (M€)</th>
      <th>Allocation (M€)</th>
      <th>Allocation (%)</th>
      <th>VNC (M€)</th>
      <th>VNC (%)</th>
    </tr></thead><tbody>
    """
    for idx, (label, rtype, detail, _, _, alloc_cible, marge) in enumerate(ROW_DEFS):
        r = res.get(idx)
        if not r:
            continue
        s  = STYLE[rtype]
        cs = f'background:{s["bg"]};color:{s["fg"]};font-weight:{s["fw"]};'

        d_s = fmt_m(r["D"]) if (rtype == "white" and detail == "normal") else ""
        g_s = fmt_m(r["G"])
        h_s = fmt_pct(r["H"])
        j_s = fmt_m(r["J"]) if r.get("J") is not None else ""
        k_s = fmt_pct(r["K"]) if r.get("K") is not None else ""

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
        "blue":   PatternFill(fgColor="2E75B6", fill_type="solid"),
        "white":  PatternFill(fgColor="FFFFFF", fill_type="solid"),
        "total":  PatternFill(fgColor="1F3864", fill_type="solid"),
        "header": PatternFill(fgColor="1F3864", fill_type="solid"),
    }
    FONTS = {
        "blue":   Font(bold=True, color="FFFFFF", name="Calibri", size=11),
        "white":  Font(color="000000",            name="Calibri", size=11),
        "total":  Font(bold=True, color="FFFFFF", name="Calibri", size=11),
        "header": Font(bold=True, color="FFFFFF", name="Calibri", size=11),
    }

    headers = [
        "Catégorie d'investissement", "Alloc. cible", "Marge de manoeuvre",
        "Retraitement (M€)", "Allocation (M€)", "Allocation (%)", "VNC (M€)", "VNC (%)",
    ]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = FILLS["header"]; c.font = FONTS["header"]; c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center")

    for erow, (idx, (label, rtype, detail, _, _, alloc_cible, marge)) in enumerate(
        enumerate(ROW_DEFS), 2
    ):
        r = res.get(idx)
        if not r:
            continue
        d_val = r["D"] if (rtype == "white" and detail == "normal") else None
        vals  = [label, alloc_cible, marge, d_val, r["G"], r["H"],
                 r["J"] if r.get("J") is not None else None,
                 r["K"] if r.get("K") is not None else None]
        fmts  = [None, "@", "@", "#,##0", "#,##0", "0.0%", "#,##0", "0.0%"]
        aligns = ["left"] + ["right"] * 7

        for col, (val, fmt, align) in enumerate(zip(vals, fmts, aligns), 1):
            c = ws.cell(row=erow, column=col, value=val)
            c.fill = FILLS[rtype]; c.font = FONTS[rtype]; c.border = border
            c.alignment = Alignment(horizontal=align, vertical="center")
            if fmt and val is not None and val != "":
                c.number_format = fmt

    ws.column_dimensions["A"].width = 42
    for col in ["B", "C", "D", "E", "F", "G", "H"]:
        ws.column_dimensions[col].width = 18
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
  Dépose le fichier Excel → Allocation, VNC et Retraitements recalculés automatiquement.
</p><hr style='border:1px solid #ddd;'>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        file_bytes = uploaded.read()

        wb_chk = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True)
        missing = [s for s in REQUIRED_SHEETS if s not in wb_chk.sheetnames]
        wb_chk.close()
        if missing:
            st.error(f"Onglets manquants : `{'`, `'.join(missing)}`")
            st.stop()

        with st.spinner("Calcul en cours..."):
            res = compute(file_bytes)

        # ── Controle : classes Retraitements non matchees ────────
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
            st.warning(
                f"Ces classes dans Retraitements n'ont matche aucune categorie "
                f"et sont ignorees : **{', '.join(sorted(non_matches))}**"
            )

        g_tot = res[TOTAL_IDX]["G"]
        j_tot = res[TOTAL_IDX]["J"]

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Allocation", f"{g_tot:,.0f} M€")
        c2.metric("Total VNC",        f"{j_tot:,.0f} M€")
        c3.metric("Ecart Alloc - VNC", f"{g_tot - j_tot:,.0f} M€")

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
| **Portefeuille** | F (catégorie instrument), W (VNC), Y (valeur marché), AK (classification) |
| **Retraitements** | A (ISIN), B (nom fonds), C (classe d'actifs), D (montant) |
        """)