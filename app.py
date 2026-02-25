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
#  (numéro de ligne dans Alloc, label, type, detail)
#  type   : blue=groupe, white=détail, total=ligne total
#  detail : normal | nanties | nanties_no_formula
# ══════════════════════════════════════════════════════════════════
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

BLUE_CHILDREN = {
    3:  [4, 5],
    6:  [7, 8],
    9:  [10, 11],
    12: [13, 14, 15],
    16: [17, 18],
    19: [20, 21, 22],
    23: [24],
}
BLUE_ROWS   = [3, 6, 9, 12, 16, 19, 23]
TOTAL_ROW   = 25
REQUIRED_SHEETS = ["Portefeuille", "Retraitements", "Alloc"]

# ══════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════
def somme_si(df, crit_col, crit_val, sum_col):
    if crit_val is None:
        return 0.0
    if isinstance(crit_val, float) and np.isnan(crit_val):
        return 0.0
    mask = (
        df[crit_col].astype(str).str.strip().str.upper()
        == str(crit_val).strip().upper()
    )
    return float(df.loc[mask, sum_col].sum())


def fmt_m(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    return f"{val:,.0f}"


def fmt_pct(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    return f"{val * 100:.1f}%"


def fmt_alloc_cible(val):
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        v = float(val)
        return f"{v * 100:.0f}%" if abs(v) <= 1.5 else f"{v:.0f}%"
    return str(val)


# ══════════════════════════════════════════════════════════════════
#  CALCUL
# ══════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def compute(file_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)

    # ── Portefeuille (colonnes 0-indexées)
    #    F=5  Catégorie instrument
    #    W=22 VNC comptable
    #    Y=24 Valeur de marché hors cc
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
    #    B=1 Classe d'actifs   C=2 Montant
    ws_r = wb["Retraitements"]
    rows_r = []
    for row in ws_r.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 3 or row[1] is None:
            continue
        rows_r.append({"B": row[1], "C": row[2] or 0})
    df_r = pd.DataFrame(rows_r) if rows_r else pd.DataFrame(columns=["B", "C"])
    if not df_r.empty:
        df_r["C"] = pd.to_numeric(df_r["C"], errors="coerce").fillna(0)

    # ── Critères depuis Alloc
    #    col2=B (critère AK)  col3=C (critère F)
    #    col5=E (alloc cible) col6=F (marge manœuvre)
    ws_a = wb["Alloc"]
    crit = {}
    for rn, _, _, _ in ROW_DEFS:
        crit[rn] = {
            "B": ws_a.cell(row=rn, column=2).value,
            "C": ws_a.cell(row=rn, column=3).value,
            "E": ws_a.cell(row=rn, column=5).value,
            "F": ws_a.cell(row=rn, column=6).value,
        }

    # ── Calcul lignes blanches
    res = {}
    for rn, label, rtype, detail in ROW_DEFS:
        if rtype != "white":
            continue

        B = crit[rn]["B"]
        C = crit[rn]["C"]

        # Col D – Retraitements
        if detail == "normal" and not df_r.empty:
            d = somme_si(df_r, "B", B, "C") / 1e6
        else:
            d = 0.0

        # Col G – Allocation M€
        if detail == "nanties_no_formula":
            g = 0.0
        else:
            g = (somme_si(df_p, "F", C, "Y") + somme_si(df_p, "AK", B, "Y")) / 1e6 + d

        # Col J – VNC M€
        if detail in ("nanties", "nanties_no_formula"):
            j = None
        else:
            j = (somme_si(df_p, "F", C, "W") + somme_si(df_p, "AK", B, "W")) / 1e6

        res[rn] = {
            "label": label, "type": rtype, "detail": detail,
            "D": d, "E": crit[rn]["E"], "F": crit[rn]["F"], "G": g, "J": j,
        }

    # ── Lignes bleues = somme enfants
    for br, children in BLUE_CHILDREN.items():
        g_sum = sum(res[c]["G"] for c in children)
        j_sum = sum(res[c]["J"] for c in children if res[c]["J"] is not None)
        lbl   = next(l for r, l, t, _ in ROW_DEFS if r == br)
        res[br] = {
            "label": lbl, "type": "blue", "detail": None,
            "D": 0.0, "E": crit[br]["E"], "F": crit[br]["F"], "G": g_sum, "J": j_sum,
        }

    # ── Total
    g_tot = sum(res[r]["G"] for r in BLUE_ROWS)
    j_tot = sum(res[r]["J"] for r in BLUE_ROWS if isinstance(res[r]["J"], (int, float)))
    res[TOTAL_ROW] = {
        "label": "Total Général", "type": "total", "detail": None,
        "D": 0.0, "E": crit[TOTAL_ROW]["E"], "F": crit[TOTAL_ROW]["F"],
        "G": g_tot, "J": j_tot,
    }

    # ── Pourcentages
    for r in res.values():
        r["H"] = r["G"] / g_tot if g_tot else 0.0
        r["K"] = (r["J"] / j_tot) if (j_tot and r["J"] is not None) else None

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
      .r{text-align:right} .l{text-align:left} .c{text-align:center}
    </style>
    <table class="at"><thead><tr>
      <th class="l">Catégorie d'investissement</th>
      <th>Alloc. cible</th>
      <th>Marge de manœuvre</th>
      <th>Retraitement (M€)</th>
      <th>Allocation (M€)</th>
      <th>Allocation (%)</th>
      <th>VNC (M€)</th>
      <th>VNC (%)</th>
    </tr></thead><tbody>
    """
    for rn, label, rtype, detail in ROW_DEFS:
        r = res.get(rn)
        if not r:
            continue
        s  = STYLE[rtype]
        cs = f'background:{s["bg"]};color:{s["fg"]};font-weight:{s["fw"]};'

        d_s = fmt_m(r["D"])   if (rtype == "white" and detail == "normal") else ""
        e_s = fmt_alloc_cible(r["E"])
        f_s = str(r["F"])     if r["F"] is not None else ""
        g_s = fmt_m(r["G"])
        h_s = fmt_pct(r["H"])
        j_s = fmt_m(r["J"])   if r["J"] is not None else ""
        k_s = fmt_pct(r["K"]) if r["K"] is not None else ""

        html += (
            f'<tr>'
            f'<td class="l" style="{cs}">{label}</td>'
            f'<td class="c" style="{cs}">{e_s}</td>'
            f'<td class="c" style="{cs}">{f_s}</td>'
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
        "Catégorie d'investissement", "Alloc. cible", "Marge de manœuvre",
        "Retraitement (M€)", "Allocation (M€)", "Allocation (%)", "VNC (M€)", "VNC (%)",
    ]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = FILLS["header"]; c.font = FONTS["header"]; c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center")

    for i, (rn, label, rtype, detail) in enumerate(ROW_DEFS):
        r = res.get(rn)
        if not r:
            continue
        erow = i + 2
        d_val = r["D"] if (rtype == "white" and detail == "normal") else None
        vals  = [label, r["E"], r["F"], d_val, r["G"], r["H"],
                 r["J"] if r["J"] is not None else None,
                 r["K"] if r["K"] is not None else None]
        fmts  = [None, "0%", "@", "#,##0", "#,##0", "0.0%", "#,##0", "0.0%"]
        aligns= ["left"] + ["right"] * 7

        for col, (val, fmt, align) in enumerate(zip(vals, fmts, aligns), 1):
            c = ws.cell(row=erow, column=col, value=val)
            c.fill = FILLS[rtype]; c.font = FONTS[rtype]; c.border = border
            c.alignment = Alignment(horizontal=align, vertical="center")
            if fmt and val is not None:
                c.number_format = fmt

    ws.column_dimensions["A"].width = 42
    for col in ["B","C","D","E","F","G","H"]:
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

        # Vérif onglets
        wb_chk = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True)
        missing = [s for s in REQUIRED_SHEETS if s not in wb_chk.sheetnames]
        wb_chk.close()
        if missing:
            st.error(f"Onglets manquants : `{'`, `'.join(missing)}`")
            st.stop()

        with st.spinner("Calcul en cours..."):
            res = compute(file_bytes)

        g_tot = res[TOTAL_ROW]["G"]
        j_tot = res[TOTAL_ROW]["J"]

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Allocation", f"{g_tot:,.0f} M€")
        c2.metric("Total VNC",        f"{j_tot:,.0f} M€")
        c3.metric("Écart Alloc − VNC", f"{g_tot - j_tot:,.0f} M€")

        st.markdown("<br>", unsafe_allow_html=True)
        components.html(render_table(res), height=700, scrolling=True)
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
        with st.expander("Détails"):
            st.exception(e)
else:
    st.info("En attente du fichier Excel...")
    with st.expander("Onglets et colonnes requis"):
        st.markdown("""
| Onglet | Colonnes |
|---|---|
| **Portefeuille** | F (Catégorie instrument), W (VNC), Y (Valeur marché), AK (Classification) |
| **Retraitements** | B (Classe d'actifs), C (Montant) |
| **Alloc** | B, C (critères), E (Alloc cible), F (Marge de manœuvre) |
        """)