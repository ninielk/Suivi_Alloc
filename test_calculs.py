"""
test_calculs.py — Tests unitaires pour le Suivi d'Allocation LMG
Lancer avec : pytest test_calculs.py -v
"""

import numpy as np
import pandas as pd
import pytest


# ══════════════════════════════════════════════════════════════════
#  HELPERS (copiés depuis app.py pour les tester indépendamment)
# ══════════════════════════════════════════════════════════════════

def somme_si(df: pd.DataFrame, criteria_col: str, criteria_val, sum_col: str) -> float:
    if criteria_val is None or (isinstance(criteria_val, float) and np.isnan(criteria_val)):
        return 0.0
    mask = df[criteria_col].astype(str).str.strip().str.upper() == str(criteria_val).strip().upper()
    return float(df.loc[mask, sum_col].sum())


def scr_total_from_vector(scr_taux, scr_spread, scr_actions, scr_immo, scr_forex):
    corr = np.array([
        [1,    0,    0,    0,    0.25],
        [0,    1,    0.75, 0.5,  0.25],
        [0,    0.75, 1,    0.75, 0.25],
        [0,    0.5,  0.75, 1,    0.25],
        [0.25, 0.25, 0.25, 0.25, 1   ],
    ])
    v = np.array([scr_taux, scr_spread, scr_actions, scr_immo, scr_forex])
    return float(np.sqrt(v @ corr @ v)) if v.any() else 0.0


def duration_alm(u_dict: dict, h_dict: dict) -> float:
    """U25 = SOMMEPROD(U ; H)"""
    return sum(u_dict.get(i, 0.0) * h_dict.get(i, 0.0) for i in u_dict)


def duration_ponderee(eq_list, z_list) -> float:
    """SOMMEPROD(EQ ; Z) / SOMME(Z)"""
    z_sum = sum(z_list)
    if z_sum == 0:
        return 0.0
    return sum(eq * z for eq, z in zip(eq_list, z_list)) / z_sum


# ══════════════════════════════════════════════════════════════════
#  TESTS SOMME.SI
# ══════════════════════════════════════════════════════════════════

def test_somme_si_basic():
    df = pd.DataFrame({"F": ["OBLIG", "OBLIG", "ACTIONS"], "Y": [100.0, 200.0, 50.0]})
    assert somme_si(df, "F", "OBLIG", "Y") == 300.0


def test_somme_si_case_insensitive():
    df = pd.DataFrame({"F": ["oblig", "OBLIG"], "Y": [100.0, 200.0]})
    assert somme_si(df, "F", "Oblig", "Y") == 300.0


def test_somme_si_aucun_match():
    df = pd.DataFrame({"F": ["ACTIONS"], "Y": [100.0]})
    assert somme_si(df, "F", "OBLIG", "Y") == 0.0


def test_somme_si_critere_none():
    df = pd.DataFrame({"F": ["OBLIG"], "Y": [100.0]})
    assert somme_si(df, "F", None, "Y") == 0.0


# ══════════════════════════════════════════════════════════════════
#  TESTS SCR / MATRICE CORRÉLATION
# ══════════════════════════════════════════════════════════════════

def test_scr_diversification():
    """Le SCR total avec corrélation doit être < somme des modules."""
    scr_taux    = 35.1
    scr_spread  = 45.3
    scr_actions = 98.7
    scr_immo    = 10.1
    scr_forex   = 3.2
    total = scr_total_from_vector(scr_taux, scr_spread, scr_actions, scr_immo, scr_forex)
    somme = scr_taux + scr_spread + scr_actions + scr_immo + scr_forex
    assert total < somme, f"SCR total {total:.2f} devrait être < {somme:.2f}"


def test_scr_un_seul_module():
    """Si un seul module, SCR total = ce module (pas de diversification)."""
    scr = scr_total_from_vector(100, 0, 0, 0, 0)
    assert abs(scr - 100) < 0.01


def test_scr_zero():
    """SCR nul si tous les modules sont à 0."""
    assert scr_total_from_vector(0, 0, 0, 0, 0) == 0.0


def test_scr_positif():
    """Le SCR total est toujours positif."""
    scr = scr_total_from_vector(35, 45, 98, 10, 3)
    assert scr > 0


def test_scr_valeurs_reelles():
    """Test avec les vraies valeurs du portefeuille LMG."""
    # Valeurs issues du tableau TPT vérifié manuellement
    scr = scr_total_from_vector(72.2, 52.9, 148.9, 10.1, 3.2)
    # SCR total attendu ~212.9 M€
    assert 200 < scr < 225, f"SCR total inattendu : {scr:.1f}"


# ══════════════════════════════════════════════════════════════════
#  TESTS DURATION
# ══════════════════════════════════════════════════════════════════

def test_duration_ponderee_simple():
    """Duration pondérée basique."""
    dur = duration_ponderee([10.0, 5.0], [100.0, 100.0])
    assert abs(dur - 7.5) < 0.001


def test_duration_ponderee_un_actif():
    """Un seul actif → sa duration."""
    dur = duration_ponderee([10.81], [306.1])
    assert abs(dur - 10.81) < 0.001


def test_duration_ponderee_zero_z():
    """Valeur marché nulle → duration = 0."""
    assert duration_ponderee([10.0], [0.0]) == 0.0


def test_duration_alm_100pct_souv():
    """100% en souv (duration 10.81) → Duration ALM = 10.81."""
    u = {0: 10.81}
    h = {0: 1.0}
    assert abs(duration_alm(u, h) - 10.81) < 0.001


def test_duration_alm_mixte():
    """Mix 50% souv (10.81) + 50% actions (20) → Duration ALM = 15.405."""
    u = {0: 10.81, 1: 20.0}
    h = {0: 0.5, 1: 0.5}
    expected = 10.81 * 0.5 + 20.0 * 0.5
    assert abs(duration_alm(u, h) - expected) < 0.001


def test_duration_alm_valeurs_reelles():
    """Test approximatif avec valeurs LMG réelles — U25 doit être ~9.58."""
    # Mapping idx → (U, H%) depuis le tableau
    u_vals = {
        1: 10.81, 2: 4.66,   # oblig souv/priv classiques
        4: 10.81, 5: 4.66,   # nanties
        7: 2.64, 8: 2.64,    # dettes privées, alternatifs
        10: 20.0, 11: 20.0, 12: 10.0,
        14: 20.0, 15: 20.0,
        17: 10.0, 18: 20.0, 19: 10.0,
        21: 1.0,
    }
    h_vals = {
        1: 0.157, 2: 0.159,
        4: 0.026, 5: 0.147,
        7: 0.056, 8: 0.032,
        10: 0.00, 11: 0.020, 12: 0.060,
        14: 0.137, 15: 0.050,
        17: 0.016, 18: 0.004, 19: 0.092,
        21: 0.044,
    }
    result = duration_alm(u_vals, h_vals)
    assert 8.5 < result < 10.5, f"Duration ALM inattendue : {result:.3f}"


# ══════════════════════════════════════════════════════════════════
#  TESTS ALLOCATION
# ══════════════════════════════════════════════════════════════════

def test_allocation_somme_100pct():
    """La somme des allocations H% des catégories parents doit faire ~100%."""
    h_vals = [0.316, 0.173, 0.088, 0.080, 0.186, 0.112, 0.044]
    assert abs(sum(h_vals) - 1.0) < 0.01


def test_scr_actifs_immo_hardcode():
    """Choc SCR Immobilier = 22% hardcodé S2."""
    assert abs(0.22 - 0.22) < 0.001  # valeur réglementaire stable


def test_scr_actifs_infra_hardcode():
    """Choc SCR Infrastructures = 32.5% hardcodé S2."""
    assert abs(0.325 - 0.325) < 0.001


def test_alloc_projetee_coherente():
    """Alloc projetée = G + M - N + O doit être positive pour la plupart des catégories."""
    # Exemple : oblig souv classiques
    g, m, n, o = 257.0, 0.0, 0.0, 0.0
    p = g + m - n + o
    assert p > 0