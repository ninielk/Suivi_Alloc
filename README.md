# 📊 Suivi d'Allocation

Application Streamlit de suivi d'allocation d'actifs — automatise les calculs SOMME.SI du tableau Alloc Excel.

## Fonctionnement

L'utilisateur dépose le fichier Excel template → l'app recalcule automatiquement les colonnes suivantes :

| Colonne | Description |
|---|---|
| **D** | Retraitements (M€) — SOMME.SI sur onglet Retraitements |
| **G** | Allocation (M€) — double SOMME.SI sur colonnes F et AK du Portefeuille |
| **H** | Allocation (%) — G / Total |
| **J** | VNC (M€) — double SOMME.SI colonne W du Portefeuille |
| **K** | VNC (%) — J / Total |

Les colonnes **Allocation cible** et **Marge de manœuvre** sont lues directement depuis l'onglet Alloc du template.

## Structure du projet

```
Suivi_Alloc/
├── app.py                  # Point d'entrée Streamlit
├── src/
│   ├── __init__.py
│   ├── calculs.py          # Logique métier (SOMME.SI, structure des lignes)
│   └── export.py           # Export Excel formaté
├── .streamlit/
│   └── config.toml         # Config Streamlit (thème)
├── requirements.txt
├── .gitignore
└── README.md
```

## Onglets requis dans le fichier Excel

| Onglet | Colonnes utilisées |
|---|---|
| **Portefeuille** | F (Catégorie instrument), W (VNC comptable), Y (Valeur marché hors cc), AK (Classification) |
| **Retraitements** | B (Classe d'actifs), C (Montant — valeurs issues de Feuil1) |
| **Alloc** | B, C (critères matching), E (Alloc cible), F (Marge de manœuvre) |
| Feuil1, TPT, ListeActifs… | Présence requise pour cohérence du classeur |

## Installation & lancement

```bash
# Cloner le repo
git clone https://github.com/ninielk/Suivi_Alloc.git
cd Suivi_Alloc

# Installer les dépendances
pip install -r requirements.txt

# Lancer l'app
streamlit run app.py
```

## Déploiement Streamlit Cloud

1. Push le repo sur GitHub
2. Aller sur [share.streamlit.io](https://share.streamlit.io)
3. Connecter le repo → `app.py` comme point d'entrée
4. Deploy 🚀
