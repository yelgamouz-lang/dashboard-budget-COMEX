"""
═══════════════════════════════════════════════════════════════════════════════
 COMEX BUDGET DASHBOARD — GROUPE INDUSTRIE MAROC — N+1
 Glassmorphism Dark / Mauve Néon · Direction Financière
═══════════════════════════════════════════════════════════════════════════════

Lance avec :
    streamlit run app.py

Pré-requis :
    pip install streamlit pandas plotly openpyxl

Le dashboard détecte automatiquement les modifications du fichier Excel
(Budget_Groupe_Industrie_NPlus1.xlsx) via os.path.getmtime et invalide le cache
pour rafraîchir tous les KPIs en temps réel.
"""

import os
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook

# ═══════════════════════════════════════════════════════════════════════════════
# 0. CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

EXCEL_PATH = Path(__file__).parent / "Budget_Groupe_Industrie_NPlus1.xlsx"

st.set_page_config(
    page_title="COMEX · Budget N+1 · Industrie Maroc",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Couleurs charte (cohérence avec l'Excel #800080 / Gold)
PURPLE_NEON = "#B026FF"
PURPLE_DEEP = "#6A0DAD"
PURPLE_DARK = "#4B004B"
GOLD_ACCENT = "#D4AF37"
CYAN_GLOW = "#00E5FF"
RED_NEON = "#FF3D71"
GREEN_NEON = "#00E676"
AMBER_NEON = "#FFB300"
BG_DARK = "#0A0014"


# ═══════════════════════════════════════════════════════════════════════════════
# 1. DESIGN — INJECTION CSS GLASSMORPHISM
# ═══════════════════════════════════════════════════════════════════════════════

def inject_css():
    """Injecte le CSS glassmorphism dark/mauve avec animations staggered reveal."""
    st.markdown("""
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@300;400;500;600;700;800&family=Space+Grotesk:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">

    <style>
    /* ════════ RESET & GLOBAL ════════ */
    .stApp {
        background: linear-gradient(135deg, #000000 0%, #0A0014 60%, #1A0033 85%, #6A0DAD 100%);
        background-attachment: fixed;
        font-family: 'Manrope', -apple-system, sans-serif;
        color: #F0F0FF;
    }

    /* Suppression des marges streamlit par défaut */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 3rem !important;
        max-width: 1600px;
    }

    /* Hide streamlit branding */
    #MainMenu, footer, header { visibility: hidden; }

    /* Halo radial décoratif en arrière-plan */
    .stApp::before {
        content: "";
        position: fixed;
        top: -10%;
        right: -10%;
        width: 800px;
        height: 800px;
        background: radial-gradient(circle, rgba(176, 38, 255, 0.15) 0%, transparent 70%);
        filter: blur(80px);
        z-index: 0;
        pointer-events: none;
    }
    .stApp::after {
        content: "";
        position: fixed;
        bottom: -20%;
        left: -10%;
        width: 700px;
        height: 700px;
        background: radial-gradient(circle, rgba(0, 229, 255, 0.08) 0%, transparent 70%);
        filter: blur(80px);
        z-index: 0;
        pointer-events: none;
    }

    /* ════════ TITRE PRINCIPAL ════════ */
    .hero-title {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 2.4rem;
        font-weight: 700;
        letter-spacing: -0.02em;
        background: linear-gradient(120deg, #FFFFFF 0%, #B026FF 70%, #D4AF37 100%);
        -webkit-background-clip: text;
        background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 0 0 0.3rem 0;
        animation: fadeInDown 0.8s cubic-bezier(0.2, 0.8, 0.2, 1) both;
    }
    .hero-subtitle {
        font-family: 'Manrope', sans-serif;
        font-size: 0.85rem;
        font-weight: 400;
        color: rgba(240, 240, 255, 0.55);
        letter-spacing: 0.18em;
        text-transform: uppercase;
        margin: 0 0 2rem 0;
        animation: fadeInDown 0.8s 0.1s cubic-bezier(0.2, 0.8, 0.2, 1) both;
    }
    .hero-divider {
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(176, 38, 255, 0.5), transparent);
        margin: 0 0 1.8rem 0;
    }

    /* ════════ KEYFRAMES ════════ */
    @keyframes fadeInDown {
        from { opacity: 0; transform: translateY(-15px); }
        to { opacity: 1; transform: translateY(0); }
    }
    @keyframes fadeInUp {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }
    @keyframes glow {
        0%, 100% { box-shadow: 0 0 20px rgba(176, 38, 255, 0.15); }
        50% { box-shadow: 0 0 30px rgba(176, 38, 255, 0.35); }
    }

    /* ════════ KPI CARDS ════════ */
    .kpi-card {
        background: rgba(255, 255, 255, 0.04);
        backdrop-filter: blur(14px);
        -webkit-backdrop-filter: blur(14px);
        border: 1px solid rgba(176, 38, 255, 0.18);
        border-radius: 14px;
        padding: 1.2rem 1.3rem;
        position: relative;
        overflow: hidden;
        transition: all 0.4s cubic-bezier(0.2, 0.8, 0.2, 1);
        height: 100%;
        animation: fadeInUp 0.6s cubic-bezier(0.2, 0.8, 0.2, 1) both;
    }
    .kpi-card:hover {
        border-color: rgba(176, 38, 255, 0.5);
        background: rgba(255, 255, 255, 0.06);
        transform: translateY(-3px);
        box-shadow: 0 8px 32px rgba(176, 38, 255, 0.25);
    }
    .kpi-card::before {
        content: "";
        position: absolute;
        top: 0; left: 0;
        width: 100%; height: 2px;
        background: linear-gradient(90deg, transparent, #B026FF, transparent);
        opacity: 0.6;
    }
    /* Stagger animation per child */
    .kpi-row > div:nth-child(1) .kpi-card { animation-delay: 0.05s; }
    .kpi-row > div:nth-child(2) .kpi-card { animation-delay: 0.10s; }
    .kpi-row > div:nth-child(3) .kpi-card { animation-delay: 0.15s; }
    .kpi-row > div:nth-child(4) .kpi-card { animation-delay: 0.20s; }
    .kpi-row > div:nth-child(5) .kpi-card { animation-delay: 0.25s; }

    .kpi-label {
        font-family: 'Manrope', sans-serif;
        font-size: 0.7rem;
        font-weight: 500;
        color: rgba(240, 240, 255, 0.55);
        letter-spacing: 0.12em;
        text-transform: uppercase;
        margin: 0 0 0.5rem 0;
    }
    .kpi-value {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 1.75rem;
        font-weight: 600;
        color: #FFFFFF;
        letter-spacing: -0.02em;
        line-height: 1.1;
        margin: 0 0 0.3rem 0;
        text-shadow: 0 0 20px rgba(176, 38, 255, 0.3);
    }
    .kpi-value.accent {
        background: linear-gradient(120deg, #FFFFFF 0%, #B026FF 100%);
        -webkit-background-clip: text;
        background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .kpi-value.gold {
        background: linear-gradient(120deg, #FFFFFF 0%, #D4AF37 100%);
        -webkit-background-clip: text;
        background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .kpi-unit {
        font-size: 0.85rem;
        font-weight: 400;
        color: rgba(240, 240, 255, 0.5);
        margin-left: 0.3rem;
    }
    .kpi-trend {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.75rem;
        font-weight: 500;
        display: inline-flex;
        align-items: center;
        gap: 0.3rem;
        padding: 0.2rem 0.55rem;
        border-radius: 999px;
        margin-top: 0.2rem;
    }
    .trend-up {
        background: rgba(0, 230, 118, 0.12);
        color: #00E676;
        border: 1px solid rgba(0, 230, 118, 0.3);
    }
    .trend-down {
        background: rgba(255, 61, 113, 0.12);
        color: #FF3D71;
        border: 1px solid rgba(255, 61, 113, 0.3);
    }
    .trend-neutral {
        background: rgba(255, 179, 0, 0.12);
        color: #FFB300;
        border: 1px solid rgba(255, 179, 0, 0.3);
    }
    .kpi-sub {
        font-size: 0.7rem;
        color: rgba(240, 240, 255, 0.4);
        margin-top: 0.4rem;
        font-style: italic;
    }

    /* ════════ SECTION HEADERS ════════ */
    .section-title {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 1.05rem;
        font-weight: 600;
        color: #FFFFFF;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        margin: 2rem 0 1rem 0;
        padding-bottom: 0.6rem;
        border-bottom: 1px solid rgba(176, 38, 255, 0.2);
        display: flex;
        align-items: center;
        gap: 0.6rem;
    }
    .section-title::before {
        content: "◆";
        color: #B026FF;
        font-size: 0.7rem;
        text-shadow: 0 0 10px #B026FF;
    }

    /* ════════ CHART CONTAINERS ════════ */
    .chart-container {
        background: rgba(255, 255, 255, 0.03);
        backdrop-filter: blur(14px);
        -webkit-backdrop-filter: blur(14px);
        border: 1px solid rgba(176, 38, 255, 0.15);
        border-radius: 14px;
        padding: 1.2rem;
        margin-bottom: 1rem;
    }

    /* ════════ SIDEBAR ════════ */
    section[data-testid="stSidebar"] {
        background: rgba(10, 0, 20, 0.85);
        backdrop-filter: blur(20px);
        border-right: 1px solid rgba(176, 38, 255, 0.2);
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 2rem;
    }
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label {
        color: #F0F0FF !important;
        font-family: 'Manrope', sans-serif !important;
    }
    section[data-testid="stSidebar"] label {
        font-size: 0.7rem !important;
        font-weight: 500 !important;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: rgba(240, 240, 255, 0.6) !important;
    }
    .sidebar-brand {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 1.1rem;
        font-weight: 700;
        background: linear-gradient(120deg, #FFFFFF 0%, #B026FF 100%);
        -webkit-background-clip: text;
        background-clip: text;
        -webkit-text-fill-color: transparent;
        letter-spacing: 0.05em;
        margin-bottom: 0.2rem;
    }
    .sidebar-tagline {
        font-size: 0.7rem;
        color: rgba(240, 240, 255, 0.4);
        letter-spacing: 0.15em;
        text-transform: uppercase;
        margin-bottom: 1.5rem;
    }

    /* Multiselect & selectbox styling */
    div[data-baseweb="select"] > div {
        background: rgba(255, 255, 255, 0.04) !important;
        border: 1px solid rgba(176, 38, 255, 0.25) !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="select"]:hover > div {
        border-color: rgba(176, 38, 255, 0.5) !important;
    }
    div[data-baseweb="tag"] {
        background: rgba(176, 38, 255, 0.2) !important;
        border: 1px solid rgba(176, 38, 255, 0.4) !important;
    }

    /* Slider styling */
    div[data-baseweb="slider"] > div > div > div {
        background: linear-gradient(90deg, #B026FF, #D4AF37) !important;
    }

    /* Buttons */
    .stButton > button {
        background: rgba(176, 38, 255, 0.12);
        color: #FFFFFF;
        border: 1px solid rgba(176, 38, 255, 0.4);
        border-radius: 8px;
        font-family: 'Manrope', sans-serif;
        font-weight: 500;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        font-size: 0.75rem;
        padding: 0.5rem 1.2rem;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background: rgba(176, 38, 255, 0.25);
        border-color: rgba(176, 38, 255, 0.7);
        box-shadow: 0 0 20px rgba(176, 38, 255, 0.3);
        transform: translateY(-1px);
    }

    /* Status badges */
    .status-pill {
        display: inline-block;
        padding: 0.25rem 0.7rem;
        border-radius: 999px;
        font-size: 0.7rem;
        font-weight: 500;
        font-family: 'JetBrains Mono', monospace;
        letter-spacing: 0.05em;
    }
    .status-live {
        background: rgba(0, 230, 118, 0.12);
        color: #00E676;
        border: 1px solid rgba(0, 230, 118, 0.3);
    }
    .status-live::before {
        content: "●";
        margin-right: 0.4rem;
        animation: blink 1.5s ease-in-out infinite;
    }
    @keyframes blink {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.3; }
    }

    /* Dataframes */
    .stDataFrame, [data-testid="stDataFrameResizable"] {
        background: rgba(255, 255, 255, 0.02) !important;
        border-radius: 12px;
        border: 1px solid rgba(176, 38, 255, 0.15);
    }

    /* Footer */
    .footer-classify {
        text-align: center;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.7rem;
        color: rgba(240, 240, 255, 0.3);
        letter-spacing: 0.18em;
        text-transform: uppercase;
        margin-top: 3rem;
        padding-top: 1.5rem;
        border-top: 1px solid rgba(176, 38, 255, 0.1);
    }

    /* Plotly modebar mauve */
    .modebar { background: transparent !important; }
    .modebar-btn svg path { fill: rgba(176, 38, 255, 0.6) !important; }
    </style>
    """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# 2. DATA LAYER — Lecture Excel + cache + détection mtime
# ═══════════════════════════════════════════════════════════════════════════════

def get_file_mtime(path: Path) -> float:
    """Retourne le mtime du fichier, ou -1 si introuvable."""
    try:
        return os.path.getmtime(path)
    except OSError:
        return -1.0


@st.cache_data(show_spinner=False)
def load_workbook_data(path_str: str, mtime: float) -> dict:
    """
    Charge l'ensemble des données nécessaires depuis l'Excel.
    Le paramètre mtime est utilisé comme clé de cache : dès que le fichier
    est modifié, le cache est invalidé automatiquement.

    Retourne un dict structuré, ou un dict {"error": ...} en cas d'échec.
    """
    try:
        wb = load_workbook(path_str, data_only=True, read_only=True)
    except FileNotFoundError:
        return {"error": f"Fichier introuvable : {path_str}"}
    except PermissionError:
        return {"error": "Le fichier est ouvert par un autre utilisateur (verrou Excel). Fermez-le et réessayez."}
    except Exception as e:
        return {"error": f"Erreur de lecture : {type(e).__name__} — {e}"}

    data = {"error": None}

    # ── 6_CPC_Normalisé ──
    try:
        ws = wb["6_CPC_Normalisé"]
        # Mapping ligne par ligne (cf. inspection)
        # cols : D=N-1, E=N estimé, F=N+1
        def cell(r, c):
            v = ws.cell(row=r, column=c).value
            return float(v) if isinstance(v, (int, float)) else 0.0

        data["cpc"] = {
            "ventes_711": {"nm1": cell(7, 4), "n": cell(7, 5), "np1": cell(7, 6)},
            "ventes_712": {"nm1": cell(8, 4), "n": cell(8, 5), "np1": cell(8, 6)},
            "achats_611": {"nm1": cell(16, 4), "n": cell(16, 5), "np1": cell(16, 6)},
            "achats_612": {"nm1": cell(17, 4), "n": cell(17, 5), "np1": cell(17, 6)},
            "charges_externes": {"nm1": cell(18, 4), "n": cell(18, 5), "np1": cell(18, 6)},
            "personnel": {"nm1": cell(20, 4), "n": cell(20, 5), "np1": cell(20, 6)},
            "dotations": {"nm1": cell(22, 4), "n": cell(22, 5), "np1": cell(22, 6)},
            "total_charges_expl": {"nm1": cell(23, 4), "n": cell(23, 5), "np1": cell(23, 6)},
            "rex": {"nm1": cell(25, 4), "n": cell(25, 5), "np1": cell(25, 6)},
            "resultat_financier": {"nm1": cell(41, 4), "n": cell(41, 5), "np1": cell(41, 6)},
            "is_du": {"nm1": cell(64, 4), "n": cell(64, 5), "np1": cell(64, 6)},
            "rn": {"nm1": cell(66, 4), "n": cell(66, 5), "np1": cell(66, 6)},
            "marge_brute": {"nm1": cell(69, 4), "n": cell(69, 5), "np1": cell(69, 6)},
            "marge_brute_pct": {"nm1": cell(70, 4), "n": cell(70, 5), "np1": cell(70, 6)},
            "ebitda": {"nm1": cell(71, 4), "n": cell(71, 5), "np1": cell(71, 6)},
            "ebitda_pct": {"nm1": cell(72, 4), "n": cell(72, 5), "np1": cell(72, 6)},
            "marge_nette_pct": {"nm1": cell(73, 4), "n": cell(73, 5), "np1": cell(73, 6)},
        }
        # CA total = 711 + 712
        for period in ("nm1", "n", "np1"):
            data["cpc"][f"ca_total_{period}"] = data["cpc"]["ventes_711"][period] + data["cpc"]["ventes_712"][period]
    except Exception as e:
        return {"error": f"Erreur sur 6_CPC_Normalisé : {e}"}

    # ── 2_Dashboard ──
    try:
        ws = wb["2_Dashboard"]
        def c(r, col):
            v = ws.cell(row=r, column=col).value
            return float(v) if isinstance(v, (int, float)) else 0.0

        data["dashboard"] = {
            "ca": c(7, 2),
            "ebitda": c(7, 4),
            "marge_ebitda": c(7, 6),
            "rn": c(11, 2),
            "bfr": c(11, 4),
            "gearing": c(11, 6),
            "dso": c(17, 3),
            "dso_cible": c(17, 4),
            "gearing_cible": c(18, 4),
            "marge_ebitda_cible": c(19, 4),
            "croissance_ca": c(20, 3),
            "croissance_ca_cible": c(20, 4),
            "var_bfr": c(21, 3),
        }
        # Waterfall data
        wf = []
        for r in range(26, 34):
            label = ws.cell(row=r, column=2).value
            kind = ws.cell(row=r, column=3).value
            val = ws.cell(row=r, column=4).value
            if label and isinstance(val, (int, float)):
                wf.append({"label": str(label), "type": str(kind or ""), "value": float(val)})
        data["waterfall"] = wf
    except Exception as e:
        return {"error": f"Erreur sur 2_Dashboard : {e}"}

    # ── 7_Bilan_Synthèse ──
    try:
        ws = wb["7_Bilan_Synthèse"]
        def c(r, col):
            v = ws.cell(row=r, column=col).value
            return float(v) if isinstance(v, (int, float)) else 0.0
        data["bilan"] = {
            "total_actif_np1": c(26, 6),
            "bfr_np1": c(60, 6),
            "dso_np1": c(61, 6),
            "dpo_np1": c(62, 6),
            "dio_np1": c(63, 6),
            "gearing_np1": c(64, 6),
            "autonomie_np1": c(65, 6),
        }
    except Exception as e:
        return {"error": f"Erreur sur 7_Bilan_Synthèse : {e}"}

    # ── 10_Heatmap_Ecarts ──
    try:
        ws = wb["10_Heatmap_Ecarts"]
        # En-têtes en R5 col C..H (centres de coût)
        centers = []
        for col in range(3, 10):
            v = ws.cell(row=5, column=col).value
            if v and v != "Total" and isinstance(v, str):
                centers.append(v)
        # Comptes en col B de R6 à R19
        rows = []
        for r in range(6, 20):
            code = ws.cell(row=r, column=2).value
            if code and code != "TOTAL" and isinstance(code, str):
                row_vals = {}
                for j, cc in enumerate(centers):
                    v = ws.cell(row=r, column=3 + j).value
                    row_vals[cc] = float(v) if isinstance(v, (int, float)) else 0.0
                rows.append({"compte": str(code), "values": row_vals})
        data["heatmap"] = {"centers": centers, "rows": rows}
    except Exception as e:
        return {"error": f"Erreur sur 10_Heatmap_Ecarts : {e}"}

    # ── 5_Personnel_CNSS — détail par site/centre pour filtres ──
    try:
        ws = wb["5_Personnel_CNSS"]
        rows = []
        # Données de R7 à R22 (16 équipes)
        for r in range(7, 23):
            entity = ws.cell(row=r, column=2).value
            site = ws.cell(row=r, column=3).value
            cc = ws.cell(row=r, column=4).value
            effectif = ws.cell(row=r, column=5).value
            cout_emp = ws.cell(row=r, column=13).value
            if entity and site and cc and isinstance(effectif, (int, float)):
                rows.append({
                    "entity": str(entity),
                    "site": str(site),
                    "centre": str(cc),
                    "effectif": int(effectif),
                    "cout_employeur": float(cout_emp or 0),
                })
        data["personnel"] = rows
    except Exception as e:
        return {"error": f"Erreur sur 5_Personnel_CNSS : {e}"}

    # ── 3_Drivers_CA — pour filtre dynamique CA par site ──
    try:
        ws = wb["3_Drivers_CA"]
        # Bloc 3 (CA résultant) lignes 41..52
        rows = []
        for r in range(41, 53):
            entity = ws.cell(row=r, column=2).value
            site = ws.cell(row=r, column=3).value
            line = ws.cell(row=r, column=4).value
            ca_total = ws.cell(row=r, column=18).value
            if entity and site and isinstance(ca_total, (int, float)):
                rows.append({
                    "entity": str(entity),
                    "site": str(site),
                    "ligne": str(line or ""),
                    "ca_np1": float(ca_total),
                })
        data["drivers_ca"] = rows
    except Exception as e:
        return {"error": f"Erreur sur 3_Drivers_CA : {e}"}

    wb.close()
    return data


# ═══════════════════════════════════════════════════════════════════════════════
# 3. UTILITIES — Formatage MAD & helpers KPI
# ═══════════════════════════════════════════════════════════════════════════════

def fmt_mad(value: float, decimals: int = 0, suffix: str = "MAD") -> str:
    """Formate un nombre en style MAD : 309 382 839 MAD (espaces fins comme séparateurs)."""
    if value is None or pd.isna(value):
        return "—"
    abs_v = abs(value)
    if abs_v >= 1e9:
        return f"{value / 1e9:,.2f}".replace(",", " ").replace(".", ",") + f" Md {suffix}"
    if abs_v >= 1e6:
        return f"{value / 1e6:,.2f}".replace(",", " ").replace(".", ",") + f" M{suffix}"
    if abs_v >= 1e3:
        return f"{value / 1e3:,.0f}".replace(",", " ") + f" k{suffix}"
    fmt = f"{{:,.{decimals}f}}".format(value).replace(",", " ")
    return f"{fmt} {suffix}"


def fmt_full_mad(value: float) -> str:
    """Format complet sans abrégé : 309 382 839 MAD."""
    if value is None or pd.isna(value):
        return "—"
    return f"{value:,.0f}".replace(",", " ") + " MAD"


def fmt_pct(value: float, decimals: int = 1) -> str:
    if value is None or pd.isna(value):
        return "—"
    return f"{value * 100:,.{decimals}f}".replace(".", ",") + " %"


def trend_pill(current: float, previous: float, kpi_kind: str = "revenue") -> str:
    """
    Génère le HTML d'un badge de tendance.
    kpi_kind = "revenue" : positif = good ; "cost" : positif = bad
    """
    if previous == 0 or previous is None:
        return ""
    delta = (current - previous) / abs(previous)
    arrow = "▲" if delta > 0 else ("▼" if delta < 0 else "—")
    if kpi_kind == "revenue":
        cls = "trend-up" if delta > 0.005 else ("trend-down" if delta < -0.005 else "trend-neutral")
    else:
        cls = "trend-down" if delta > 0.005 else ("trend-up" if delta < -0.005 else "trend-neutral")
    pct = f"{delta * 100:+.1f}".replace(".", ",") + " %"
    return f'<span class="kpi-trend {cls}">{arrow} {pct}</span>'


def kpi_card_html(label: str, value: str, value_class: str = "",
                  trend_html: str = "", sub: str = "") -> str:
    """Génère le HTML d'une carte KPI."""
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value {value_class}">{value}</div>
        {trend_html}
        {f'<div class="kpi-sub">{sub}</div>' if sub else ''}
    </div>
    """


# ═══════════════════════════════════════════════════════════════════════════════
# 4. CHART BUILDERS — Plotly avec thème dark/mauve
# ═══════════════════════════════════════════════════════════════════════════════

PLOTLY_LAYOUT_BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Manrope, sans-serif", color="#F0F0FF", size=12),
    margin=dict(l=20, r=20, t=40, b=20),
    xaxis=dict(gridcolor="rgba(176, 38, 255, 0.1)", linecolor="rgba(176, 38, 255, 0.2)",
               zerolinecolor="rgba(176, 38, 255, 0.2)"),
    yaxis=dict(gridcolor="rgba(176, 38, 255, 0.1)", linecolor="rgba(176, 38, 255, 0.2)",
               zerolinecolor="rgba(176, 38, 255, 0.2)"),
)


def build_gauge(value_pct: float, target_pct: float, title: str = "MARGE EBITDA") -> go.Figure:
    """Jauge néon mauve — affiche le taux EBITDA vs cible."""
    val = value_pct * 100
    tgt = target_pct * 100
    max_range = max(val, tgt) * 1.5 if max(val, tgt) > 0 else 25

    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=val,
        number={"suffix": " %", "font": {"size": 38, "color": "#FFFFFF",
                                         "family": "Space Grotesk"}},
        delta={"reference": tgt, "relative": False,
               "valueformat": ".1f", "suffix": " pts vs cible",
               "increasing": {"color": GREEN_NEON},
               "decreasing": {"color": RED_NEON},
               "font": {"size": 13, "family": "JetBrains Mono"}},
        domain={"x": [0, 1], "y": [0, 1]},
        gauge={
            "axis": {"range": [0, max_range], "tickwidth": 1,
                     "tickcolor": "rgba(176, 38, 255, 0.4)",
                     "tickfont": {"color": "rgba(240, 240, 255, 0.55)",
                                  "size": 11, "family": "JetBrains Mono"}},
            "bar": {"color": PURPLE_NEON, "thickness": 0.32},
            "bgcolor": "rgba(255, 255, 255, 0.03)",
            "borderwidth": 1,
            "bordercolor": "rgba(176, 38, 255, 0.25)",
            "steps": [
                {"range": [0, tgt * 0.7], "color": "rgba(255, 61, 113, 0.18)"},
                {"range": [tgt * 0.7, tgt * 0.95], "color": "rgba(255, 179, 0, 0.18)"},
                {"range": [tgt * 0.95, max_range], "color": "rgba(0, 230, 118, 0.15)"},
            ],
            "threshold": {
                "line": {"color": GOLD_ACCENT, "width": 3},
                "thickness": 0.85,
                "value": tgt,
            },
        },
        title={"text": f"<span style='font-size:0.75rem; letter-spacing:0.15em; "
                       f"color:rgba(240,240,255,0.55)'>{title}</span><br>"
                       f"<span style='font-size:0.65rem; color:rgba(240,240,255,0.4)'>"
                       f"CIBLE : {tgt:.1f} % · OR ─ ─</span>",
               "font": {"family": "Manrope"}},
    ))
    fig.update_layout(
        **{k: v for k, v in PLOTLY_LAYOUT_BASE.items() if k not in ("xaxis", "yaxis")},
        height=320,
    )
    return fig


def build_waterfall_cpc(cpc: dict) -> go.Figure:
    """Cascade CA → Marge brute → EBITDA → REX → RN (N+1)."""
    ca = cpc["ca_total_np1"]
    achats = -(cpc["achats_611"]["np1"] + cpc["achats_612"]["np1"])
    autres_charges = -(cpc["charges_externes"]["np1"] + cpc["personnel"]["np1"]
                       + cpc["dotations"]["np1"] + 980000 + 480000)  # 616 + 618
    autres_charges_to_rex = -(cpc["charges_externes"]["np1"] + cpc["personnel"]["np1"]
                              + cpc["dotations"]["np1"])
    autres_produits = (1800000 + 500000 + 950000 + 250000)  # 713+714+718+719
    impots_taxes = -(980000 + 480000)
    res_fin = cpc["resultat_financier"]["np1"]
    is_du = -cpc["is_du"]["np1"]

    labels = ["CA HT N+1", "Achats consommés", "Marge brute",
              "Autres charges expl.", "EBITDA",
              "Dotations", "REX",
              "Résultat financier", "Impôt (IS)", "Résultat Net"]
    measures = ["absolute", "relative", "total",
                "relative", "total",
                "relative", "total",
                "relative", "relative", "total"]
    values = [
        ca,
        achats,
        0,  # total → calcul auto
        autres_charges_to_rex - (-cpc["dotations"]["np1"]) + impots_taxes,  # autres charges hors dotations + impôts/taxes + 618
        0,
        -cpc["dotations"]["np1"],
        0,
        res_fin,
        is_du,
        0,
    ]
    # Recompose proprement avec les vraies valeurs CGNC :
    # CA + autres_produits_expl - achats - autres_charges_expl - dotations = REX
    # REX + dotations = EBITDA  (donc on inverse l'ordre : MB → EBITDA → REX → RN)
    # Définition simplifiée :
    mb = ca - (cpc["achats_611"]["np1"] + cpc["achats_612"]["np1"])
    ebitda = cpc["ebitda"]["np1"]
    rex = cpc["rex"]["np1"]
    rn = cpc["rn"]["np1"]

    # Cascade CA → MB → EBITDA → REX → RN
    labels_clean = ["CA HT N+1", "Achats consommés (611+612)", "MARGE BRUTE",
                    "Autres ch. expl. (hors dotations)", "EBITDA",
                    "Dotations (619)", "REX",
                    "Résultat financier + IS", "RÉSULTAT NET"]
    measures_clean = ["absolute", "relative", "total",
                      "relative", "total",
                      "relative", "total",
                      "relative", "total"]
    autres_ch_hors_dot = mb - ebitda  # négatif
    impact_fin_is = res_fin + is_du   # négatif
    values_clean = [
        ca,
        -(cpc["achats_611"]["np1"] + cpc["achats_612"]["np1"]),
        0,
        -autres_ch_hors_dot if autres_ch_hors_dot < 0 else -autres_ch_hors_dot,
        0,
        -cpc["dotations"]["np1"],
        0,
        impact_fin_is,
        0,
    ]
    # Correction : autres_ch_hors_dot = mb - ebitda. Si mb=99M et ebitda=34M, écart = 65M positif → on doit retirer 65M
    autres_ch_hors_dot = -(mb - ebitda)
    values_clean[3] = autres_ch_hors_dot

    fig = go.Figure(go.Waterfall(
        name="P&L",
        orientation="v",
        measure=measures_clean,
        x=labels_clean,
        y=values_clean,
        text=[fmt_mad(v) if v != 0 else "" for v in values_clean],
        textposition="outside",
        textfont=dict(family="JetBrains Mono", size=10, color="#F0F0FF"),
        connector={"line": {"color": "rgba(176, 38, 255, 0.4)", "width": 1, "dash": "dot"}},
        increasing={"marker": {"color": PURPLE_NEON,
                               "line": {"color": "#FFFFFF", "width": 0.5}}},
        decreasing={"marker": {"color": RED_NEON,
                               "line": {"color": "#FFFFFF", "width": 0.5}}},
        totals={"marker": {"color": GOLD_ACCENT,
                           "line": {"color": "#FFFFFF", "width": 0.5}}},
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT_BASE,
        height=460,
        showlegend=False,
        title=dict(
            text="<b>CASCADE P&amp;L</b> · CA → Marge → EBITDA → REX → RN",
            font=dict(family="Space Grotesk", size=14, color="#F0F0FF"),
            x=0.02,
        ),
    )
    fig.update_xaxes(tickangle=-25, tickfont=dict(size=10))
    fig.update_yaxes(tickformat=",.0f", title="MAD")
    return fig


def build_heatmap(heatmap_data: dict, centers_filter: list = None) -> go.Figure:
    """Heatmap des écarts budgétaires par compte × centre de coût."""
    centers = heatmap_data["centers"]
    if centers_filter:
        keep_idx = [i for i, c in enumerate(centers) if c in centers_filter]
        centers = [centers[i] for i in keep_idx]
    else:
        keep_idx = list(range(len(centers)))

    if not centers:
        return go.Figure().update_layout(**PLOTLY_LAYOUT_BASE,
                                         annotations=[dict(text="Aucun centre sélectionné", x=0.5, y=0.5,
                                                           showarrow=False, font=dict(color="#888"))])

    rows = heatmap_data["rows"]
    z = []
    accounts = []
    for row in rows:
        line = [row["values"].get(c, 0) for c in centers]
        if any(v != 0 for v in line):
            z.append(line)
            accounts.append(row["compte"])

    if not z:
        return go.Figure().update_layout(**PLOTLY_LAYOUT_BASE,
                                         annotations=[dict(text="Aucun écart à afficher", x=0.5, y=0.5,
                                                           showarrow=False, font=dict(color="#888"))])

    text_matrix = [[fmt_mad(v) if v != 0 else "" for v in row] for row in z]

    fig = go.Figure(go.Heatmap(
        z=z,
        x=centers,
        y=accounts,
        text=text_matrix,
        texttemplate="%{text}",
        textfont=dict(family="JetBrains Mono", size=10, color="#FFFFFF"),
        colorscale=[
            [0.0, "rgba(0, 230, 118, 0.6)"],     # vert (économie)
            [0.5, "rgba(20, 0, 40, 0.3)"],       # neutre sombre
            [1.0, "rgba(255, 61, 113, 0.85)"],   # rouge (dérapage)
        ],
        zmid=0,
        showscale=True,
        colorbar=dict(
            title=dict(text="Écart MAD", font=dict(color="#F0F0FF", family="Manrope", size=11)),
            tickfont=dict(color="#F0F0FF", family="JetBrains Mono", size=10),
            outlinecolor="rgba(176, 38, 255, 0.3)",
            bgcolor="rgba(0,0,0,0)",
            tickformat=",.0f",
        ),
        hovertemplate="<b>%{y}</b> × <b>%{x}</b><br>Écart : %{z:,.0f} MAD<extra></extra>",
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT_BASE,
        height=460,
        title=dict(
            text="<b>HEATMAP</b> · Écarts Budget N+1 vs Réel N estimé · Compte × Centre de coût",
            font=dict(family="Space Grotesk", size=14, color="#F0F0FF"),
            x=0.02,
        ),
    )
    fig.update_xaxes(side="top", tickfont=dict(family="Manrope", size=11, color="#F0F0FF"))
    fig.update_yaxes(tickfont=dict(family="JetBrains Mono", size=11, color="#F0F0FF"),
                     autorange="reversed")
    return fig


# ═══════════════════════════════════════════════════════════════════════════════
# 5. KPI ENGINE — Calcul des 10 KPIs (avec filtres)
# ═══════════════════════════════════════════════════════════════════════════════

def compute_kpis(data: dict, filter_sites: list, filter_centers: list,
                 objectif_ca: float = 320_000_000) -> dict:
    """
    Calcule les 10 KPIs du DAF avec filtres sites + centres de coût.

    Méthodo : Les KPIs financiers (CPC, Bilan) sont au niveau Groupe consolidé
    et NON ventilés par site/CC dans le modèle (le modèle Excel n'a pas
    de CPC par site). On applique les filtres uniquement sur les composantes
    qui sont ventilées : personnel (5_Personnel) et CA drivers (3_Drivers_CA).
    Le ratio Masse Salariale/CA et le poids CA sont donc impactés ; les autres
    KPIs restent au niveau consolidé (badge "Groupe" affiché en sous-titre).
    """
    cpc = data["cpc"]
    dash = data["dashboard"]
    bilan = data["bilan"]

    # Périmètre filtré : par défaut tous les sites/centres
    pers_rows = data["personnel"]
    if filter_sites:
        pers_rows = [r for r in pers_rows if r["site"] in filter_sites]
    if filter_centers:
        pers_rows = [r for r in pers_rows if r["centre"] in filter_centers]
    masse_sal_filtered = sum(r["cout_employeur"] for r in pers_rows)

    drivers_rows = data["drivers_ca"]
    if filter_sites:
        drivers_rows = [r for r in drivers_rows if r["site"] in filter_sites]
    ca_filtered = sum(r["ca_np1"] for r in drivers_rows)

    # Si pas de filtre, on prend les valeurs Groupe officielles
    is_full_perimeter = (not filter_sites) and (not filter_centers)

    ca_np1 = cpc["ca_total_np1"] if is_full_perimeter else ca_filtered
    ca_n = cpc["ca_total_n"]
    ca_nm1 = cpc["ca_total_nm1"]

    # Masse salariale : si filtre site/centre, on prend la valeur filtrée
    masse_sal = cpc["personnel"]["np1"] if is_full_perimeter else masse_sal_filtered

    # Ratio masse salariale / CA Groupe (ou CA filtré si filtre actif)
    ca_for_ratio = ca_np1 if ca_np1 > 0 else cpc["ca_total_np1"]

    # Point mort : Charges fixes / Taux marge sur coûts variables
    # Approche simplifiée : Charges fixes ≈ personnel + charges externes + dotations + autres
    # Coûts variables ≈ achats consommés (611+612)
    charges_fixes = (cpc["personnel"]["np1"] + cpc["charges_externes"]["np1"]
                     + cpc["dotations"]["np1"] + 980000 + 480000)
    couts_variables = cpc["achats_611"]["np1"] + cpc["achats_612"]["np1"]
    ca_groupe = cpc["ca_total_np1"]
    taux_marge_cv = (ca_groupe - couts_variables) / ca_groupe if ca_groupe > 0 else 0
    point_mort = charges_fixes / taux_marge_cv if taux_marge_cv > 0 else 0

    return {
        # 1. CA Global HT
        "ca_np1": ca_np1,
        "ca_n": ca_n,
        "ca_nm1": ca_nm1,
        # 2. EBITDA
        "ebitda_np1": cpc["ebitda"]["np1"],
        "ebitda_n": cpc["ebitda"]["n"],
        "ebitda_nm1": cpc["ebitda"]["nm1"],
        "marge_ebitda": cpc["ebitda_pct"]["np1"],
        "marge_ebitda_cible": dash["marge_ebitda_cible"],
        # 3. Résultat Net
        "rn_np1": cpc["rn"]["np1"],
        "rn_n": cpc["rn"]["n"],
        "rn_nm1": cpc["rn"]["nm1"],
        # 4. Marge brute
        "marge_brute_pct": cpc["marge_brute_pct"]["np1"],
        "marge_brute_pct_n": cpc["marge_brute_pct"]["n"],
        # 5. Poids masse salariale / CA
        "ratio_personnel": masse_sal / ca_for_ratio if ca_for_ratio > 0 else 0,
        "ratio_personnel_n": cpc["personnel"]["n"] / cpc["ca_total_n"] if cpc["ca_total_n"] > 0 else 0,
        "masse_salariale": masse_sal,
        # 6. Poids charges externes / CA
        "ratio_charges_ext": cpc["charges_externes"]["np1"] / ca_groupe if ca_groupe > 0 else 0,
        "ratio_charges_ext_n": cpc["charges_externes"]["n"] / cpc["ca_total_n"] if cpc["ca_total_n"] > 0 else 0,
        # 7. DSO
        "dso": bilan["dso_np1"],
        "dso_cible": dash["dso_cible"],
        # 8. Gearing
        "gearing": bilan["gearing_np1"],
        "gearing_cible": dash["gearing_cible"],
        # 9. Point mort
        "point_mort": point_mort,
        "ratio_point_mort": point_mort / ca_groupe if ca_groupe > 0 else 0,
        # 10. Atteinte objectif
        "objectif_ca": objectif_ca,
        "atteinte_objectif": ca_np1 / objectif_ca if objectif_ca > 0 else 0,
        # Méta
        "is_full_perimeter": is_full_perimeter,
        "perimeter_label": "GROUPE CONSOLIDÉ" if is_full_perimeter else f"{len(filter_sites or [])} SITE(S) · {len(filter_centers or [])} CC",
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 6. RENDER LAYERS
# ═══════════════════════════════════════════════════════════════════════════════

def render_header(perimeter_label: str, mtime: float):
    """Bandeau d'en-tête avec titre, périmètre et statut live."""
    from datetime import datetime
    ts = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y · %H:%M") if mtime > 0 else "—"

    col1, col2 = st.columns([5, 2])
    with col1:
        st.markdown(f"""
        <h1 class="hero-title">COMEX · Budget N+1</h1>
        <p class="hero-subtitle">Direction Financière Groupe · Industrie Maroc · CGNC</p>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div style="text-align:right; padding-top:1rem;">
            <span class="status-pill status-live">LIVE · {ts}</span>
            <div style="font-family:'JetBrains Mono'; font-size:0.7rem;
                        color:rgba(240,240,255,0.45); margin-top:0.5rem;
                        letter-spacing:0.1em;">
                PÉRIMÈTRE · {perimeter_label}
            </div>
        </div>
        """, unsafe_allow_html=True)
    st.markdown('<div class="hero-divider"></div>', unsafe_allow_html=True)


def render_kpi_row_1(k: dict):
    """Première ligne : 5 KPIs principaux (CA, EBITDA, RN, Marge brute, Atteinte obj)."""
    st.markdown('<div class="section-title">Indicateurs Financiers Clés</div>',
                unsafe_allow_html=True)

    cols = st.columns(5, gap="medium")
    cards = [
        # 1. CA Global
        kpi_card_html(
            "Chiffre d'affaires HT · N+1",
            fmt_full_mad(k["ca_np1"]).replace(" MAD", '<span class="kpi-unit">MAD</span>'),
            value_class="accent",
            trend_html=trend_pill(k["ca_np1"], k["ca_n"], "revenue"),
            sub=f"vs N estimé · {fmt_mad(k['ca_n'])}",
        ),
        # 2. EBITDA
        kpi_card_html(
            "EBITDA · N+1",
            fmt_full_mad(k["ebitda_np1"]).replace(" MAD", '<span class="kpi-unit">MAD</span>'),
            value_class="gold",
            trend_html=trend_pill(k["ebitda_np1"], k["ebitda_nm1"], "revenue"),
            sub=f"Marge {fmt_pct(k['marge_ebitda'])} · cible {fmt_pct(k['marge_ebitda_cible'])}",
        ),
        # 3. Résultat Net
        kpi_card_html(
            "Résultat Net · N+1",
            fmt_full_mad(k["rn_np1"]).replace(" MAD", '<span class="kpi-unit">MAD</span>'),
            value_class="accent",
            trend_html=trend_pill(k["rn_np1"], k["rn_nm1"], "revenue"),
            sub=f"vs N-1 · {fmt_mad(k['rn_nm1'])}",
        ),
        # 4. Marge brute
        kpi_card_html(
            "Taux de Marge Brute",
            f"{k['marge_brute_pct'] * 100:.1f}".replace(".", ",") + '<span class="kpi-unit">%</span>',
            value_class="gold",
            trend_html=trend_pill(k["marge_brute_pct"], k["marge_brute_pct_n"], "revenue"),
            sub="Ventes − Achats consommés / CA",
        ),
        # 10. Atteinte objectif
        kpi_card_html(
            "Atteinte Objectif CA",
            f"{k['atteinte_objectif'] * 100:.1f}".replace(".", ",") + '<span class="kpi-unit">%</span>',
            value_class="accent" if k["atteinte_objectif"] >= 1 else "",
            trend_html=(
                f'<span class="kpi-trend trend-up">▲ Objectif atteint</span>'
                if k["atteinte_objectif"] >= 1 else
                f'<span class="kpi-trend trend-down">▼ Sous l\'objectif</span>'
            ),
            sub=f"Cible · {fmt_mad(k['objectif_ca'])}",
        ),
    ]
    st.markdown('<div class="kpi-row" style="display:contents">', unsafe_allow_html=True)
    for col, card in zip(cols, cards):
        col.markdown(card, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)


def render_kpi_row_2(k: dict):
    """Deuxième ligne : 5 KPIs structure (ratios, DSO, Gearing, point mort)."""
    st.markdown('<div class="section-title">Structure & Ratios de Pilotage</div>',
                unsafe_allow_html=True)

    cols = st.columns(5, gap="medium")
    cards = [
        # 5. Poids Personnel / CA
        kpi_card_html(
            "Masse Salariale / CA",
            f"{k['ratio_personnel'] * 100:.1f}".replace(".", ",") + '<span class="kpi-unit">%</span>',
            trend_html=trend_pill(k["ratio_personnel"], k["ratio_personnel_n"], "cost"),
            sub=f"Compte 617 · {fmt_mad(k['masse_salariale'])}",
        ),
        # 6. Poids Charges externes / CA
        kpi_card_html(
            "Charges Externes / CA",
            f"{k['ratio_charges_ext'] * 100:.1f}".replace(".", ",") + '<span class="kpi-unit">%</span>',
            trend_html=trend_pill(k["ratio_charges_ext"], k["ratio_charges_ext_n"], "cost"),
            sub="Comptes 613-614 · ZBB",
        ),
        # 7. DSO
        kpi_card_html(
            "DSO · Délai Recouvrement",
            f"{k['dso']:.0f}".replace(",", " ") + '<span class="kpi-unit">jours</span>',
            value_class="accent" if k["dso"] <= k["dso_cible"] else "",
            trend_html=(
                f'<span class="kpi-trend trend-up">▲ Sous cible</span>'
                if k["dso"] <= k["dso_cible"] else
                f'<span class="kpi-trend trend-down">▼ Dépassement</span>'
            ),
            sub=f"Cible Loi 69-21 · {k['dso_cible']:.0f} j",
        ),
        # 8. Gearing
        kpi_card_html(
            "Gearing · Endettement",
            f"{k['gearing'] * 100:.1f}".replace(".", ",") + '<span class="kpi-unit">%</span>',
            value_class="accent" if k["gearing"] <= k["gearing_cible"] else "",
            trend_html=(
                f'<span class="kpi-trend trend-up">▲ Levier maîtrisé</span>'
                if k["gearing"] <= k["gearing_cible"] else
                f'<span class="kpi-trend trend-down">▼ Sur-endettement</span>'
            ),
            sub=f"Dettes fin. / Capitaux propres · cible {k['gearing_cible'] * 100:.0f} %",
        ),
        # 9. Point mort
        kpi_card_html(
            "Point Mort · Seuil Rentabilité",
            fmt_mad(k["point_mort"]).replace(" MAD", '<span class="kpi-unit">MAD</span>')
                                    .replace(" MMAD", '<span class="kpi-unit">MMAD</span>'),
            trend_html=(
                f'<span class="kpi-trend trend-up">▲ Au-dessus du seuil</span>'
                if k["ca_np1"] >= k["point_mort"] else
                f'<span class="kpi-trend trend-down">▼ Sous le seuil</span>'
            ),
            sub=f"{fmt_pct(k['ratio_point_mort'])} du CA",
        ),
    ]
    st.markdown('<div class="kpi-row" style="display:contents">', unsafe_allow_html=True)
    for col, card in zip(cols, cards):
        col.markdown(card, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)


def render_visuals(data: dict, k: dict, filter_centers: list):
    """Rangée jauge + waterfall, puis heatmap."""
    st.markdown('<div class="section-title">Vues Stratégiques</div>',
                unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1.6], gap="medium")

    with col_left:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        gauge_fig = build_gauge(k["marge_ebitda"], k["marge_ebitda_cible"])
        st.plotly_chart(gauge_fig, use_container_width=True,
                        config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    with col_right:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        wf_fig = build_waterfall_cpc(data["cpc"])
        st.plotly_chart(wf_fig, use_container_width=True,
                        config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    # Heatmap pleine largeur
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    hm_fig = build_heatmap(data["heatmap"], filter_centers)
    st.plotly_chart(hm_fig, use_container_width=True,
                    config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)


def render_sidebar(data: dict) -> tuple:
    """Sidebar avec filtres et bouton refresh. Retourne (sites, centers, objectif)."""
    with st.sidebar:
        st.markdown("""
        <div class="sidebar-brand">◆ COMEX · DAF</div>
        <div class="sidebar-tagline">Pilotage Budgétaire N+1</div>
        """, unsafe_allow_html=True)

        # Filtres
        st.markdown("**Périmètre d'analyse**")

        # Sites disponibles
        sites_avail = sorted({r["site"] for r in data["personnel"]})
        sites = st.multiselect(
            "Sites industriels",
            options=sites_avail,
            default=[],
            placeholder="Tous les sites (Casa + Tanger)",
            key="filter_sites",
        )

        # Centres de coût
        centers_avail = sorted({r["centre"] for r in data["personnel"]})
        centers = st.multiselect(
            "Centres de coût",
            options=centers_avail,
            default=[],
            placeholder="Tous les centres",
            key="filter_centers",
        )

        st.markdown("---")
        st.markdown("**Paramètres de pilotage**")

        objectif = st.number_input(
            "Objectif CA N+1 (MAD)",
            min_value=0,
            value=320_000_000,
            step=10_000_000,
            format="%d",
            help="Objectif fixé par le Conseil d'Administration",
        )

        st.markdown("---")

        # Bouton refresh
        if st.button("⟳ ACTUALISER LES DONNÉES", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        # Auto-refresh info
        st.markdown("""
        <div style="font-size:0.7rem; color:rgba(240,240,255,0.4);
                    margin-top:1rem; line-height:1.5;
                    font-family:'JetBrains Mono', monospace;">
            <div style="color:#00E676; margin-bottom:0.3rem;">● AUTO-REFRESH ACTIF</div>
            Détection automatique<br>des modifications<br>via os.path.getmtime
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div style="margin-top:2rem; padding-top:1rem;
                    border-top:1px solid rgba(176,38,255,0.15);
                    font-size:0.65rem; color:rgba(240,240,255,0.3);
                    letter-spacing:0.1em; text-transform:uppercase;">
            Référentiel CGNC<br>Loi 69-21 · Article 144 CGI<br>
            v1.0 · Confidentiel DAF
        </div>
        """, unsafe_allow_html=True)

    return sites, centers, objectif


def render_footer():
    st.markdown("""
    <div class="footer-classify">
        ◆ Document Confidentiel · Direction Financière · © 2026 ·
        Diffusion restreinte au COMEX et au Conseil d'Administration ◆
    </div>
    """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# 7. MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    inject_css()

    # Vérification fichier
    if not EXCEL_PATH.exists():
        st.error(f"Fichier introuvable : `{EXCEL_PATH}`. "
                 f"Placez le fichier Excel dans le même dossier que `app.py`.")
        st.stop()

    mtime = get_file_mtime(EXCEL_PATH)
    data = load_workbook_data(str(EXCEL_PATH), mtime)

    if data.get("error"):
        st.error(f"⚠ {data['error']}")
        st.info("Le dashboard se réactualisera automatiquement dès que le problème sera résolu.")
        st.stop()

    # Sidebar (avec filtres)
    sites, centers, objectif = render_sidebar(data)

    # KPIs avec filtres
    k = compute_kpis(data, sites, centers, objectif)

    # Render
    render_header(k["perimeter_label"], mtime)
    render_kpi_row_1(k)
    render_kpi_row_2(k)
    render_visuals(data, k, centers)
    render_footer()


if __name__ == "__main__":
    main()