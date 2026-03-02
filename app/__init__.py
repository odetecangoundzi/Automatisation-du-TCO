"""
styles.py — CSS centralisé pour l'interface Streamlit Export du TCO.

Fournit le thème clair/sombre avec des variables CSS.
Extrait de app.py pour améliorer la maintenabilité.
"""


def get_theme_vars(is_dark: bool) -> str:
    """Retourne les variables CSS :root pour le thème choisi."""
    if is_dark:
        return """
    :root {
        --bg:           #0e1117;
        --surface:      #1a1f2e;
        --surface-alt:  #232940;
        --border:       #2d3553;
        --border-hover: #4472C4;
        --text:         #e8ecf1;
        --text-muted:   #8a94a8;
        --accent:       #5b9bd5;
        --accent-deep:  #4472C4;
        --accent-dark:  #2F5496;
        --shadow:       rgba(0,0,0,0.35);
        --shadow-light: rgba(0,0,0,0.18);
        --card-bg:      linear-gradient(145deg, #1a1f2e 0%, #232940 100%);
        --metric-bg:    linear-gradient(145deg, #1a1f2e, #1e2538);
        --legend-bg:    #1a1f2e;
        --legend-border:#2d3553;
        --legend-text:  #8a94a8;
    }
    /* Force Streamlit dark overrides */
    .stApp, [data-testid="stAppViewContainer"] { background-color: #0e1117 !important; }
    .stMarkdown, .stMarkdown p, .stText { color: #e8ecf1 !important; }
    [data-testid="stSidebar"] { background-color: #151923 !important; }
    """
    else:
        return """
    :root {
        --bg:           #ffffff;
        --surface:      #ffffff;
        --surface-alt:  #f4f7fb;
        --border:       #d8e2ef;
        --border-hover: #4472C4;
        --text:         #1a1a2e;
        --text-muted:   #8899aa;
        --accent:       #4472C4;
        --accent-deep:  #2F5496;
        --accent-dark:  #1a3a6e;
        --shadow:       rgba(47,84,150,0.25);
        --shadow-light: rgba(0,0,0,0.04);
        --card-bg:      linear-gradient(145deg, #ffffff 0%, #f4f7fb 100%);
        --metric-bg:    linear-gradient(145deg, #ffffff, #f8fafd);
        --legend-bg:    #f9fafb;
        --legend-border:#eee;
        --legend-text:  #555;
    }
    """


def get_full_css(is_dark: bool, hide_sidebar: bool = False) -> str:
    """
    Retourne le bloc <style> complet pour l'application.

    Args:
        is_dark: True pour le thème sombre
        hide_sidebar: True pour masquer la sidebar (Step 0)
    """
    theme_vars = get_theme_vars(is_dark)
    sidebar_display = "none" if hide_sidebar else "block"

    css = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

{theme_vars}

/* ── Global ───────────────────────────────────────────── */
html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
.block-container {{ max-width: 1200px; padding-top: 2rem; }}

/* ── Title ────────────────────────────────────────────── */
.main-title {{
    text-align: center;
    font-size: 2.4rem;
    font-weight: 700;
    background: linear-gradient(135deg, var(--accent-dark) 0%, var(--accent) 60%, #6ca0dc 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: .2rem;
    letter-spacing: -.5px;
}}
.subtitle {{
    text-align: center;
    color: var(--text-muted);
    font-size: .95rem;
    margin-bottom: 1.2rem;
}}

/* ── Step headers ─────────────────────────────────────── */
.step-header {{
    background: linear-gradient(135deg, var(--accent-dark) 0%, var(--accent-deep) 50%, var(--accent) 100%);
    color: white;
    padding: 14px 24px;
    border-radius: 10px;
    margin: 1.5rem 0 .8rem;
    font-size: 1.05rem;
    font-weight: 600;
    letter-spacing: .3px;
    box-shadow: 0 4px 15px var(--shadow);
}}

/* ── Company cards ────────────────────────────────────── */
.company-card {{
    background: var(--card-bg);
    border: 1px solid var(--border);
    border-left: 4px solid var(--accent);
    border-radius: 10px;
    padding: 12px 18px;
    margin: 6px 0;
    transition: all .2s ease;
    box-shadow: 0 1px 4px var(--shadow-light);
    color: var(--text);
}}
.company-card:hover {{
    transform: translateY(-1px);
    box-shadow: 0 4px 12px var(--shadow);
    border-left-color: var(--accent-dark);
}}

/* ── Buttons ──────────────────────────────────────────── */
.stButton > button {{
    border-radius: 8px;
    font-weight: 600;
    padding: 0.5rem 1.2rem;
    transition: all .2s ease;
    border: none;
}}
.stButton > button:hover {{
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}}
.stButton > button[kind="primary"] {{
    background: linear-gradient(135deg, var(--accent-deep) 0%, var(--accent) 100%);
}}

/* ── Final Export Button ─────────────────────────────── */
[data-testid="stDownloadButton"] button {{
    background: linear-gradient(135deg, #28a745 0%, #1e7e34 100%) !important;
    color: white !important;
    border: none !important;
    padding: 0.7rem 2rem !important;
    font-size: 1rem !important;
    font-weight: 700 !important;
    box-shadow: 0 4px 15px rgba(40, 167, 69, 0.25) !important;
    transition: all .2s ease !important;
}}
[data-testid="stDownloadButton"] button:hover {{
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 20px rgba(40, 167, 69, 0.4) !important;
    background: linear-gradient(135deg, #218838 0%, #19692c 100%) !important;
}}

/* ── Sidebar ───────────────────────────────────────────── */
[data-testid="stSidebar"] {{
    border-right: 1px solid var(--border);
}}
[data-testid="stSidebar"] .stMarkdown h3 {{
    font-size: 0.9rem !important;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: var(--accent) !important;
    margin-bottom: 0.5rem !important;
    margin-top: 1.5rem !important;
}}
[data-testid="stSidebar"] .stButton button {{
    background-color: transparent !important;
    color: var(--text) !important;
    border: 1px solid transparent !important;
    text-align: left !important;
    display: flex !important;
    justify-content: flex-start !important;
    padding: 10px 15px !important;
    font-weight: 500 !important;
    font-size: 0.95rem !important;
    border-radius: 8px !important;
    margin-bottom: 4px !important;
    width: 100% !important;
    transition: all 0.2s ease !important;
}}
[data-testid="stSidebar"] .stButton button:hover {{
    background-color: var(--surface-alt) !important;
    border-color: var(--border) !important;
    color: var(--accent) !important;
    transform: translateX(4px) !important;
}}
[data-testid="stSidebar"] .stButton button[kind="primary"] {{
    background: linear-gradient(135deg, var(--accent-deep) 0%, var(--accent) 100%) !important;
    color: white !important;
    border: none !important;
    margin-top: 10px !important;
    justify-content: center !important;
}}
[data-testid="stSidebar"] .stButton button[kind="primary"]:hover {{
    transform: translateY(-2px) !important;
    box-shadow: 0 4px 12px var(--shadow) !important;
}}

/* ── Metrics ──────────────────────────────────────────── */
[data-testid="stMetric"] {{
    background: var(--metric-bg);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 12px 16px;
    box-shadow: 0 1px 6px var(--shadow-light);
}}
[data-testid="stMetricLabel"] {{ font-size: .85rem; font-weight: 600; color: var(--text-muted); }}
[data-testid="stMetricValue"] {{ font-size: 1.4rem; font-weight: 700; color: var(--accent-dark); }}

/* ── Alerts legend ────────────────────────────────────── */
.legend-box {{
    display: flex; gap: 1.2rem; flex-wrap: wrap;
    margin: .6rem 0; padding: 8px 14px;
    background: var(--legend-bg); border-radius: 8px;
    border: 1px solid var(--legend-border);
}}
.legend-item {{ display: flex; align-items: center; gap: .4rem; font-size: .82rem; color: var(--legend-text); }}
.color-dot   {{ width: 12px; height: 12px; border-radius: 50%; display: inline-block; box-shadow: inset 0 -1px 2px rgba(0,0,0,.1); }}

/* ── File uploader ────────────────────────────────────── */
[data-testid="stFileUploader"] {{
    border: 2px dashed var(--border);
    border-radius: 10px;
    padding: 12px;
    transition: border-color .2s, box-shadow .2s, background .2s, transform .2s;
}}
[data-testid="stFileUploader"]:hover {{ border-color: var(--accent); }}
/* Drag-over : utilisation du focus-within natif ou hover state */
[data-testid="stFileUploader"]:focus-within,
[data-testid="stFileUploader"]:has([data-testid="stFileUploadDropzone"]:hover) {{
    border-color: var(--accent) !important;
    border-style: solid !important;
    background: rgba(68, 114, 196, 0.07) !important;
    box-shadow: 0 0 0 4px rgba(68, 114, 196, 0.2) !important;
    transform: scale(1.012);
}}
[data-testid="stFileUploader"]:has([data-testid="stFileUploadDropzone"]:hover) section > div::after {{
    content: "⬇️ Relâchez pour importer";
    color: var(--accent);
    font-weight: 600;
}}
[data-testid="stFileUploader"] section {{
    padding: 0 !important;
}}
[data-testid="stFileUploader"] section > div {{
    padding: 1rem !important;
    min-height: 100px !important;
}}
[data-testid="stFileUploader"] section [data-testid="stMarkdownContainer"],
[data-testid="stFileUploader"] section small,
[data-testid="stFileUploader"] section p {{
    display: none !important;
}}
[data-testid="stFileUploader"] section > div::after {{
    content: "📁 Déposer le fichier Excel ici";
    font-size: 0.9rem;
    color: var(--text-muted);
}}
[data-testid="stFileUploader"] button[kind="secondary"] {{
    font-size: 0 !important;
    padding: 0.4rem 1.2rem !important;
    border: 1px solid var(--border) !important;
    background: transparent !important;
    border-radius: 6px !important;
}}
[data-testid="stFileUploader"] button[kind="secondary"]::after {{
    content: "Parcourir";
    font-size: 0.85rem;
    color: var(--text);
}}
[data-testid="stFileUploader"] button[aria-label*="Help"]::after,
[data-testid="stFileUploader"] button[aria-label*="aide"]::after {{
    content: "" !important;
    display: none !important;
}}
[data-testid="stFileUploader"] button[aria-label*="Remove"],
[data-testid="stFileUploader"] button[aria-label*="supprimer"] {{
    font-size: 0 !important;
}}
[data-testid="stFileUploader"] button[aria-label*="Remove"]::after,
[data-testid="stFileUploader"] button[aria-label*="supprimer"]::after {{
    content: "Retirer";
    font-size: 0.85rem;
    margin-left: 0.5rem;
    color: var(--text-muted);
}}
[data-testid="stFileUploader"] button:hover {{
    border-color: var(--accent) !important;
    background: var(--surface-alt) !important;
}}

/* ── Dataframe ────────────────────────────────────────── */
[data-testid="stDataFrame"] {{
    border-radius: 8px;
    border: 1px solid var(--border);
}}

/* ── Divider ──────────────────────────────────────────── */
hr {{ border: none; border-top: 1px solid var(--border); margin: 1.2rem 0; }}

/* ── Footer ───────────────────────────────────────────── */
footer {{ visibility: hidden; }}

/* ── Sidebar visibility ──────────────────────────────── */
[data-testid="stSidebar"] {{
    display: {sidebar_display} !important;
}}
/* Masque aussi le bouton d'expansion flottant (step 0) */
[data-testid="collapsedControl"] {{
    display: {sidebar_display} !important;
}}

/* ── Champ "Nom du projet" — comportement de focus et remplissage ── */
/* 1. Neutralise le rouge Streamlit au clic (champ vide) */
[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stTextInput"] input:focus {{
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 2px rgba(68, 114, 196, 0.25) !important;
    transition: border-color 0.2s ease, box-shadow 0.2s ease;
}}
/* 2. Bordure verte quand le champ contient une valeur */
[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stTextInput"] input:not(:placeholder-shown) {{
    border-color: #28a745 !important;
    box-shadow: 0 0 0 2px rgba(40, 167, 69, 0.25) !important;
    transition: border-color 0.2s ease, box-shadow 0.2s ease;
}}
/* 3. Verte aussi quand rempli ET en focus (priorité sur la règle bleue) */
[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stTextInput"] input:not(:placeholder-shown):focus {{
    border-color: #28a745 !important;
    box-shadow: 0 0 0 2px rgba(40, 167, 69, 0.4) !important;
}}

</style>
"""

    js = ""
    return css + js
