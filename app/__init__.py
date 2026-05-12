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
        --success:      #4ade80;
        --success-bg:   rgba(74, 222, 128, 0.12);
        --warning:      #f59e0b;
        --warning-bg:   rgba(245, 158, 11, 0.14);
        --danger:       #ef4444;
        --danger-bg:    rgba(239, 68, 68, 0.13);
        --info-bg:      rgba(91, 155, 213, 0.14);
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
        --success:      #198754;
        --success-bg:   #eaf7ef;
        --warning:      #b7791f;
        --warning-bg:   #fff7e6;
        --danger:       #c2410c;
        --danger-bg:    #fff1ed;
        --info-bg:      #edf5ff;
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
{theme_vars}

/* ── Global ───────────────────────────────────────────── */
html, body, [class*="css"] {{
    font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}}
.block-container {{ max-width: 1320px; padding-top: 1.4rem; padding-bottom: 2rem; }}

.app-shell {{
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 14px 18px;
    background: var(--surface);
    box-shadow: 0 1px 8px var(--shadow-light);
}}

/* ── Title ────────────────────────────────────────────── */
.main-title {{
    text-align: center;
    font-size: 2.35rem;
    font-weight: 700;
    background: linear-gradient(135deg, var(--accent-dark) 0%, var(--accent) 60%, #6ca0dc 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: .2rem;
    letter-spacing: 0;
}}
.subtitle {{
    text-align: center;
    color: var(--text-muted);
    font-size: .95rem;
    margin-bottom: 1.2rem;
}}
.home-description {{
    max-width: 880px;
    margin: 0 auto 1.15rem;
    padding: 14px 18px;
    border: 1px solid var(--border);
    border-left: 4px solid var(--accent);
    border-radius: 8px;
    background: var(--surface);
    color: var(--text);
    box-shadow: 0 1px 8px var(--shadow-light);
    font-size: .95rem;
    line-height: 1.55;
}}
.home-description strong {{
    color: var(--accent-dark);
}}

/* ── Step headers ─────────────────────────────────────── */
.step-header {{
    background: linear-gradient(135deg, var(--accent-dark) 0%, var(--accent-deep) 68%, var(--success) 160%);
    color: white;
    padding: 14px 18px;
    border-radius: 8px;
    margin: 1.5rem 0 .8rem;
    font-size: 1.05rem;
    font-weight: 600;
    letter-spacing: 0;
    box-shadow: 0 4px 15px var(--shadow);
}}
.workflow-steps {{
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 8px;
    margin: .3rem 0 1.2rem;
}}
.workflow-step {{
    display: flex;
    align-items: center;
    gap: 10px;
    min-height: 48px;
    padding: 10px 12px;
    border: 1px solid var(--border);
    border-radius: 8px;
    background: var(--surface);
    color: var(--text-muted);
}}
.workflow-step.done {{
    border-color: rgba(25, 135, 84, .42);
    background: var(--success-bg);
    color: var(--text);
}}
.workflow-step.active {{
    border-color: var(--accent);
    box-shadow: 0 0 0 3px rgba(68, 114, 196, .14);
    color: var(--text);
}}
.workflow-badge {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 26px;
    height: 26px;
    border-radius: 50%;
    font-weight: 700;
    font-size: .78rem;
    background: var(--surface-alt);
    color: var(--text-muted);
    flex: 0 0 auto;
}}
.workflow-step.done .workflow-badge {{
    background: var(--success);
    color: white;
}}
.workflow-step.active .workflow-badge {{
    background: var(--accent-deep);
    color: white;
}}
.workflow-label {{
    display: block;
    font-size: .86rem;
    font-weight: 700;
    line-height: 1.15;
}}
.workflow-caption {{
    display: block;
    font-size: .74rem;
    color: var(--text-muted);
    margin-top: 2px;
}}

/* ── Company cards ────────────────────────────────────── */
.company-card {{
    background: var(--card-bg);
    border: 1px solid var(--border);
    border-left: 4px solid var(--accent);
    border-radius: 8px;
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
.company-meta {{
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-top: 6px;
    color: var(--text-muted);
    font-size: .82rem;
}}
.status-pill {{
    display: inline-flex;
    align-items: center;
    gap: 6px;
    border-radius: 999px;
    padding: 3px 9px;
    font-size: .76rem;
    font-weight: 700;
    border: 1px solid transparent;
}}
.status-ready {{ color: var(--success); background: var(--success-bg); border-color: rgba(25, 135, 84, .24); }}
.status-progress {{ color: var(--accent-dark); background: var(--info-bg); border-color: rgba(68, 114, 196, .24); }}
.status-empty {{ color: var(--text-muted); background: var(--surface-alt); border-color: var(--border); }}

/* ── Lots / project dashboard ─────────────────────────── */
.summary-strip {{
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 10px;
    margin: 1rem 0 1.2rem;
}}
.summary-card {{
    border: 1px solid var(--border);
    border-radius: 8px;
    background: var(--surface);
    padding: 12px 14px;
    box-shadow: 0 1px 6px var(--shadow-light);
}}
.summary-label {{
    color: var(--text-muted);
    font-size: .76rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .06em;
}}
.summary-value {{
    color: var(--text);
    font-size: 1.35rem;
    font-weight: 800;
    line-height: 1.2;
    margin-top: 2px;
}}
.lot-card {{
    border: 1px solid var(--border);
    border-left: 4px solid var(--text-muted);
    border-radius: 8px;
    background: var(--surface);
    padding: 11px 14px;
    margin-bottom: 8px;
    color: var(--text);
}}
.lot-card.status-ready {{ border-left-color: var(--success); }}
.lot-card.status-progress {{ border-left-color: var(--accent); }}
.lot-card.status-empty {{ border-left-color: var(--text-muted); }}
.lot-title {{
    font-weight: 800;
    line-height: 1.25;
    overflow-wrap: anywhere;
}}
.lot-meta {{
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-top: 6px;
    color: var(--text-muted);
    font-size: .8rem;
}}
.empty-state {{
    border: 1px dashed var(--border);
    border-radius: 8px;
    padding: 16px;
    background: var(--surface-alt);
    color: var(--text-muted);
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
.sidebar-project {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-left: 4px solid var(--accent);
    color: var(--text);
    padding: 10px 12px;
    border-radius: 8px;
    font-weight: 800;
    font-size: .94rem;
    overflow-wrap: anywhere;
    margin-bottom: .65rem;
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
    border-radius: 8px;
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
.alert-summary {{
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin: .4rem 0 .9rem;
}}
.alert-chip {{
    border-radius: 999px;
    padding: 4px 10px;
    font-size: .8rem;
    font-weight: 700;
    border: 1px solid var(--border);
}}
.alert-chip.error {{ background: var(--danger-bg); color: var(--danger); }}
.alert-chip.warning {{ background: var(--warning-bg); color: var(--warning); }}
.alert-chip.info {{ background: var(--info-bg); color: var(--accent-dark); }}
.alert-row {{
    border: 1px solid var(--border);
    border-left: 4px solid var(--accent);
    border-radius: 8px;
    padding: 9px 12px;
    margin: 7px 0;
    background: var(--surface);
    color: var(--text);
}}
.alert-row.error {{ border-left-color: var(--danger); background: var(--danger-bg); }}
.alert-row.warning {{ border-left-color: var(--warning); background: var(--warning-bg); }}
.alert-row.info {{ border-left-color: var(--accent); background: var(--info-bg); }}
.alert-code {{
    font-weight: 800;
    margin-right: 6px;
}}
.alert-meta {{
    color: var(--text-muted);
    font-size: .78rem;
    margin-top: 3px;
}}

/* ── File uploader ────────────────────────────────────── */
[data-testid="stFileUploader"] {{
    border: 2px dashed var(--border);
    border-radius: 8px;
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
.preview-toolbar {{
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 10px 12px;
    background: var(--surface);
    margin-bottom: 10px;
}}
.export-panel {{
    border: 1px solid var(--border);
    border-radius: 8px;
    background: var(--surface);
    padding: 16px;
    box-shadow: 0 1px 8px var(--shadow-light);
}}
.export-filename {{
    color: var(--text-muted);
    font-size: .86rem;
    overflow-wrap: anywhere;
    margin-bottom: 10px;
}}

@media (max-width: 760px) {{
    .workflow-steps,
    .summary-strip {{
        grid-template-columns: 1fr;
    }}
    .main-title {{
        font-size: 1.9rem;
    }}
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
