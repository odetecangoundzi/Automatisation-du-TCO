"""
app.py — Interface Streamlit pour TCO Automator (production-ready).

Application en 3 étapes :
1. Import du TCO modèle
2. Import des DPGF entreprises (un ou plusieurs, avec suppression)
3. Visualisation du résultat et export

Corrections production :
  SEC-1 : noms de fichiers sanitisés + UUID
  SEC-2 : limite de 10 entreprises max
  SEC-3 : fichiers uploadés supprimés après parsing
  BUG-1 : excel_row indépendant (dans exporter)
  UX-1  : taux TVA paramétrable
  UX-2  : validation du nom d'entreprise
  UX-3  : compteur matched corrigé (article/sub_section)
  UX-4  : confirmation suppression entreprise
  UX-6  : export via BytesIO sans sauvegarde disque
  ARCH-3/4 : config.py + logger.py
"""

import re
import uuid
import os
import pickle
import streamlit as st
import pandas as pd

from core.parser_tco import parse_tco
from core.parser_dpgf import parse_dpgf
from core.merger import merge_company_into_tco
from core.exporter import export_tco
from config import (
    UPLOAD_DIR, ALLOWED_EXTENSIONS, MAX_FILE_SIZE_MB, MAX_COMPANIES,
    COMPANY_NAME_MAX_LEN, TVA_OPTIONS, TVA_DEFAULT, APP_TITLE, APP_ICON,
    PROJECTS_DIR
)
from logger import get_logger

log = get_logger("app")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title=APP_TITLE,
    page_icon=APP_ICON,
    layout="wide",
    initial_sidebar_state="expanded",
)

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PROJECTS_DIR, exist_ok=True)
COMPANY_PATTERN = re.compile(r"^[A-Za-zÀ-ÖØ-öø-ÿ0-9 \-_&'\.]+$")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _safe_save(uploaded_file):
    name = uploaded_file.name
    ext  = os.path.splitext(name)[1].lower()
    
    if ext not in ALLOWED_EXTENSIONS:
        st.error(f"❌ Format non accepté : {ext}. Seul .xlsx est autorisé.")
        log.warning("Upload refusé (extension) : %s", name)
        return None
        
    if uploaded_file.size / (1024 * 1024) > MAX_FILE_SIZE_MB:
        st.error(f"❌ Fichier trop volumineux (> {MAX_FILE_SIZE_MB} MB).")
        log.warning("Upload refusé (taille) : %s", name)
        return None

    safe_name = f"{uuid.uuid4().hex}{ext}"
    path = os.path.join(UPLOAD_DIR, safe_name)
    try:
        with open(path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        log.info("Fichier sauvegardé temporairement : %s -> %s", name, safe_name)
        return path
    except Exception as e:
        log.error("Erreur sauvegarde fichier : %s", e)
        st.error("Erreur technique lors de la sauvegarde.")
        return None


def _validate_company_name(name):
    name = name.strip()
    if not name:
        return None, "Le nom ne peut pas être vide."
    if len(name) > COMPANY_NAME_MAX_LEN:
        return None, f"Nom trop long (max {COMPANY_NAME_MAX_LEN} caractères)."
    if not COMPANY_PATTERN.match(name):
        return None, "Nom invalide (caractères spéciaux interdits)."
    return name, None


def rebuild_merged_tco(tva_rate=TVA_DEFAULT):
    if st.session_state.tco_df is None:
        return
    
    log.info("Reconstruction TCO. TVA=%.2f. Entreprises=%s", 
             tva_rate, list(st.session_state.company_data.keys()))
             
    merged     = st.session_state.tco_df.copy()
    all_alerts = []
    
    for comp_name, comp_data in st.session_state.company_data.items():
        merged, merge_alerts = merge_company_into_tco(
            merged, comp_data["dpgf_df"], comp_name
        )
        all_alerts.extend(comp_data["parse_alerts"])
        all_alerts.extend(merge_alerts)
        
    st.session_state.merged_df  = merged
    st.session_state.all_alerts = all_alerts


def display_alerts(alerts, title="Alertes"):
    if not alerts:
        st.success("✅ Aucune anomalie détectée")
        return
        
    counts = {}
    for a in alerts:
        t = a.get("type", "info")
        counts[t] = counts.get(t, 0) + 1
        
    cols = st.columns(3)
    if counts.get("error"):   cols[0].error(f"🔴 {counts['error']} erreur(s)")
    if counts.get("warning"): cols[1].warning(f"🟡 {counts['warning']} avertissement(s)")
    if counts.get("info"):    cols[2].info(f"🔵 {counts['info']} info(s)")
        
    with st.expander(f"📋 {title} — détail", expanded=False):
        for a in alerts:
            icon = {"error": "🔴", "warning": "🟡", "info": "🔵"}.get(a["type"], "ℹ️")
            st.write(f"{icon} **{a.get('code', '')}** — {a.get('message', '')}")


def display_preview(df, title="Aperçu"):
    hidden = {"Entete", "row_type", "original_row", "parent_code"}
    cols   = [c for c in df.columns if c not in hidden]
    hidden_types = {"empty", "recap", "recap_summary", "total_line", "total_text"}
    visible = df[~df["row_type"].isin(hidden_types)][cols]
    st.write(f"**{title}** ({len(visible)} lignes)")
    st.dataframe(visible, width="stretch", hide_index=True, height=500)


# ---------------------------------------------------------------------------
# Persistence Logic
# ---------------------------------------------------------------------------

def save_project(name):
    """Sauvegarde l'état actuel dans un fichier pickle."""
    if not name:
        return False, "Le nom du projet est vide."
    
    path = os.path.join(PROJECTS_DIR, f"{name}.tco")
    data = {
        "tco_df":       st.session_state.get("tco_df"),
        "company_data": st.session_state.get("company_data"),
        "tco_meta":     st.session_state.get("tco_meta"),
        "step":         st.session_state.get("step"),
        "all_alerts":   st.session_state.get("all_alerts"),
        "merged_df":    st.session_state.get("merged_df"),
        "project_name": name
    }
    try:
        with open(path, "wb") as f:
            pickle.dump(data, f)
        log.info("Projet sauvegardé : %s", name)
        return True, f"Projet '{name}' sauvegardé avec succès."
    except Exception as e:
        log.error("Erreur sauvegarde projet %s : %s", name, e)
        return False, f"Erreur technique : {e}"


def load_project(name):
    """Charge un projet depuis un fichier pickle."""
    path = os.path.join(PROJECTS_DIR, f"{name}.tco")
    if not os.path.exists(path):
        return False, "Le fichier de projet n'existe plus."
    
    try:
        with open(path, "rb") as f:
            data = pickle.load(f)
        
        # Restauration sélective pour éviter de briser la session
        st.session_state.tco_df       = data.get("tco_df")
        st.session_state.company_data = data.get("company_data", {})
        st.session_state.tco_meta     = data.get("tco_meta", {})
        st.session_state.step         = data.get("step", 1)
        st.session_state.all_alerts   = data.get("all_alerts", [])
        st.session_state.merged_df    = data.get("merged_df")
        st.session_state.current_project = name
        
        log.info("Projet chargé : %s", name)
        return True, f"Projet '{name}' chargé."
    except Exception as e:
        log.error("Erreur chargement projet %s : %s", name, e)
        return False, f"Erreur de lecture : {e}"


def list_projects():
    """Liste les noms de projets disponibles."""
    if not os.path.exists(PROJECTS_DIR):
        return []
    files = [f for f in os.listdir(PROJECTS_DIR) if f.endswith(".tco")]
    return sorted([os.path.splitext(f)[0] for f in files])


def delete_project(name):
    """Supprime un fichier projet."""
    path = os.path.join(PROJECTS_DIR, f"{name}.tco")
    if os.path.exists(path):
        os.remove(path)
        log.info("Projet supprimé : %s", name)
        return True
    return False


# ---------------------------------------------------------------------------
# Theme toggle
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Sidebar & Logo
# ---------------------------------------------------------------------------

with st.sidebar:
    # Logo Odetec
    if os.path.exists("odetec_logo.png"):
        st.image("odetec_logo.png", width="stretch")
    
    st.markdown("---")
    st.markdown("### 📋 Mes Projets")
    
    # Liste des projets existants (Direct Click)
    projects = list_projects()
    if projects:
        for p in projects:
            # Button style handled by CSS
            if st.button(f"📄  {p}", key=f"load_{p}", use_container_width=True):
                ok, msg = load_project(p)
                if ok: st.rerun()
                else: st.error(msg)
    else:
        st.caption("Aucun projet sauvegardé.")

    st.markdown("---")
    st.markdown("### 🏗️ Gestion")
    
    curr_name = st.session_state.get("current_project", "")
    proj_name = st.text_input(
        "Titre du projet", 
        value=curr_name,
        placeholder="Entrez le nom...",
        key="proj_title_input",
        label_visibility="collapsed"
    )
    
    if st.button("💾 Sauvegarder", use_container_width=True, type="primary"):
        if proj_name:
            ok, msg = save_project(proj_name)
            if ok: st.success(msg)
            else: st.error(msg)
            st.rerun()
        else:
            st.warning("Nom requis.")

    if st.button("🆕 Nouveau Projet", use_container_width=True):
        for k in ["tco_df", "company_data", "tco_meta", "step", "merged_df", "all_alerts", "current_project"]:
            if k in st.session_state: del st.session_state[k]
        st.session_state.step = 0
        st.rerun()
    
    if projects:
        with st.expander("🗑️ Administration"):
            to_del = st.selectbox("Supprimer un projet", [""] + projects, key="del_select")
            if to_del and st.button(f"Confirmer la suppression"):
                if delete_project(to_del): st.rerun()
    
    st.markdown("---")
    st.markdown("### ⚙️ Paramètres")
    
    if "dark_mode" not in st.session_state:
        st.session_state.dark_mode = False
    
    dark = st.toggle("🌙 Mode sombre", value=st.session_state.dark_mode)
    if dark != st.session_state.dark_mode:
        st.session_state.dark_mode = dark
        st.rerun()

is_dark = st.session_state.dark_mode

# ---------------------------------------------------------------------------
# CSS — Design system avec variables de thème
# ---------------------------------------------------------------------------

if is_dark:
    theme_vars = """
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
    theme_vars = """
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

st.markdown(f"""
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
/* Style pour les boutons de navigation (projets) dans la sidebar */
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
/* Style spécifique pour le bouton 'Nouveau' ou 'Sauvegarder' pour les distinguer */
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
    transition: border-color .2s;
}}
[data-testid="stFileUploader"]:hover {{ border-color: var(--accent); }}
/* Simplification extrême de l'uploader */
[data-testid="stFileUploader"] section {{
    padding: 0 !important;
}}
[data-testid="stFileUploader"] section > div {{
    padding: 1rem !important;
    min-height: 100px !important;
}}

/* Masquer tout ce qui est texte natif (Anglais) */
[data-testid="stFileUploader"] section [data-testid="stMarkdownContainer"],
[data-testid="stFileUploader"] section small,
[data-testid="stFileUploader"] section p {{
    display: none !important;
}}

/* Juste une icône et un petit rappel de format en français */
[data-testid="stFileUploader"] section > div::after {{
    content: "� Déposer le fichier Excel ici";
    font-size: 0.9rem;
    color: var(--text-muted);
}}

/* Style minimaliste du bouton 'Parcourir' uniquement */
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

/* Masquer le texte parasite sur le bouton d'aide (?) s'il existe */
[data-testid="stFileUploader"] button[aria-label*="Help"]::after,
[data-testid="stFileUploader"] button[aria-label*="aide"]::after {{
    content: "" !important;
    display: none !important;
}}

/* Style du bouton 'Retirer' (le X) */
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
    overflow: hidden;
    border: 1px solid var(--border);
}}

/* ── Divider ──────────────────────────────────────────── */
hr {{ border: none; border-top: 1px solid var(--border); margin: 1.2rem 0; }}

/* ── Footer ───────────────────────────────────────────── */
footer {{ visibility: hidden; }}
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Session state init
# ---------------------------------------------------------------------------

defaults = {
    "tco_df":        None,
    "tco_meta":      None,
    "merged_df":     None,
    "all_alerts":    [],
    "company_data":  {},
    "step":          0,
    "upload_counter":0,
    "tva_rate":      TVA_DEFAULT,
    "confirm_remove":None,  # UX-4 : stocke le nom de l'entreprise à supprimer
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ---------------------------------------------------------------------------
# STEP 0 — Landing Page
# ---------------------------------------------------------------------------

if st.session_state.step == 0:
    st.markdown("<div style='height: 4rem;'></div>", unsafe_allow_html=True)
    
    # Hero Section
    col_logo_left, col_logo_mid, col_logo_right = st.columns([1, 2, 1])
    with col_logo_mid:
        if os.path.exists("odetec_logo.png"):
            st.image("odetec_logo.png", use_container_width=True)
        st.markdown(f"<h1 class='main-title'>{APP_TITLE}</h1>", unsafe_allow_html=True)
        st.markdown("<p class='subtitle'>Solution intelligente pour la consolidation des DPGF et le remplissage du TCO.</p>", unsafe_allow_html=True)

    st.markdown("<div style='height: 2rem;'></div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class='company-card' style='height: 100%; min-height: 250px; display: flex; flex-direction: column; justify-content: space-between;'>
            <div>
                <h3>🆕 Nouveau Projet</h3>
                <p style='color: var(--text-muted);'>Commencez une nouvelle analyse en important un modèle de TCO vierge.</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        new_proj_name = st.text_input("Nom du futur projet", placeholder="Ex: Chantier Bordeaux - Lot 04", key="landing_new_proj_name", label_visibility="collapsed")
        if st.button("🚀 Créer le projet", type="primary", use_container_width=True):
            if new_proj_name:
                st.session_state.current_project = new_proj_name
                st.session_state.step = 1
                st.rerun()
            else:
                st.warning("Veuillez saisir un nom de projet.")

    with col2:
        st.markdown("""
        <div class='company-card' style='height: 100%; min-height: 250px; display: flex; flex-direction: column; justify-content: space-between;'>
            <div>
                <h3>📂 Ouvrir un Projet</h3>
                <p style='color: var(--text-muted);'>Reprenez un travail en cours depuis vos sauvegardes locales.</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        projects = list_projects()
        if projects:
            selected_proj = st.selectbox("Choisir un projet", [""] + projects, key="landing_open_proj_select", label_visibility="collapsed")
            if st.button("📂 Ouvrir", use_container_width=True) and selected_proj:
                ok, msg = load_project(selected_proj)
                if ok: 
                    # If loaded, the step is usually >= 1. If it was 0, force it to 1
                    if st.session_state.step == 0:
                        st.session_state.step = 1
                    st.rerun()
                else: 
                    st.error(msg)
        else:
            st.info("Aucun projet sauvegardé pour le moment.")


# ---------------------------------------------------------------------------
# Header + progress (Visible only after Step 0)
# ---------------------------------------------------------------------------

if st.session_state.step > 0:
    st.markdown(f"<h1 class='main-title' style='font-size: 1.8rem; text-align: left;'>{APP_TITLE} <span style='color: var(--text-muted); font-size: 1.2rem; font-weight: 400;'>| {st.session_state.get('current_project', 'Sans titre')}</span></h1>", unsafe_allow_html=True)
    st.divider()


# ---------------------------------------------------------------------------
# STEP 1 — Import TCO
# ---------------------------------------------------------------------------

if st.session_state.step >= 1:
    st.markdown(
        "<div class='step-header'>📥 Etape 1 : importer le modèle vierge</div>",
        unsafe_allow_html=True,
    )
    st.caption("Fichier DPGF LOT (.xlsx) — Colonnes : Code | Désignation | Qu. | U. | Px U. | Px tot.")

    tco_file = st.file_uploader(
        "Charger Le DPGF Modèle", type=["xlsx"], key="tco_upload",
        help="Fichier DPGF LOT servant de base",
        label_visibility="visible"
    )

    if tco_file and st.session_state.tco_df is None:
        path = _safe_save(tco_file)
        if path:
            with st.spinner("🔄 Lecture du TCO..."):
                try:
                    tco_df, meta = parse_tco(path)
                    st.session_state.tco_df      = tco_df
                    st.session_state.tco_meta    = meta
                    st.session_state.merged_df   = tco_df.copy()
                    
                    info = meta["project_info"]
                    if info:
                        st.info(
                            f"📋 **Projet :** {info.get('projet','N/A')} — "
                            f"**Lot :** {info.get('lot','N/A')}"
                        )
                    st.success(f"✅ TCO chargé — {len(tco_df)} lignes")
                except Exception as e:
                    log.error("Erreur parsing TCO", exc_info=True)
                    st.error(f"❌ Erreur de lecture : {e}")
                finally:
                    try: os.remove(path)
                    except: pass

    if st.session_state.tco_df is not None:
        # Enlever l'aperçu à l'étape 1 (demandé)
        # display_preview(st.session_state.tco_df, "Aperçu du TCO")
        if st.session_state.step == 1:
            if st.button("✅ Valider le TCO et continuer", type="primary"):
                st.session_state.step = 2
                st.rerun()


# ---------------------------------------------------------------------------
# STEP 2 — Import DPGF Entreprises
# ---------------------------------------------------------------------------

if st.session_state.step >= 2:
    st.markdown(
        "<div class='step-header'>📥 Etape 2 : charger les DPGF fournis par les entreprises</div>",
        unsafe_allow_html=True,
    )

    st.divider()
    
    
    # UX-4 : Confirmation de suppression (Dialogue modal simulé)
    if st.session_state.confirm_remove:
        to_remove = st.session_state.confirm_remove
        st.warning(f"⚠️ Voulez-vous vraiment supprimer **{to_remove}** ?")
        col_y, col_n = st.columns([1, 5])
        with col_y:
            if st.button("✅ Oui, supprimer", type="primary"):
                del st.session_state.company_data[to_remove]
                st.session_state.confirm_remove = None
                rebuild_merged_tco(st.session_state.tva_rate)
                st.success(f"Entreprise {to_remove} supprimée.")
                st.rerun()
        with col_n:
            if st.button("❌ Annuler"):
                st.session_state.confirm_remove = None
                st.rerun()
        st.divider()

    n_companies = len(st.session_state.company_data)
    if n_companies:
        st.write(f"**{n_companies} / {MAX_COMPANIES} entreprise(s) importée(s) :**")
        for comp_name in list(st.session_state.company_data.keys()):
            comp   = st.session_state.company_data[comp_name]
            n_art  = len(comp["dpgf_df"][comp["dpgf_df"]["row_type"] == "article"])
            n_alrt = len(comp["parse_alerts"])
            
            col_inf, col_btn = st.columns([4, 1])
            with col_inf:
                st.markdown(
                    f"<div class='company-card'>🏢 <b>{comp_name}</b> — "
                    f"{n_art} articles, {n_alrt} alerte(s) "
                    f"<i>({comp['filename']})</i></div>",
                    unsafe_allow_html=True,
                )
            with col_btn:
                # UX-4 : Déclenche la confirmation
                if st.button("🗑️ Retirer", key=f"rm_{comp_name}"):
                    st.session_state.confirm_remove = comp_name
                    st.rerun()
        st.divider()

    if n_companies >= MAX_COMPANIES:
        st.warning(f"⚠️ Limite de {MAX_COMPANIES} entreprises atteinte.")
    else:
        st.write("**Ajouter une entreprise :**")
        dpgf_key = f"dpgf_{st.session_state.upload_counter}"
        name_key = f"name_{st.session_state.upload_counter}"

        col1, col2 = st.columns([2, 1])
        with col1:
            dpgf_file = st.file_uploader(
                "Charger un DPGF entreprise", type=["xlsx"],
                key=dpgf_key,
            )
        
        # UX : Auto-fill du nom de l'entreprise si un fichier est chargé
        # On utilise une clé de suivi pour détecter le changement de fichier
        file_tracker_key = f"tracker_{st.session_state.upload_counter}"
        if dpgf_file:
            if st.session_state.get(file_tracker_key) != dpgf_file.name:
                filename_clean = os.path.splitext(dpgf_file.name)[0]
                suggested_name = re.sub(r"^DPGF\s+", "", filename_clean, flags=re.IGNORECASE)
                st.session_state[name_key] = suggested_name.upper()
                st.session_state[file_tracker_key] = dpgf_file.name

        with col2:
            raw_name = st.text_input("Nom de l'entreprise", key=name_key, placeholder="Ex: MAB SUD-OUEST")

        if dpgf_file and raw_name:
            company_name, name_err = _validate_company_name(raw_name)
            if name_err:
                st.error(f"❌ {name_err}")
            elif company_name in st.session_state.company_data:
                st.warning(f"⚠️ **{company_name}** est déjà importée.")
            else:
                # Automatisme : Fusion au chargement
                path = _safe_save(dpgf_file)
                if path:
                    with st.spinner(f"🔄 Traitement de {company_name}..."):
                        try:
                            dpgf_df, parse_alerts = parse_dpgf(path)
                            n_matched = len(dpgf_df[dpgf_df["row_type"].isin(["article", "sub_section"])])
                            
                            st.session_state.company_data[company_name] = {
                                "dpgf_df":      dpgf_df,
                                "parse_alerts": parse_alerts,
                                "filename":     dpgf_file.name,
                            }
                            rebuild_merged_tco(st.session_state.tva_rate)
                            st.session_state.upload_counter += 1
                            # Succès affiché seulement si pas de message d'erreur persistant
                            # st.success(f"✅ **{company_name}** fusionnée — {n_matched} postes")
                            # if parse_alerts: display_alerts(parse_alerts, company_name)
                            st.rerun()
                        except Exception as e:
                            log.error("Erreur fusion DPGF %s", company_name, exc_info=True)
                            st.error(f"❌ Erreur : {e}")
                        finally:
                            try: os.remove(path)
                            except: pass

    # Enlever le preview à l'étape 2 (demandé)
    # if st.session_state.company_data:
    #     display_preview(st.session_state.merged_df, f"TCO fusionné ({n_companies} entreprise(s))")

    col_nav1, col_nav2 = st.columns(2)
    with col_nav1:
        if st.session_state.company_data:
            if st.button("➡️ Passer au résultat final", type="primary"):
                st.session_state.step = 3
                # Reset export buffer to ensure fresh data
                if "export_buffer" in st.session_state:
                    del st.session_state.export_buffer
                st.session_state.export_done = False
                st.rerun()
    with col_nav2:
        if st.button("🔄 Tout réinitialiser"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()


# ---------------------------------------------------------------------------
# STEP 3 — Résultat & Export
# ---------------------------------------------------------------------------

if st.session_state.step >= 3:
    st.markdown(
        "<div class='step-header'>📊 Étape 3 — Résultat Final</div>",
        unsafe_allow_html=True,
    )

    st.markdown("""
    <div class='legend-box'>
        <div class='legend-item'><span class='color-dot' style='background:#FFC7CE'></span> Erreur</div>
        <div class='legend-item'><span class='color-dot' style='background:#FFE4B5'></span> Avertissement</div>
        <div class='legend-item'><span class='color-dot' style='background:#FFFFCC'></span> Note</div>
        <div class='legend-item'><span class='color-dot' style='background:#D6EAF8'></span> Info</div>
    </div>
    """, unsafe_allow_html=True)

    merged = st.session_state.merged_df

    # Stats
    art_rows = merged[merged["row_type"] == "article"]

    cols = st.columns(3)
    with cols[0]: st.metric("📋 Articles",   len(art_rows))
    with cols[1]: st.metric("🏢 Entreprises", len(st.session_state.company_data))
    with cols[2]: st.metric("⚠️ Alertes",    len(st.session_state.all_alerts))

    display_preview(merged, "TCO Final Consolidé")

    if st.session_state.all_alerts:
        display_alerts(st.session_state.all_alerts, "Toutes les alertes")

    if st.button("⬅️ Retour — Modifier les entreprises"):
        st.session_state.step = 2
        st.session_state.export_done = False
        st.rerun()

    st.divider()
    from datetime import datetime
    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename   = f"TCO_FINAL_{timestamp}.xlsx"

    # Pré-génération du buffer pour téléchargement immédiat
    try:
        if "export_buffer" not in st.session_state:
            with st.spinner("🔄 Préparation du fichier..."):
                st.session_state.export_buffer = export_tco(
                    st.session_state.merged_df,
                    st.session_state.tco_meta,
                    output_path=None,
                    alerts=st.session_state.all_alerts,
                )
        
        def on_export_click():
            st.session_state.export_done = True

        st.download_button(
            label="📥 Exporter le TCO Final (.xlsx)",
            data=st.session_state.export_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            on_click=on_export_click,
            width="stretch"
        )
        
        if st.session_state.get("export_done"):
            st.success("✅ Téléchargement OK")
            # Optionnel : masquer après quelques secondes ou au prochain rerun ? 
            # Streamlit garde le message jusqu'au prochain changement d'état.
            log.info("Export téléchargé : %s", filename)

    except Exception as e:
        log.error("Erreur préparation export", exc_info=True)
        st.error(f"❌ Erreur de génération : {e}")

st.divider()
st.caption(f"{APP_TITLE} v{2.1} — TCO Automator")
