"""
app.py — Interface Streamlit pour TCO Automator (production-ready).

Application en 3 étapes :
1. Import du TCO modèle
2. Import des DPGF entreprises (un ou plusieurs, avec suppression)
3. Visualisation du résultat et export

Corrections production :
  SEC-1 : noms de fichiers sanitisés + UUID
  SEC-2 : limite d'entreprises max
  SEC-3 : fichiers uploadés supprimés après parsing
  SEC-4 : pickle remplacé par JSON+gzip (services/persistence.py)
  SEC-5 : validation magic bytes XLSX (services/file_validator.py)
  SEC-6 : bouton kill protégé par ADMIN_MODE
  BUG-1 : excel_row indépendant (dans exporter)
  UX-1  : taux TVA paramétrable
  UX-2  : validation du nom d'entreprise
  UX-3  : compteur matched corrigé (article/sub_section)
  UX-4  : confirmation suppression entreprise
  UX-6  : export via BytesIO sans sauvegarde disque
  ARCH-3/4 : config.py + logger.py
  ARCH-5 : CSS extrait dans app/__init__.py
  ARCH-6 : persistence extraite dans services/persistence.py
"""

import os
import re
import signal
import uuid
from datetime import datetime

import pandas as pd
import streamlit as st

from app import get_full_css
from config import (
    ADMIN_MODE,
    APP_ICON,
    APP_TITLE,
    APP_VERSION,
    ALLOWED_EXTENSIONS,
    COMPANY_NAME_MAX_LEN,
    MAX_COMPANIES,
    MAX_FILE_SIZE_MB,
    PROJECTS_DIR,
    TVA_DEFAULT,
    TVA_OPTIONS,
    UPLOAD_DIR,
)
from core.exporter import export_tco
from core.merger import merge_company_into_tco, merge_all_companies, compute_section_totals
from core.parser_dpgf import parse_dpgf
from core.parser_tco import parse_tco
from logger import get_logger
from services.file_validator import validate_uploaded_file
from services.persistence import (
    delete_project,
    list_projects,
    load_project,
    save_project,
)

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
    "dark_mode":     False,
    "export_done":   False,
    "_flash_msg":    None,   # P12 : message court affiché après rerun
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PROJECTS_DIR, exist_ok=True)
COMPANY_PATTERN = re.compile(r"^[A-Za-zÀ-ÖØ-öø-ÿ0-9 \-_&'\.]+$")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _safe_save(uploaded_file):
    """Sauvegarde un fichier uploadé après validation complète."""
    # SEC-5 : Validation extension + taille + magic bytes
    is_valid, error_msg = validate_uploaded_file(
        uploaded_file,
        max_mb=MAX_FILE_SIZE_MB,
        allowed_extensions=ALLOWED_EXTENSIONS,
    )
    if not is_valid:
        st.error(f"❌ {error_msg}")
        return None

    # SEC-1 : nom sécurisé UUID
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    safe_name = f"{uuid.uuid4().hex}{ext}"
    path = os.path.join(UPLOAD_DIR, safe_name)
    try:
        with open(path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        log.info("Fichier sauvegardé : %s (taille=%d Ko)", safe_name, uploaded_file.size // 1024)
        return path
    except OSError as e:
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
    """Reconstruit le TCO fusionné depuis l'état de session.
    Délègue la logique métier à merge_all_companies (core/merger.py).
    """
    if st.session_state.tco_df is None:
        return
    merged_df, all_alerts = merge_all_companies(
        st.session_state.tco_df,
        st.session_state.company_data,
        tva_rate=tva_rate,
    )
    st.session_state.merged_df  = merged_df
    st.session_state.all_alerts = all_alerts


def display_alerts(alerts, title="Alertes"):
    if not alerts:
        st.success("✅ Aucune anomalie détectée")
        return

    with st.expander(f"📋 {title} — détails", expanded=False):
        for a in alerts:
            icon = {"error": "🔴", "warning": "🟡", "info": "🔵"}.get(a["type"], "ℹ️")
            st.write(f"{icon} **{a.get('code', '')}** — {a.get('message', '')}")


def display_preview(df, title="Aperçu"):
    hidden = {"Entete", "row_type", "original_row", "parent_code"}
    cols   = [c for c in df.columns if c not in hidden]
    hidden_types = {"empty", "recap", "recap_summary", "total_line", "total_text"}
    visible = df[~df["row_type"].isin(hidden_types)][cols]
    st.write(f"**{title}** ({len(visible)} lignes)")
    st.dataframe(visible, use_container_width=True, hide_index=True, height=500)


def _cleanup_file(path):
    """Supprime un fichier temporaire de manière sûre."""
    try:
        os.remove(path)
    except OSError as e:
        log.warning("Impossible de supprimer %s : %s", path, e)


# ---------------------------------------------------------------------------
# Theme toggle
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Sidebar & Logo
# ---------------------------------------------------------------------------

if st.session_state.step > 0:
    with st.sidebar:
        # Logo
        if os.path.exists("odetec_logo.png"):
            st.image("odetec_logo.png", use_container_width=True)

        st.markdown("---")

        # Nom du projet en cours — affiché, non éditable
        curr_name = st.session_state.get("current_project", "Sans titre")
        st.markdown(
            f"<div style='"
            f"background: linear-gradient(135deg, #2F5496, #4472C4);"
            f"color: white; padding: 12px 16px; border-radius: 10px;"
            f"font-weight: 600; font-size: 0.95rem; word-break: break-word;"
            f"margin-bottom: 0.5rem;'>"
            f"📁 {curr_name}</div>",
            unsafe_allow_html=True,
        )

        # Point 2 : bouton Enregistrer juste sous le nom du projet
        if st.button("💾 Enregistrer", use_container_width=True, key="sidebar_save"):
            ok, msg = save_project(curr_name, st.session_state)
            if ok:
                st.success(msg)
            else:
                st.error(msg)

        st.markdown("---")

        # Retour à l'accueil
        if st.button("🏠 Retour à l'accueil", use_container_width=True, type="primary"):
            for k in ["tco_df", "company_data", "tco_meta", "merged_df", "all_alerts", "current_project"]:
                if k in st.session_state:
                    del st.session_state[k]
            st.session_state.step = 0
            st.rerun()

        st.markdown("---")

        # Fermer l'application — réservé à l'administrateur
        if ADMIN_MODE:
            if st.button("❌ Fermer l'application", use_container_width=True, help="Arrête le serveur Streamlit (admin uniquement)"):
                st.warning("Arrêt de l'application...")
                os.kill(os.getpid(), signal.SIGTERM)

is_dark = st.session_state.dark_mode

# ---------------------------------------------------------------------------
# CSS — Design system (extrait dans app/__init__.py)
# ---------------------------------------------------------------------------

st.markdown(
    get_full_css(is_dark, hide_sidebar=(st.session_state.step == 0)),
    unsafe_allow_html=True,
)



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
        with st.container(border=True):
            st.markdown("### 🆕 Nouveau Projet")
            st.markdown("<p>Commencez une nouvelle analyse en important un modèle de TCO vierge.</p>", unsafe_allow_html=True)
            
            new_proj_name = st.text_input("Nom du futur projet", placeholder="Ex: Chantier Bordeaux - Lot 04", key="landing_new_proj_name", label_visibility="collapsed")
            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            if st.button("🚀 Créer le projet", type="primary", use_container_width=True):
                if new_proj_name:
                    st.session_state.current_project = new_proj_name
                    st.session_state.step = 1
                    st.rerun()
                else:
                    st.warning("Veuillez saisir un nom de projet.")

    with col2:
        with st.container(border=True):
            st.markdown("### 📂 Ouvrir un Projet")
            st.markdown("<p>Reprenez un travail en cours depuis vos sauvegardes locales.</p>", unsafe_allow_html=True)
            
            projects = list_projects()
            if projects:
                for p in projects:
                    rcol_name, rcol_del = st.columns([7, 1])
                    with rcol_name:
                        if st.button(f"📄 {p}", key=f"landing_load_{p}", use_container_width=True):
                            ok, msg = load_project(p, st.session_state)
                            if ok:
                                if st.session_state.step == 0:
                                    st.session_state.step = 1
                                if "export_buffer" in st.session_state:
                                    del st.session_state.export_buffer
                                st.rerun()
                            else:
                                st.error(msg)
                    with rcol_del:
                        if st.button("🗑️", key=f"landing_del_{p}", help=f"Supprimer {p}"):
                            if delete_project(p):
                                st.rerun()
            else:
                st.caption("Aucun projet sauvegardé pour le moment.")


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
                    
                    # SÉCURITÉ : Recaler les totaux de l'estimation (colonne de base)
                    # Cela garantit que même si le fichier source a des erreurs de calcul
                    # ou des sous-totaux manquants, le TCO interne est cohérent.
                    compute_section_totals(tco_df, "Px_Tot_HT", tva_rate=st.session_state.tva_rate)
                    
                    st.session_state.tco_df      = tco_df
                    st.session_state.tco_meta    = meta
                    st.session_state.merged_df   = tco_df.copy()
                    
                    info = meta["project_info"]
                    if info:
                        st.info(
                            f"📋 **Projet :** {info.get('projet','N/A')} — "
                            f"**Lot :** {info.get('lot','N/A')}"
                        )
                    template_name = os.path.splitext(tco_file.name)[0]
                    st.success(f"✅ Template {template_name} chargé — {len(tco_df)} lignes")
                except Exception as e:
                    log.error("Erreur parsing TCO", exc_info=True)
                    st.error(f"❌ Erreur de lecture : {e}")
                finally:
                    _cleanup_file(path)

    if st.session_state.tco_df is not None:
        if st.session_state.step == 1:
            if st.button("➡️ Passer à l'étape suivante", type="primary"):
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

    # P12 : affichage du message de confirmation après rerun
    if st.session_state._flash_msg:
        st.success(st.session_state._flash_msg)
        st.session_state._flash_msg = None

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
                if "export_buffer" in st.session_state:
                    del st.session_state.export_buffer
                # BUG-A FIX : incrémenter upload_counter pour réinitialiser le widget file_uploader
                st.session_state.upload_counter += 1
                st.session_state._flash_msg = f"✅ Entreprise **{to_remove}** supprimée."
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
        # UX-5 : Multi-upload
        # BUG-A FIX : clé dynamique via upload_counter pour forcer la réinitialisation
        # du widget après suppression d'une entreprise (évite que Streamlit garde le
        # fichier précédent en mémoire et ignore un re-upload du même fichier)
        dpgf_files = st.file_uploader(
            "Importer un ou plusieurs DPGF entreprise",
            type=["xlsx"],
            key=f"multi_dpgf_upload_{st.session_state.upload_counter}",
            accept_multiple_files=True,
            help="Sélectionnez tous les fichiers DPGF des entreprises à fusionner (Format .xlsx)"
        )

        if dpgf_files:
            # Traitement automatique : seuls les fichiers non encore importés sont traités
            processed_filenames = {v["filename"] for v in st.session_state.company_data.values()}
            new_files = [f for f in dpgf_files if f.name not in processed_filenames]

            if new_files:
                success_count = 0
                for dpgf_file in new_files:
                    filename_clean = os.path.splitext(dpgf_file.name)[0]
                    company_name = re.sub(
                        r"^DPGF\s+", "", filename_clean, flags=re.IGNORECASE
                    ).strip().upper()

                    path = _safe_save(dpgf_file)
                    if path:
                        with st.spinner(f"🔄 Fusion de {company_name}..."):
                            try:
                                dpgf_df, parse_alerts = parse_dpgf(path)
                                st.session_state.company_data[company_name] = {
                                    "dpgf_df":      dpgf_df,
                                    "parse_alerts": parse_alerts,
                                    "filename":     dpgf_file.name,
                                }
                                success_count += 1
                            except Exception as e:
                                log.error("Erreur fusion %s", company_name, exc_info=True)
                                st.error(f"❌ Erreur sur {company_name}: {e}")
                            finally:
                                _cleanup_file(path)

                if success_count > 0:
                    rebuild_merged_tco(st.session_state.tva_rate)
                    if "export_buffer" in st.session_state:
                        del st.session_state.export_buffer
                    st.session_state.step = 3
                    st.session_state.export_done = False
                    st.rerun()

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

    # Point 7 : résumé des anomalies par catégorie
    all_alerts = st.session_state.all_alerts
    if all_alerts:
        n_err  = sum(1 for a in all_alerts if a.get("type") == "error")
        n_warn = sum(1 for a in all_alerts if a.get("type") == "warning")
        n_info = sum(1 for a in all_alerts if a.get("type") == "info")
        parts  = []
        if n_err:
            parts.append(f"🔴 {n_err} erreur(s)")
        if n_warn:
            parts.append(f"🟡 {n_warn} avertissement(s)")
        if n_info:
            parts.append(f"🔵 {n_info} info(s)")
        st.caption("Anomalies : " + " — ".join(parts))

    display_preview(merged, "TCO Final Consolidé")

    if st.session_state.all_alerts:
        display_alerts(st.session_state.all_alerts, "Toutes les alertes")

    if st.button("⬅️ Retour — Modifier les entreprises"):
        st.session_state.step = 2
        st.session_state.export_done = False
        st.rerun()

    st.divider()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Point 1 : nom du fichier basé sur le lot détecté dans le template
    _meta_info = st.session_state.tco_meta or {}
    _lot_raw   = _meta_info.get("project_info", {}).get("lot", "") or ""
    if not _lot_raw:
        _lot_raw = st.session_state.get("current_project", "")
    # Normalisation : majuscules, espaces → underscore, caractères spéciaux supprimés
    _lot_norm = re.sub(r"[^A-Z0-9]", "_", _lot_raw.upper())
    _lot_norm = re.sub(r"_+", "_", _lot_norm).strip("_")
    filename  = (
        f"TCO_FINAL_{_lot_norm}_{timestamp}.xlsx" if _lot_norm
        else f"TCO_FINAL_{timestamp}.xlsx"
    )

    # Pré-génération du buffer pour téléchargement immédiat
    try:
        if "export_buffer" not in st.session_state:
            with st.spinner("🔄 Préparation du fichier..."):
                st.session_state.export_buffer = export_tco(
                    st.session_state.merged_df,
                    st.session_state.tco_meta,
                    output_path=None,
                    alerts=st.session_state.all_alerts,
                    tva_rate=st.session_state.tva_rate
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
            use_container_width=True,
        )
        
        if st.session_state.get("export_done"):
            st.success("✅ Téléchargement OK")
            log.info("Export téléchargé : %s", filename)

    except Exception as e:
        log.error("Erreur préparation export", exc_info=True)
        st.error(f"❌ Erreur de génération : {e}")

st.divider()
st.caption(f"{APP_TITLE} v{APP_VERSION} — TCO Automator")
