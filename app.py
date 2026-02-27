"""
app.py — Interface Streamlit pour Export du TCO (production-ready).

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

import html as html_mod
import os
import re
import signal
import uuid
from datetime import datetime

import streamlit as st

from app import get_full_css
from config import (
    ADMIN_MODE,
    ALLOWED_EXTENSIONS,
    APP_ICON,
    APP_TITLE,
    APP_VERSION,
    COMPANY_NAME_MAX_LEN,
    MAX_COMPANIES,
    MAX_FILE_SIZE_MB,
    PROJECTS_DIR,
    TVA_DEFAULT,
    TVA_OPTIONS,
    UPLOAD_DIR,
)
from core.exporter import export_tco
from core.merger import compute_section_totals, merge_all_companies, merge_company_into_tco
from core.parser_dpgf import parse_dpgf
from core.parser_dpgf_pdf import parse_dpgf_pdf
from core.parser_tco import parse_tco
from logger import get_logger
from services.file_validator import DPGF_ALLOWED_EXTENSIONS, validate_uploaded_file
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
    "active_project": None,  # dict {project_id, project_name, created_at, lots:[...]}
    "active_lot_id": None,  # str — lot_id du lot courant dans active_project
    "step": 0,
    "upload_counter": 0,
    "confirm_remove": None,  # UX-4 : stocke le nom de l'entreprise à supprimer
    "confirm_shutdown": False,  # ADMIN : confirmation deux étapes avant arrêt serveur
    "dark_mode": False,
    "export_done": False,
    "_flash_msg": None,  # P12 : message court affiché après rerun
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PROJECTS_DIR, exist_ok=True)
# APP-1 : & et ' retirés — ils peuvent servir à s'échapper de contextes HTML/SQL
COMPANY_PATTERN = re.compile(r"^[A-Za-zÀ-ÖØ-öø-ÿ0-9 \-_.()]+$")
# Extensions sans le point, pour le paramètre type= des widgets file_uploader Streamlit
_UPLOADER_TYPES = [ext.lstrip(".") for ext in ALLOWED_EXTENSIONS]
# DPGF entreprise : Excel + PDF
_DPGF_UPLOADER_TYPES = [ext.lstrip(".") for ext in DPGF_ALLOWED_EXTENSIONS]


# ---------------------------------------------------------------------------
# Helpers multi-lots
# ---------------------------------------------------------------------------


def _get_active_lot() -> dict | None:
    """Retourne le lot actif ou None si aucun lot/projet selectionne."""
    proj = st.session_state.get("active_project")
    lot_id = st.session_state.get("active_lot_id")
    if not proj or not lot_id:
        return None
    return next((lot for lot in proj.get("lots", []) if lot.get("lot_id") == lot_id), None)


def _active_lot_get(key: str, default=None):
    """Lit une valeur dans le lot actif (retourne default si pas de lot actif)."""
    lot = _get_active_lot()
    return lot.get(key, default) if lot is not None else default


def _active_lot_set(key: str, value) -> None:
    """Ecrit une valeur dans le lot actif (no-op si pas de lot actif)."""
    lot = _get_active_lot()
    if lot is not None:
        lot[key] = value


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _safe_save(uploaded_file, allowed_extensions=None):
    """Sauvegarde un fichier uploadé après validation complète."""
    # SEC-5 : Validation extension + taille + magic bytes
    is_valid, error_msg = validate_uploaded_file(
        uploaded_file,
        max_mb=MAX_FILE_SIZE_MB,
        allowed_extensions=allowed_extensions or ALLOWED_EXTENSIONS,
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


def rebuild_merged_tco(tva_rate=TVA_DEFAULT, new_companies: list[str] | None = None):
    """Reconstruit ou met a jour le TCO fusionne pour le lot actif.

    Si new_companies est fourni ET que merged_df existe deja,
    seules ces nouvelles entreprises sont ajoutees (fusion incrementale O(1)
    au lieu de tout refusionner depuis zero).
    Si new_companies est None ou merged_df absent, reconstruction complete.
    """
    tco_df = _active_lot_get("tco_df")
    if tco_df is None:
        return

    companies = _active_lot_get("companies", {})
    merged_df_cur = _active_lot_get("merged_df")

    if new_companies and merged_df_cur is not None:
        # Fusion incrementale : ajoute uniquement les nouvelles entreprises
        merged = merged_df_cur.copy()
        all_alerts = list(_active_lot_get("all_alerts") or [])
        for comp_name in new_companies:
            comp_data = companies[comp_name]
            merged, merge_alerts = merge_company_into_tco(
                merged, comp_data["dpgf_df"], comp_name, tva_rate=tva_rate
            )
            for alert in comp_data.get("parse_alerts", []):
                alert["company"] = comp_name
            for alert in merge_alerts:
                alert["company"] = comp_name
            all_alerts.extend(comp_data.get("parse_alerts", []))
            all_alerts.extend(merge_alerts)
        _active_lot_set("merged_df", merged)
        _active_lot_set("all_alerts", all_alerts)
    else:
        # Reconstruction complete depuis le TCO de base
        merged_df, all_alerts = merge_all_companies(
            tco_df,
            companies,
            tva_rate=tva_rate,
        )
        _active_lot_set("merged_df", merged_df)
        _active_lot_set("all_alerts", all_alerts)


@st.cache_data(ttl=5)
def _cached_list_projects() -> list[str]:
    """Liste des projets avec cache TTL=5 s pour éviter le scan disque à chaque rerun."""
    return list_projects()


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
    cols = [c for c in df.columns if c not in hidden]
    hidden_types = {"empty", "recap", "recap_summary", "total_line", "total_text"}
    visible = df[~df["row_type"].isin(hidden_types)][cols]
    st.write(f"**{title}** ({len(visible)} lignes)")
    st.dataframe(visible, width="stretch", hide_index=True, height=500)


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

if st.session_state.get("active_project") is not None:
    with st.sidebar:
        # Logo
        if os.path.exists("odetec_logo.png"):
            st.image("odetec_logo.png", width="stretch")

        # Nom du projet + lot actif — bloc uni en haut de la sidebar
        curr_name = (st.session_state.get("active_project") or {}).get("project_name", "Sans titre")
        lot_label = _active_lot_get("lot_label", "")

        if lot_label:
            # Projet en haut (grand), lot en dessous (plus petit) — bloc uni
            st.markdown(
                f"<div style='margin-bottom: 0.75rem;'>"
                f"<div style='"
                f"background: linear-gradient(135deg, #2F5496, #4472C4);"
                f"color: white; padding: 12px 16px;"
                f"border-radius: 10px;"
                f"font-weight: 700; font-size: 1.0rem; word-break: break-word;'>"
                f"📁 {html_mod.escape(curr_name)}</div>"
                f"<div style='height: 5px;'></div>"
                f"<div style='"
                f"background: #17375E;"
                f"color: #FFD700; padding: 7px 16px;"
                f"border-radius: 10px;"
                f"font-weight: 600; font-size: 0.78rem; word-break: break-word;"
                f"border-left: 4px solid #FFD700;"
                f"letter-spacing: 0.02em;'>"
                f"🏗️ {html_mod.escape(lot_label)}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )
        else:
            # Pas de lot actif : projet seul
            st.markdown(
                f"<div style='margin-bottom: 0.75rem;'>"
                f"<div style='"
                f"background: linear-gradient(135deg, #2F5496, #4472C4);"
                f"color: white; padding: 12px 16px; border-radius: 10px;"
                f"font-weight: 700; font-size: 1.0rem; word-break: break-word;'>"
                f"📁 {html_mod.escape(curr_name)}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )

        st.markdown("---")

        # Bouton Enregistrer
        if st.button("💾 Enregistrer", width="stretch", key="sidebar_save"):
            ok, msg = save_project(curr_name, st.session_state)
            if ok:
                _cached_list_projects.clear()
                st.success(msg)
            else:
                st.error(msg)

        st.markdown("---")

        # Retour a l'accueil
        if st.button("🏠 Retour a l'accueil", width="stretch", type="primary"):
            st.session_state.active_project = None
            st.session_state.active_lot_id = None
            st.session_state.step = 0
            st.session_state.pop("export_buffer", None)
            st.rerun()

        st.markdown("---")

        # Fermer l'application — réservé à l'administrateur
        # Confirmation deux étapes : le warning s'affiche AVANT le kill
        # (st.warning() est bufférisé et ne s'affiche pas si os.kill() est appelé
        # dans le même cycle — d'où le rerun intermédiaire)
        if ADMIN_MODE:
            if not st.session_state.confirm_shutdown:
                if st.button(
                    "❌ Fermer l'application",
                    width="stretch",
                    help="Arrête le serveur Streamlit (admin uniquement)",
                ):
                    st.session_state.confirm_shutdown = True
                    st.rerun()
            else:
                st.warning("⚠️ Arrêt du serveur — sauvegardez vos données.")
                col_ok, col_no = st.columns(2)
                with col_ok:
                    if st.button("✅ Confirmer", width="stretch"):
                        os.kill(os.getpid(), signal.SIGTERM)
                with col_no:
                    if st.button("✗ Annuler", width="stretch"):
                        st.session_state.confirm_shutdown = False
                        st.rerun()

is_dark = st.session_state.dark_mode

# ---------------------------------------------------------------------------
# CSS — Design system (extrait dans app/__init__.py)
# ---------------------------------------------------------------------------

st.markdown(
    get_full_css(
        is_dark,
        hide_sidebar=(
            st.session_state.step == 0 and st.session_state.get("active_project") is None
        ),
    ),
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# STEP 0 — Landing Page
# ---------------------------------------------------------------------------


def _render_project_lots_view(proj: dict) -> None:
    """Affiche la vue gestion des lots du projet actif."""
    proj_name = proj.get("project_name", "Sans titre")
    st.markdown(f"## 📁 {html_mod.escape(proj_name)}")
    st.markdown("---")

    lots = proj.get("lots", [])
    if lots:
        st.write(f"**{len(lots)} lot(s) dans ce projet :**")
        for lot in list(lots):
            col_lbl, col_del = st.columns([8, 1])
            with col_lbl:
                label = lot.get("lot_label", "Lot sans nom")
                has_data = lot.get("tco_df") is not None
                icon = "✅" if has_data else "📋"
                if st.button(
                    f"{icon} {html_mod.escape(label)}",
                    key=f"open_lot_{lot['lot_id']}",
                    width="stretch",
                ):
                    st.session_state.active_lot_id = lot["lot_id"]
                    st.session_state.pop("export_buffer", None)
                    # Atterrir sur la bonne etape selon l'avancement du lot
                    if lot.get("tco_df") is not None:
                        st.session_state.step = 3 if lot.get("companies") else 1
                    else:
                        st.session_state.step = 1
                    st.rerun()
            with col_del:
                if st.button("🗑️", key=f"del_lot_{lot['lot_id']}", help=f"Supprimer {label}"):
                    proj["lots"] = [lt for lt in lots if lt["lot_id"] != lot["lot_id"]]
                    if st.session_state.active_lot_id == lot["lot_id"]:
                        st.session_state.active_lot_id = None
                    st.rerun()
        st.divider()
    else:
        st.info("Aucun lot dans ce projet. Ajoutez un premier lot ci-dessous.")

    col_lot_name, col_lot_btn = st.columns([3, 1])
    with col_lot_name:
        new_lot_label = st.text_input(
            "Nom du lot",
            placeholder="Ex: GROS OEUVRE",
            key="new_lot_name_input",
            label_visibility="collapsed",
        )
    with col_lot_btn:
        add_lot_clicked = st.button("➕ Ajouter un lot", type="primary", width="stretch")

    if add_lot_clicked:
        new_lot_label = new_lot_label.strip() if new_lot_label.strip() else "Nouveau lot"
        new_lot = {
            "lot_id": uuid.uuid4().hex,
            "lot_label": new_lot_label,
            "lot_num": "",
            "tco_df": None,
            "tco_meta": None,
            "tva_rate": TVA_DEFAULT,
            "merged_df": None,
            "all_alerts": [],
            "companies": {},
        }
        proj["lots"].append(new_lot)
        st.session_state.active_lot_id = new_lot["lot_id"]
        st.session_state.pop("export_buffer", None)
        st.session_state.step = 1
        st.rerun()


if st.session_state.step == 0:
    if st.session_state.active_project is None:
        # --- Landing page ---
        st.markdown("<div style='height: 4rem;'></div>", unsafe_allow_html=True)

        col_logo_left, col_logo_mid, col_logo_right = st.columns([1, 2, 1])
        with col_logo_mid:
            if os.path.exists("odetec_logo.png"):
                st.image("odetec_logo.png", width="stretch")
            st.markdown(f"<h1 class='main-title'>{APP_TITLE}</h1>", unsafe_allow_html=True)
            st.markdown(
                "<p class='subtitle'>Solution intelligente pour la consolidation des DPGF et le remplissage du TCO.</p>",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 2rem;'></div>", unsafe_allow_html=True)

        col1, col2 = st.columns(2)

        with col1:
            with st.container(border=True):
                st.markdown("### 🆕 Nouveau Projet")
                st.markdown(
                    "<p>Commencez une nouvelle analyse en important un DPGF vierge.</p>",
                    unsafe_allow_html=True,
                )
                new_proj_name = st.text_input(
                    "Nom du projet",
                    placeholder="Ex: Chantier Bordeaux",
                    key="landing_new_proj_name",
                )

                st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
                if st.button("🚀 Creer le projet", type="primary", width="stretch"):
                    if new_proj_name:
                        st.session_state.active_project = {
                            "project_id": uuid.uuid4().hex,
                            "project_name": new_proj_name,
                            "created_at": datetime.now().isoformat(),
                            "lots": [],
                        }
                        st.session_state.active_lot_id = None
                        # step reste 0 -> affichera _render_project_lots_view
                        st.rerun()
                    else:
                        st.warning("Veuillez saisir un nom de projet.")

        with col2:
            with st.container(border=True):
                st.markdown("### 📂 Ouvrir un Projet")
                st.markdown(
                    "<p>Reprenez un travail en cours depuis vos sauvegardes locales.</p>",
                    unsafe_allow_html=True,
                )
                projects = _cached_list_projects()
                if projects:
                    for p in projects:
                        rcol_name, rcol_del = st.columns([7, 1])
                        with rcol_name:
                            if st.button(f"📄 {p}", key=f"landing_load_{p}", width="stretch"):
                                ok, msg = load_project(p, st.session_state)
                                if ok:
                                    st.session_state.pop("export_buffer", None)
                                    st.rerun()
                                else:
                                    st.error(msg)
                        with rcol_del:
                            if st.button("🗑️", key=f"landing_del_{p}", help=f"Supprimer {p}"):
                                if delete_project(p):
                                    _cached_list_projects.clear()
                                    st.rerun()
                else:
                    st.caption("Aucun projet sauvegarde pour le moment.")
    else:
        # --- Vue lots du projet ---
        _render_project_lots_view(st.session_state.active_project)


# ---------------------------------------------------------------------------
# Header + progress (Visible only after Step 0)
# ---------------------------------------------------------------------------

if st.session_state.step > 0:
    _proj_name_hdr = (st.session_state.get("active_project") or {}).get(
        "project_name", "Sans titre"
    )
    _lot_label_hdr = _active_lot_get("lot_label", "")
    _subtitle_hdr = f"{_proj_name_hdr} — {_lot_label_hdr}" if _lot_label_hdr else _proj_name_hdr
    st.markdown(
        f"<h1 class='main-title' style='font-size: 1.8rem; text-align: left;'>"
        f"{html_mod.escape(APP_TITLE)} "
        f"<span style='color: var(--text-muted); font-size: 1.2rem; font-weight: 400;'>"
        f"| {html_mod.escape(_subtitle_hdr)}</span></h1>",
        unsafe_allow_html=True,
    )
    st.divider()


# ---------------------------------------------------------------------------
# STEP 1 — Import TCO
# ---------------------------------------------------------------------------

if st.session_state.step >= 1:
    st.markdown(
        "<div class='step-header'>📥 Etape 1 : importer le DPGF vierge</div>",
        unsafe_allow_html=True,
    )
    st.caption(
        "Fichier DPGF LOT (.xlsx) — Colonnes : Code | Designation | Qu. | U. | Px U. | Px tot."
    )

    tco_file = st.file_uploader(
        "Charger Le DPGF Modele",
        type=_UPLOADER_TYPES,
        key="tco_upload",
        help="Fichier DPGF LOT servant de base",
        label_visibility="visible",
    )

    if tco_file and _active_lot_get("tco_df") is None:
        path = _safe_save(tco_file)
        if path:
            with st.spinner("Lecture du DPGF..."):
                try:
                    tco_df, meta = parse_tco(path)

                    # Recaler les totaux de l'estimation (colonne de base)
                    tva_rate_cur = _active_lot_get("tva_rate", TVA_DEFAULT)
                    compute_section_totals(tco_df, "Px_Tot_HT", tva_rate=tva_rate_cur)

                    _active_lot_set("tco_df", tco_df)
                    _active_lot_set("tco_meta", meta)
                    _active_lot_set("merged_df", tco_df.copy())

                    # Mettre a jour le label et le numero de lot depuis les metadonnees
                    lot_raw = ((meta.get("project_info") or {}).get("lot") or "").strip()
                    m = re.search(r"\b(\d{2})\b", lot_raw)
                    _active_lot_set("lot_label", lot_raw or "LOT INCONNU")
                    _active_lot_set("lot_num", m.group(1) if m else "")

                    info = meta["project_info"]
                    if info:
                        st.info(
                            f"📋 **Projet :** {info.get('projet', 'N/A')} — "
                            f"**Lot :** {info.get('lot', 'N/A')}"
                        )
                    template_name = os.path.splitext(tco_file.name)[0]
                    st.success(f"✅ {template_name} charge — {len(tco_df)} lignes")
                except Exception as e:
                    log.error("Erreur parsing TCO", exc_info=True)
                    st.error(f"❌ Erreur de lecture : {e}")
                finally:
                    _cleanup_file(path)

    if _active_lot_get("tco_df") is not None:
        # UX-1 : Selecteur TVA — affecte les calculs HT/TVA/TTC du lot
        tva_labels = list(TVA_OPTIONS.keys())
        tva_rate_lot = _active_lot_get("tva_rate", TVA_DEFAULT)
        current_tva_label = next(
            (k for k, v in TVA_OPTIONS.items() if v == tva_rate_lot),
            tva_labels[-1],
        )
        selected_tva_label = st.selectbox(
            "Taux de TVA applicable",
            options=tva_labels,
            index=tva_labels.index(current_tva_label),
            help="5,5 % — renovation residentielle | 10 % — renovation | 20 % — neuf/defaut",
        )
        new_tva = TVA_OPTIONS[selected_tva_label]
        if new_tva != tva_rate_lot:
            _active_lot_set("tva_rate", new_tva)
            if _active_lot_get("companies", {}):
                rebuild_merged_tco(new_tva)
            st.rerun()

        if st.session_state.step == 1:
            if st.button("➡️ Passer a l'etape suivante", type="primary"):
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

    # UX-4 : Confirmation de suppression
    if st.session_state.confirm_remove:
        to_remove = st.session_state.confirm_remove
        st.warning(f"⚠️ Voulez-vous vraiment supprimer **{to_remove}** ?")
        col_y, col_n = st.columns([1, 5])
        with col_y:
            if st.button("✅ Oui, supprimer", type="primary"):
                companies_cur = _active_lot_get("companies", {})
                companies_cur.pop(to_remove, None)
                _active_lot_set("companies", companies_cur)
                st.session_state.confirm_remove = None
                rebuild_merged_tco(_active_lot_get("tva_rate", TVA_DEFAULT))
                st.session_state.pop("export_buffer", None)
                st.session_state.upload_counter += 1
                st.session_state._flash_msg = f"✅ Entreprise **{to_remove}** supprimee."
                st.rerun()
        with col_n:
            if st.button("❌ Annuler"):
                st.session_state.confirm_remove = None
                st.rerun()
        st.divider()

    companies_lot = _active_lot_get("companies", {})
    n_companies = len(companies_lot)
    if n_companies:
        st.write(f"**{n_companies} / {MAX_COMPANIES} entreprise(s) importee(s) :**")
        for comp_name in list(companies_lot.keys()):
            comp = companies_lot[comp_name]
            if "n_articles" not in comp:
                comp["n_articles"] = int((comp["dpgf_df"]["row_type"] == "article").sum())
            n_art = comp["n_articles"]
            n_alrt = len(comp["parse_alerts"])

            col_inf, col_btn = st.columns([4, 1])
            with col_inf:
                st.markdown(
                    f"<div class='company-card'>🏢 <b>{html_mod.escape(comp_name)}</b> — "
                    f"{n_art} articles, {n_alrt} alerte(s) "
                    f"<i>({html_mod.escape(comp['filename'])})</i></div>",
                    unsafe_allow_html=True,
                )
            with col_btn:
                if st.button("🗑️ Retirer", key=f"rm_{comp_name}"):
                    st.session_state.confirm_remove = comp_name
                    st.rerun()
        st.divider()

    if n_companies >= MAX_COMPANIES:
        st.warning(f"⚠️ Limite de {MAX_COMPANIES} entreprises atteinte.")
    else:
        # UX-5 : Multi-upload — cle dynamique pour forcer la reinitialisation apres suppression
        dpgf_files = st.file_uploader(
            "Importer un ou plusieurs DPGF entreprise",
            type=_DPGF_UPLOADER_TYPES,
            key=f"multi_dpgf_upload_{st.session_state.upload_counter}",
            accept_multiple_files=True,
            help="Selectionnez tous les fichiers DPGF des entreprises a fusionner (.xlsx, .xls, .xlsb ou .pdf)",
        )

        if dpgf_files:
            companies_lot = _active_lot_get("companies", {})
            processed_filenames = {v["filename"] for v in companies_lot.values()}
            new_files = [f for f in dpgf_files if f.name not in processed_filenames]

            if new_files:
                success_count = 0
                added_companies: list[str] = []
                for dpgf_file in new_files:
                    filename_clean = os.path.splitext(dpgf_file.name)[0]
                    company_name = (
                        re.sub(r"^DPGF\s+", "", filename_clean, flags=re.IGNORECASE).strip().upper()
                    )

                    path = _safe_save(dpgf_file, allowed_extensions=DPGF_ALLOWED_EXTENSIONS)
                    if path:
                        with st.spinner(f"Fusion de {company_name}..."):
                            try:
                                if dpgf_file.name.lower().endswith(".pdf"):
                                    dpgf_df, parse_alerts = parse_dpgf_pdf(path)
                                else:
                                    dpgf_df, parse_alerts = parse_dpgf(path)
                                if dpgf_df.empty:
                                    err_msgs = [
                                        a["message"] for a in parse_alerts if a["type"] == "error"
                                    ]
                                    msg = (
                                        err_msgs[0]
                                        if err_msgs
                                        else "Fichier invalide ou non reconnu"
                                    )
                                    st.error(f"❌ {company_name} : {msg}")
                                else:
                                    companies_lot[company_name] = {
                                        "dpgf_df": dpgf_df,
                                        "parse_alerts": parse_alerts,
                                        "filename": dpgf_file.name,
                                        "n_articles": int((dpgf_df["row_type"] == "article").sum()),
                                    }
                                    added_companies.append(company_name)
                                    success_count += 1
                            except Exception as e:
                                log.error("Erreur fusion %s", company_name, exc_info=True)
                                st.error(f"❌ Erreur sur {company_name}: {e}")
                            finally:
                                _cleanup_file(path)

                if success_count > 0:
                    _active_lot_set("companies", companies_lot)
                    rebuild_merged_tco(
                        _active_lot_get("tva_rate", TVA_DEFAULT),
                        new_companies=added_companies,
                    )
                    st.session_state.pop("export_buffer", None)
                    st.session_state.step = 3
                    st.session_state.export_done = False
                    st.rerun()

    if st.button("🔄 Reinitialiser ce lot"):
        # Reinitialise uniquement le lot actif — les autres lots du projet sont preserves
        lot_cur = _get_active_lot()
        if lot_cur is not None:
            lot_cur.update(
                {
                    "tco_df": None,
                    "tco_meta": None,
                    "lot_label": "Nouveau lot",
                    "lot_num": "",
                    "merged_df": None,
                    "all_alerts": [],
                    "companies": {},
                    "tva_rate": TVA_DEFAULT,
                }
            )
        st.session_state.step = 1
        st.session_state.upload_counter = 0
        st.session_state.confirm_remove = None
        st.session_state.export_done = False
        st.session_state.pop("export_buffer", None)
        st.rerun()


# ---------------------------------------------------------------------------
# STEP 3 — Résultat & Export
# ---------------------------------------------------------------------------

if st.session_state.step >= 3:
    st.markdown(
        "<div class='step-header'>📊 Étape 3 — Résultat Final</div>",
        unsafe_allow_html=True,
    )

    st.markdown(
        """
    <div class='legend-box'>
        <div class='legend-item'><span class='color-dot' style='background:#FFC7CE'></span> Erreur</div>
        <div class='legend-item'><span class='color-dot' style='background:#FFE4B5'></span> Avertissement</div>
        <div class='legend-item'><span class='color-dot' style='background:#FFFFCC'></span> Note</div>
        <div class='legend-item'><span class='color-dot' style='background:#D6EAF8'></span> Info</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    merged = _active_lot_get("merged_df")
    all_alerts = _active_lot_get("all_alerts", [])
    tco_meta = _active_lot_get("tco_meta") or {}
    companies_s3 = _active_lot_get("companies", {})

    # Stats
    art_rows = merged[merged["row_type"] == "article"]

    cols = st.columns(3)
    with cols[0]:
        st.metric("📋 Articles", len(art_rows))
    with cols[1]:
        st.metric("🏢 Entreprises", len(companies_s3))
    with cols[2]:
        st.metric("⚠️ Alertes", len(all_alerts))

    # Resume des anomalies par categorie
    if all_alerts:
        n_err = sum(1 for a in all_alerts if a.get("type") == "error")
        n_warn = sum(1 for a in all_alerts if a.get("type") == "warning")
        n_info = sum(1 for a in all_alerts if a.get("type") == "info")
        parts = []
        if n_err:
            parts.append(f"🔴 {n_err} erreur(s)")
        if n_warn:
            parts.append(f"🟡 {n_warn} avertissement(s)")
        if n_info:
            parts.append(f"🔵 {n_info} info(s)")
        st.caption("Anomalies : " + " — ".join(parts))

    display_preview(merged, "TCO Final Consolide")

    if all_alerts:
        display_alerts(all_alerts, "Toutes les alertes")

    col_b1, col_b2 = st.columns(2)
    with col_b1:
        if st.button("📁 Retour aux lots du projet"):
            st.session_state.active_lot_id = None
            st.session_state.step = 0
            st.session_state.pop("export_buffer", None)
            st.rerun()
    with col_b2:
        if st.button("⬅️ Retour — Modifier les entreprises"):
            st.session_state.step = 2
            st.session_state.export_done = False
            st.rerun()

    st.divider()
    # Nom du fichier : TCO_FINAL_<projet>_<lot>.xlsx
    _proj_raw = (st.session_state.get("active_project") or {}).get("project_name", "") or ""
    _lot_raw = tco_meta.get("project_info", {}).get("lot", "") or ""
    if not _lot_raw:
        _lot_raw = _active_lot_get("lot_label", "") or ""

    def _normalize(s: str) -> str:
        norm = re.sub(r"[^A-Z0-9]", "_", s.upper())
        return re.sub(r"_+", "_", norm).strip("_")

    _proj_norm = _normalize(_proj_raw)
    _lot_norm = _normalize(_lot_raw)

    if _proj_norm and _lot_norm:
        filename = f"TCO_FINAL_{_proj_norm}_{_lot_norm}.xlsx"
    elif _proj_norm:
        filename = f"TCO_FINAL_{_proj_norm}.xlsx"
    elif _lot_norm:
        filename = f"TCO_FINAL_{_lot_norm}.xlsx"
    else:
        filename = "TCO_FINAL.xlsx"

    # Pre-generation du buffer pour telechargement immediat
    try:
        if "export_buffer" not in st.session_state:
            with st.spinner("Preparation du fichier..."):
                # Project metadata injection removed

                st.session_state.export_buffer = export_tco(
                    merged,
                    tco_meta,
                    output_path=None,
                    alerts=all_alerts,
                    tva_rate=_active_lot_get("tva_rate", TVA_DEFAULT),
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
            width="stretch",
        )

        if st.session_state.get("export_done"):
            st.success("✅ Telechargement OK")
            log.info("Export telecharge : %s", filename)

    except Exception as e:
        log.error("Erreur preparation export", exc_info=True)
        st.error(f"❌ Erreur de generation : {e}")

st.divider()
st.caption(f"{APP_TITLE} v{APP_VERSION} — Export du TCO")
