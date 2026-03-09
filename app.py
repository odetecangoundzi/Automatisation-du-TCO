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
from app.controllers import normalize_filename as _normalize_filename
from app.controllers import rebuild_merged_tco as _ctrl_rebuild
from config import (
    ADMIN_MODE,
    ALLOWED_EXTENSIONS,
    APP_ICON,
    APP_TITLE,
    APP_VERSION,
    MAX_COMPANIES,
    MAX_FILE_SIZE_MB,
    PROJECTS_DIR,
    TVA_DEFAULT,
    TVA_OPTIONS,
    UPLOAD_DIR,
)
from core.exporter import export_tco
from core.merger import compute_section_totals
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
    "_last_autosave": None,  # Horodatage de la dernière auto-sauvegarde
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PROJECTS_DIR, exist_ok=True)
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


def _safe_save(uploaded_file, allowed_extensions: set[str] | None = None) -> str | None:
    """Sauvegarde un fichier uploadé après validation complète (extension, taille, magic bytes).

    Returns:
        Chemin absolu du fichier sauvegardé, ou None si la validation échoue.
    """
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


def rebuild_merged_tco(tva_rate=TVA_DEFAULT, new_companies: list[str] | None = None) -> None:
    """Wrapper Streamlit : délègue au controller et écrit dans le lot actif.

    Appelle app.controllers.rebuild_merged_tco() (pur Python, testable)
    puis met à jour session_state via _active_lot_set().
    """
    tco_df = _active_lot_get("tco_df")

    if tco_df is None:
        # Mode comparatif : pas de DPGF estimation — on utilise le 1er DPGF entreprise
        # comme structure de base (codes + désignations + hiérarchie), prix vidés.
        companies_data = _active_lot_get("companies", {})
        if not companies_data:
            return
        first_comp = next(iter(companies_data.values()))
        tco_df = first_comp["dpgf_df"].copy()
        for col in ("Qu.", "Px_U_HT", "Px_Tot_HT"):
            if col in tco_df.columns:
                tco_df[col] = None

    merged_df, alerts = _ctrl_rebuild(
        tco_df,
        _active_lot_get("companies", {}),
        tva_rate,
        merged_df=_active_lot_get("merged_df"),
        new_companies=new_companies,
    )
    _active_lot_set("merged_df", merged_df)
    if new_companies:
        existing = list(_active_lot_get("all_alerts") or [])
        _active_lot_set("all_alerts", existing + alerts)
    else:
        _active_lot_set("all_alerts", alerts)


def _on_export_click() -> None:
    """Callback bouton téléchargement — marque l'export comme effectué."""
    st.session_state.export_done = True


def _autosave() -> None:
    """Sauvegarde silencieuse du projet actif.

    Appelée après chaque action importante (import, fusion, ajout entreprise)
    pour éviter la perte de données en cas de déconnexion ou de timeout Render.
    """
    proj = st.session_state.get("active_project")
    if proj is None:
        return
    name = proj.get("project_name", "")
    if not name:
        return
    ok, _ = save_project(name, st.session_state)
    if ok:
        _cached_list_projects.clear()
        st.session_state["_last_autosave"] = datetime.now().strftime("%H:%M:%S")


@st.cache_data(ttl=5)
def _cached_list_projects() -> list[str]:
    """Liste des projets avec cache TTL=5 s pour éviter le scan disque à chaque rerun."""
    return list_projects()


def display_alerts(alerts: list[dict], title: str = "Alertes") -> None:
    """Affiche les alertes dans un expander Streamlit avec icône par sévérité."""
    if not alerts:
        st.success("✅ Aucune anomalie détectée")
        return

    with st.expander(f"📋 {title} — détails", expanded=False):
        for a in alerts:
            icon = {"error": "🔴", "warning": "🟡", "info": "🔵"}.get(a["type"], "ℹ️")
            st.write(f"{icon} **{a.get('code', '')}** — {a.get('message', '')}")


def display_preview(df, title: str = "Aperçu") -> None:
    """Affiche un aperçu éditable du DataFrame TCO."""
    hidden = {
        "Entete",
        "row_type",
        "original_row",
        "parent_code",
        "is_extra_line",
        "skip_sum",
        "is_added",
    }
    cols = [c for c in df.columns if c not in hidden]
    hidden_types = {"empty", "recap", "recap_summary", "total_line", "total_text"}
    # Masquer les lignes techniques pour l'affichage
    display_df = df[~df["row_type"].isin(hidden_types)][cols]
    st.write(f"**{title}** ({len(display_df)} lignes)")

    # Configuration des colonnes éditables (Uniquement Qu, PU et Commentaire pour les entreprises)
    # Les colonnes de base (A-F) sont en lecture seule pour préserver le modèle.
    base_cols = {"Code", "Désignation", "Qu.", "U", "Px_U_HT", "Px_Tot_HT", "Commentaire"}
    column_config = {}
    for c in cols:
        if c in base_cols:
            column_config[c] = st.column_config.Column(disabled=True)
        elif "_Px_Tot_HT" in c:
            # Px_Tot_HT est calculé, donc lecture seule
            column_config[c] = st.column_config.NumberColumn(disabled=True)
        elif "_Qu." in c or "_Px_U_HT" in c:
            column_config[c] = st.column_config.NumberColumn(disabled=False)
        elif str(display_df[c].dtype) == "bool" or (
            display_df[c].dtype == object and display_df[c].dropna().map(type).eq(bool).all()
        ):
            column_config[c] = st.column_config.CheckboxColumn(disabled=True)
        else:
            column_config[c] = st.column_config.TextColumn(disabled=False)

    st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        height=500,
        column_config=column_config,
        key="tco_main_editor",
    )

    # Analyse des changements de st.data_editor via session_state
    if "tco_main_editor" in st.session_state:
        changes = st.session_state["tco_main_editor"].get("edited_rows")
        if changes:
            # Appliquer les changements au DataFrame source (merged_df du lot actif)
            # On a besoin de mapper l'index de display_df à l'index de merged_df
            # L'index original est préservé dans display_df (it's a view)
            lot = _get_active_lot()
            if lot is not None and lot.get("merged_df") is not None:
                source_df = lot["merged_df"]
                any_change = False
                for row_idx_display, row_changes in changes.items():
                    # row_idx_display est l'index relatif dans display_df (0 à N)
                    # On récupère l'index réel de Pandas
                    real_idx = display_df.index[row_idx_display]
                    for col_name, new_val in row_changes.items():
                        source_df.at[real_idx, col_name] = new_val
                        any_change = True

                if any_change:
                    import pandas as pd_internal

                    # Recalculer les totaux après modification
                    compute_section_totals(
                        source_df, "Px_Tot_HT", tva_rate=lot.get("tva_rate", TVA_DEFAULT)
                    )
                    # Recalculer pour chaque entreprise (colonne Px_Tot_HT)
                    for col in source_df.columns:
                        if col.endswith("_Px_Tot_HT"):
                            # Recalculer ligne à ligne Qu * PU pour la ligne modifiée ?
                            # Plus simple : recalculer tout le bloc pour l'entreprise
                            company = col.replace("_Px_Tot_HT", "")
                            qu_col = f"{company}_Qu."
                            pu_col = f"{company}_Px_U_HT"
                            source_df[col] = pd_internal.to_numeric(
                                source_df[qu_col], errors="coerce"
                            ).fillna(0) * pd_internal.to_numeric(
                                source_df[pu_col], errors="coerce"
                            ).fillna(0)
                            compute_section_totals(
                                source_df, col, tva_rate=lot.get("tva_rate", TVA_DEFAULT)
                            )

                    _active_lot_set("merged_df", source_df)
                    st.rerun()


def _cleanup_file(path: str) -> None:
    """Supprime un fichier temporaire de manière sûre (SEC-3)."""
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
            with open("odetec_logo.png", "rb") as f:
                st.image(f.read(), use_container_width=True)

        # Nom du projet — bloc en haut de la sidebar
        curr_name = (st.session_state.get("active_project") or {}).get("project_name", "Sans titre")
        st.markdown(
            f"<div style='margin-bottom: 0.5rem;'>"
            f"<div style='"
            f"background: linear-gradient(135deg, #2F5496, #4472C4);"
            f"color: white; padding: 10px 14px; border-radius: 10px;"
            f"font-weight: 700; font-size: 0.95rem; word-break: break-word;'>"
            f"📁 {html_mod.escape(curr_name)}</div>"
            f"</div>",
            unsafe_allow_html=True,
        )

        # ── Liste des lots ──────────────────────────────────
        lots_all = (st.session_state.get("active_project") or {}).get("lots", [])
        active_lid = st.session_state.get("active_lot_id")

        if lots_all:
            st.markdown(
                "<div style='font-size:0.75rem; text-transform:uppercase; "
                "letter-spacing:0.08em; color:#8899aa; margin: 0.4rem 0 0.25rem;'>"
                "Lots du projet</div>",
                unsafe_allow_html=True,
            )
            for lot in lots_all:
                is_active = lot["lot_id"] == active_lid
                label = lot.get("lot_label", "Lot sans nom")
                has_data = lot.get("tco_df") is not None
                icon = "✅" if has_data else "📋"

                col_btn, col_del = st.sidebar.columns([7, 1])
                with col_btn:
                    btn_type = "primary" if is_active else "secondary"
                    if st.button(
                        f"{icon} {label}",
                        key=f"sb_lot_{lot['lot_id']}",
                        use_container_width=True,
                        type=btn_type,
                    ):
                        if not is_active:
                            st.session_state.active_lot_id = lot["lot_id"]
                            st.session_state.pop("export_buffer", None)
                            if lot.get("tco_df") is not None:
                                st.session_state.step = 3 if lot.get("companies") else 1
                            else:
                                st.session_state.step = 1
                            st.rerun()
                with col_del:
                    if st.button(
                        "🗑️",
                        key=f"sb_del_lot_{lot['lot_id']}",
                        help=f"Supprimer le lot {label}",
                    ):
                        proj_cur = st.session_state.get("active_project", {})
                        proj_cur["lots"] = [lt for lt in lots_all if lt["lot_id"] != lot["lot_id"]]
                        if active_lid == lot["lot_id"]:
                            st.session_state.active_lot_id = None
                            st.session_state.step = 0
                        st.rerun()

        st.markdown("---")

        # Nouveau lot
        if st.button("➕ Nouveau lot", use_container_width=True, key="sidebar_new_lot"):
            st.session_state.active_lot_id = None
            st.session_state.step = 0
            st.session_state.pop("export_buffer", None)
            st.rerun()

        # Bouton Enregistrer (manuel)
        if st.button("💾 Enregistrer", use_container_width=True, key="sidebar_save"):
            ok, msg = save_project(curr_name, st.session_state)
            if ok:
                _cached_list_projects.clear()
                st.success(msg)
            else:
                st.error(msg)

        # Indicateur de dernière sauvegarde automatique
        if st.session_state.get("_last_autosave"):
            st.caption(f"🟢 Auto-sauvegardé à {st.session_state['_last_autosave']}")

        st.markdown("---")

        # Retour a l'accueil
        if st.button("🏠 Retour a l'accueil", use_container_width=True, type="primary"):
            save_project(curr_name, st.session_state)  # save before leaving
            _cached_list_projects.clear()
            st.session_state.active_project = None
            st.session_state.active_lot_id = None
            st.session_state.step = 0
            st.session_state.pop("export_buffer", None)
            st.rerun()

        st.markdown("---")

        # Fermer l'application — réservé à l'administrateur
        if ADMIN_MODE:
            if not st.session_state.confirm_shutdown:
                if st.button(
                    "❌ Fermer l'application",
                    use_container_width=True,
                    help="Arrête le serveur Streamlit (admin uniquement)",
                ):
                    st.session_state.confirm_shutdown = True
                    st.rerun()
            else:
                st.warning("⚠️ Arrêt du serveur — sauvegardez vos données.")
                col_ok, col_no = st.columns(2)
                with col_ok:
                    if st.button("✅ Confirmer", use_container_width=True):
                        os.kill(os.getpid(), signal.SIGTERM)
                with col_no:
                    if st.button("✗ Annuler", use_container_width=True):
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
                    use_container_width=True,
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
        add_lot_clicked = st.button("➕ Ajouter un lot", type="primary", use_container_width=True)

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
        _autosave()  # Auto-sauvegarde à la création du lot
        st.rerun()


if st.session_state.step == 0:
    if st.session_state.active_project is None:
        # --- Landing page ---
        st.markdown("<div style='height: 4rem;'></div>", unsafe_allow_html=True)

        col_logo_left, col_logo_mid, col_logo_right = st.columns([1, 2, 1])
        with col_logo_mid:
            if os.path.exists("odetec_logo.png"):
                with open("odetec_logo.png", "rb") as f:
                    st.image(f.read(), use_container_width=True)
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
                if st.button("🚀 Creer le projet", type="primary", use_container_width=True):
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
                            if st.button(
                                f"📄 {p}", key=f"landing_load_{p}", use_container_width=True
                            ):
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

    # Détection du retrait via le X du widget : si le fichier est supprimé du widget
    # mais que tco_df est encore en session, on remet à zéro pour que le bouton
    # "Mode comparatif" réapparaisse.
    # Garde step == 1 : ne pas effacer les données d'un projet chargé quand l'utilisateur
    # est déjà à l'étape 2 ou 3 (le widget est toujours vide au rechargement de session).
    if not tco_file and _active_lot_get("tco_df") is not None and st.session_state.step == 1:
        _active_lot_set("tco_df", None)
        _active_lot_set("tco_meta", None)
        _active_lot_set("merged_df", None)
        st.rerun()

    if tco_file and _active_lot_get("tco_df") is None:
        path = _safe_save(tco_file)
        if path:
            with st.status(f"Import du DPGF modèle : {tco_file.name}...", expanded=True) as status:
                try:
                    status.write("Analyse de la structure du fichier...")
                    tco_df, meta = parse_tco(path)

                    status.write("Recalcul des totaux par section...")
                    # Recaler les totaux de l'estimation (colonne de base)
                    tva_rate_cur = _active_lot_get("tva_rate", TVA_DEFAULT)
                    compute_section_totals(tco_df, "Px_Tot_HT", tva_rate=tva_rate_cur)

                    status.write("Mise à jour de la session...")

                    _active_lot_set("tco_df", tco_df)
                    _active_lot_set("tco_meta", meta)
                    _active_lot_set("merged_df", tco_df.copy())

                    # Mettre a jour le label et le numero de lot depuis les metadonnees
                    # N'ecrase le label que si le TCO contient une info lot — sinon
                    # on conserve le nom saisi par l'utilisateur lors de la creation du lot.
                    lot_raw = ((meta.get("project_info") or {}).get("lot") or "").strip()
                    if lot_raw:
                        _active_lot_set("lot_label", lot_raw)
                        m = re.search(r"\b(\d{2})\b", lot_raw)
                        _active_lot_set("lot_num", m.group(1) if m else "")

                    template_name = os.path.splitext(tco_file.name)[0]
                    status.update(
                        label=f"✅ {template_name} chargé ({len(tco_df)} lignes)",
                        state="complete",
                        expanded=False,
                    )
                except Exception as e:
                    status.update(label="❌ Erreur lors de l'import", state="error", expanded=True)
                    log.error("Erreur parsing TCO", exc_info=True)
                    st.error(f"❌ Erreur de lecture : {e}")
                finally:
                    _cleanup_file(path)

    if _active_lot_get("tco_df") is not None:
        # UX-1 : Selecteur TVA — affecte les calculs HT/TVA/TTC du lot
        tva_labels = list(TVA_OPTIONS.keys()) + ["Autre (personnalisé)"]
        tva_rate_lot = _active_lot_get("tva_rate", TVA_DEFAULT)
        current_tva_label = next(
            (k for k, v in TVA_OPTIONS.items() if v == tva_rate_lot),
            "Autre (personnalisé)",
        )
        selected_tva_label = st.selectbox(
            "Taux de TVA applicable",
            options=tva_labels,
            index=tva_labels.index(current_tva_label),
            help="Sélectionnez un taux prédéfini ou choisissez 'Autre' pour entrer une valeur manuelle.",
        )
        if selected_tva_label == "Autre (personnalisé)":
            custom_pct = st.number_input(
                "Taux de TVA personnalisé (%)",
                min_value=0.0,
                max_value=100.0,
                value=float(tva_rate_lot * 100)
                if current_tva_label == "Autre (personnalisé)"
                else 20.0,
                step=0.1,
                format="%.1f",
            )
            new_tva = round(custom_pct / 100.0, 4)
        else:
            new_tva = TVA_OPTIONS[selected_tva_label]

        if new_tva != tva_rate_lot:
            _active_lot_set("tva_rate", new_tva)
            if _active_lot_get("companies", {}):
                rebuild_merged_tco(new_tva)
            st.rerun()

    # Boutons de navigation — visibles quelle que soit l'étape de chargement
    if st.session_state.step == 1:
        if _active_lot_get("tco_df") is not None:
            if st.button("➡️ Passer a l'etape suivante", type="primary"):
                st.session_state.step = 2
                st.rerun()
        else:
            st.info(
                "💡 Pas de DPGF estimation disponible ? "
                "Importez directement les DPGFs entreprises pour les comparer."
            )
            if st.button(
                "📊 Mode comparatif — comparer les offres sans estimation",
                type="primary",
                use_container_width=True,
            ):
                _active_lot_set("comparatif_mode", True)
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

    if _active_lot_get("comparatif_mode", False):
        st.info("📊 **Mode comparatif** — pas de colonne Estimation dans l'export.")

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
            help="Sélectionnez tous les fichiers DPGF des entreprises à fusionner (.xlsx, .xls, .xlsb ou .pdf)",
        )

        if dpgf_files:
            companies_lot = _active_lot_get("companies", {})
            processed_filenames = {v["filename"] for v in companies_lot.values()}
            new_files = [f for f in dpgf_files if f.name not in processed_filenames]

            if new_files:
                success_count = 0
                added_companies: list[str] = []
                with st.status(
                    f"Importation de {len(new_files)} fichier(s)...", expanded=True
                ) as status:
                    for dpgf_file in new_files:
                        filename_clean = os.path.splitext(dpgf_file.name)[0]
                        company_name = (
                            re.sub(r"^DPGF\s+", "", filename_clean, flags=re.IGNORECASE)
                            .strip()
                            .upper()
                        )

                        status.write(f"Traitement de **{company_name}**...")
                        path = _safe_save(dpgf_file, allowed_extensions=DPGF_ALLOWED_EXTENSIONS)
                        if path:
                            try:
                                # Détection du type réel par magic bytes (prioritaire
                                # sur l'extension — certains PDF sont renommés .xlsx)
                                _is_pdf = False
                                try:
                                    with open(path, "rb") as _fh:
                                        _is_pdf = _fh.read(4) == b"%PDF"
                                except OSError:
                                    _is_pdf = dpgf_file.name.lower().endswith(".pdf")
                                if _is_pdf:
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
                                    # Vérification : au moins un article doit avoir un code valide.
                                    # Si tous les codes sont vides, le fichier n'est probablement
                                    # pas un DPGF entreprise (mauvais document importé).
                                    art_mask = dpgf_df["row_type"].isin(["article", "sub_section"])
                                    has_any_code = (
                                        (
                                            dpgf_df.loc[art_mask, "Code"]
                                            .astype(str)
                                            .str.strip()
                                            .ne("")
                                            .any()
                                        )
                                        if art_mask.any()
                                        else False
                                    )

                                    if not has_any_code:
                                        st.error(
                                            f"❌ **{company_name}** : Document incorrect — "
                                            "aucun code article n'a été trouvé dans ce fichier. "
                                            "Vérifiez que vous importez bien un DPGF entreprise "
                                            "(colonne « Code » obligatoire)."
                                        )
                                        log.warning(
                                            "DPGF %s rejeté : aucun code article trouvé — "
                                            "document probablement invalide.",
                                            company_name,
                                        )
                                    else:
                                        companies_lot[company_name] = {
                                            "dpgf_df": dpgf_df,
                                            "parse_alerts": parse_alerts,
                                            "filename": dpgf_file.name,
                                            "n_articles": int(
                                                (dpgf_df["row_type"] == "article").sum()
                                            ),
                                        }
                                        added_companies.append(company_name)
                                        success_count += 1
                                        status.write(f"✅ {company_name} intégré.")
                            except Exception:
                                log.error("Erreur fusion %s", company_name, exc_info=True)
                                status.write(f"❌ Erreur sur **{company_name}**")
                                st.error(f"❌ Erreur sur {company_name}")
                            finally:
                                _cleanup_file(path)

                    if success_count > 0:
                        status.update(
                            label=f"✅ {success_count} entreprise(s) importée(s) avec succès.",
                            state="complete",
                            expanded=False,
                        )
                    else:
                        status.update(
                            label="⚠️ Aucun fichier valide n'a été importé.", state="error"
                        )

                if success_count > 0:
                    _active_lot_set("companies", companies_lot)
                    rebuild_merged_tco(
                        _active_lot_get("tva_rate", TVA_DEFAULT),
                        new_companies=added_companies,
                    )
                    st.session_state.pop("export_buffer", None)
                    st.session_state.upload_counter += 1  # réinitialise le widget uploader
                    st.session_state.step = 3
                    st.session_state.export_done = False
                    _autosave()  # Auto-sauvegarde après fusion entreprises
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
                    "comparatif_mode": False,
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

    if merged is None or merged.empty:
        st.warning(
            "⚠️ Aucune donnée fusionnée disponible. Revenez à l'étape 2 et importez au moins un DPGF."
        )
        st.stop()

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

    _proj_norm = _normalize_filename(_proj_raw)
    _lot_norm = _normalize_filename(_lot_raw)

    _date_stamp = datetime.now().strftime("%Y-%m-%d")

    if _proj_norm and _lot_norm:
        filename = f"TCO_FINAL_{_proj_norm}_{_lot_norm}_{_date_stamp}.xlsx"
    elif _proj_norm:
        filename = f"TCO_FINAL_{_proj_norm}_{_date_stamp}.xlsx"
    elif _lot_norm:
        filename = f"TCO_FINAL_{_lot_norm}_{_date_stamp}.xlsx"
    else:
        filename = f"TCO_FINAL_{_date_stamp}.xlsx"

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
                    comparatif_mode=_active_lot_get("comparatif_mode", False),
                )

        st.download_button(
            label="📥 Exporter le TCO Final (.xlsx)",
            data=st.session_state.export_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            on_click=_on_export_click,
            use_container_width=True,
        )

        if st.session_state.get("export_done"):
            st.success("✅ Telechargement OK")
            log.info("Export telecharge : %s", filename)

    except Exception as e:
        log.error("Erreur preparation export", exc_info=True)
        st.error(f"❌ Erreur de generation : {e}")

st.divider()
st.caption(f"{APP_TITLE} v{APP_VERSION} — Export du TCO")
