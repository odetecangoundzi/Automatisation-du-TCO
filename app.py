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
import streamlit as st
import pandas as pd

from core.parser_tco import parse_tco
from core.parser_dpgf import parse_dpgf
from core.merger import merge_company_into_tco
from core.exporter import export_tco
from config import (
    UPLOAD_DIR, ALLOWED_EXTENSIONS, MAX_FILE_SIZE_MB, MAX_COMPANIES,
    COMPANY_NAME_MAX_LEN, TVA_OPTIONS, TVA_DEFAULT, APP_TITLE, APP_ICON
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
    initial_sidebar_state="collapsed",
)

os.makedirs(UPLOAD_DIR, exist_ok=True)
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


def display_preview(df, title="Aperçu", n_rows=20):
    hidden = {"Entete", "row_type", "original_row", "parent_code"}
    cols   = [c for c in df.columns if c not in hidden]
    st.write(f"**{title}** ({len(df)} lignes)")
    preview = df[df["row_type"] != "empty"][cols].head(n_rows)
    st.dataframe(preview, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------

st.markdown("""
<style>
.main-title   { text-align:center; color:#2F5496; margin-bottom:.5rem; }
.step-header  {
    background: linear-gradient(135deg,#2F5496,#4472C4);
    color:white; padding:12px 20px; border-radius:8px; margin:1rem 0 .5rem;
}
.company-card {
    background:#f8f9fa; border:1px solid #dee2e6;
    border-radius:8px; padding:10px 14px; margin:3px 0;
}
.legend-box { display:flex; gap:1rem; flex-wrap:wrap; margin:.5rem 0; }
.legend-item { display:flex; align-items:center; gap:.3rem; font-size:.85rem; }
.color-dot   { width:14px; height:14px; border-radius:3px; display:inline-block; }
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
    "step":          1,
    "upload_counter":0,
    "tva_rate":      TVA_DEFAULT,
    "confirm_remove":None,  # UX-4 : stocke le nom de l'entreprise à supprimer
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ---------------------------------------------------------------------------
# Header + progress
# ---------------------------------------------------------------------------

st.markdown(f"<h1 class='main-title'>{APP_ICON} {APP_TITLE}</h1>", unsafe_allow_html=True)
st.markdown(
    "<p style='text-align:center;color:#666;'>"
    "Automatisez l'intégration des DPGF entreprises dans votre TCO"
    "</p>", unsafe_allow_html=True,
)

labels = ["1️⃣ TCO Modèle", "2️⃣ DPGF Entreprises", "3️⃣ Résultat & Export"]
step_cols = st.columns(3)
for i, (col, label) in enumerate(zip(step_cols, labels)):
    with col:
        if   st.session_state.step > i + 1:  st.success(label)
        elif st.session_state.step == i + 1: st.info(label)
        else:                                 st.write(label)

st.divider()


# ---------------------------------------------------------------------------
# STEP 1 — Import TCO
# ---------------------------------------------------------------------------

if st.session_state.step >= 1:
    st.markdown(
        "<div class='step-header'>📥 Étape 1 — Importer le TCO Modèle</div>",
        unsafe_allow_html=True,
    )
    st.caption("Fichier DPGF LOT (.xlsx) — Colonnes : Code | Désignation | Qu. | U. | Px U. | Px tot.")

    tco_file = st.file_uploader(
        "Charger le TCO modèle", type=["xlsx"], key="tco_upload",
        help="Fichier DPGF LOT servant de base",
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
        display_preview(st.session_state.tco_df, "Aperçu du TCO")
        if st.session_state.step == 1:
            if st.button("✅ Valider le TCO et continuer", type="primary"):
                st.session_state.step = 2
                st.rerun()


# ---------------------------------------------------------------------------
# STEP 2 — Import DPGF Entreprises
# ---------------------------------------------------------------------------

if st.session_state.step >= 2:
    st.markdown(
        "<div class='step-header'>📥 Étape 2 — Gérer les DPGF Entreprises</div>",
        unsafe_allow_html=True,
    )

    col_tva, _ = st.columns([1, 3])
    with col_tva:
        tva_label = st.selectbox(
            "Taux de TVA", list(TVA_OPTIONS.keys()),
            index=list(TVA_OPTIONS.values()).index(st.session_state.tva_rate),
        )
        new_tva = TVA_OPTIONS[tva_label]
        if new_tva != st.session_state.tva_rate:
            st.session_state.tva_rate = new_tva
            rebuild_merged_tco(new_tva)
            st.rerun()

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
        col1, col2 = st.columns([2, 1])
        with col1:
            dpgf_file = st.file_uploader(
                "Charger un DPGF entreprise", type=["xlsx"],
                key=f"dpgf_{st.session_state.upload_counter}",
            )
        with col2:
            raw_name = st.text_input("Nom de l'entreprise", placeholder="Ex: MAB SUD-OUEST")

        if dpgf_file and raw_name:
            company_name, name_err = _validate_company_name(raw_name)
            if name_err:
                st.error(f"❌ {name_err}")
            elif company_name in st.session_state.company_data:
                st.warning(f"⚠️ **{company_name}** est déjà importée.")
            else:
                if st.button("🔗 Fusionner ce DPGF", type="primary"):
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
                                st.success(f"✅ **{company_name}** fusionnée — {n_matched} postes")
                                if parse_alerts: display_alerts(parse_alerts, company_name)
                                st.rerun()
                            except Exception as e:
                                log.error("Erreur fusion DPGF %s", company_name, exc_info=True)
                                st.error(f"❌ Erreur : {e}")
                            finally:
                                try: os.remove(path)
                                except: pass

    if st.session_state.company_data:
        display_preview(st.session_state.merged_df, f"TCO fusionné ({n_companies} entreprise(s))")

    col_nav1, col_nav2 = st.columns(2)
    with col_nav1:
        if st.session_state.company_data:
            if st.button("➡️ Passer au résultat final", type="primary"):
                st.session_state.step = 3
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
    montant_ht = None
    for _, row in merged.iterrows():
        if row["row_type"] == "total_line":
            desig = str(row.get("Désignation", "")).strip().lower()
            if "montant ht" in desig:
                for comp in st.session_state.company_data:
                    v = row.get(f"{comp}_Px_Tot_HT")
                    if v and v != 0:
                        montant_ht = v
                        break
                break

    cols = st.columns(4)
    with cols[0]: st.metric("📋 Articles",   len(art_rows))
    with cols[1]: st.metric("🏢 Entreprises", len(st.session_state.company_data))
    with cols[2]: st.metric("⚠️ Alertes",    len(st.session_state.all_alerts))
    with cols[3]: st.metric("💰 Montant HT", f"{montant_ht:,.2f} €" if montant_ht else "N/A")

    display_preview(merged, "TCO Final Consolidé", n_rows=50)

    if st.session_state.all_alerts:
        display_alerts(st.session_state.all_alerts, "Toutes les alertes")

    if st.button("⬅️ Retour — Modifier les entreprises"):
        st.session_state.step = 2
        st.rerun()

    st.divider()
    from datetime import datetime
    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename   = f"TCO_FINAL_{timestamp}.xlsx"

    if st.button("📥 Exporter le TCO Final (.xlsx)", type="primary"):
        with st.spinner("🔄 Génération du fichier Excel..."):
            try:
                buffer = export_tco(
                    st.session_state.merged_df,
                    st.session_state.tco_meta,
                    output_path=None,
                    alerts=st.session_state.all_alerts,
                )
                st.download_button(
                    label=f"⬇️ Télécharger {filename}",
                    data=buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
                st.success("✅ Fichier prêt au téléchargement.")
                log.info("Export réussi : %s", filename)
            except Exception as e:
                log.error("Erreur export Excel", exc_info=True)
                st.error(f"❌ Erreur d'export : {e}")

st.divider()
st.caption(f"{APP_TITLE} v{2.1} — TCO Automator")
