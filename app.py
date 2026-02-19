"""
app.py — Interface Streamlit pour TCO Automator.

Application en 3 étapes :
1. Import du TCO modèle
2. Import des DPGF entreprises (un ou plusieurs)
3. Visualisation du résultat et export
"""

import streamlit as st
import pandas as pd
import os
import tempfile
from datetime import datetime

from core.parser_tco import parse_tco
from core.parser_dpgf import parse_dpgf
from core.merger import merge_company_into_tco
from core.exporter import export_tco


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="TCO Automator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = [".xlsx"]
MAX_FILE_SIZE_MB = 20


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def save_uploaded_file(uploaded_file, directory):
    """Sauvegarde un fichier uploadé de manière sécurisée."""
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        st.error(f"❌ Format non accepté : {ext}. Seul .xlsx est autorisé.")
        return None

    size_mb = uploaded_file.size / (1024 * 1024)
    if size_mb > MAX_FILE_SIZE_MB:
        st.error(f"❌ Fichier trop volumineux ({size_mb:.1f} MB > {MAX_FILE_SIZE_MB} MB)")
        return None

    filepath = os.path.join(directory, uploaded_file.name)
    with open(filepath, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return filepath


def display_alerts(alerts, title="Alertes détectées"):
    """Affiche les alertes sous forme de badges colorés."""
    if not alerts:
        st.success("✅ Aucune anomalie détectée")
        return

    # Compter les alertes par type
    counts = {"error": 0, "warning": 0, "info": 0}
    for a in alerts:
        counts[a.get("type", "info")] = counts.get(a.get("type", "info"), 0) + 1

    cols = st.columns(3)
    with cols[0]:
        if counts["error"] > 0:
            st.error(f"🔴 {counts['error']} erreur(s)")
    with cols[1]:
        if counts["warning"] > 0:
            st.warning(f"🟡 {counts['warning']} avertissement(s)")
    with cols[2]:
        if counts["info"] > 0:
            st.info(f"🔵 {counts['info']} info(s)")

    # Détails dans un expander
    with st.expander("📋 Détail des alertes", expanded=False):
        for alert in alerts:
            icon = {
                "error": "🔴",
                "warning": "🟡",
                "info": "🔵",
            }.get(alert["type"], "ℹ️")
            st.write(
                f"{icon} **{alert.get('code', '')}** — {alert.get('message', '')}"
            )


def display_preview(df, title="Aperçu", n_rows=20):
    """Affiche un aperçu d'un DataFrame."""
    st.write(f"**{title}** ({len(df)} lignes)")

    # Colonnes à afficher (masquer les colonnes techniques)
    display_cols = [
        c for c in df.columns
        if c not in ("Entete", "row_type", "original_row")
    ]

    # Filtrer les lignes vides
    preview = df[df["row_type"] != "empty"][display_cols].head(n_rows)
    st.dataframe(preview, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# CSS personnalisé
# ---------------------------------------------------------------------------

st.markdown("""
<style>
    .stApp {
        max-width: 1400px;
        margin: 0 auto;
    }
    .main-title {
        text-align: center;
        color: #2F5496;
        margin-bottom: 0.5rem;
    }
    .step-header {
        background: linear-gradient(135deg, #2F5496, #4472C4);
        color: white;
        padding: 12px 20px;
        border-radius: 8px;
        margin: 1rem 0 0.5rem;
    }
    .legend-box {
        display: flex;
        gap: 1rem;
        flex-wrap: wrap;
        margin: 0.5rem 0;
    }
    .legend-item {
        display: flex;
        align-items: center;
        gap: 0.3rem;
        font-size: 0.85rem;
    }
    .color-dot {
        width: 14px;
        height: 14px;
        border-radius: 3px;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# State Management
# ---------------------------------------------------------------------------

if "tco_df" not in st.session_state:
    st.session_state.tco_df = None
if "tco_meta" not in st.session_state:
    st.session_state.tco_meta = None
if "merged_df" not in st.session_state:
    st.session_state.merged_df = None
if "all_alerts" not in st.session_state:
    st.session_state.all_alerts = []
if "companies" not in st.session_state:
    st.session_state.companies = []
if "step" not in st.session_state:
    st.session_state.step = 1


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

st.markdown("<h1 class='main-title'>📊 TCO Automator</h1>", unsafe_allow_html=True)
st.markdown(
    "<p style='text-align:center; color:#666;'>"
    "Automatisez l'intégration des DPGF entreprises dans votre TCO"
    "</p>",
    unsafe_allow_html=True,
)

# Progress bar
progress_labels = ["1️⃣ TCO Modèle", "2️⃣ DPGF Entreprise", "3️⃣ Résultat & Export"]
progress_cols = st.columns(3)
for i, (col, label) in enumerate(zip(progress_cols, progress_labels)):
    with col:
        if st.session_state.step > i + 1:
            st.success(label)
        elif st.session_state.step == i + 1:
            st.info(label)
        else:
            st.write(label)

st.divider()

# ---------------------------------------------------------------------------
# STEP 1 — Import TCO
# ---------------------------------------------------------------------------

if st.session_state.step >= 1:
    st.markdown(
        "<div class='step-header'>📥 Étape 1 — Importer le TCO Modèle</div>",
        unsafe_allow_html=True,
    )
    st.caption(
        "Sélectionnez le fichier DPGF LOT modèle (.xlsx) — "
        "Colonnes attendues : Code | Désignation | Qu. | U. | Px U. | Px tot."
    )

    tco_file = st.file_uploader(
        "Charger le TCO modèle",
        type=["xlsx"],
        key="tco_upload",
        help="Fichier DPGF LOT servant de base (ex: DPGF LOT 01)",
    )

    if tco_file and st.session_state.tco_df is None:
        filepath = save_uploaded_file(tco_file, UPLOAD_DIR)
        if filepath:
            with st.spinner("🔄 Lecture du TCO..."):
                try:
                    tco_df, meta = parse_tco(filepath)
                    st.session_state.tco_df = tco_df
                    st.session_state.tco_meta = meta
                    st.session_state.merged_df = tco_df.copy()

                    # Afficher les infos du projet
                    info = meta["project_info"]
                    if info:
                        st.info(
                            f"📋 **Projet :** {info.get('projet', 'N/A')} — "
                            f"**Lot :** {info.get('lot', 'N/A')}"
                        )

                    st.success(
                        f"✅ TCO chargé avec succès — "
                        f"{len(tco_df)} lignes, "
                        f"{len(tco_df[tco_df['row_type']=='data'])} postes"
                    )
                except Exception as e:
                    st.error(f"❌ Erreur de lecture : {e}")

    if st.session_state.tco_df is not None:
        display_preview(st.session_state.tco_df, "Aperçu du TCO")

        if st.session_state.step == 1:
            if st.button("✅ Valider le TCO et continuer", type="primary"):
                st.session_state.step = 2
                st.rerun()


# ---------------------------------------------------------------------------
# STEP 2 — Import DPGF Entreprise
# ---------------------------------------------------------------------------

if st.session_state.step >= 2:
    st.markdown(
        "<div class='step-header'>📥 Étape 2 — Importer un DPGF Entreprise</div>",
        unsafe_allow_html=True,
    )

    # Afficher les entreprises déjà importées
    if st.session_state.companies:
        st.success(
            "✅ Entreprises importées : " +
            ", ".join(f"**{c}**" for c in st.session_state.companies)
        )

    # Formulaire d'import
    col1, col2 = st.columns([2, 1])
    with col1:
        dpgf_file = st.file_uploader(
            "Charger un DPGF entreprise",
            type=["xlsx"],
            key=f"dpgf_upload_{len(st.session_state.companies)}",
            help="Fichier DPGF d'une entreprise (ex: MAB_SUD_OUEST.xlsx)",
        )
    with col2:
        company_name = st.text_input(
            "Nom de l'entreprise",
            placeholder="Ex: MAB SUD-OUEST",
            help="Nom qui apparaîtra dans les colonnes du TCO final",
        )

    if dpgf_file and company_name:
        if st.button("🔗 Fusionner ce DPGF", type="primary"):
            filepath = save_uploaded_file(dpgf_file, UPLOAD_DIR)
            if filepath:
                with st.spinner(f"🔄 Traitement de {company_name}..."):
                    try:
                        dpgf_df, parse_alerts = parse_dpgf(filepath)

                        # Afficher les alertes de parsing
                        display_alerts(parse_alerts, f"Alertes — {company_name}")

                        # Fusionner
                        merged_df, merge_alerts = merge_company_into_tco(
                            st.session_state.merged_df,
                            dpgf_df,
                            company_name,
                        )

                        st.session_state.merged_df = merged_df
                        st.session_state.all_alerts.extend(parse_alerts)
                        st.session_state.all_alerts.extend(merge_alerts)
                        st.session_state.companies.append(company_name)

                        # Afficher les alertes de fusion
                        if merge_alerts:
                            display_alerts(merge_alerts, "Alertes de fusion")

                        matched = len(dpgf_df[
                            (dpgf_df["row_type"] == "data") &
                            (dpgf_df["Code"] != "")
                        ])
                        st.success(
                            f"✅ **{company_name}** fusionnée — "
                            f"{matched} postes traités, "
                            f"{len(merge_alerts)} anomalie(s)"
                        )
                        st.rerun()

                    except Exception as e:
                        st.error(f"❌ Erreur de traitement : {e}")

    # Aperçu du TCO fusionné
    if st.session_state.companies:
        display_preview(
            st.session_state.merged_df,
            f"TCO fusionné ({len(st.session_state.companies)} entreprise(s))",
        )

    # Boutons de navigation
    col_nav1, col_nav2 = st.columns(2)
    with col_nav1:
        if st.session_state.companies:
            if st.button("➡️ Passer au résultat final", type="primary"):
                st.session_state.step = 3
                st.rerun()
    with col_nav2:
        if st.button("🔄 Réinitialiser"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()


# ---------------------------------------------------------------------------
# STEP 3 — Résultat & Export
# ---------------------------------------------------------------------------

if st.session_state.step >= 3:
    st.markdown(
        "<div class='step-header'>📊 Étape 3 — Résultat Final</div>",
        unsafe_allow_html=True,
    )

    # Légende des couleurs
    st.markdown("""
    <div class='legend-box'>
        <div class='legend-item'>
            <span class='color-dot' style='background:#FFC7CE'></span>
            Erreur (total incohérent)
        </div>
        <div class='legend-item'>
            <span class='color-dot' style='background:#FFE4B5'></span>
            Avertissement (code inconnu)
        </div>
        <div class='legend-item'>
            <span class='color-dot' style='background:#FFFFCC'></span>
            Note (texte dans numérique)
        </div>
        <div class='legend-item'>
            <span class='color-dot' style='background:#D6EAF8'></span>
            Info (mot-clé détecté)
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Statistiques
    merged = st.session_state.merged_df
    data_rows = merged[merged["row_type"] == "data"]

    stat_cols = st.columns(4)
    with stat_cols[0]:
        st.metric("📋 Postes", len(data_rows))
    with stat_cols[1]:
        st.metric("🏢 Entreprises", len(st.session_state.companies))
    with stat_cols[2]:
        st.metric("⚠️ Alertes", len(st.session_state.all_alerts))
    with stat_cols[3]:
        total_ht = data_rows["Px_Tot_HT"].sum() if "Px_Tot_HT" in data_rows else 0
        st.metric(
            "💰 Total estimation HT",
            f"{total_ht:,.2f} €" if total_ht else "N/A",
        )

    # Tableau complet
    display_preview(merged, "TCO Final Consolidé", n_rows=50)

    # Alertes consolidées
    if st.session_state.all_alerts:
        display_alerts(st.session_state.all_alerts, "Toutes les alertes")

    # Export
    st.divider()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"TCO_FINAL_{timestamp}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    if st.button("📥 Exporter le TCO Final (.xlsx)", type="primary"):
        with st.spinner("🔄 Génération du fichier Excel..."):
            try:
                export_tco(
                    st.session_state.merged_df,
                    st.session_state.tco_meta,
                    output_path,
                    st.session_state.all_alerts,
                )

                with open(output_path, "rb") as f:
                    st.download_button(
                        label=f"⬇️ Télécharger {output_filename}",
                        data=f,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                    )

                st.success(f"✅ Fichier exporté : `{output_filename}`")
            except Exception as e:
                st.error(f"❌ Erreur d'export : {e}")


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------

st.divider()
st.caption(
    "TCO Automator v1.0 — Python + Pandas + OpenPyXL + Streamlit | "
    "Fichiers .xlsx uniquement"
)
