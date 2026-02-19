"""
app.py — Interface Streamlit pour TCO Automator.

Application en 3 étapes :
1. Import du TCO modèle
2. Import des DPGF entreprises (un ou plusieurs), avec suppression possible
3. Visualisation du résultat et export

Gestion d'état : chaque DPGF entreprise est stocké individuellement
dans session_state pour permettre l'ajout/suppression et la reconstruction.
"""

import streamlit as st
import pandas as pd
import os
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


def rebuild_merged_tco():
    """
    Reconstruit le TCO fusionné à partir du TCO de base et de toutes
    les entreprises stockées dans session_state.
    Appelée après chaque ajout ou suppression d'entreprise.
    """
    if st.session_state.tco_df is None:
        return

    merged = st.session_state.tco_df.copy()
    all_alerts = []

    for comp_name, comp_data in st.session_state.company_data.items():
        dpgf_df = comp_data["dpgf_df"]
        parse_alerts = comp_data["parse_alerts"]

        merged, merge_alerts = merge_company_into_tco(
            merged, dpgf_df, comp_name
        )
        all_alerts.extend(parse_alerts)
        all_alerts.extend(merge_alerts)

    st.session_state.merged_df = merged
    st.session_state.all_alerts = all_alerts


def display_alerts(alerts, title="Alertes détectées"):
    """Affiche les alertes sous forme de badges colorés."""
    if not alerts:
        st.success("✅ Aucune anomalie détectée")
        return

    counts = {"error": 0, "warning": 0, "info": 0}
    for a in alerts:
        t = a.get("type", "info")
        counts[t] = counts.get(t, 0) + 1

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

    with st.expander("📋 Détail des alertes", expanded=False):
        for alert in alerts:
            icon = {"error": "🔴", "warning": "🟡", "info": "🔵"}.get(
                alert["type"], "ℹ️"
            )
            st.write(
                f"{icon} **{alert.get('code', '')}** — {alert.get('message', '')}"
            )


def display_preview(df, title="Aperçu", n_rows=20):
    """Affiche un aperçu d'un DataFrame."""
    st.write(f"**{title}** ({len(df)} lignes)")
    display_cols = [
        c for c in df.columns
        if c not in ("Entete", "row_type", "original_row", "parent_code")
    ]
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
    .company-card {
        background: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 12px 16px;
        margin: 4px 0;
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
# Stockage individuel par entreprise : {nom: {dpgf_df, parse_alerts, filename}}
if "company_data" not in st.session_state:
    st.session_state.company_data = {}
if "step" not in st.session_state:
    st.session_state.step = 1
if "upload_counter" not in st.session_state:
    st.session_state.upload_counter = 0


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
progress_labels = ["1️⃣ TCO Modèle", "2️⃣ DPGF Entreprises", "3️⃣ Résultat & Export"]
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

                    info = meta["project_info"]
                    if info:
                        st.info(
                            f"📋 **Projet :** {info.get('projet', 'N/A')} — "
                            f"**Lot :** {info.get('lot', 'N/A')}"
                        )

                    # Compter les articles
                    article_count = len(
                        tco_df[tco_df["row_type"] == "article"]
                    )
                    st.success(
                        f"✅ TCO chargé — {len(tco_df)} lignes, "
                        f"{article_count} articles"
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
# STEP 2 — Import DPGF Entreprises
# ---------------------------------------------------------------------------

if st.session_state.step >= 2:
    st.markdown(
        "<div class='step-header'>📥 Étape 2 — Gérer les DPGF Entreprises</div>",
        unsafe_allow_html=True,
    )

    # --- Liste des entreprises importées avec bouton supprimer ---
    if st.session_state.company_data:
        st.write(
            f"**{len(st.session_state.company_data)} entreprise(s) importée(s) :**"
        )

        for comp_name in list(st.session_state.company_data.keys()):
            comp_info = st.session_state.company_data[comp_name]
            n_alerts = len(comp_info["parse_alerts"])
            n_articles = len(
                comp_info["dpgf_df"][
                    comp_info["dpgf_df"]["row_type"] == "article"
                ]
            )

            col_info, col_btn = st.columns([4, 1])
            with col_info:
                st.markdown(
                    f"<div class='company-card'>"
                    f"🏢 <b>{comp_name}</b> — "
                    f"{n_articles} articles, {n_alerts} alerte(s) "
                    f"<i>({comp_info['filename']})</i>"
                    f"</div>",
                    unsafe_allow_html=True,
                )
            with col_btn:
                if st.button(
                    "🗑️ Retirer",
                    key=f"remove_{comp_name}",
                    help=f"Supprimer {comp_name} et recalculer le TCO",
                ):
                    del st.session_state.company_data[comp_name]
                    rebuild_merged_tco()
                    st.success(f"✅ **{comp_name}** supprimée. TCO recalculé.")
                    st.rerun()

        st.divider()

    # --- Formulaire d'ajout ---
    st.write("**Ajouter une entreprise :**")
    col1, col2 = st.columns([2, 1])
    with col1:
        dpgf_file = st.file_uploader(
            "Charger un DPGF entreprise",
            type=["xlsx"],
            key=f"dpgf_upload_{st.session_state.upload_counter}",
            help="Fichier DPGF d'une entreprise (ex: MAB_SUD_OUEST.xlsx)",
        )
    with col2:
        company_name = st.text_input(
            "Nom de l'entreprise",
            placeholder="Ex: MAB SUD-OUEST",
            help="Nom qui apparaîtra dans les colonnes du TCO final",
        )

    if dpgf_file and company_name:
        # Vérifier que le nom n'est pas déjà pris
        if company_name in st.session_state.company_data:
            st.warning(
                f"⚠️ **{company_name}** est déjà importée. "
                "Supprimez-la d'abord pour la remplacer."
            )
        else:
            if st.button("🔗 Fusionner ce DPGF", type="primary"):
                filepath = save_uploaded_file(dpgf_file, UPLOAD_DIR)
                if filepath:
                    with st.spinner(f"🔄 Traitement de {company_name}..."):
                        try:
                            dpgf_df, parse_alerts = parse_dpgf(filepath)

                            # Stocker les données brutes
                            st.session_state.company_data[company_name] = {
                                "dpgf_df": dpgf_df,
                                "parse_alerts": parse_alerts,
                                "filename": dpgf_file.name,
                            }

                            # Reconstruire le TCO complet
                            rebuild_merged_tco()

                            # Incrémenter le compteur pour rafraîchir le widget
                            st.session_state.upload_counter += 1

                            st.success(
                                f"✅ **{company_name}** fusionnée — "
                                f"{len(parse_alerts)} alerte(s)"
                            )
                            st.rerun()

                        except Exception as e:
                            st.error(f"❌ Erreur de traitement : {e}")

    # --- Aperçu du TCO fusionné ---
    if st.session_state.company_data:
        display_preview(
            st.session_state.merged_df,
            f"TCO fusionné ({len(st.session_state.company_data)} entreprise(s))",
        )

    # --- Navigation ---
    col_nav1, col_nav2 = st.columns(2)
    with col_nav1:
        if st.session_state.company_data:
            if st.button("➡️ Passer au résultat final", type="primary"):
                st.session_state.step = 3
                st.rerun()
    with col_nav2:
        if st.button("🔄 Tout réinitialiser"):
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
    article_rows = merged[merged["row_type"] == "article"]

    stat_cols = st.columns(4)
    with stat_cols[0]:
        st.metric("📋 Articles", len(article_rows))
    with stat_cols[1]:
        st.metric("🏢 Entreprises", len(st.session_state.company_data))
    with stat_cols[2]:
        st.metric("⚠️ Alertes", len(st.session_state.all_alerts))
    with stat_cols[3]:
        # Montant HT depuis les total_line
        montant_ht = None
        for _, row in merged.iterrows():
            if row["row_type"] == "total_line":
                desig = str(row.get("Désignation", "")).strip().lower()
                if "montant ht" in desig:
                    # Chercher le premier Px_Tot_HT d'entreprise non nul
                    for comp in st.session_state.company_data:
                        val = row.get(f"{comp}_Px_Tot_HT")
                        if val is not None and val != 0:
                            montant_ht = val
                            break
                    break
        st.metric(
            "💰 Montant HT",
            f"{montant_ht:,.2f} €" if montant_ht else "N/A",
        )

    # Tableau complet
    display_preview(merged, "TCO Final Consolidé", n_rows=50)

    # Alertes consolidées
    if st.session_state.all_alerts:
        display_alerts(st.session_state.all_alerts, "Toutes les alertes")

    # --- Bouton retour pour modifier les entreprises ---
    if st.button("⬅️ Retour — Modifier les entreprises"):
        st.session_state.step = 2
        st.rerun()

    # --- Export ---
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
