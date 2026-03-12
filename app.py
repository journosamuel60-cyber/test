"""
app.py
Interface Streamlit pour l'analyseur de contrats d'assurance-crédit.
"""

import os
import time
import logging
from pathlib import Path

import pandas as pd
import streamlit as st

from extractor import extract_text, clean_text
from nlp_parser import extract_with_llm, merge_results
from regex_rules import extract_all_fields_regex
from excel_exporter import results_to_row, export_to_excel
from utils import (ProcessingLog, count_found_fields, format_confidence,
                   save_uploaded_file, setup_logging)

# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────

setup_logging()
logger = logging.getLogger(__name__)

OUTPUT_EXCEL = "contrats_analyses.xlsx"
TMP_DIR = "/tmp/cia_uploads"

FIELD_LABELS = {
    "assure": "Assuré",
    "assureur": "Assureur",
    "taux_prime": "Taux de prime",
    "minimum_prime": "Minimum de prime",
    "limite_decaissement": "Limite de décaissement",
    "zone_discretionnaire": "Zone discrétionnaire",
    "quotites_garanties": "Quotités garanties",
    "date_echeance": "Date d'échéance",
}

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Analyseur de Contrats d'Assurance-Crédit",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────
# CSS CUSTOM
# ─────────────────────────────────────────────

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
        padding: 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
    }
    .field-card {
        background: #F8F9FA;
        border: 1px solid #DEE2E6;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 0.5rem;
    }
    .confidence-high { color: #1E8449; font-weight: bold; }
    .confidence-med  { color: #D4AC0D; font-weight: bold; }
    .confidence-low  { color: #C0392B; font-weight: bold; }
    .not-found { color: #999; font-style: italic; }
    .section-header {
        font-size: 1.2rem;
        font-weight: 700;
        color: #1F3864;
        border-bottom: 2px solid #2E75B6;
        padding-bottom: 0.3rem;
        margin: 1.5rem 0 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# EN-TÊTE
# ─────────────────────────────────────────────

st.markdown("""
<div class="main-header">
    <h1>📋 Analyseur de Contrats d'Assurance-Crédit</h1>
    <p>Extraction automatique des données contractuelles · FR / EN · PDF · DOCX · TXT</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR — CONFIGURATION
# ─────────────────────────────────────────────

with st.sidebar:
    st.markdown("## ⚙️ Configuration")

    use_llm = st.toggle("Activer l'analyse IA (Claude)", value=True)

    api_key = None
    if use_llm:
        api_key_input = st.text_input(
            "Clé API Anthropic (optionnelle)",
            type="password",
            help="Laissez vide si la variable ANTHROPIC_API_KEY est définie dans l'environnement."
        )
        api_key = api_key_input if api_key_input else os.environ.get("ANTHROPIC_API_KEY")

    st.divider()
    st.markdown("### 📊 Journal")
    log = ProcessingLog()
    summary = log.get_summary()
    if summary:
        st.metric("Contrats traités", summary.get("total_processed", 0))
        st.metric("Taux de succès", summary.get("success_rate", "N/A"))
    else:
        st.info("Aucun traitement enregistré")

    st.divider()
    st.markdown("""
    **Champs extraits**
    - Assuré / Assureur
    - Taux de prime
    - Minimum de prime
    - Limite de décaissement
    - Zone discrétionnaire
    - Quotités garanties
    - Date d'échéance
    """)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────

if "all_results" not in st.session_state:
    st.session_state.all_results = []  # Liste de dicts {filename, merged, edited_values}

if "current_text" not in st.session_state:
    st.session_state.current_text = {}

# ─────────────────────────────────────────────
# ONGLETS
# ─────────────────────────────────────────────

tab_upload, tab_results, tab_export = st.tabs([
    "📤 Import & Analyse",
    "📊 Résultats",
    "📥 Export Excel"
])

# ════════════════════════════════════════════
# TAB 1 — IMPORT & ANALYSE
# ════════════════════════════════════════════

with tab_upload:
    st.markdown('<div class="section-header">Import des contrats</div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Déposez vos contrats ici",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=True,
        help="Formats supportés : PDF (natif ou scanné), DOCX, TXT"
    )

    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} fichier(s) chargé(s)")

        col1, col2 = st.columns([2, 1])
        with col1:
            analyze_btn = st.button(
                f"🔍 Analyser {len(uploaded_files)} contrat(s)",
                type="primary",
                use_container_width=True
            )
        with col2:
            show_text = st.checkbox("Afficher le texte extrait", value=False)

        if analyze_btn:
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, uploaded_file in enumerate(uploaded_files):
                status_text.text(f"Analyse de {uploaded_file.name}...")
                progress_bar.progress((i) / len(uploaded_files))

                start_time = time.time()
                errors = []

                try:
                    # 1. Sauvegarder temporairement
                    tmp_path = save_uploaded_file(uploaded_file, TMP_DIR)

                    # 2. Extraire le texte
                    raw_text, file_type = extract_text(tmp_path)
                    clean = clean_text(raw_text)
                    st.session_state.current_text[uploaded_file.name] = clean

                    if not clean.strip():
                        st.warning(f"⚠️ Aucun texte extrait de {uploaded_file.name}")
                        errors.append("Texte vide")
                        continue

                    # 3. Regex
                    regex_results = extract_all_fields_regex(clean)

                    # 4. LLM (si activé)
                    llm_results = {}
                    if use_llm and (api_key or os.environ.get("ANTHROPIC_API_KEY")):
                        with st.spinner(f"Analyse IA de {uploaded_file.name}..."):
                            llm_results = extract_with_llm(clean, api_key=api_key)

                    # 5. Fusion
                    merged = merge_results(regex_results, llm_results)

                    # 6. Enregistrement
                    found, total = count_found_fields(merged)
                    processing_time = time.time() - start_time

                    # Vérifier si déjà analysé
                    existing = next((r for r in st.session_state.all_results
                                     if r["filename"] == uploaded_file.name), None)
                    if existing:
                        existing["merged"] = merged
                        existing["edited_values"] = {
                            f: v["value"] for f, v in merged.items()
                        }
                    else:
                        st.session_state.all_results.append({
                            "filename": uploaded_file.name,
                            "merged": merged,
                            "edited_values": {f: v["value"] for f, v in merged.items()}
                        })

                    log.add_entry(
                        uploaded_file.name, "success", found, total,
                        errors, processing_time
                    )

                    st.success(f"✅ {uploaded_file.name} — {found}/{total} champs trouvés")

                except Exception as e:
                    st.error(f"❌ Erreur sur {uploaded_file.name} : {e}")
                    logger.error(f"Erreur analyse {uploaded_file.name}: {e}")
                    log.add_entry(uploaded_file.name, "error", 0, 8, [str(e)])

            progress_bar.progress(1.0)
            status_text.text("Analyse terminée !")

        # Affichage du texte extrait
        if show_text and st.session_state.current_text:
            st.markdown('<div class="section-header">Texte extrait</div>', unsafe_allow_html=True)
            selected = st.selectbox("Fichier", list(st.session_state.current_text.keys()))
            if selected:
                st.text_area(
                    "Texte brut",
                    st.session_state.current_text[selected],
                    height=300
                )

# ════════════════════════════════════════════
# TAB 2 — RÉSULTATS & VALIDATION
# ════════════════════════════════════════════

with tab_results:
    if not st.session_state.all_results:
        st.info("💡 Aucun contrat analysé. Utilisez l'onglet **Import & Analyse**.")
    else:
        st.markdown('<div class="section-header">Validation et correction des résultats</div>',
                    unsafe_allow_html=True)

        for result in st.session_state.all_results:
            with st.expander(f"📄 {result['filename']}", expanded=True):
                merged = result["merged"]
                found, total = count_found_fields(merged)

                # Barre de progression des champs
                pct = found / total if total > 0 else 0
                st.progress(pct, text=f"{found}/{total} champs trouvés")

                # Grille d'édition 2 colonnes
                cols = st.columns(2)
                for idx, (field, label) in enumerate(FIELD_LABELS.items()):
                    col = cols[idx % 2]
                    with col:
                        field_data = merged.get(field, {"value": "Non trouvé", "confidence": 0.0})
                        current_val = result["edited_values"].get(field, field_data["value"])
                        confidence = field_data.get("confidence", 0.0)
                        method = field_data.get("method", "none")

                        new_val = st.text_input(
                            f"{label} {format_confidence(confidence)}",
                            value=current_val,
                            key=f"{result['filename']}_{field}",
                            help=f"Méthode : {method}"
                        )
                        result["edited_values"][field] = new_val

# ════════════════════════════════════════════
# TAB 3 — EXPORT EXCEL
# ════════════════════════════════════════════

with tab_export:
    if not st.session_state.all_results:
        st.info("💡 Aucun résultat à exporter.")
    else:
        st.markdown('<div class="section-header">Export vers Excel</div>', unsafe_allow_html=True)

        # Aperçu tableau
        preview_rows = []
        for result in st.session_state.all_results:
            row = {"Fichier": result["filename"]}
            for field, label in FIELD_LABELS.items():
                row[label] = result["edited_values"].get(field, "Non trouvé")
            preview_rows.append(row)

        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)

        col1, col2 = st.columns(2)
        with col1:
            export_btn = st.button(
                "📥 Générer le fichier Excel",
                type="primary",
                use_container_width=True
            )

        if export_btn:
            # Reconstruire les résultats avec les valeurs éditées
            rows_for_export = []
            for result in st.session_state.all_results:
                # Reconstruire merged avec valeurs éditées
                merged_edited = {}
                for field in FIELD_LABELS:
                    edited_val = result["edited_values"].get(field, "Non trouvé")
                    orig = result["merged"].get(field, {})
                    merged_edited[field] = {
                        "value": edited_val,
                        "confidence": orig.get("confidence", 0.0) if edited_val == orig.get("value") else 1.0,
                        "method": orig.get("method", "manual") if edited_val == orig.get("value") else "manual"
                    }

                rows_for_export.append(results_to_row(result["filename"], merged_edited))

            try:
                output_path = export_to_excel(rows_for_export, OUTPUT_EXCEL)

                with open(output_path, "rb") as f:
                    st.download_button(
                        label="⬇️ Télécharger le fichier Excel",
                        data=f.read(),
                        file_name=OUTPUT_EXCEL,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                st.success(f"✅ Fichier Excel généré : {len(rows_for_export)} contrat(s)")

            except Exception as e:
                st.error(f"❌ Erreur lors de l'export : {e}")
