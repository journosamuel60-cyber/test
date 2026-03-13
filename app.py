"""
app.py
Interface Streamlit pour l'analyseur de contrats d'assurance-crédit.
Visuel original + fonctionnalités db_manager (base persistante, onglet BDD, export).
"""

import os
import io
import time
import logging
from datetime import datetime

import pandas as pd
import streamlit as st

from extractor import extract_text, clean_text
from nlp_parser import extract_with_llm, merge_results, SCALAR_FIELDS
from regex_rules import extract_all_fields
from excel_exporter import results_to_row, export_to_excel
from db_manager import (load_db, upsert_row, delete_rows,
                        exists_in_db, merged_to_db_row, DB_COLUMNS)
from utils import (ProcessingLog, count_found_fields, format_confidence,
                   save_uploaded_file, setup_logging)

# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────

setup_logging()
logger = logging.getLogger(__name__)

DEFAULT_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "contrats_base.xlsx")
OUTPUT_EXCEL    = "contrats_analyses.xlsx"
TMP_DIR         = "/tmp/cia_uploads"

FIELD_LABELS = {
    "numero_police":        "N° Police",
    "assure":               "Assuré",
    "assureur":             "Assureur",
    "courtier":             "Courtier",
    "taux_prime":           "Taux de prime",
    "prime_provisionnelle": "Prime provisionnelle",
    "quotites_garanties":   "Pourcentage assuré",
    "delai_indemnisation":  "Délai d'indemnisation",
    "delai_max_credit":     "Délai max crédit",
    "date_prise_effet":     "Date de prise d'effet",
    "date_echeance":        "Date d'échéance",
    "duree_police":         "Durée de la police",
    "devise":               "Devise",
    "limite_decaissement":  "Limite de décaissement",
    "zone_discretionnaire": "Zone discrétionnaire",
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
# CSS ORIGINAL — inchangé
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
    .groupe-table th { background-color: #1F3864; color: white; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# EN-TÊTE
# ─────────────────────────────────────────────

st.markdown("""
<div class="main-header">
    <h1>📋 Analyseur de Contrats d'Assurance-Crédit</h1>
    <p>Extraction automatique · Atradius Modula · Coface · Euler Hermes · FR / EN · PDF · DOCX · TXT</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────

if "all_results" not in st.session_state:
    st.session_state.all_results = []
if "current_text" not in st.session_state:
    st.session_state.current_text = {}
if "db" not in st.session_state:
    st.session_state.db = None

if st.session_state.db is None:
    st.session_state.db = load_db(DEFAULT_DB_PATH)

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────

with st.sidebar:
    st.markdown("## ⚙️ Configuration")

    use_llm = st.toggle("Activer l'analyse IA (Claude)", value=True)

    api_key = None
    if use_llm:
        api_key_input = st.text_input(
            "Clé API Anthropic (optionnelle)",
            type="password",
            help="Laissez vide si ANTHROPIC_API_KEY est définie dans l'environnement."
        )
        api_key = api_key_input if api_key_input else os.environ.get("ANTHROPIC_API_KEY")

    st.divider()

    # ── Base de données persistante ──
    st.markdown("### 🗄️ Base de données")

    uploaded_db = st.file_uploader(
        "Charger contrats_base.xlsx",
        type=["xlsx"],
        key="db_upload",
        help="Chargez votre base en début de session"
    )
    if uploaded_db:
        st.session_state.db = load_db(uploaded_db)
        nb = len(st.session_state.db) if not st.session_state.db.empty else 0
        st.success(f"✅ Base chargée — {nb} contrat(s)")

    nb_db = len(st.session_state.db) if st.session_state.db is not None and not st.session_state.db.empty else 0
    last_maj = "—"
    if st.session_state.db is not None and not st.session_state.db.empty:
        if "Date de MAJ" in st.session_state.db.columns:
            lv = st.session_state.db["Date de MAJ"].dropna()
            if not lv.empty:
                last_maj = lv.iloc[-1]

    st.metric("Contrats en base", nb_db)
    st.caption(f"Dernière MAJ : {last_maj}")

    if st.session_state.db is not None and not st.session_state.db.empty:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            st.session_state.db.to_excel(w, index=False, sheet_name="Contrats analysés")
        st.download_button(
            "⬇️ Télécharger la base",
            data=buf.getvalue(),
            file_name="contrats_base.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.divider()
    st.markdown("### 📊 Journal")
    log = ProcessingLog()
    summary = log.get_summary()
    if summary:
        st.metric("Contrats traités", summary.get("total_processed", 0))
        st.metric("Taux de succès",   summary.get("success_rate", "N/A"))
    else:
        st.info("Aucun traitement enregistré")

    st.divider()
    st.markdown("""
    **Champs extraits**
    - N° Police / Assuré / Assureur / Courtier
    - Taux de prime / Prime provisionnelle
    - Pourcentage assuré / Délai d'indemnisation
    - Délai max crédit / Devise
    - Date d'effet / Échéance / Durée
    - Limite de décaissement / Zone discrétionnaire
    - **Groupe de polices** (liste complète)
    """)

# ─────────────────────────────────────────────
# ONGLETS
# ─────────────────────────────────────────────

tab_upload, tab_results, tab_db, tab_export = st.tabs([
    "📤 Import & Analyse",
    "📊 Résultats",
    "🗄️ Base de données",
    "📥 Export Excel",
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
            status_text  = st.empty()

            for i, uploaded_file in enumerate(uploaded_files):
                status_text.text(f"Analyse de {uploaded_file.name}...")
                progress_bar.progress(i / len(uploaded_files))

                start_time = time.time()
                errors = []

                try:
                    tmp_path = save_uploaded_file(uploaded_file, TMP_DIR)
                    raw_text, file_type = extract_text(tmp_path)
                    clean = clean_text(raw_text)
                    st.session_state.current_text[uploaded_file.name] = clean

                    if not clean.strip():
                        st.warning(f"⚠️ Aucun texte extrait de {uploaded_file.name}")
                        errors.append("Texte vide")
                        continue

                    regex_results = extract_all_fields(clean)

                    llm_results = {}
                    if use_llm and (api_key or os.environ.get("ANTHROPIC_API_KEY")):
                        with st.spinner(f"Analyse IA de {uploaded_file.name}..."):
                            llm_results = extract_with_llm(clean, api_key=api_key)

                    merged = merge_results(regex_results, llm_results)
                    found, total   = count_found_fields(merged)
                    processing_time = time.time() - start_time

                    num_police = (merged.get("numero_police") or {}).get("value", "")
                    already    = exists_in_db(st.session_state.db, num_police) if (num_police and st.session_state.db is not None) else False

                    existing = next((r for r in st.session_state.all_results
                                     if r["filename"] == uploaded_file.name), None)
                    entry = {
                        "filename":       uploaded_file.name,
                        "merged":         merged,
                        "already_exists": already,
                        "numero_police":  num_police,
                        "edited_values":  {
                            f: (v["value"] if isinstance(v, dict) and not isinstance(v.get("value"), list) else "")
                            for f, v in merged.items()
                        }
                    }
                    if existing:
                        existing.update(entry)
                    else:
                        st.session_state.all_results.append(entry)

                    log.add_entry(uploaded_file.name, "success", found, total, errors, processing_time)

                    note = " · ⚠️ existe déjà en base" if already else ""
                    st.success(f"✅ {uploaded_file.name} — {found}/{total} champs trouvés{note}")

                except Exception as e:
                    st.error(f"❌ Erreur sur {uploaded_file.name} : {e}")
                    logger.error(f"Erreur analyse {uploaded_file.name}: {e}")
                    log.add_entry(uploaded_file.name, "error", 0, len(FIELD_LABELS) + 1, [str(e)])

            progress_bar.progress(1.0)
            status_text.text("Analyse terminée !")

        if show_text and st.session_state.current_text:
            st.markdown('<div class="section-header">Texte extrait</div>', unsafe_allow_html=True)
            selected = st.selectbox("Fichier", list(st.session_state.current_text.keys()))
            if selected:
                st.text_area("Texte brut", st.session_state.current_text[selected], height=300)

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
            already = result.get("already_exists", False)
            with st.expander(f"📄 {result['filename']}{' 🔄' if already else ''}", expanded=True):
                merged = result["merged"]
                found, total = count_found_fields(merged)

                st.progress(found / total if total > 0 else 0,
                            text=f"{found}/{total} champs trouvés")

                if already:
                    st.warning(f"⚠️ Ce contrat existe déjà en base (N° {result.get('numero_police', '')}). Enregistrer écrasera la ligne existante.")

                # ── Champs scalaires (grille 2 colonnes) ──
                st.markdown("#### Données du contrat")
                cols = st.columns(2)
                for idx, (field, label) in enumerate(FIELD_LABELS.items()):
                    col = cols[idx % 2]
                    with col:
                        field_data  = merged.get(field, {"value": "", "confidence": 0.0})
                        current_val = result["edited_values"].get(field, field_data.get("value", ""))
                        if isinstance(current_val, list):
                            current_val = ""
                        confidence = field_data.get("confidence", 0.0)
                        method     = field_data.get("method", "none")

                        new_val = st.text_input(
                            f"{label} {format_confidence(confidence)}",
                            value=str(current_val) if current_val else "",
                            key=f"{result['filename']}_{field}",
                            help=f"Méthode : {method}"
                        )
                        result["edited_values"][field] = new_val

                # ── Groupe de polices ──
                gp_data = merged.get("groupe_polices", {})
                polices = gp_data.get("value")
                if polices and isinstance(polices, list):
                    st.markdown(f"#### 🏢 Groupe de polices ({len(polices)} polices)")
                    df_groupe = pd.DataFrame(polices).rename(columns={
                        "numero": "N° Police",
                        "devise": "Devise",
                        "assure": "Assuré"
                    })
                    df_groupe["Devise"] = df_groupe["Devise"].fillna("—")
                    st.dataframe(
                        df_groupe,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "N° Police": st.column_config.TextColumn(width="small"),
                            "Devise":    st.column_config.TextColumn(width="small"),
                            "Assuré":    st.column_config.TextColumn(width="large"),
                        }
                    )
                else:
                    st.markdown("#### 🏢 Groupe de polices")
                    st.caption("Aucun groupe de polices détecté dans ce document.")

                # ── Enregistrer en base ──
                st.divider()
                c1, c2, _ = st.columns([1, 1, 3])
                with c1:
                    if st.button("💾 Enregistrer en base", key=f"save_{result['filename']}", type="primary", use_container_width=True):
                        me = dict(merged)
                        for field in FIELD_LABELS:
                            ev   = result["edited_values"].get(field, "")
                            orig = merged.get(field, {})
                            me[field] = {
                                "value":      ev,
                                "confidence": orig.get("confidence", 0.0) if isinstance(orig, dict) else 0.0,
                                "method":     orig.get("method", "manual") if isinstance(orig, dict) else "manual"
                            }
                        me["groupe_polices"] = merged.get("groupe_polices", {})
                        db_row = merged_to_db_row(result["filename"], me)
                        if st.session_state.db is None:
                            st.session_state.db = load_db(DEFAULT_DB_PATH)
                        st.session_state.db, action = upsert_row(st.session_state.db, db_row)
                        verb = "mis à jour" if action == "updated" else "ajouté"
                        st.success(f"✅ Contrat {verb} en base")
                        result["already_exists"] = True
                with c2:
                    st.button("Ignorer", key=f"skip_{result['filename']}", type="secondary", use_container_width=True)

# ════════════════════════════════════════════
# TAB 3 — BASE DE DONNÉES
# ════════════════════════════════════════════

with tab_db:
    db = st.session_state.db
    if db is None or db.empty:
        st.info("💡 Aucune donnée en base. Chargez une base depuis la sidebar ou enregistrez des contrats depuis l'onglet Résultats.")
    else:
        st.markdown('<div class="section-header">Parcourir et modifier la base</div>', unsafe_allow_html=True)

        # Filtres
        fc1, fc2, fc3, fc4 = st.columns([1, 1, 1, 1])
        f_police   = fc1.text_input("Filtrer N° Police", "")
        f_assure   = fc2.text_input("Filtrer Assuré", "")
        f_assureur = fc3.text_input("Filtrer Assureur", "")
        with fc4:
            st.markdown("<div style='margin-top:1.75rem'></div>", unsafe_allow_html=True)
            if st.button("🔄 Recharger", type="secondary", use_container_width=True):
                st.session_state.db = load_db(DEFAULT_DB_PATH)
                st.rerun()

        mask = pd.Series([True] * len(db), index=db.index)
        if f_police:   mask &= db["N° Police"].str.contains(f_police,   case=False, na=False)
        if f_assure:   mask &= db["Assuré"].str.contains(f_assure,      case=False, na=False)
        if f_assureur: mask &= db["Assureur"].str.contains(f_assureur,  case=False, na=False)
        db_view = db[mask].copy()

        st.caption(f"{len(db_view)} résultat(s) sur {len(db)} contrat(s)")
        edited_df = st.data_editor(db_view, use_container_width=True, hide_index=False,
                                   num_rows="fixed", key="db_editor")

        st.divider()
        da1, da2, da3 = st.columns([1, 2, 1])
        with da1:
            if st.button("💾 Sauvegarder les modifications", type="primary", use_container_width=True):
                now = datetime.now().strftime("%d/%m/%Y %H:%M")
                for idx in edited_df.index:
                    for col in DB_COLUMNS:
                        if col in edited_df.columns and idx in st.session_state.db.index:
                            st.session_state.db.at[idx, col] = edited_df.at[idx, col]
                    if idx in st.session_state.db.index:
                        st.session_state.db.at[idx, "Date de MAJ"] = now
                st.success("✅ Modifications sauvegardées — téléchargez la base depuis la sidebar")

        with da2:
            to_delete = st.multiselect(
                "Supprimer des polices",
                options=db_view["N° Police"].tolist(),
                placeholder="Sélectionner N° Police(s) à supprimer…"
            )
        with da3:
            if to_delete:
                if st.button(f"🗑️ Supprimer ({len(to_delete)})", type="primary", use_container_width=True):
                    indices = db[db["N° Police"].isin(to_delete)].index.tolist()
                    st.session_state.db = delete_rows(st.session_state.db, indices)
                    st.success(f"✅ {len(indices)} ligne(s) supprimée(s)")
                    st.rerun()

# ════════════════════════════════════════════
# TAB 4 — EXPORT EXCEL
# ════════════════════════════════════════════

with tab_export:
    st.markdown('<div class="section-header">Export et sauvegarde</div>', unsafe_allow_html=True)

    # Workflow
    st.markdown("""
    > **Workflow recommandé**
    > 1. Chargez la base via la sidebar en début de session
    > 2. Analysez les nouveaux contrats (Import & Analyse)
    > 3. Validez et enregistrez chaque contrat (Résultats)
    > 4. Téléchargez la base mise à jour ci-dessous
    > 5. Sauvegardez sur OneDrive / SharePoint
    """)

    st.divider()

    col1, col2 = st.columns(2)

    # ── Export base persistante ──
    with col1:
        st.markdown("#### 🗄️ Base complète (persistante)")
        db_final = st.session_state.db
        if db_final is not None and not db_final.empty:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                db_final.to_excel(w, index=False, sheet_name="Contrats analysés")
            st.download_button(
                "⬇️ Télécharger contrats_base.xlsx",
                data=buf.getvalue(),
                file_name="contrats_base.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
            st.caption(f"{len(db_final)} contrat(s) · {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        else:
            st.info("Aucune donnée en base.")

    # ── Export Excel session (format original) ──
    with col2:
        st.markdown("#### 📊 Export session (format original)")
        if not st.session_state.all_results:
            st.info("💡 Aucun résultat à exporter.")
        else:
            # Aperçu
            preview_rows = []
            for result in st.session_state.all_results:
                row = {"Fichier": result["filename"]}
                for field, label in FIELD_LABELS.items():
                    val = result["edited_values"].get(field, "Non trouvé")
                    row[label] = val if not isinstance(val, list) else "—"
                gp = result["merged"].get("groupe_polices", {}).get("value")
                row["Groupe (nb polices)"] = len(gp) if gp else 0
                preview_rows.append(row)
            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)

            if st.button("📥 Générer le fichier Excel", type="primary", use_container_width=True):
                rows_for_export = []
                for result in st.session_state.all_results:
                    merged_edited = {}
                    for field in FIELD_LABELS:
                        edited_val = result["edited_values"].get(field, "Non trouvé")
                        orig = result["merged"].get(field, {})
                        merged_edited[field] = {
                            "value":      edited_val,
                            "confidence": orig.get("confidence", 0.0) if edited_val == orig.get("value") else 1.0,
                            "method":     orig.get("method", "manual") if edited_val == orig.get("value") else "manual"
                        }
                    merged_edited["groupe_polices"] = result["merged"].get("groupe_polices", {})
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
