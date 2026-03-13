"""
app.py - Analyseur de Contrats d'Assurance-Crédit
"""
import os, time, logging, tempfile, shutil
from datetime import datetime
import pandas as pd
import streamlit as st

from extractor import extract_text, clean_text
from nlp_parser import extract_with_llm, merge_results, SCALAR_FIELDS
from regex_rules import extract_all_fields
from excel_exporter import results_to_row, export_to_excel
from db_manager import (load_db, save_db, upsert_row, delete_rows,
                         exists_in_db, merged_to_db_row, DB_COLUMNS)
from utils import (ProcessingLog, count_found_fields, format_confidence,
                   save_uploaded_file, setup_logging)

setup_logging()
logger = logging.getLogger(__name__)
DEFAULT_DB_PATH = "contrats_base.xlsx"
TMP_DIR = "/tmp/cia_uploads"

FIELD_LABELS = {
    "numero_police": "N° Police", "assure": "Assuré", "assureur": "Assureur",
    "courtier": "Courtier", "taux_prime": "Taux de prime",
    "prime_provisionnelle": "Prime provisionnelle", "quotites_garanties": "Pourcentage assuré",
    "delai_indemnisation": "Délai d'indemnisation", "delai_max_credit": "Délai max crédit",
    "date_prise_effet": "Date de prise d'effet", "date_echeance": "Date d'échéance",
    "duree_police": "Durée", "devise": "Devise",
    "limite_decaissement": "Limite de décaissement", "zone_discretionnaire": "Zone discrétionnaire",
}

st.set_page_config(page_title="Analyseur Assurance-Crédit", page_icon="📋", layout="wide")
st.markdown("""<style>
.main-header{background:linear-gradient(135deg,#1F3864 0%,#2E75B6 100%);padding:1.5rem 2rem;
border-radius:12px;color:white;margin-bottom:1.5rem;}
.section-header{font-size:1.1rem;font-weight:700;color:#1F3864;border-bottom:2px solid #2E75B6;
padding-bottom:.3rem;margin:1.2rem 0 .8rem 0;}
</style>""", unsafe_allow_html=True)
st.markdown("""<div class="main-header"><h1 style="margin:0">📋 Analyseur de Contrats d'Assurance-Crédit</h1>
<p style="margin:.3rem 0 0 0;opacity:.85">Atradius Modula · Coface · Euler Hermes · FR/EN · PDF · DOCX · TXT</p>
</div>""", unsafe_allow_html=True)

# ── Sidebar ──
with st.sidebar:
    st.markdown("## ⚙️ Configuration")
    use_llm = st.toggle("Activer l'analyse IA (Claude)", value=True)
    api_key = None
    if use_llm:
        api_key_input = st.text_input("Clé API Anthropic", type="password")
        api_key = api_key_input or os.environ.get("ANTHROPIC_API_KEY")
    st.divider()
    st.markdown("### 🗄️ Base de données")
    db_path = st.text_input("Chemin fichier base", value=os.environ.get("CIA_DB_PATH", DEFAULT_DB_PATH))
    uploaded_db = st.file_uploader("Charger une base existante", type=["xlsx"])
    if uploaded_db:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_db.read())
        shutil.copy(tmp.name, db_path)
        st.success(f"✅ Base importée → {db_path}")
        if "db" in st.session_state:
            del st.session_state["db"]
    st.divider()
    log = ProcessingLog()

# ── Session state ──
for k, v in [("all_results", []), ("current_text", {}), ("pending_upserts", [])]:
    if k not in st.session_state:
        st.session_state[k] = v
if "db" not in st.session_state:
    st.session_state.db = load_db(db_path)

tab_upload, tab_results, tab_db, tab_export = st.tabs([
    "📤 Import & Analyse", "📊 Résultats", "🗄️ Base de données", "📥 Export Excel"
])

# ═══ TAB 1 ════════════════════════════════════════════════════
with tab_upload:
    st.markdown('<div class="section-header">Import des contrats</div>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader("Déposez vos contrats", type=["pdf","docx","txt"], accept_multiple_files=True)

    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} fichier(s) chargé(s)")
        c1, c2 = st.columns([2,1])
        with c1: analyze_btn = st.button(f"🔍 Analyser {len(uploaded_files)} contrat(s)", type="primary", use_container_width=True)
        with c2: show_text = st.checkbox("Afficher le texte extrait")

        if analyze_btn:
            pb = st.progress(0); st_txt = st.empty()
            st.session_state.pending_upserts = []
            for i, uf in enumerate(uploaded_files):
                st_txt.text(f"Analyse de {uf.name}...")
                pb.progress(i / len(uploaded_files))
                try:
                    tmp = save_uploaded_file(uf, TMP_DIR)
                    raw, _ = extract_text(tmp)
                    clean = clean_text(raw)
                    st.session_state.current_text[uf.name] = clean
                    if not clean.strip():
                        st.warning(f"⚠️ Aucun texte dans {uf.name}"); continue

                    regex_r = extract_all_fields(clean)
                    llm_r = {}
                    if use_llm and (api_key or os.environ.get("ANTHROPIC_API_KEY")):
                        with st.spinner(f"Analyse IA {uf.name}..."):
                            llm_r = extract_with_llm(clean, api_key=api_key)
                    merged = merge_results(regex_r, llm_r)
                    found, total = count_found_fields(merged)
                    num_police = (merged.get("numero_police") or {}).get("value", "")
                    already = exists_in_db(st.session_state.db, num_police) if num_police else False

                    entry = {
                        "filename": uf.name, "merged": merged,
                        "already_exists": already, "numero_police": num_police,
                        "edited_values": {
                            f: (v.get("value","") if isinstance(v,dict) and not isinstance(v.get("value"),list) else "")
                            for f,v in merged.items()
                        }
                    }
                    existing = next((r for r in st.session_state.all_results if r["filename"]==uf.name), None)
                    if existing: existing.update(entry)
                    else: st.session_state.all_results.append(entry)
                    st.session_state.pending_upserts.append(uf.name)
                    log.add_entry(uf.name, "success", found, total, [], time.time())
                    icon = "🔄" if already else "✅"
                    note = " — **existe déjà en base**" if already else ""
                    st.success(f"{icon} {uf.name} — {found}/{total} champs{note}")
                except Exception as e:
                    st.error(f"❌ {uf.name} : {e}"); logger.error(e)
            pb.progress(1.0); st_txt.text("Analyse terminée !")
            if st.session_state.pending_upserts:
                st.info("👉 Allez dans **Résultats** pour valider et enregistrer en base.")

        if show_text and st.session_state.current_text:
            st.markdown('<div class="section-header">Texte extrait</div>', unsafe_allow_html=True)
            sel = st.selectbox("Fichier", list(st.session_state.current_text.keys()))
            if sel: st.text_area("Texte brut", st.session_state.current_text[sel], height=300)

# ═══ TAB 2 ════════════════════════════════════════════════════
with tab_results:
    if not st.session_state.all_results:
        st.info("💡 Aucun contrat analysé. Utilisez **Import & Analyse**.")
    else:
        st.markdown('<div class="section-header">Validation, correction et enregistrement</div>', unsafe_allow_html=True)
        for result in st.session_state.all_results:
            already = result.get("already_exists", False)
            badge = "🔄 Mise à jour" if already else "🆕 Nouveau"
            with st.expander(f"📄 {result['filename']}  —  {badge}", expanded=True):
                merged = result["merged"]
                found, total = count_found_fields(merged)
                st.progress(found/total if total>0 else 0, text=f"{found}/{total} champs trouvés")
                if already:
                    st.warning(f"⚠️ Le contrat **{result.get('numero_police','')}** existe déjà. Enregistrer écrasera la ligne.")

                cols = st.columns(2)
                for idx, (field, label) in enumerate(FIELD_LABELS.items()):
                    with cols[idx % 2]:
                        fd = merged.get(field, {})
                        cur = result["edited_values"].get(field, fd.get("value","") if isinstance(fd,dict) else "")
                        if isinstance(cur, list): cur = ""
                        conf = fd.get("confidence", 0.0) if isinstance(fd,dict) else 0.0
                        new_val = st.text_input(
                            f"{label} {format_confidence(conf)}", value=str(cur) if cur else "",
                            key=f"edit_{result['filename']}_{field}",
                            help=f"Méthode : {fd.get('method','') if isinstance(fd,dict) else ''}"
                        )
                        result["edited_values"][field] = new_val

                gp_d = merged.get("groupe_polices", {})
                polices = gp_d.get("value") if isinstance(gp_d,dict) else None
                if polices and isinstance(polices, list):
                    st.markdown(f"**🏢 Groupe de polices ({len(polices)} polices)**")
                    df_g = pd.DataFrame(polices).rename(columns={"numero":"N° Police","devise":"Devise","assure":"Assuré"})
                    df_g["Devise"] = df_g["Devise"].fillna("—")
                    st.dataframe(df_g, use_container_width=True, hide_index=True)

                st.divider()
                if st.button("💾 Enregistrer en base", key=f"save_{result['filename']}", type="primary"):
                    merged_edited = dict(merged)
                    for field in FIELD_LABELS:
                        ev = result["edited_values"].get(field, "")
                        orig = merged.get(field, {})
                        merged_edited[field] = {
                            "value": ev,
                            "confidence": (orig.get("confidence",0.0) if isinstance(orig,dict) else 0.0),
                            "method": (orig.get("method","manual") if isinstance(orig,dict) else "manual")
                        }
                    db_row = merged_to_db_row(result["filename"], merged_edited)
                    st.session_state.db, action = upsert_row(st.session_state.db, db_row)
                    if save_db(st.session_state.db, db_path):
                        verb = "mis à jour" if action=="updated" else "ajouté"
                        st.success(f"✅ Contrat {verb} dans `{db_path}`")
                        result["already_exists"] = True
                    else:
                        st.error("❌ Erreur de sauvegarde.")

# ═══ TAB 3 ════════════════════════════════════════════════════
with tab_db:
    st.markdown('<div class="section-header">Parcourir et modifier la base</div>', unsafe_allow_html=True)

    c_reload, c_info = st.columns([1, 3])
    with c_reload:
        if st.button("🔄 Recharger", use_container_width=True):
            st.session_state.db = load_db(db_path)
            st.success("Base rechargée.")

    db: pd.DataFrame = st.session_state.db

    if db.empty:
        st.info(f"La base est vide. Analysez et enregistrez des contrats pour les voir ici.\n\nFichier : `{db_path}`")
    else:
        with c_info:
            st.caption(f"📁 `{db_path}`  —  **{len(db)} contrat(s)**")

        # Filtres
        with st.expander("🔍 Filtrer", expanded=False):
            fc1, fc2, fc3 = st.columns(3)
            f_police  = fc1.text_input("N° Police contient", "")
            f_assure  = fc2.text_input("Assuré contient", "")
            f_assureur = fc3.text_input("Assureur contient", "")

        mask = pd.Series([True]*len(db), index=db.index)
        if f_police:  mask &= db["N° Police"].str.contains(f_police, case=False, na=False)
        if f_assure:  mask &= db["Assuré"].str.contains(f_assure, case=False, na=False)
        if f_assureur: mask &= db["Assureur"].str.contains(f_assureur, case=False, na=False)
        db_view = db[mask].copy()
        st.caption(f"{len(db_view)} résultat(s)")

        # Tableau éditable
        st.markdown("**Modifier directement dans le tableau, puis sauvegarder :**")
        edited_df = st.data_editor(
            db_view, use_container_width=True, hide_index=False, num_rows="fixed",
            key="db_editor"
        )

        # Actions
        st.divider()
        a1, a2, a3 = st.columns([1, 2, 1])

        with a1:
            if st.button("💾 Sauvegarder modifications", type="primary", use_container_width=True):
                now = datetime.now().strftime("%d/%m/%Y %H:%M")
                for idx in edited_df.index:
                    for col in DB_COLUMNS:
                        if col in edited_df.columns and idx in st.session_state.db.index:
                            st.session_state.db.at[idx, col] = edited_df.at[idx, col]
                    if idx in st.session_state.db.index:
                        st.session_state.db.at[idx, "Date de MAJ"] = now
                if save_db(st.session_state.db, db_path):
                    st.success(f"✅ Sauvegardé dans `{db_path}`")
                else:
                    st.error("❌ Erreur de sauvegarde.")

        with a2:
            to_delete = st.multiselect(
                "Supprimer des polices :",
                options=db_view["N° Police"].tolist(),
                placeholder="Sélectionner N° Police(s)…"
            )

        with a3:
            if to_delete:
                if st.button(f"🗑️ Supprimer {len(to_delete)} ligne(s)", type="primary", use_container_width=True):
                    indices = db[db["N° Police"].isin(to_delete)].index.tolist()
                    st.session_state.db = delete_rows(st.session_state.db, indices)
                    if save_db(st.session_state.db, db_path):
                        st.success(f"✅ {len(indices)} ligne(s) supprimée(s).")
                        st.rerun()
                    else:
                        st.error("❌ Erreur de sauvegarde.")

        # Export filtré
        st.divider()
        exp_name = st.text_input("Nom du fichier d'export", value="export_contrats.xlsx", key="exp_name_db")
        if st.button("📥 Exporter la vue filtrée"):
            tmp_f = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            tmp_f.close()
            if save_db(db_view.reset_index(drop=True), tmp_f.name):
                with open(tmp_f.name, "rb") as f:
                    st.download_button(
                        label=f"⬇️ Télécharger {exp_name}",
                        data=f.read(), file_name=exp_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

# ═══ TAB 4 ════════════════════════════════════════════════════
with tab_export:
    if not st.session_state.all_results:
        st.info("💡 Aucun résultat à exporter.")
    else:
        st.markdown('<div class="section-header">Export ponctuel (session en cours)</div>', unsafe_allow_html=True)
        st.caption("Exporte les contrats de cette session uniquement, indépendamment de la base persistante.")

        preview_rows = []
        for result in st.session_state.all_results:
            row = {"Fichier": result["filename"]}
            for field, label in FIELD_LABELS.items():
                val = result["edited_values"].get(field, "Non trouvé")
                row[label] = val if not isinstance(val, list) else "—"
            gp = result["merged"].get("groupe_polices", {})
            gp_v = gp.get("value") if isinstance(gp, dict) else None
            row["Groupe (nb polices)"] = len(gp_v) if gp_v else 0
            preview_rows.append(row)
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)

        if st.button("📥 Générer le fichier Excel", type="primary"):
            rows_for_export = []
            for result in st.session_state.all_results:
                me = {}
                for field in FIELD_LABELS:
                    ev = result["edited_values"].get(field, "Non trouvé")
                    orig = result["merged"].get(field, {})
                    me[field] = {"value": ev, "confidence": orig.get("confidence",0.0) if isinstance(orig,dict) else 0.0,
                                 "method": orig.get("method","manual") if isinstance(orig,dict) else "manual"}
                gp = result["merged"].get("groupe_polices", {})
                me["groupe_polices"] = gp
                rows_for_export.append(results_to_row(result["filename"], me))
            try:
                out = export_to_excel(rows_for_export, "contrats_session.xlsx")
                with open(out, "rb") as f:
                    st.download_button(label="⬇️ Télécharger", data=f.read(),
                        file_name="contrats_session.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
            except Exception as e:
                st.error(f"❌ Erreur export : {e}")
