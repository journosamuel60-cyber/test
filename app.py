"""
app.py - Credit Insurance Analyzer
Design: Modèle B — fond blanc/lavande, sidebar blanche, accents violet #7C3AED
"""
import os, logging, io
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

st.set_page_config(
    page_title="CIA · Analyseur de Contrats",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=Inter:wght@400;500;600&display=swap');

/* === FOND GÉNÉRAL === */
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewContainer"] > section.main {
    background-color: #F0EDFB !important;
}
.main .block-container {
    background-color: #F0EDFB !important;
    padding-top: 1.8rem !important;
    padding-bottom: 4rem !important;
    max-width: 980px !important;
}

/* === SIDEBAR === */
[data-testid="stSidebar"],
[data-testid="stSidebar"] > div:first-child {
    background-color: #FFFFFF !important;
    border-right: 1px solid #E8E2F8 !important;
}
[data-testid="stSidebar"] { min-width: 250px !important; max-width: 250px !important; }

/* === MASQUER CHROME === */
#MainMenu, footer, header,
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"] { display: none !important; }

/* === TABS === */
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 2px solid #E2DBF5 !important;
    gap: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    color: #B0A3CC !important;
    padding: 10px 18px !important;
    border-bottom: 2px solid transparent !important;
    margin-bottom: -2px !important;
    border-radius: 0 !important;
}
.stTabs [aria-selected="true"] {
    color: #7C3AED !important;
    border-bottom-color: #7C3AED !important;
    font-weight: 600 !important;
}
.stTabs [data-baseweb="tab-panel"] { padding-top: 1.5rem !important; background: transparent !important; }

/* === BOUTONS === */
.stButton > button {
    background-color: #7C3AED !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    padding: 0.5rem 1.3rem !important;
    box-shadow: 0 2px 10px rgba(124,58,237,0.28) !important;
    transition: all 0.15s ease !important;
}
.stButton > button:hover {
    background-color: #6D28D9 !important;
    box-shadow: 0 4px 16px rgba(124,58,237,0.38) !important;
    transform: translateY(-1px) !important;
}
.stButton > button[kind="secondary"] {
    background-color: #ffffff !important;
    color: #9B8FBE !important;
    border: 1px solid #DDD0F5 !important;
    box-shadow: none !important;
}
.stButton > button[kind="secondary"]:hover {
    color: #7C3AED !important;
    border-color: #A78BFA !important;
}
.stDownloadButton > button {
    background-color: #F3EEFF !important;
    color: #7C3AED !important;
    border: 1px solid #DDD0F5 !important;
    border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    box-shadow: none !important;
}
.stDownloadButton > button:hover {
    background-color: #EBE0FF !important;
    border-color: #A78BFA !important;
}

/* === INPUTS === */
.stTextInput > label,
.stFileUploader > label,
.stToggle > label,
.stSelectbox > label,
.stMultiSelect > label,
.stTextArea > label {
    font-family: 'Inter', sans-serif !important;
    font-size: 11px !important;
    font-weight: 600 !important;
    letter-spacing: 0.06em !important;
    color: #9B8FBE !important;
    text-transform: uppercase !important;
}
.stTextInput input, .stTextArea textarea {
    background-color: #ffffff !important;
    border: 1px solid #DDD0F5 !important;
    border-radius: 7px !important;
    color: #1A1030 !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 13px !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: #7C3AED !important;
    box-shadow: 0 0 0 3px rgba(124,58,237,0.12) !important;
}

/* === FILE UPLOADER === */
[data-testid="stFileUploader"] section {
    background-color: #ffffff !important;
    border: 1.5px dashed #C9BBEE !important;
    border-radius: 10px !important;
}

/* === PROGRESS BAR === */
.stProgress > div > div {
    background: linear-gradient(90deg, #7C3AED, #A78BFA) !important;
    border-radius: 4px !important;
}
.stProgress > div {
    background: #EDE8FF !important;
    border-radius: 4px !important;
    height: 6px !important;
}

/* === EXPANDER === */
details {
    background-color: #ffffff !important;
    border: 1px solid #E8E2F8 !important;
    border-radius: 10px !important;
    box-shadow: 0 1px 4px rgba(124,58,237,0.06) !important;
    margin-bottom: 0.5rem !important;
}
details summary {
    font-family: 'Syne', sans-serif !important;
    font-size: 13px !important;
    font-weight: 700 !important;
    color: #1A1030 !important;
    padding: 0.9rem 1rem !important;
}
.streamlit-expanderContent {
    background-color: #FAF8FF !important;
    border-top: 1px solid #E8E2F8 !important;
    padding: 1rem !important;
}

/* === ALERTS === */
.stSuccess { background-color: #F0FBF5 !important; border: 1px solid rgba(52,199,133,0.3) !important; border-radius: 8px !important; }
.stWarning { background-color: #FFFBEB !important; border: 1px solid rgba(234,179,8,0.3) !important;  border-radius: 8px !important; }
.stError   { background-color: #FFF5F5 !important; border: 1px solid rgba(220,38,38,0.25) !important; border-radius: 8px !important; }
.stInfo    { background-color: #F3EEFF !important; border: 1px solid #DDD0F5 !important;               border-radius: 8px !important; color: #5B21B6 !important; }

/* === DATAFRAME === */
[data-testid="stDataFrame"] {
    border: 1px solid #E8E2F8 !important;
    border-radius: 10px !important;
    overflow: hidden !important;
    box-shadow: 0 1px 4px rgba(124,58,237,0.05) !important;
}

/* === METRICS === */
[data-testid="stMetric"] {
    background: #fff !important;
    border: 1px solid #E8E2F8 !important;
    border-radius: 10px !important;
    padding: 0.8rem 1rem !important;
}
[data-testid="stMetricValue"] {
    font-family: 'Syne', sans-serif !important;
    color: #7C3AED !important;
}

/* === MULTISELECT / SELECTBOX === */
[data-baseweb="select"] > div {
    background-color: #ffffff !important;
    border-color: #DDD0F5 !important;
    border-radius: 7px !important;
}

/* === TOGGLE === */
.stToggle p { color: #3A2870 !important; font-weight: 500 !important; font-size: 13px !important; }

/* === DIVIDER === */
hr { border-color: #E8E2F8 !important; margin: 1rem 0 !important; }

/* === CUSTOM COMPONENTS === */
.cia-card {
    background: #ffffff;
    border: 1px solid #E8E2F8;
    border-radius: 10px;
    padding: 16px 18px;
    margin-bottom: 12px;
    box-shadow: 0 1px 4px rgba(124,58,237,0.06);
}
.cia-card.upd { border-left: 3px solid #7C3AED; }

.field-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 6px;
    margin-bottom: 10px;
}
.field-box {
    background: #F5F2FF;
    border: 1px solid #E8E2F8;
    border-radius: 6px;
    padding: 8px 10px;
}
.field-box.has-value { border-left: 2px solid #A78BFA; }
.field-box .f-label {
    font-size: 8px; color: #B0A3CC;
    letter-spacing: 0.12em; text-transform: uppercase;
    margin-bottom: 3px; font-weight: 600;
}
.field-box .f-value { font-size: 11px; color: #3A1F8C; font-weight: 500; }
.field-box .f-value.empty { color: #D6CEF0; font-style: italic; }

.prog-wrap { display: flex; align-items: center; gap: 10px; margin-top: 10px; }
.prog-track { flex: 1; height: 4px; background: #EDE8FF; border-radius: 2px; }
.prog-fill { height: 100%; border-radius: 2px; background: linear-gradient(90deg, #7C3AED, #A78BFA); }
.prog-label { font-size: 10px; color: #B0A3CC; white-space: nowrap; font-weight: 500; }

.badge {
    display: inline-block;
    font-size: 9px; padding: 3px 10px;
    border-radius: 20px; letter-spacing: 0.1em;
    text-transform: uppercase; font-weight: 700;
}
.badge-new { background: #EDFAF4; color: #1A7A4A; border: 1px solid #A7E8C8; }
.badge-upd { background: #F3EEFF; color: #7C3AED; border: 1px solid #DDD0F5; }

.r-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; }
.r-filename { font-family: 'Syne', sans-serif; font-size: 13px; font-weight: 700; color: #1A1030; }

.groupe-header {
    font-size: 10px; font-weight: 700; color: #7C3AED;
    letter-spacing: 0.16em; text-transform: uppercase;
    margin: 14px 0 8px;
    display: flex; align-items: center; gap: 8px;
}
.groupe-header::after { content: ''; flex: 1; height: 1px; background: #E8E2F8; }

.wf-card {
    background: #fff; border: 1px solid #E8E2F8;
    border-left: 3px solid #7C3AED;
    border-radius: 10px; padding: 20px 22px;
    box-shadow: 0 1px 4px rgba(124,58,237,0.06);
}
.wf-title { font-size: 9px; font-weight: 700; color: #7C3AED; letter-spacing: .18em; text-transform: uppercase; margin-bottom: 14px; }
.wf-step { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 10px; font-size: 13px; color: #9B8FBE; line-height: 1.5; }
.wf-step:last-child { margin-bottom: 0; }
.wf-num { min-width: 22px; height: 22px; border-radius: 50%; background: #F3EEFF; color: #7C3AED; font-size: 10px; font-weight: 700; display: flex; align-items: center; justify-content: center; flex-shrink: 0; border: 1px solid #DDD0F5; }
.wf-step strong { color: #3A2870; font-weight: 600; }

.empty-state { text-align: center; padding: 3rem 1rem; color: #C8BEE0; font-size: 13px; font-family: 'Inter', sans-serif; }
.empty-state .e-icon { font-size: 2rem; margin-bottom: .8rem; }
</style>
""", unsafe_allow_html=True)

# ── Session state ──
for k, v in [("all_results", []), ("current_text", {}), ("db", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

if st.session_state.db is None:
    st.session_state.db = load_db(DEFAULT_DB_PATH)

db      = st.session_state.db
nb_db   = len(db) if db is not None and not db.empty else 0
nb_sess = len(st.session_state.all_results)

# ════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════
with st.sidebar:
    # Logo
    st.markdown("""
    <div style="padding:1.6rem 1.4rem 1.3rem; border-bottom:1px solid #EEE8FF; margin-bottom:.4rem">
      <div style="width:34px;height:34px;border-radius:9px;background:#7C3AED;display:flex;align-items:center;justify-content:center;font-family:Syne,sans-serif;font-size:11px;font-weight:700;color:#fff;letter-spacing:.04em;margin-bottom:.8rem">CIA</div>
      <div style="font-family:Syne,sans-serif;font-size:15px;font-weight:700;color:#1A1030;line-height:1.3;margin-bottom:.3rem">Credit Insurance<br>Analyzer</div>
      <div style="font-size:10px;color:#BEB3D8;letter-spacing:.05em">Direction Risques · Outil interne</div>
    </div>
    """, unsafe_allow_html=True)

    # Config
    st.markdown('<p style="font-size:9px;font-weight:700;letter-spacing:.2em;text-transform:uppercase;color:#D4CCE8;padding:.9rem 0 .3rem">Configuration</p>', unsafe_allow_html=True)
    use_llm = st.toggle("Analyse IA (Claude)", value=True)
    api_key = None
    if use_llm:
        api_key_input = st.text_input("Clé API Anthropic", type="password", placeholder="sk-ant-...")
        api_key = api_key_input or os.environ.get("ANTHROPIC_API_KEY")

    # Base
    st.markdown('<p style="font-size:9px;font-weight:700;letter-spacing:.2em;text-transform:uppercase;color:#D4CCE8;padding:.9rem 0 .3rem">Base de données</p>', unsafe_allow_html=True)
    uploaded_db = st.file_uploader("Charger contrats_base.xlsx", type=["xlsx"], key="db_upload")
    if uploaded_db:
        st.session_state.db = load_db(uploaded_db)
        db    = st.session_state.db
        nb_db = len(db) if db is not None and not db.empty else 0
        st.success(f"✓ {nb_db} contrat(s) chargé(s)")

    last_maj = "—"
    if st.session_state.db is not None and not st.session_state.db.empty:
        if "Date de MAJ" in st.session_state.db.columns:
            lv = st.session_state.db["Date de MAJ"].dropna()
            if not lv.empty: last_maj = lv.iloc[-1]
    nb_db = len(st.session_state.db) if st.session_state.db is not None and not st.session_state.db.empty else 0

    st.markdown(f"""
    <div style="margin:.5rem 0 1rem;background:#F3EEFF;border:1px solid #DDD0F5;border-radius:10px;padding:14px 16px">
      <div style="font-size:9px;font-weight:700;color:#7C3AED;letter-spacing:.16em;text-transform:uppercase;margin-bottom:8px">Base contrats</div>
      <div style="font-family:Syne,sans-serif;font-size:34px;font-weight:700;color:#1A1030;line-height:1">{nb_db}</div>
      <div style="font-size:10px;color:#9B8FBE;margin-top:3px">contrats enregistrés</div>
      <div style="font-size:9px;color:#C8BEE0;margin-top:4px">MAJ · {last_maj}</div>
    </div>
    """, unsafe_allow_html=True)

    if st.session_state.db is not None and not st.session_state.db.empty:
        buf_sb = io.BytesIO()
        with pd.ExcelWriter(buf_sb, engine="openpyxl") as w:
            st.session_state.db.to_excel(w, index=False, sheet_name="Contrats analysés")
        st.download_button("⬇  Télécharger la base",
            data=buf_sb.getvalue(), file_name="contrats_base.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    # Journal
    st.markdown('<p style="font-size:9px;font-weight:700;letter-spacing:.2em;text-transform:uppercase;color:#D4CCE8;padding:.9rem 0 .3rem">Journal</p>', unsafe_allow_html=True)
    log = ProcessingLog()
    summary = log.get_summary()
    if summary:
        st.metric("Contrats traités", summary.get("total_processed", 0))
        st.metric("Taux de succès",   summary.get("success_rate", "N/A"))
    else:
        st.caption("Aucun traitement enregistré")

# ════════════════════════════════════════════
# HEADER + STATS
# ════════════════════════════════════════════
st.markdown("""
<div style="margin-bottom:6px">
  <span style="font-size:9px;font-weight:700;letter-spacing:.22em;text-transform:uppercase;color:#7C3AED">Assurance-crédit · Extraction automatique</span>
</div>
<div style="font-family:Syne,sans-serif;font-size:26px;font-weight:700;color:#1A1030;margin-bottom:3px;letter-spacing:-.01em">Credit Insurance Analyzer</div>
<div style="font-size:10px;color:#C8BEE0;letter-spacing:.06em;margin-bottom:1.4rem">ATRADIUS MODULA · COFACE · EULER HERMES · FR / EN · PDF · DOCX · TXT</div>
""", unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown("""<div style="background:#fff;border:1px solid #E8E2F8;border-radius:10px;padding:14px 16px;box-shadow:0 1px 4px rgba(124,58,237,.06)">
      <div style="font-family:Syne,sans-serif;font-size:26px;font-weight:700;color:#7C3AED;line-height:1">16</div>
      <div style="font-size:9px;color:#C8BEE0;margin-top:4px;letter-spacing:.12em;text-transform:uppercase">Champs extraits</div>
    </div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""<div style="background:#fff;border:1px solid #E8E2F8;border-radius:10px;padding:14px 16px;box-shadow:0 1px 4px rgba(124,58,237,.06)">
      <div style="font-family:Syne,sans-serif;font-size:26px;font-weight:700;color:#1A1030;line-height:1">{nb_sess}</div>
      <div style="font-size:9px;color:#C8BEE0;margin-top:4px;letter-spacing:.12em;text-transform:uppercase">Session en cours</div>
    </div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""<div style="background:#fff;border:1px solid #E8E2F8;border-radius:10px;padding:14px 16px;box-shadow:0 1px 4px rgba(124,58,237,.06)">
      <div style="font-family:Syne,sans-serif;font-size:26px;font-weight:700;color:#1A1030;line-height:1">{nb_db}</div>
      <div style="font-size:9px;color:#C8BEE0;margin-top:4px;letter-spacing:.12em;text-transform:uppercase">Total en base</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)

# ════════════════════════════════════════════
# TABS
# ════════════════════════════════════════════
tab_import, tab_results, tab_db, tab_export = st.tabs([
    "📤  Import & Analyse",
    f"📊  Résultats ({nb_sess})",
    f"🗄  Base de données ({nb_db})",
    "📥  Export",
])

# ═══ TAB 1 — IMPORT ══════════════════════════
with tab_import:
    uploaded_files = st.file_uploader(
        "Déposer les contrats à analyser",
        type=["pdf","docx","txt"],
        accept_multiple_files=True,
        key="contracts_upload"
    )

    if uploaded_files:
        st.markdown(f"<p style='font-size:11px;color:#7C3AED;font-weight:600;margin:.3rem 0 .8rem'>→ {len(uploaded_files)} fichier(s) sélectionné(s)</p>", unsafe_allow_html=True)
        col1, col2 = st.columns([3, 1])
        with col1:
            analyze_btn = st.button(f"Analyser {len(uploaded_files)} contrat(s)", type="primary", use_container_width=True)
        with col2:
            show_text = st.checkbox("Voir texte extrait")

        if analyze_btn:
            pb   = st.progress(0)
            stxt = st.empty()
            for i, uf in enumerate(uploaded_files):
                stxt.markdown(f"<p style='font-size:11px;color:#B0A3CC'>→ Analyse de {uf.name}…</p>", unsafe_allow_html=True)
                pb.progress(i / len(uploaded_files))
                try:
                    tmp = save_uploaded_file(uf, TMP_DIR)
                    raw, _ = extract_text(tmp)
                    clean  = clean_text(raw)
                    st.session_state.current_text[uf.name] = clean
                    if not clean.strip():
                        st.warning(f"Aucun texte extrait — {uf.name}"); continue
                    regex_r = extract_all_fields(clean)
                    llm_r   = {}
                    if use_llm and (api_key or os.environ.get("ANTHROPIC_API_KEY")):
                        with st.spinner("Analyse IA…"):
                            llm_r = extract_with_llm(clean, api_key=api_key)
                    merged  = merge_results(regex_r, llm_r)
                    found, total = count_found_fields(merged)
                    num_police   = (merged.get("numero_police") or {}).get("value", "")
                    already      = exists_in_db(st.session_state.db, num_police) if (num_police and st.session_state.db is not None) else False
                    entry = {
                        "filename": uf.name, "merged": merged,
                        "already_exists": already, "numero_police": num_police,
                        "edited_values": {
                            f: (v.get("value","") if isinstance(v,dict) and not isinstance(v.get("value"),list) else "")
                            for f,v in merged.items()
                        }
                    }
                    ex = next((r for r in st.session_state.all_results if r["filename"]==uf.name), None)
                    if ex: ex.update(entry)
                    else:  st.session_state.all_results.append(entry)
                    note = " · existe déjà en base" if already else ""
                    st.success(f"{'🔄' if already else '✓'} **{uf.name}** · {found}/{total} champs extraits{note}")
                except Exception as e:
                    st.error(f"Erreur — {uf.name} : {e}"); logger.error(e)
            pb.progress(1.0); stxt.empty()
            st.info("→ Allez dans l'onglet **Résultats** pour valider et enregistrer.")

        if show_text and st.session_state.current_text:
            sel = st.selectbox("Fichier", list(st.session_state.current_text.keys()))
            if sel:
                st.text_area("Texte extrait", st.session_state.current_text[sel], height=250)
    else:
        st.markdown("""
        <div class="empty-state">
          <div class="e-icon">📄</div>
          Glissez vos contrats ci-dessus pour commencer
        </div>""", unsafe_allow_html=True)

# ═══ TAB 2 — RÉSULTATS ═══════════════════════
with tab_results:
    if not st.session_state.all_results:
        st.markdown("""
        <div class="empty-state">
          <div class="e-icon">📊</div>
          Aucun contrat analysé. Utilisez l'onglet <strong>Import & Analyse</strong>.
        </div>""", unsafe_allow_html=True)
    else:
        for result in st.session_state.all_results:
            already = result.get("already_exists", False)
            merged  = result["merged"]
            found, total = count_found_fields(merged)
            pct = found/total if total>0 else 0

            fhtml = '<div class="field-grid">'
            for field, label in FIELD_LABELS.items():
                fd  = merged.get(field, {})
                val = result["edited_values"].get(field, fd.get("value","") if isinstance(fd,dict) else "")
                if isinstance(val, list): val = ""
                has = "has-value" if val else ""
                emp = "empty" if not val else ""
                fhtml += f'<div class="field-box {has}"><div class="f-label">{label}</div><div class="f-value {emp}">{val if val else "—"}</div></div>'
            fhtml += '</div>'
            fhtml += f'<div class="prog-wrap"><div class="prog-track"><div class="prog-fill" style="width:{pct*100:.0f}%"></div></div><div class="prog-label">{found} / {total} champs</div></div>'

            badge_cls = "badge-upd" if already else "badge-new"
            badge_txt = "Mise à jour" if already else "Nouveau"
            card_cls  = "cia-card upd" if already else "cia-card"

            st.markdown(f"""
            <div class="{card_cls}">
              <div class="r-header">
                <div class="r-filename">📄 {result['filename']}</div>
                <span class="badge {badge_cls}">{badge_txt}</span>
              </div>
              {fhtml}
            </div>""", unsafe_allow_html=True)

            if already:
                st.warning(f"⚠️ Le contrat **{result.get('numero_police','')}** existe déjà en base. Enregistrer écrasera la ligne existante.")

            # Groupe de polices
            gp_d    = merged.get("groupe_polices", {})
            polices = gp_d.get("value") if isinstance(gp_d,dict) else None
            if polices and isinstance(polices, list):
                st.markdown(f'<div class="groupe-header">Groupe de polices — {len(polices)} membres</div>', unsafe_allow_html=True)
                df_g = pd.DataFrame(polices).rename(columns={"numero":"N° Police","devise":"Devise","assure":"Assuré"})
                df_g["Devise"] = df_g.get("Devise", pd.Series()).fillna("—")
                st.dataframe(df_g, use_container_width=True, hide_index=True)

            with st.expander("✏️  Modifier les valeurs extraites"):
                ec = st.columns(2)
                for idx, (field, label) in enumerate(FIELD_LABELS.items()):
                    with ec[idx%2]:
                        fd  = merged.get(field,{})
                        cur = result["edited_values"].get(field, fd.get("value","") if isinstance(fd,dict) else "")
                        if isinstance(cur,list): cur=""
                        conf = fd.get("confidence",0.0) if isinstance(fd,dict) else 0.0
                        nv = st.text_input(
                            f"{label} {format_confidence(conf)}",
                            value=str(cur) if cur else "",
                            key=f"e_{result['filename']}_{field}",
                            help=f"Méthode : {fd.get('method','—') if isinstance(fd,dict) else '—'}"
                        )
                        result["edited_values"][field] = nv

            sc1, sc2, _ = st.columns([1,1,4])
            with sc1:
                if st.button("💾  Enregistrer en base", key=f"save_{result['filename']}", type="primary", use_container_width=True):
                    me = dict(merged)
                    for field in FIELD_LABELS:
                        ev   = result["edited_values"].get(field,"")
                        orig = merged.get(field,{})
                        me[field] = {"value":ev,
                            "confidence": orig.get("confidence",0.0) if isinstance(orig,dict) else 0.0,
                            "method":     orig.get("method","manual") if isinstance(orig,dict) else "manual"}
                    db_row = merged_to_db_row(result["filename"], me)
                    if st.session_state.db is None:
                        st.session_state.db = load_db(DEFAULT_DB_PATH)
                    st.session_state.db, action = upsert_row(st.session_state.db, db_row)
                    st.success(f"✓ Contrat {'mis à jour' if action=='updated' else 'ajouté'} en base")
                    result["already_exists"] = True
            with sc2:
                st.button("Ignorer", key=f"skip_{result['filename']}", type="secondary", use_container_width=True)

            st.markdown("<hr>", unsafe_allow_html=True)

# ═══ TAB 3 — BASE ════════════════════════════
with tab_db:
    db = st.session_state.db
    if db is None or db.empty:
        st.markdown("""
        <div class="empty-state">
          <div class="e-icon">🗄</div>
          Aucune donnée. Chargez une base depuis la sidebar ou enregistrez des contrats.
        </div>""", unsafe_allow_html=True)
    else:
        fc1, fc2, fc3, fc4 = st.columns([1,1,1,1])
        f_pol = fc1.text_input("N° Police","")
        f_ass = fc2.text_input("Assuré","")
        f_asr = fc3.text_input("Assureur","")
        with fc4:
            st.markdown("<div style='height:1.7rem'></div>", unsafe_allow_html=True)
            if st.button("🔄  Recharger", type="secondary", use_container_width=True):
                st.session_state.db = load_db(DEFAULT_DB_PATH); st.rerun()

        mask = pd.Series([True]*len(db), index=db.index)
        if f_pol: mask &= db["N° Police"].str.contains(f_pol, case=False, na=False)
        if f_ass: mask &= db["Assuré"].str.contains(f_ass,   case=False, na=False)
        if f_asr: mask &= db["Assureur"].str.contains(f_asr, case=False, na=False)
        db_view = db[mask].copy()
        st.markdown(f"<p style='font-size:10px;color:#B0A3CC;font-weight:500;margin:.3rem 0 .8rem'>{len(db_view)} résultat(s) sur {len(db)} contrat(s)</p>", unsafe_allow_html=True)

        edited_df = st.data_editor(db_view, use_container_width=True, hide_index=False, num_rows="fixed", key="db_editor")

        st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)
        da1, da2, da3 = st.columns([1,2,1])
        with da1:
            if st.button("💾  Sauvegarder", type="primary", use_container_width=True):
                now = datetime.now().strftime("%d/%m/%Y %H:%M")
                for idx in edited_df.index:
                    for col in DB_COLUMNS:
                        if col in edited_df.columns and idx in st.session_state.db.index:
                            st.session_state.db.at[idx,col] = edited_df.at[idx,col]
                    if idx in st.session_state.db.index:
                        st.session_state.db.at[idx,"Date de MAJ"] = now
                st.success("✓ Modifications sauvegardées · Téléchargez la base via la sidebar")
        with da2:
            to_delete = st.multiselect("Supprimer des polices",
                options=db_view["N° Police"].tolist(), placeholder="Sélectionner N° Police(s)…")
        with da3:
            if to_delete:
                if st.button(f"🗑  Supprimer ({len(to_delete)})", type="primary", use_container_width=True):
                    indices = db[db["N° Police"].isin(to_delete)].index.tolist()
                    st.session_state.db = delete_rows(st.session_state.db, indices)
                    st.success(f"✓ {len(indices)} ligne(s) supprimée(s)"); st.rerun()

# ═══ TAB 4 — EXPORT ══════════════════════════
with tab_export:
    e1, e2 = st.columns([3,2])
    with e1:
        st.markdown("""
        <div class="wf-card">
          <div class="wf-title">Workflow recommandé</div>
          <div class="wf-step"><div class="wf-num">1</div><div><strong>Charger</strong> la base via la sidebar en début de session</div></div>
          <div class="wf-step"><div class="wf-num">2</div><div><strong>Analyser</strong> les nouveaux contrats dans Import & Analyse</div></div>
          <div class="wf-step"><div class="wf-num">3</div><div><strong>Valider</strong> et enregistrer chaque contrat dans Résultats</div></div>
          <div class="wf-step"><div class="wf-num">4</div><div><strong>Télécharger</strong> la base mise à jour via la sidebar ou ci-contre</div></div>
          <div class="wf-step"><div class="wf-num">5</div><div><strong>Sauvegarder</strong> sur OneDrive / SharePoint pour partager avec l'équipe</div></div>
        </div>""", unsafe_allow_html=True)

    with e2:
        st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
        db_final = st.session_state.db
        if db_final is not None and not db_final.empty:
            # Export base persistante
            buf_ex = io.BytesIO()
            with pd.ExcelWriter(buf_ex, engine="openpyxl") as w:
                db_final.to_excel(w, index=False, sheet_name="Contrats analysés")
            st.download_button("⬇  Télécharger contrats_base.xlsx",
                data=buf_ex.getvalue(), file_name="contrats_base.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary")
            st.markdown(f"<p style='font-size:9px;color:#C8BEE0;text-align:center;margin-top:.5rem'>{len(db_final)} contrat(s) · {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>", unsafe_allow_html=True)
            st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)

        # Export Excel session (format original)
        if st.session_state.all_results:
            st.markdown('<p style="font-size:10px;font-weight:600;color:#9B8FBE;letter-spacing:.08em;text-transform:uppercase;margin-bottom:.5rem">Export session (format original)</p>', unsafe_allow_html=True)
            if st.button("📥  Générer Excel session", use_container_width=True, type="secondary"):
                rows = []
                for r in st.session_state.all_results:
                    me = {}
                    for f in FIELD_LABELS:
                        ev   = r["edited_values"].get(f,"")
                        orig = r["merged"].get(f,{})
                        me[f] = {"value":ev,
                            "confidence": orig.get("confidence",0.0) if isinstance(orig,dict) else 0.0,
                            "method":     orig.get("method","manual") if isinstance(orig,dict) else "manual"}
                    me["groupe_polices"] = r["merged"].get("groupe_polices",{})
                    rows.append(results_to_row(r["filename"], me))
                try:
                    path = export_to_excel(rows, OUTPUT_EXCEL)
                    with open(path,"rb") as f:
                        st.download_button("⬇  Télécharger Excel session",
                            data=f.read(), file_name=OUTPUT_EXCEL,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True)
                    st.success(f"✓ {len(rows)} contrat(s) exporté(s)")
                except Exception as e:
                    st.error(f"Erreur export : {e}")
        else:
            st.markdown("""
            <div class="empty-state" style="padding:1.5rem 0">
              <div class="e-icon">📭</div>
              Aucune donnée à exporter.
            </div>""", unsafe_allow_html=True)

st.markdown("""
<div style="margin-top:3rem;padding-top:1.2rem;border-top:1px solid #E8E2F8;
text-align:center;font-size:9px;color:#DDD5F0;letter-spacing:.18em;font-family:Inter,sans-serif">
CIA · CREDIT INSURANCE ANALYZER · POWERED BY CLAUDE
</div>""", unsafe_allow_html=True)
