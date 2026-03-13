"""
app.py - Credit Insurance Analyzer
Design: Model B — light white/violet
"""
import os, logging, io
from datetime import datetime
import pandas as pd
import streamlit as st

from extractor import extract_text, clean_text
from nlp_parser import extract_with_llm, merge_results
from regex_rules import extract_all_fields
from db_manager import (load_db, upsert_row, delete_rows,
                        exists_in_db, merged_to_db_row, DB_COLUMNS)
from utils import count_found_fields, save_uploaded_file, setup_logging

setup_logging()
logger = logging.getLogger(__name__)
DEFAULT_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "contrats_base.xlsx")
TMP_DIR = "/tmp/cia_uploads"

FIELD_LABELS = {
    "numero_police": "N° Police", "assure": "Assuré", "assureur": "Assureur",
    "courtier": "Courtier", "taux_prime": "Taux de prime",
    "prime_provisionnelle": "Prime provisionnelle", "quotites_garanties": "% Assuré",
    "delai_indemnisation": "Délai indemn.", "delai_max_credit": "Délai crédit",
    "date_prise_effet": "Date d'effet", "date_echeance": "Date d'échéance",
    "duree_police": "Durée", "devise": "Devise",
    "limite_decaissement": "Limite décaiss.", "zone_discretionnaire": "Zone discr.",
}

st.set_page_config(
    page_title="CIA · Credit Insurance Analyzer",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=Inter:wght@300;400;500;600&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    background: #F8F6FF !important;
    font-family: 'Inter', sans-serif !important;
    color: #1A1030 !important;
}

#MainMenu, footer, header,
[data-testid="stToolbar"], [data-testid="stDecoration"],
[data-testid="stStatusWidget"] { display: none !important; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: #FFFFFF !important;
    border-right: 1px solid #EDE8FF !important;
    min-width: 240px !important;
    max-width: 240px !important;
}
[data-testid="stSidebar"] > div { padding: 0 !important; }

/* ── Main content ── */
.main .block-container {
    max-width: 960px !important;
    padding: 2rem 2.5rem 4rem !important;
    margin: 0 auto !important;
}

/* ── Sidebar logo ── */
.sb-logo {
    padding: 1.6rem 1.4rem 1.3rem;
    border-bottom: 1px solid #F0EBF8;
    margin-bottom: .4rem;
}
.sb-gem {
    display: inline-flex; align-items: center; justify-content: center;
    width: 34px; height: 34px; border-radius: 8px;
    background: #7C3AED;
    font-family: 'Syne', sans-serif; font-size: 11px; font-weight: 700;
    color: #fff; letter-spacing: .04em; margin-bottom: .7rem;
}
.sb-title {
    font-family: 'Syne', sans-serif; font-size: 15px; font-weight: 700;
    color: #1A1030; line-height: 1.25; margin-bottom: .25rem;
}
.sb-sub { font-size: 10px; color: #C8BEE0; letter-spacing: .05em; }
.sb-section {
    font-size: 9px; font-weight: 600; letter-spacing: .2em;
    text-transform: uppercase; color: #D4CCE8;
    padding: .9rem 1.4rem .4rem;
}
.sb-dbcard {
    margin: 1rem 1rem 0;
    background: #F3EEFF;
    border: 1px solid #DDD0F5;
    border-radius: 10px; padding: 14px;
}
.sb-dbcard-title {
    font-size: 9px; font-weight: 600; color: #7C3AED;
    letter-spacing: .16em; text-transform: uppercase; margin-bottom: 8px;
}
.sb-dbnum {
    font-family: 'Syne', sans-serif; font-size: 32px;
    font-weight: 700; color: #1A1030; line-height: 1;
}
.sb-dblbl { font-size: 10px; color: #9B8FBE; margin-top: 2px; }
.sb-dbdate { font-size: 9px; color: #C8BEE0; margin-top: 4px; }

/* ── Page header ── */
.pg-eyebrow {
    font-size: 9px; font-weight: 600; letter-spacing: .22em;
    text-transform: uppercase; color: #7C3AED; margin-bottom: 5px;
}
.pg-title {
    font-family: 'Syne', sans-serif; font-size: 26px; font-weight: 700;
    color: #1A1030; margin-bottom: 3px; letter-spacing: -.01em;
}
.pg-sub { font-size: 10px; color: #C8BEE0; letter-spacing: .06em; margin-bottom: 1.5rem; }

/* ── Stat cards ── */
.stat-row { display: flex; gap: 10px; margin-bottom: 1.5rem; }
.stat-card {
    flex: 1; background: #fff;
    border: 1px solid #EDE8FF;
    border-radius: 10px; padding: 14px 16px;
    box-shadow: 0 1px 3px rgba(124,58,237,.06);
}
.stat-num {
    font-family: 'Syne', sans-serif; font-size: 26px;
    font-weight: 700; color: #7C3AED; line-height: 1;
}
.stat-num.dark { color: #1A1030; }
.stat-lbl {
    font-size: 9px; color: #C8BEE0; margin-top: 4px;
    letter-spacing: .12em; text-transform: uppercase;
}

/* ── Upload zone ── */
[data-testid="stFileUploader"] {
    background: #fff !important;
    border: 1.5px dashed #DDD0F5 !important;
    border-radius: 10px !important;
    transition: border-color .2s !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: #7C3AED !important;
    background: #F9F7FF !important;
}

/* ── Result cards ── */
.r-card {
    background: #fff;
    border: 1px solid #EDE8FF;
    border-radius: 10px; padding: 18px; margin-bottom: 14px;
    box-shadow: 0 1px 4px rgba(124,58,237,.05);
}
.r-card.upd { border-left: 3px solid #7C3AED; }
.r-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; }
.r-name { font-size: 13px; font-weight: 600; color: #1A1030; }
.badge {
    font-size: 9px; padding: 3px 9px; border-radius: 20px;
    letter-spacing: .1em; text-transform: uppercase; font-weight: 600;
}
.badge-new { background: rgba(52,199,133,.1); color: #1A7A4A; border: 1px solid rgba(52,199,133,.3); }
.badge-upd { background: #F3EEFF; color: #7C3AED; border: 1px solid #DDD0F5; }

/* ── Field grid ── */
.fg { display: grid; grid-template-columns: repeat(5,1fr); gap: 6px; }
.fi {
    background: #F8F6FF;
    border: 1px solid #EDE8FF;
    border-radius: 6px; padding: 8px 10px;
}
.fi.ok { border-left: 2px solid rgba(124,58,237,.35); }
.fi .fl { font-size: 8px; color: #C8BEE0; letter-spacing: .12em; text-transform: uppercase; margin-bottom: 3px; }
.fi .fv { font-size: 11px; color: #3A2870; font-weight: 500; }
.fi .fv.empty { color: #DDD5F0; }

/* ── Progress ── */
.prog-row { display: flex; align-items: center; gap: 10px; margin-top: 12px; }
.prog-bg { flex: 1; height: 3px; background: #EDE8FF; border-radius: 2px; }
.prog-fill { height: 100%; border-radius: 2px; background: linear-gradient(90deg,#7C3AED,#A78BFA); }
.prog-txt { font-size: 9px; color: #B8A9D4; white-space: nowrap; }

/* ── Groupe header ── */
.gt-hdr {
    font-size: 9px; font-weight: 600; color: #7C3AED;
    letter-spacing: .16em; text-transform: uppercase;
    margin: 14px 0 8px; display: flex; align-items: center; gap: 8px;
}
.gt-hdr::after { content:''; flex:1; height:1px; background: #EDE8FF; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 2px solid #EDE8FF !important;
    gap: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: #B8A9D4 !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 12px !important; font-weight: 500 !important;
    padding: .65rem 1.2rem !important;
    border-bottom: 2px solid transparent !important;
    margin-bottom: -2px !important;
    border-radius: 0 !important;
}
.stTabs [aria-selected="true"] {
    color: #7C3AED !important;
    border-bottom-color: #7C3AED !important;
    font-weight: 600 !important;
}
.stTabs [data-baseweb="tab-panel"] { padding: 1.5rem 0 0 !important; }

/* ── Buttons ── */
.stButton > button {
    background: #7C3AED !important; color: #fff !important;
    border: none !important; border-radius: 7px !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 12px !important; font-weight: 500 !important;
    padding: .5rem 1.2rem !important;
    transition: background .15s, transform .1s !important;
    box-shadow: 0 2px 8px rgba(124,58,237,.25) !important;
}
.stButton > button:hover {
    background: #6D28D9 !important; transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(124,58,237,.35) !important;
}
.stButton > button[kind="secondary"] {
    background: #fff !important; color: #9B8FBE !important;
    border: 1px solid #DDD0F5 !important;
    box-shadow: none !important;
}
.stButton > button[kind="secondary"]:hover {
    color: #7C3AED !important; border-color: #B8A0E8 !important;
}
.stDownloadButton > button {
    background: #F3EEFF !important; color: #7C3AED !important;
    border: 1px solid #DDD0F5 !important; border-radius: 7px !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 12px !important; font-weight: 500 !important;
    box-shadow: none !important;
}
.stDownloadButton > button:hover {
    background: #EDE5FF !important; border-color: #B8A0E8 !important;
}

/* ── Inputs ── */
.stTextInput > label, .stFileUploader > label,
.stToggle > label, .stSelectbox > label,
.stMultiSelect > label, .stTextArea > label {
    font-family: 'Inter', sans-serif !important;
    font-size: 10px !important; letter-spacing: .1em !important;
    text-transform: uppercase !important; color: #C8BEE0 !important;
    font-weight: 600 !important;
}
.stTextInput input, .stTextArea textarea {
    background: #fff !important;
    border: 1px solid #DDD0F5 !important;
    border-radius: 7px !important; color: #1A1030 !important;
    font-family: 'Inter', sans-serif !important; font-size: 13px !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: #7C3AED !important;
    box-shadow: 0 0 0 3px rgba(124,58,237,.1) !important;
}

/* ── Progress bar widget ── */
.stProgress > div > div {
    background: linear-gradient(90deg,#7C3AED,#A78BFA) !important;
    border-radius: 4px !important;
}
.stProgress > div { background: #EDE8FF !important; border-radius: 4px !important; }

/* ── Expander ── */
.streamlit-expanderHeader {
    background: #F8F6FF !important;
    border: 1px solid #EDE8FF !important; border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 12px !important; color: #9B8FBE !important;
}
.streamlit-expanderContent {
    background: #F8F6FF !important;
    border: 1px solid #EDE8FF !important;
    border-top: none !important; border-radius: 0 0 8px 8px !important;
}

/* ── Alerts ── */
.stSuccess { background: rgba(52,199,133,.07) !important; border: 1px solid rgba(52,199,133,.25) !important; border-radius: 8px !important; color: #1A5C38 !important; }
.stWarning { background: rgba(234,179,8,.07) !important; border: 1px solid rgba(234,179,8,.25) !important; border-radius: 8px !important; }
.stError   { background: rgba(220,38,38,.06) !important; border: 1px solid rgba(220,38,38,.2) !important; border-radius: 8px !important; }
.stInfo    { background: #F3EEFF !important; border: 1px solid #DDD0F5 !important; border-radius: 8px !important; color: #5B21B6 !important; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
    border: 1px solid #EDE8FF !important;
    border-radius: 9px !important; overflow: hidden !important;
    box-shadow: 0 1px 4px rgba(124,58,237,.05) !important;
}

/* ── Multiselect / Selectbox ── */
[data-testid="stMultiSelect"] > div > div,
[data-testid="stSelectbox"] > div > div {
    background: #fff !important;
    border-color: #DDD0F5 !important;
    border-radius: 7px !important;
}

/* ── Toggle ── */
.stToggle label { color: #3A2870 !important; }

/* ── Divider ── */
hr { border-color: #EDE8FF !important; }

/* ── Workflow card ── */
.wf-card {
    background: #fff; border: 1px solid #EDE8FF;
    border-left: 3px solid #7C3AED;
    border-radius: 10px; padding: 20px 22px;
    box-shadow: 0 1px 4px rgba(124,58,237,.06);
}
.wf-title {
    font-size: 9px; font-weight: 600; color: #7C3AED;
    letter-spacing: .16em; text-transform: uppercase; margin-bottom: 14px;
}
.wf-step {
    display: flex; align-items: flex-start; gap: 12px;
    margin-bottom: 10px; font-size: 13px; color: #9B8FBE; line-height: 1.5;
}
.wf-step:last-child { margin-bottom: 0; }
.wf-num {
    min-width: 22px; height: 22px; border-radius: 50%;
    background: #F3EEFF; color: #7C3AED;
    font-size: 10px; font-weight: 700;
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
    border: 1px solid #DDD0F5;
}
.wf-step strong { color: #3A2870; font-weight: 600; }

/* ── Empty state ── */
.empty-state {
    text-align: center; padding: 3rem 1rem; color: #C8BEE0; font-size: 13px;
}
.empty-state .emoji { font-size: 2rem; margin-bottom: .8rem; }
</style>
""", unsafe_allow_html=True)

# ── Session state ──
for k, v in [("all_results", []), ("current_text", {}), ("db", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

if st.session_state.db is None:
    st.session_state.db = load_db(DEFAULT_DB_PATH)

db = st.session_state.db
nb_db = len(db) if db is not None and not db.empty else 0
nb_session = len(st.session_state.all_results)

# ════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div class="sb-logo">
        <div class="sb-gem">CIA</div>
        <div class="sb-title">Credit Insurance<br>Analyzer</div>
        <div class="sb-sub">Direction Risques · Outil interne</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="sb-section">Configuration</div>', unsafe_allow_html=True)
    use_llm = st.toggle("Analyse IA (Claude)", value=True)
    api_key = None
    if use_llm:
        api_key_input = st.text_input("Clé API Anthropic", type="password", placeholder="sk-ant-...")
        api_key = api_key_input or os.environ.get("ANTHROPIC_API_KEY")

    st.markdown('<div class="sb-section">Base de données</div>', unsafe_allow_html=True)
    uploaded_db = st.file_uploader("Charger contrats_base.xlsx", type=["xlsx"], key="db_upload")
    if uploaded_db:
        st.session_state.db = load_db(uploaded_db)
        db = st.session_state.db
        nb_db = len(db) if db is not None and not db.empty else 0
        st.success(f"✓ {nb_db} contrat(s) chargé(s)")

    last_maj = "—"
    if db is not None and not db.empty and "Date de MAJ" in db.columns:
        lv = db["Date de MAJ"].dropna()
        if not lv.empty: last_maj = lv.iloc[-1]

    nb_db = len(st.session_state.db) if st.session_state.db is not None and not st.session_state.db.empty else 0

    st.markdown(f"""
    <div class="sb-dbcard">
        <div class="sb-dbcard-title">Base contrats</div>
        <div class="sb-dbnum">{nb_db}</div>
        <div class="sb-dblbl">contrats enregistrés</div>
        <div class="sb-dbdate">MAJ · {last_maj}</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)

    if st.session_state.db is not None and not st.session_state.db.empty:
        buf_sb = io.BytesIO()
        with pd.ExcelWriter(buf_sb, engine="openpyxl") as w:
            st.session_state.db.to_excel(w, index=False, sheet_name="Contrats analysés")
        st.download_button(
            "⬇  Télécharger la base",
            data=buf_sb.getvalue(), file_name="contrats_base.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ════════════════════════════════════════════
# MAIN — HEADER + STATS
# ════════════════════════════════════════════
st.markdown('<div class="pg-eyebrow">Assurance-crédit · Extraction automatique</div>', unsafe_allow_html=True)
st.markdown('<div class="pg-title">Credit Insurance Analyzer</div>', unsafe_allow_html=True)
st.markdown('<div class="pg-sub">ATRADIUS MODULA · COFACE · EULER HERMES · FR / EN · PDF · DOCX · TXT</div>', unsafe_allow_html=True)

st.markdown(f"""
<div class="stat-row">
  <div class="stat-card"><div class="stat-num">16</div><div class="stat-lbl">Champs extraits</div></div>
  <div class="stat-card"><div class="stat-num dark">{nb_session}</div><div class="stat-lbl">Session en cours</div></div>
  <div class="stat-card"><div class="stat-num dark">{nb_db}</div><div class="stat-lbl">Total en base</div></div>
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════
# TABS
# ════════════════════════════════════════════
tab_import, tab_results, tab_db, tab_export = st.tabs([
    "📤  Import & Analyse",
    f"📊  Résultats ({nb_session})",
    f"🗄  Base de données ({nb_db})",
    "📥  Export",
])

# ═══════════════════════════════════════
# TAB 1 — IMPORT
# ═══════════════════════════════════════
with tab_import:
    uploaded_files = st.file_uploader(
        "Déposer les contrats à analyser",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=True,
        key="contracts_upload"
    )

    if uploaded_files:
        st.markdown(f"<div style='font-size:11px;color:#7C3AED;margin:.3rem 0 .8rem;font-weight:500'>→ {len(uploaded_files)} fichier(s) sélectionné(s)</div>", unsafe_allow_html=True)
        c1, c2 = st.columns([2, 1])
        with c1:
            analyze_btn = st.button(f"Analyser {len(uploaded_files)} contrat(s)", type="primary", use_container_width=True)
        with c2:
            show_text = st.checkbox("Voir texte extrait")

        if analyze_btn:
            pb = st.progress(0)
            stxt = st.empty()
            for i, uf in enumerate(uploaded_files):
                stxt.markdown(f"<div style='font-size:11px;color:#B8A9D4'>→ Analyse de {uf.name}…</div>", unsafe_allow_html=True)
                pb.progress(i / len(uploaded_files))
                try:
                    tmp = save_uploaded_file(uf, TMP_DIR)
                    raw, _ = extract_text(tmp)
                    clean = clean_text(raw)
                    st.session_state.current_text[uf.name] = clean
                    if not clean.strip():
                        st.warning(f"Aucun texte extrait — {uf.name}"); continue
                    regex_r = extract_all_fields(clean)
                    llm_r = {}
                    if use_llm and (api_key or os.environ.get("ANTHROPIC_API_KEY")):
                        with st.spinner("Analyse IA…"):
                            llm_r = extract_with_llm(clean, api_key=api_key)
                    merged = merge_results(regex_r, llm_r)
                    found, total = count_found_fields(merged)
                    num_police = (merged.get("numero_police") or {}).get("value", "")
                    already = exists_in_db(st.session_state.db, num_police) if (num_police and st.session_state.db is not None) else False
                    entry = {
                        "filename": uf.name, "merged": merged,
                        "already_exists": already, "numero_police": num_police,
                        "edited_values": {
                            f: (v.get("value", "") if isinstance(v, dict) and not isinstance(v.get("value"), list) else "")
                            for f, v in merged.items()
                        }
                    }
                    ex = next((r for r in st.session_state.all_results if r["filename"] == uf.name), None)
                    if ex: ex.update(entry)
                    else: st.session_state.all_results.append(entry)
                    icon = "🔄" if already else "✓"
                    note = " · existe déjà en base" if already else ""
                    st.success(f"{icon} **{uf.name}** · {found}/{total} champs extraits{note}")
                except Exception as e:
                    st.error(f"Erreur — {uf.name} : {e}"); logger.error(e)
            pb.progress(1.0); stxt.empty()
            st.info("→ Allez dans l'onglet **Résultats** pour valider et enregistrer.")

        if show_text and st.session_state.current_text:
            sel = st.selectbox("Fichier", list(st.session_state.current_text.keys()))
            if sel:
                st.text_area("Texte extrait", st.session_state.current_text[sel], height=200)
    else:
        st.markdown("""
        <div class="empty-state">
            <div class="emoji">📄</div>
            Glissez vos contrats ci-dessus pour commencer
        </div>
        """, unsafe_allow_html=True)

# ═══════════════════════════════════════
# TAB 2 — RÉSULTATS
# ═══════════════════════════════════════
with tab_results:
    if not st.session_state.all_results:
        st.markdown("""
        <div class="empty-state">
            <div class="emoji">📊</div>
            Aucun contrat analysé. Utilisez l'onglet Import & Analyse.
        </div>
        """, unsafe_allow_html=True)
    else:
        for result in st.session_state.all_results:
            already = result.get("already_exists", False)
            merged = result["merged"]
            found, total = count_found_fields(merged)
            pct = found / total if total > 0 else 0

            fhtml = '<div class="fg">'
            for field, label in FIELD_LABELS.items():
                fd = merged.get(field, {})
                val = result["edited_values"].get(field, fd.get("value", "") if isinstance(fd, dict) else "")
                if isinstance(val, list): val = ""
                fhtml += f'<div class="fi {"ok" if val else ""}"><div class="fl">{label}</div><div class="fv {"empty" if not val else ""}">{val if val else "—"}</div></div>'
            fhtml += '</div>'

            st.markdown(f"""
            <div class="r-card {"upd" if already else ""}">
                <div class="r-head">
                    <div class="r-name">{result["filename"]}</div>
                    <span class="badge {"badge-upd" if already else "badge-new"}">{"Mise à jour" if already else "Nouveau"}</span>
                </div>
                {fhtml}
                <div class="prog-row">
                    <div class="prog-bg"><div class="prog-fill" style="width:{pct*100:.0f}%"></div></div>
                    <div class="prog-txt">{found} / {total} champs</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            if already:
                st.warning(f"⚠️ Le contrat **{result.get('numero_police','')}** existe déjà. Enregistrer écrasera la ligne existante.")

            # Groupe de polices
            gp_d = merged.get("groupe_polices", {})
            polices = gp_d.get("value") if isinstance(gp_d, dict) else None
            if polices and isinstance(polices, list):
                st.markdown(f'<div class="gt-hdr">Groupe de polices — {len(polices)} membres</div>', unsafe_allow_html=True)
                df_g = pd.DataFrame(polices).rename(columns={"numero": "N° Police", "devise": "Devise", "assure": "Assuré"})
                df_g["Devise"] = df_g.get("Devise", pd.Series()).fillna("—")
                st.dataframe(df_g, use_container_width=True, hide_index=True)

            with st.expander("✏️  Modifier les valeurs extraites"):
                ec = st.columns(3)
                for idx, (field, label) in enumerate(FIELD_LABELS.items()):
                    with ec[idx % 3]:
                        fd = merged.get(field, {})
                        cur = result["edited_values"].get(field, fd.get("value", "") if isinstance(fd, dict) else "")
                        if isinstance(cur, list): cur = ""
                        nv = st.text_input(label, value=str(cur) if cur else "", key=f"e_{result['filename']}_{field}")
                        result["edited_values"][field] = nv

            sc1, sc2, _ = st.columns([1, 1, 4])
            with sc1:
                if st.button("💾  Enregistrer en base", key=f"save_{result['filename']}", type="primary", use_container_width=True):
                    me = dict(merged)
                    for field in FIELD_LABELS:
                        ev = result["edited_values"].get(field, "")
                        orig = merged.get(field, {})
                        me[field] = {
                            "value": ev,
                            "confidence": orig.get("confidence", 0.0) if isinstance(orig, dict) else 0.0,
                            "method": orig.get("method", "manual") if isinstance(orig, dict) else "manual"
                        }
                    db_row = merged_to_db_row(result["filename"], me)
                    if st.session_state.db is None:
                        st.session_state.db = load_db(DEFAULT_DB_PATH)
                    st.session_state.db, action = upsert_row(st.session_state.db, db_row)
                    st.success(f"✓ Contrat {'mis à jour' if action=='updated' else 'ajouté'} en base")
                    result["already_exists"] = True
            with sc2:
                st.button("Ignorer", key=f"skip_{result['filename']}", type="secondary", use_container_width=True)

            st.markdown("<hr>", unsafe_allow_html=True)

# ═══════════════════════════════════════
# TAB 3 — BASE
# ═══════════════════════════════════════
with tab_db:
    db = st.session_state.db
    if db is None or db.empty:
        st.markdown("""
        <div class="empty-state">
            <div class="emoji">🗄</div>
            Aucune donnée. Chargez une base depuis la sidebar ou enregistrez des contrats.
        </div>
        """, unsafe_allow_html=True)
    else:
        fc1, fc2, fc3, fc4 = st.columns([1, 1, 1, 1])
        f_police   = fc1.text_input("N° Police", "")
        f_assure   = fc2.text_input("Assuré", "")
        f_assureur = fc3.text_input("Assureur", "")
        with fc4:
            st.markdown("<div style='height:1.65rem'></div>", unsafe_allow_html=True)
            if st.button("🔄  Recharger", type="secondary", use_container_width=True):
                st.session_state.db = load_db(DEFAULT_DB_PATH)
                st.rerun()

        mask = pd.Series([True] * len(db), index=db.index)
        if f_police:   mask &= db["N° Police"].str.contains(f_police, case=False, na=False)
        if f_assure:   mask &= db["Assuré"].str.contains(f_assure, case=False, na=False)
        if f_assureur: mask &= db["Assureur"].str.contains(f_assureur, case=False, na=False)
        db_view = db[mask].copy()

        st.markdown(f"<div style='font-size:10px;color:#B8A9D4;margin:.2rem 0 .8rem;font-weight:500'>{len(db_view)} résultat(s) sur {len(db)} contrat(s)</div>", unsafe_allow_html=True)
        edited_df = st.data_editor(db_view, use_container_width=True, hide_index=False, num_rows="fixed", key="db_editor")

        st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)
        da1, da2, da3 = st.columns([1, 2, 1])
        with da1:
            if st.button("💾  Sauvegarder", type="primary", use_container_width=True):
                now = datetime.now().strftime("%d/%m/%Y %H:%M")
                for idx in edited_df.index:
                    for col in DB_COLUMNS:
                        if col in edited_df.columns and idx in st.session_state.db.index:
                            st.session_state.db.at[idx, col] = edited_df.at[idx, col]
                    if idx in st.session_state.db.index:
                        st.session_state.db.at[idx, "Date de MAJ"] = now
                st.success("✓ Modifications sauvegardées · Téléchargez la base depuis la sidebar")
        with da2:
            to_delete = st.multiselect("Supprimer des polices",
                options=db_view["N° Police"].tolist(), placeholder="Sélectionner N° Police(s)…")
        with da3:
            if to_delete:
                if st.button(f"🗑  Supprimer ({len(to_delete)})", type="primary", use_container_width=True):
                    indices = db[db["N° Police"].isin(to_delete)].index.tolist()
                    st.session_state.db = delete_rows(st.session_state.db, indices)
                    st.success(f"✓ {len(indices)} ligne(s) supprimée(s)")
                    st.rerun()

# ═══════════════════════════════════════
# TAB 4 — EXPORT
# ═══════════════════════════════════════
with tab_export:
    e1, e2 = st.columns([3, 2])
    with e1:
        st.markdown("""
        <div class="wf-card">
            <div class="wf-title">Workflow recommandé</div>
            <div class="wf-step"><div class="wf-num">1</div><div><strong>Charger</strong> la base via la sidebar en début de session</div></div>
            <div class="wf-step"><div class="wf-num">2</div><div><strong>Analyser</strong> les nouveaux contrats dans Import & Analyse</div></div>
            <div class="wf-step"><div class="wf-num">3</div><div><strong>Valider</strong> et enregistrer chaque contrat dans Résultats</div></div>
            <div class="wf-step"><div class="wf-num">4</div><div><strong>Télécharger</strong> la base mise à jour ci-contre ou via la sidebar</div></div>
            <div class="wf-step"><div class="wf-num">5</div><div><strong>Sauvegarder</strong> sur OneDrive / SharePoint</div></div>
        </div>
        """, unsafe_allow_html=True)

    with e2:
        db_final = st.session_state.db
        st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
        if db_final is not None and not db_final.empty:
            buf_ex = io.BytesIO()
            with pd.ExcelWriter(buf_ex, engine="openpyxl") as w:
                db_final.to_excel(w, index=False, sheet_name="Contrats analysés")
            st.download_button(
                "⬇  Télécharger contrats_base.xlsx",
                data=buf_ex.getvalue(), file_name="contrats_base.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary"
            )
            st.markdown(f"<div style='font-size:9px;color:#C8BEE0;text-align:center;margin-top:.5rem'>{len(db_final)} contrat(s) · {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>", unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="empty-state" style="padding:1.5rem 0">
                <div class="emoji">📭</div>
                Aucune donnée à exporter.<br>Enregistrez d'abord des contrats.
            </div>
            """, unsafe_allow_html=True)

# ── Footer ──
st.markdown("""
<div style='margin-top:4rem;padding-top:1.2rem;border-top:1px solid #EDE8FF;
text-align:center;font-size:9px;color:#DDD5F0;letter-spacing:.18em;font-family:Inter,sans-serif'>
CIA · CREDIT INSURANCE ANALYZER · POWERED BY CLAUDE
</div>
""", unsafe_allow_html=True)
