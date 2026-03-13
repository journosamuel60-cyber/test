"""
app.py - Credit Insurance Analyzer
Design: dark violet / white, Syne + Inter
"""
import os, time, logging, io
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
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=Inter:wght@300;400;500;600&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    background: #0C0A14 !important;
    font-family: 'Inter', sans-serif !important;
    color: #EEE8FF !important;
}
[data-testid="stAppViewContainer"] {
    background: radial-gradient(ellipse at 30% 0%, #1A1030 0%, #0C0A14 55%) !important;
}
#MainMenu, footer, header, [data-testid="stToolbar"],
[data-testid="stDecoration"], [data-testid="stStatusWidget"],
[data-testid="stSidebar"] { display: none !important; }

.main .block-container { max-width: 1180px !important; padding: 0 2rem 5rem 2rem !important; }

/* ── Top bar ── */
.cia-topbar {
    position: sticky; top: 0; z-index: 100;
    background: rgba(17,14,28,0.92);
    backdrop-filter: blur(12px);
    border-bottom: 0.5px solid rgba(255,255,255,0.06);
    padding: 0 0 0 0;
    margin: 0 -2rem 2.5rem -2rem;
    display: flex; align-items: center;
    justify-content: space-between;
    height: 56px;
    padding: 0 2rem;
}
.cia-logo { display: flex; align-items: center; gap: 10px; }
.cia-gem {
    width: 30px; height: 30px; border-radius: 7px;
    background: #7C3AED;
    display: flex; align-items: center; justify-content: center;
    font-family: 'Syne', sans-serif; font-size: 10px; font-weight: 700;
    color: #fff; letter-spacing: 0.05em;
}
.cia-name { font-family: 'Syne', sans-serif; font-size: 14px; font-weight: 600; color: #EEE8FF; }
.cia-pipe { width: 1px; height: 16px; background: rgba(255,255,255,0.1); }
.cia-sub { font-size: 10px; color: #3D3060; letter-spacing: 0.06em; }
.cia-nav { display: flex; gap: 4px; }
.cia-nav a {
    font-size: 11px; padding: 5px 13px; border-radius: 20px;
    border: 0.5px solid transparent; color: #6B5F8A;
    text-decoration: none; font-weight: 400;
}
.cia-nav a:hover { color: #C4B5FD; }
.cia-nav a.active {
    background: rgba(124,58,237,0.15);
    border-color: rgba(124,58,237,0.35);
    color: #A78BFA; font-weight: 500;
}

/* ── Section headers ── */
.s-eyebrow {
    font-size: 9px; letter-spacing: 0.22em; text-transform: uppercase;
    color: #7C3AED; margin-bottom: 5px; font-weight: 600;
}
.s-title {
    font-family: 'Syne', sans-serif; font-size: 20px; font-weight: 700;
    color: #EEE8FF; margin-bottom: 3px; letter-spacing: -0.01em;
}
.s-sub { font-size: 10px; color: #3D3060; letter-spacing: 0.06em; margin-bottom: 20px; }
.s-divider {
    height: 1px;
    background: linear-gradient(90deg, rgba(124,58,237,0.3) 0%, transparent 70%);
    margin: 3rem 0 2rem 0;
}

/* ── Stat cards ── */
.stat-row { display: flex; gap: 10px; margin-bottom: 20px; }
.stat-card {
    flex: 1; background: #110E1C;
    border: 0.5px solid rgba(255,255,255,0.06);
    border-radius: 10px; padding: 14px 16px;
}
.stat-card .sv {
    font-family: 'Syne', sans-serif; font-size: 28px;
    font-weight: 700; color: #A78BFA; line-height: 1;
}
.stat-card .sv.w { color: #EEE8FF; }
.stat-card .sl {
    font-size: 9px; color: #3D3060; margin-top: 4px;
    letter-spacing: 0.12em; text-transform: uppercase;
}

/* ── DB sidebar card ── */
.db-card {
    background: rgba(124,58,237,0.08);
    border: 0.5px solid rgba(124,58,237,0.2);
    border-radius: 10px; padding: 16px; margin-bottom: 1.5rem;
}
.db-card-title {
    font-size: 9px; font-weight: 600; color: #A78BFA;
    letter-spacing: 0.18em; text-transform: uppercase; margin-bottom: 8px;
}
.db-card-num {
    font-family: 'Syne', sans-serif; font-size: 32px;
    font-weight: 700; color: #EEE8FF; line-height: 1;
}
.db-card-lbl { font-size: 10px; color: #5A4F7A; }
.db-card-date { font-size: 9px; color: #3D3060; margin-top: 4px; }

/* ── Upload zone ── */
[data-testid="stFileUploader"] {
    background: #110E1C !important;
    border: 1px dashed rgba(124,58,237,0.25) !important;
    border-radius: 10px !important;
    padding: 0.5rem !important;
    transition: border-color 0.2s !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: rgba(124,58,237,0.5) !important;
}

/* ── Result cards ── */
.r-card {
    background: #110E1C;
    border: 0.5px solid rgba(255,255,255,0.06);
    border-radius: 10px; padding: 18px; margin-bottom: 12px;
}
.r-card.updated { border-left: 2px solid #7C3AED; }
.r-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; }
.r-filename { font-size: 12px; font-weight: 500; color: #D4C8FF; font-family: 'Inter', sans-serif; }
.badge {
    font-size: 9px; padding: 3px 9px; border-radius: 20px;
    letter-spacing: 0.1em; text-transform: uppercase; font-weight: 500;
}
.badge-new { background: rgba(52,199,133,0.1); color: #34C785; border: 0.5px solid rgba(52,199,133,0.2); }
.badge-upd { background: rgba(124,58,237,0.12); color: #A78BFA; border: 0.5px solid rgba(124,58,237,0.25); }

/* ── Fields grid ── */
.fg { display: grid; grid-template-columns: repeat(5, 1fr); gap: 5px; }
.fi {
    background: rgba(255,255,255,0.02);
    border: 0.5px solid rgba(255,255,255,0.05);
    border-radius: 5px; padding: 7px 9px;
}
.fi.ok { border-left: 2px solid rgba(167,139,250,0.35); }
.fi .fl { font-size: 8px; color: #3D3060; letter-spacing: 0.12em; text-transform: uppercase; margin-bottom: 3px; }
.fi .fv { font-size: 11px; color: #C4B5FD; }
.fi .fv.empty { color: #1E1830; }

/* ── Progress ── */
.prog-row { display: flex; align-items: center; gap: 10px; margin-top: 12px; }
.prog-bg { flex: 1; height: 2px; background: rgba(255,255,255,0.05); border-radius: 1px; }
.prog-fill { height: 100%; background: linear-gradient(90deg, #7C3AED, #A78BFA); border-radius: 1px; }
.prog-txt { font-size: 9px; color: #3D3060; white-space: nowrap; }

/* ── Groupe table header ── */
.gt-header {
    font-size: 9px; font-weight: 600; color: #A78BFA;
    letter-spacing: 0.16em; text-transform: uppercase;
    margin: 14px 0 8px 0;
    display: flex; align-items: center; gap: 8px;
}
.gt-header::after { content: ''; flex: 1; height: 0.5px; background: rgba(124,58,237,0.2); }

/* ── Buttons ── */
.stButton > button {
    background: #7C3AED !important; color: #fff !important;
    border: none !important; border-radius: 6px !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 12px !important; font-weight: 500 !important;
    letter-spacing: 0.03em !important;
    padding: 0.5rem 1.2rem !important;
    transition: all 0.15s !important;
}
.stButton > button:hover {
    background: #6D28D9 !important;
    transform: translateY(-1px) !important;
}
.stButton > button[kind="secondary"] {
    background: rgba(255,255,255,0.04) !important;
    color: #5A4F7A !important;
    border: 0.5px solid rgba(255,255,255,0.08) !important;
}
.stButton > button[kind="secondary"]:hover {
    color: #A78BFA !important;
    border-color: rgba(124,58,237,0.3) !important;
}
.stDownloadButton > button {
    background: rgba(124,58,237,0.12) !important;
    color: #A78BFA !important;
    border: 0.5px solid rgba(124,58,237,0.25) !important;
    border-radius: 6px !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 12px !important; font-weight: 500 !important;
}
.stDownloadButton > button:hover {
    background: rgba(124,58,237,0.22) !important;
}

/* ── Inputs ── */
.stTextInput > label, .stFileUploader > label,
.stToggle > label, .stSelectbox > label,
.stMultiSelect > label, .stTextArea > label {
    font-family: 'Inter', sans-serif !important;
    font-size: 10px !important; letter-spacing: 0.1em !important;
    text-transform: uppercase !important; color: #3D3060 !important;
    font-weight: 500 !important;
}
.stTextInput input, .stTextArea textarea {
    background: rgba(255,255,255,0.04) !important;
    border: 0.5px solid rgba(255,255,255,0.08) !important;
    border-radius: 6px !important; color: #EEE8FF !important;
    font-family: 'Inter', sans-serif !important; font-size: 13px !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: rgba(124,58,237,0.5) !important;
    box-shadow: 0 0 0 2px rgba(124,58,237,0.1) !important;
}

/* ── Progress bar ── */
.stProgress > div > div {
    background: linear-gradient(90deg, #7C3AED, #A78BFA) !important;
    border-radius: 4px !important;
}
.stProgress > div { background: rgba(255,255,255,0.05) !important; border-radius: 4px !important; }

/* ── Expander ── */
.streamlit-expanderHeader {
    background: rgba(255,255,255,0.02) !important;
    border: 0.5px solid rgba(255,255,255,0.06) !important;
    border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 12px !important; color: #5A4F7A !important;
}
.streamlit-expanderContent {
    background: rgba(255,255,255,0.01) !important;
    border: 0.5px solid rgba(255,255,255,0.04) !important;
    border-top: none !important; border-radius: 0 0 8px 8px !important;
}

/* ── Alerts ── */
.stSuccess { background: rgba(52,199,133,0.07) !important; border: 0.5px solid rgba(52,199,133,0.2) !important; border-radius: 7px !important; }
.stWarning { background: rgba(251,191,36,0.07) !important; border: 0.5px solid rgba(251,191,36,0.2) !important; border-radius: 7px !important; }
.stError   { background: rgba(239,68,68,0.07) !important;  border: 0.5px solid rgba(239,68,68,0.2) !important;  border-radius: 7px !important; }
.stInfo    { background: rgba(124,58,237,0.07) !important; border: 0.5px solid rgba(124,58,237,0.2) !important; border-radius: 7px !important; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
    border: 0.5px solid rgba(255,255,255,0.06) !important;
    border-radius: 9px !important; overflow: hidden !important;
}

/* ── Toggle ── */
.stToggle { color: #C4B5FD !important; }

/* ── Multiselect / Selectbox ── */
[data-testid="stMultiSelect"] > div,
[data-testid="stSelectbox"] > div > div {
    background: rgba(255,255,255,0.04) !important;
    border-color: rgba(255,255,255,0.08) !important;
    border-radius: 6px !important;
}

/* ── Workflow card ── */
.workflow-card {
    background: #110E1C;
    border: 0.5px solid rgba(255,255,255,0.06);
    border-left: 2px solid #7C3AED;
    border-radius: 8px; padding: 18px 20px;
}
.workflow-step {
    display: flex; align-items: flex-start; gap: 12px;
    margin-bottom: 10px; font-size: 13px; color: #6B5F8A;
}
.workflow-step:last-child { margin-bottom: 0; }
.step-num {
    min-width: 20px; height: 20px; border-radius: 50%;
    background: rgba(124,58,237,0.2); color: #A78BFA;
    font-size: 10px; font-weight: 600;
    display: flex; align-items: center; justify-content: center;
    margin-top: 1px;
}
.step-txt { line-height: 1.5; }
.step-txt strong { color: #C4B5FD; font-weight: 500; }
</style>
""", unsafe_allow_html=True)

# ── Session state ──
for k, v in [("all_results", []), ("current_text", {}), ("db", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── Top bar ──
nb_db = len(st.session_state.db) if st.session_state.db is not None and not st.session_state.db.empty else 0
nb_session = len(st.session_state.all_results)
st.markdown(f"""
<div class="cia-topbar">
  <div class="cia-logo">
    <div class="cia-gem">CIA</div>
    <div class="cia-name">Credit Insurance Analyzer</div>
    <div class="cia-pipe"></div>
    <div class="cia-sub">Direction Risques · Outil interne</div>
  </div>
  <div class="cia-nav">
    <a href="#base" class="active">Base</a>
    <a href="#import">Import & Analyse</a>
    <a href="#resultats">Résultats</a>
    <a href="#parcourir">Parcourir</a>
    <a href="#export">Export</a>
  </div>
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════
# SECTION — CONFIG
# ════════════════════════════════════════════
c1, c2, c3 = st.columns([1, 1, 2])
with c1:
    use_llm = st.toggle("Analyse IA (Claude)", value=True)
with c2:
    api_key = None
    if use_llm:
        api_key_input = st.text_input("Clé API Anthropic", type="password",
            placeholder="sk-ant-... ou variable ANTHROPIC_API_KEY")
        api_key = api_key_input or os.environ.get("ANTHROPIC_API_KEY")

# ════════════════════════════════════════════
# SECTION 01 — BASE DE DONNÉES
# ════════════════════════════════════════════
st.markdown('<div class="s-divider" id="base"></div>', unsafe_allow_html=True)
st.markdown('<div class="s-eyebrow">01 — Persistance</div>', unsafe_allow_html=True)
st.markdown('<div class="s-title">Base de données</div>', unsafe_allow_html=True)
st.markdown('<div class="s-sub">Chargez votre base en début de session · Téléchargez-la après avoir enregistré vos contrats</div>', unsafe_allow_html=True)

b1, b2 = st.columns([3, 2])
with b1:
    uploaded_db = st.file_uploader("Importer contrats_base.xlsx", type=["xlsx"], key="db_upload")
    if uploaded_db:
        st.session_state.db = load_db(uploaded_db)
        nb_db = len(st.session_state.db) if not st.session_state.db.empty else 0
        st.success(f"✓ Base chargée — {nb_db} contrat(s)")

with b2:
    db = st.session_state.db
    nb_db = len(db) if db is not None and not db.empty else 0
    last_maj = datetime.now().strftime("%d/%m/%Y")
    if db is not None and not db.empty and "Date de MAJ" in db.columns:
        last_vals = db["Date de MAJ"].dropna()
        if not last_vals.empty:
            last_maj = last_vals.iloc[-1]

    st.markdown(f"""
    <div class="db-card">
        <div class="db-card-title">Base contrats</div>
        <div style="display:flex;align-items:baseline;gap:8px">
            <div class="db-card-num">{nb_db}</div>
            <div class="db-card-lbl">contrats enregistrés</div>
        </div>
        <div class="db-card-date">Dernière MAJ · {last_maj}</div>
    </div>
    """, unsafe_allow_html=True)

    if db is not None and not db.empty:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            db.to_excel(w, index=False, sheet_name="Contrats analysés")
        st.download_button(
            "⬇  Télécharger contrats_base.xlsx",
            data=buf.getvalue(),
            file_name="contrats_base.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.markdown("<div style='font-size:11px;color:#3D3060;padding:.5rem 0'>Aucune base chargée</div>", unsafe_allow_html=True)

# ════════════════════════════════════════════
# SECTION 02 — IMPORT & ANALYSE
# ════════════════════════════════════════════
st.markdown('<div class="s-divider" id="import"></div>', unsafe_allow_html=True)
st.markdown('<div class="s-eyebrow">02 — Extraction</div>', unsafe_allow_html=True)
st.markdown('<div class="s-title">Import & Analyse</div>', unsafe_allow_html=True)
st.markdown('<div class="s-sub">ATRADIUS MODULA · COFACE · EULER HERMES · PDF · DOCX · TXT</div>', unsafe_allow_html=True)

# Stats
st.markdown(f"""
<div class="stat-row">
  <div class="stat-card"><div class="sv">16</div><div class="sl">Champs extraits</div></div>
  <div class="stat-card"><div class="sv">37</div><div class="sl">Polices groupe max</div></div>
  <div class="stat-card"><div class="sv w">{nb_session}</div><div class="sl">Session en cours</div></div>
  <div class="stat-card"><div class="sv w">{nb_db}</div><div class="sl">Total en base</div></div>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Déposer les contrats à analyser",
    type=["pdf", "docx", "txt"],
    accept_multiple_files=True,
    key="contracts_upload"
)

if uploaded_files:
    st.markdown(f"<div style='font-size:11px;color:#7C3AED;margin:.4rem 0 .8rem'>→ {len(uploaded_files)} fichier(s) sélectionné(s)</div>", unsafe_allow_html=True)
    ac1, ac2 = st.columns([2, 1])
    with ac1:
        analyze_btn = st.button(f"Analyser {len(uploaded_files)} contrat(s)", type="primary", use_container_width=True)
    with ac2:
        show_text = st.checkbox("Voir texte extrait")

    if analyze_btn:
        pb = st.progress(0)
        stxt = st.empty()
        for i, uf in enumerate(uploaded_files):
            stxt.markdown(f"<div style='font-size:11px;color:#5A4F7A;font-family:Inter,sans-serif'>→ Analyse de {uf.name}…</div>", unsafe_allow_html=True)
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
                tag = "🔄 Mise à jour" if already else "✓ Nouveau"
                st.success(f"{tag} — **{uf.name}** · {found}/{total} champs extraits")
            except Exception as e:
                st.error(f"Erreur — {uf.name} : {e}"); logger.error(e)
        pb.progress(1.0); stxt.empty()
        st.info("↓ Faites défiler jusqu'à Résultats pour valider et enregistrer")

    if show_text and st.session_state.current_text:
        sel = st.selectbox("Fichier", list(st.session_state.current_text.keys()))
        if sel:
            st.text_area("Texte extrait", st.session_state.current_text[sel], height=180)

# ════════════════════════════════════════════
# SECTION 03 — RÉSULTATS
# ════════════════════════════════════════════
if st.session_state.all_results:
    st.markdown('<div class="s-divider" id="resultats"></div>', unsafe_allow_html=True)
    st.markdown('<div class="s-eyebrow">03 — Validation</div>', unsafe_allow_html=True)
    st.markdown('<div class="s-title">Résultats</div>', unsafe_allow_html=True)
    st.markdown('<div class="s-sub">Vérifiez les données extraites · Corrigez si nécessaire · Enregistrez en base</div>', unsafe_allow_html=True)

    for result in st.session_state.all_results:
        already = result.get("already_exists", False)
        merged = result["merged"]
        found, total = count_found_fields(merged)
        pct = found / total if total > 0 else 0

        # Build fields HTML
        fields_html = '<div class="fg">'
        for field, label in FIELD_LABELS.items():
            fd = merged.get(field, {})
            val = result["edited_values"].get(field, fd.get("value", "") if isinstance(fd, dict) else "")
            if isinstance(val, list): val = ""
            ok = "ok" if val else ""
            em = "empty" if not val else ""
            fields_html += f'<div class="fi {ok}"><div class="fl">{label}</div><div class="fv {em}">{val if val else "—"}</div></div>'
        fields_html += '</div>'

        badge_cls = "badge-upd" if already else "badge-new"
        badge_txt = "Mise à jour" if already else "Nouveau"
        card_cls = "r-card updated" if already else "r-card"

        st.markdown(f"""
        <div class="{card_cls}">
            <div class="r-header">
                <div class="r-filename">{result['filename']}</div>
                <span class="badge {badge_cls}">{badge_txt}</span>
            </div>
            {fields_html}
            <div class="prog-row">
                <div class="prog-bg"><div class="prog-fill" style="width:{pct*100:.0f}%"></div></div>
                <div class="prog-txt">{found} / {total} champs</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if already:
            st.warning(f"Le contrat **{result.get('numero_police', '')}** existe déjà en base. Enregistrer écrasera la ligne existante.")

        # Groupe de polices
        gp_d = merged.get("groupe_polices", {})
        polices = gp_d.get("value") if isinstance(gp_d, dict) else None
        if polices and isinstance(polices, list):
            st.markdown(f'<div class="gt-header">Groupe de polices — {len(polices)} membres</div>', unsafe_allow_html=True)
            df_g = pd.DataFrame(polices).rename(columns={"numero": "N° Police", "devise": "Devise", "assure": "Assuré"})
            df_g["Devise"] = df_g["Devise"].fillna("—")
            st.dataframe(df_g, use_container_width=True, hide_index=True)

        # Edit + Save
        with st.expander("✏️  Modifier les valeurs extraites"):
            ecols = st.columns(3)
            for idx, (field, label) in enumerate(FIELD_LABELS.items()):
                with ecols[idx % 3]:
                    fd = merged.get(field, {})
                    cur = result["edited_values"].get(field, fd.get("value", "") if isinstance(fd, dict) else "")
                    if isinstance(cur, list): cur = ""
                    nv = st.text_input(label, value=str(cur) if cur else "",
                        key=f"e_{result['filename']}_{field}")
                    result["edited_values"][field] = nv

        sc1, sc2, _ = st.columns([1, 1, 3])
        with sc1:
            if st.button("💾  Enregistrer en base", key=f"save_{result['filename']}", type="primary", use_container_width=True):
                me = dict(merged)
                for field in FIELD_LABELS:
                    ev = result["edited_values"].get(field, "")
                    orig = merged.get(field, {})
                    me[field] = {"value": ev,
                        "confidence": orig.get("confidence", 0.0) if isinstance(orig, dict) else 0.0,
                        "method": orig.get("method", "manual") if isinstance(orig, dict) else "manual"}
                db_row = merged_to_db_row(result["filename"], me)
                if st.session_state.db is None:
                    st.session_state.db = load_db(DEFAULT_DB_PATH)
                st.session_state.db, action = upsert_row(st.session_state.db, db_row)
                verb = "mis à jour" if action == "updated" else "ajouté"
                st.success(f"✓ Contrat {verb} en base")
                result["already_exists"] = True
        with sc2:
            st.button("Ignorer", key=f"skip_{result['filename']}", type="secondary", use_container_width=True)

        st.markdown("<div style='height:.5rem'></div>", unsafe_allow_html=True)

# ════════════════════════════════════════════
# SECTION 04 — PARCOURIR & MODIFIER
# ════════════════════════════════════════════
st.markdown('<div class="s-divider" id="parcourir"></div>', unsafe_allow_html=True)
st.markdown('<div class="s-eyebrow">04 — Gestion</div>', unsafe_allow_html=True)
st.markdown('<div class="s-title">Parcourir & Modifier</div>', unsafe_allow_html=True)
st.markdown('<div class="s-sub">Filtrez · Éditez directement dans le tableau · Supprimez des lignes</div>', unsafe_allow_html=True)

db = st.session_state.db
if db is None or db.empty:
    st.markdown("<div style='font-size:13px;color:#3D3060;padding:1.5rem 0'>Aucune donnée en base. Chargez une base ou enregistrez des contrats.</div>", unsafe_allow_html=True)
else:
    fc1, fc2, fc3, fc4 = st.columns([1, 1, 1, 1])
    f_police   = fc1.text_input("N° Police", "")
    f_assure   = fc2.text_input("Assuré", "")
    f_assureur = fc3.text_input("Assureur", "")
    with fc4:
        if st.button("🔄  Recharger", type="secondary", use_container_width=True):
            st.session_state.db = load_db(DEFAULT_DB_PATH)
            st.success("Base rechargée.")

    mask = pd.Series([True] * len(db), index=db.index)
    if f_police:   mask &= db["N° Police"].str.contains(f_police, case=False, na=False)
    if f_assure:   mask &= db["Assuré"].str.contains(f_assure, case=False, na=False)
    if f_assureur: mask &= db["Assureur"].str.contains(f_assureur, case=False, na=False)
    db_view = db[mask].copy()

    count_txt = f"{len(db_view)} résultat(s)" if len(db_view) < len(db) else f"{len(db_view)} contrat(s)"
    st.markdown(f"<div style='font-size:10px;color:#5A4F7A;margin:.3rem 0 .8rem;letter-spacing:.06em'>{count_txt}</div>", unsafe_allow_html=True)

    edited_df = st.data_editor(db_view, use_container_width=True, hide_index=False, num_rows="fixed", key="db_editor")

    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
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
            st.success("✓ Modifications sauvegardées en mémoire — téléchargez la base pour conserver")
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

# ════════════════════════════════════════════
# SECTION 05 — EXPORT FINAL
# ════════════════════════════════════════════
st.markdown('<div class="s-divider" id="export"></div>', unsafe_allow_html=True)
st.markdown('<div class="s-eyebrow">05 — Sauvegarde</div>', unsafe_allow_html=True)
st.markdown('<div class="s-title">Télécharger la base</div>', unsafe_allow_html=True)

ex1, ex2 = st.columns([3, 2])
with ex1:
    st.markdown("""
    <div class="workflow-card">
        <div style="font-size:9px;font-weight:600;color:#A78BFA;letter-spacing:.16em;text-transform:uppercase;margin-bottom:12px">Workflow recommandé</div>
        <div class="workflow-step"><div class="step-num">1</div><div class="step-txt"><strong>Charger</strong> la base en section 01</div></div>
        <div class="workflow-step"><div class="step-num">2</div><div class="step-txt"><strong>Analyser</strong> les nouveaux contrats en section 02</div></div>
        <div class="workflow-step"><div class="step-num">3</div><div class="step-txt"><strong>Valider</strong> et enregistrer en section 03</div></div>
        <div class="workflow-step"><div class="step-num">4</div><div class="step-txt"><strong>Télécharger</strong> la base mise à jour ci-contre</div></div>
        <div class="workflow-step"><div class="step-num">5</div><div class="step-txt"><strong>Sauvegarder</strong> sur OneDrive / SharePoint</div></div>
    </div>
    """, unsafe_allow_html=True)

with ex2:
    db_final = st.session_state.db
    if db_final is not None and not db_final.empty:
        buf2 = io.BytesIO()
        with pd.ExcelWriter(buf2, engine="openpyxl") as w:
            db_final.to_excel(w, index=False, sheet_name="Contrats analysés")
        st.markdown("<div style='height:1.5rem'></div>", unsafe_allow_html=True)
        st.download_button(
            "⬇  Télécharger contrats_base.xlsx",
            data=buf2.getvalue(),
            file_name="contrats_base.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        st.markdown(f"<div style='font-size:9px;color:#3D3060;text-align:center;margin-top:.5rem'>{len(db_final)} contrat(s) · {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div style='font-size:12px;color:#3D3060;padding:2rem 0;text-align:center'>Aucune donnée à exporter</div>", unsafe_allow_html=True)

# ── Footer ──
st.markdown("""
<div style='margin-top:4rem;padding-top:1.5rem;
border-top:0.5px solid rgba(255,255,255,0.04);
text-align:center;font-size:9px;color:#1E1830;
letter-spacing:.2em;font-family:Inter,sans-serif'>
CIA · CREDIT INSURANCE ANALYZER · POWERED BY CLAUDE
</div>
""", unsafe_allow_html=True)
