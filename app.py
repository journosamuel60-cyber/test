"""
app.py - Analyseur de Contrats d'Assurance-Crédit
Design: dark fintech, single-page scroll
"""
import os, time, logging, tempfile, shutil, io
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
    "prime_provisionnelle": "Prime provisionnelle", "quotites_garanties": "Pourcentage assuré",
    "delai_indemnisation": "Délai d'indemnisation", "delai_max_credit": "Délai max crédit",
    "date_prise_effet": "Date de prise d'effet", "date_echeance": "Date d'échéance",
    "duree_police": "Durée", "devise": "Devise",
    "limite_decaissement": "Limite de décaissement", "zone_discretionnaire": "Zone discrétionnaire",
}

st.set_page_config(
    page_title="CIA — Credit Insurance Analyzer",
    page_icon="⬡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@300;400;500&display=swap');

/* ── Reset & Base ── */
*, *::before, *::after { box-sizing: border-box; }

html, body, [data-testid="stAppViewContainer"] {
    background: #080C12 !important;
    color: #E8EDF5 !important;
    font-family: 'DM Mono', monospace !important;
}

[data-testid="stAppViewContainer"] {
    background: radial-gradient(ellipse at 20% 0%, #0D1F3C 0%, #080C12 60%) !important;
}

/* Hide default streamlit chrome */
#MainMenu, footer, header, [data-testid="stToolbar"],
[data-testid="stDecoration"], [data-testid="stStatusWidget"] { display: none !important; }

[data-testid="stSidebar"] { display: none !important; }

/* Main container */
.main .block-container {
    max-width: 1200px !important;
    padding: 0 2rem 4rem 2rem !important;
}

/* ── Hero ── */
.hero {
    padding: 4rem 0 3rem 0;
    border-bottom: 1px solid rgba(99,179,237,0.15);
    margin-bottom: 3rem;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: '⬡';
    position: absolute;
    right: -2rem;
    top: -1rem;
    font-size: 18rem;
    color: rgba(99,179,237,0.03);
    line-height: 1;
    pointer-events: none;
}
.hero-tag {
    font-family: 'DM Mono', monospace;
    font-size: 0.7rem;
    font-weight: 500;
    letter-spacing: 0.25em;
    text-transform: uppercase;
    color: #63B3ED;
    margin-bottom: 1rem;
    display: flex;
    align-items: center;
    gap: 0.6rem;
}
.hero-tag::before {
    content: '';
    display: inline-block;
    width: 24px;
    height: 1px;
    background: #63B3ED;
}
.hero h1 {
    font-family: 'Syne', sans-serif !important;
    font-size: clamp(2.2rem, 4vw, 3.5rem) !important;
    font-weight: 800 !important;
    line-height: 1.1 !important;
    margin: 0 0 1rem 0 !important;
    color: #F0F4FF !important;
    letter-spacing: -0.02em;
}
.hero h1 span { color: #63B3ED; }
.hero-sub {
    font-size: 0.85rem;
    color: #6B7A99;
    letter-spacing: 0.05em;
}

/* ── Section titles ── */
.section-title {
    font-family: 'Syne', sans-serif;
    font-size: 1.3rem;
    font-weight: 700;
    color: #F0F4FF;
    margin: 3rem 0 1.5rem 0;
    display: flex;
    align-items: center;
    gap: 0.8rem;
    letter-spacing: -0.01em;
}
.section-title .num {
    font-family: 'DM Mono', monospace;
    font-size: 0.7rem;
    font-weight: 400;
    color: #63B3ED;
    background: rgba(99,179,237,0.1);
    border: 1px solid rgba(99,179,237,0.2);
    padding: 0.15rem 0.5rem;
    border-radius: 4px;
    letter-spacing: 0.1em;
}
.section-divider {
    height: 1px;
    background: linear-gradient(90deg, rgba(99,179,237,0.2) 0%, transparent 100%);
    margin: 0 0 2rem 0;
}

/* ── Cards ── */
.card {
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1rem;
    transition: border-color 0.2s;
}
.card:hover { border-color: rgba(99,179,237,0.2); }
.card-accent {
    border-left: 2px solid #63B3ED;
    background: rgba(99,179,237,0.04);
}

/* ── Config bar ── */
.config-bar {
    background: rgba(255,255,255,0.02);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    margin-bottom: 2rem;
    display: flex;
    align-items: center;
    gap: 2rem;
    flex-wrap: wrap;
}

/* ── Stat badges ── */
.stats-row {
    display: flex;
    gap: 1rem;
    margin-bottom: 2rem;
    flex-wrap: wrap;
}
.stat-badge {
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 8px;
    padding: 0.8rem 1.2rem;
    flex: 1;
    min-width: 140px;
}
.stat-badge .val {
    font-family: 'Syne', sans-serif;
    font-size: 1.8rem;
    font-weight: 700;
    color: #63B3ED;
    line-height: 1;
}
.stat-badge .lbl {
    font-size: 0.65rem;
    color: #6B7A99;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-top: 0.3rem;
}

/* ── Field grid ── */
.field-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 0.75rem;
    margin-bottom: 1.5rem;
}
.field-item {
    background: rgba(255,255,255,0.02);
    border: 1px solid rgba(255,255,255,0.05);
    border-radius: 8px;
    padding: 0.75rem 1rem;
}
.field-item .field-lbl {
    font-size: 0.6rem;
    color: #6B7A99;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin-bottom: 0.3rem;
}
.field-item .field-val {
    font-size: 0.85rem;
    color: #E8EDF5;
    font-weight: 500;
}
.field-item .field-val.empty { color: #3A4560; font-style: italic; }
.field-item.found { border-left: 2px solid rgba(99,179,237,0.4); }

/* ── Streamlit overrides ── */
.stTextInput > label, .stFileUploader > label,
.stToggle > label, .stSelectbox > label,
.stMultiSelect > label, .stTextArea > label {
    font-family: 'DM Mono', monospace !important;
    font-size: 0.7rem !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
    color: #6B7A99 !important;
}

.stTextInput input, .stTextArea textarea {
    background: rgba(255,255,255,0.04) !important;
    border: 1px solid rgba(255,255,255,0.08) !important;
    border-radius: 8px !important;
    color: #E8EDF5 !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 0.85rem !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: rgba(99,179,237,0.4) !important;
    box-shadow: 0 0 0 2px rgba(99,179,237,0.08) !important;
}

.stButton > button {
    background: #63B3ED !important;
    color: #080C12 !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 0.75rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.08em !important;
    padding: 0.6rem 1.4rem !important;
    transition: all 0.15s !important;
}
.stButton > button:hover {
    background: #90CDF4 !important;
    transform: translateY(-1px) !important;
}
.stButton > button[kind="secondary"] {
    background: rgba(255,255,255,0.04) !important;
    color: #E8EDF5 !important;
    border: 1px solid rgba(255,255,255,0.08) !important;
}
.stButton > button[kind="secondary"]:hover {
    background: rgba(255,255,255,0.08) !important;
    border-color: rgba(99,179,237,0.3) !important;
}

.stDownloadButton > button {
    background: rgba(99,179,237,0.1) !important;
    color: #63B3ED !important;
    border: 1px solid rgba(99,179,237,0.25) !important;
    border-radius: 8px !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 0.75rem !important;
    letter-spacing: 0.08em !important;
}
.stDownloadButton > button:hover {
    background: rgba(99,179,237,0.18) !important;
}

/* Progress bar */
.stProgress > div > div {
    background: linear-gradient(90deg, #63B3ED, #90CDF4) !important;
    border-radius: 4px !important;
}
.stProgress > div {
    background: rgba(255,255,255,0.05) !important;
    border-radius: 4px !important;
}

/* Expander */
.streamlit-expanderHeader {
    background: rgba(255,255,255,0.02) !important;
    border: 1px solid rgba(255,255,255,0.06) !important;
    border-radius: 10px !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 0.8rem !important;
    color: #E8EDF5 !important;
}
.streamlit-expanderContent {
    background: rgba(255,255,255,0.01) !important;
    border: 1px solid rgba(255,255,255,0.04) !important;
    border-top: none !important;
    border-radius: 0 0 10px 10px !important;
}

/* Alerts */
.stSuccess { background: rgba(72,187,120,0.08) !important; border: 1px solid rgba(72,187,120,0.2) !important; border-radius: 8px !important; }
.stWarning { background: rgba(237,137,54,0.08) !important; border: 1px solid rgba(237,137,54,0.2) !important; border-radius: 8px !important; }
.stError   { background: rgba(245,101,101,0.08) !important; border: 1px solid rgba(245,101,101,0.2) !important; border-radius: 8px !important; }
.stInfo    { background: rgba(99,179,237,0.06) !important; border: 1px solid rgba(99,179,237,0.15) !important; border-radius: 8px !important; }

/* Dataframe */
[data-testid="stDataFrame"] {
    border: 1px solid rgba(255,255,255,0.06) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: rgba(255,255,255,0.02) !important;
    border: 1px dashed rgba(99,179,237,0.2) !important;
    border-radius: 12px !important;
    padding: 1rem !important;
}

/* Toggle */
.stToggle { color: #E8EDF5 !important; }

/* Multiselect */
[data-testid="stMultiSelect"] > div {
    background: rgba(255,255,255,0.04) !important;
    border-color: rgba(255,255,255,0.08) !important;
    border-radius: 8px !important;
}

/* Selectbox */
[data-testid="stSelectbox"] > div {
    background: rgba(255,255,255,0.04) !important;
    border-color: rgba(255,255,255,0.08) !important;
}

/* Scroll divider */
.scroll-divider {
    height: 60px;
    display: flex;
    align-items: center;
    gap: 1rem;
    color: rgba(99,179,237,0.2);
    font-size: 0.65rem;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    margin: 2rem 0;
}
.scroll-divider::before, .scroll-divider::after {
    content: '';
    flex: 1;
    height: 1px;
    background: rgba(255,255,255,0.04);
}

/* Confidence dots */
.conf-high { color: #68D391; font-size: 0.65rem; }
.conf-med  { color: #F6AD55; font-size: 0.65rem; }
.conf-low  { color: #FC8181; font-size: 0.65rem; }

/* Groupe table header */
.groupe-header {
    font-family: 'DM Mono', monospace;
    font-size: 0.65rem;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: #63B3ED;
    margin: 1.5rem 0 0.75rem 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}
.groupe-header::after {
    content: '';
    flex: 1;
    height: 1px;
    background: rgba(99,179,237,0.15);
}
</style>
""", unsafe_allow_html=True)

# ── Session state ──
for k, v in [("all_results", []), ("current_text", {}), ("db", None), ("db_path", DEFAULT_DB_PATH)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ════════════════════════════════════════════
# HERO
# ════════════════════════════════════════════
st.markdown("""
<div class="hero">
    <div class="hero-tag">Credit Insurance Analyzer</div>
    <h1>Extraction<br>automatique<br><span>des données contractuelles</span></h1>
    <div class="hero-sub">Atradius Modula · Coface · Euler Hermes · FR / EN · PDF · DOCX · TXT</div>
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════
# SECTION 0 — CONFIG (inline, no sidebar)
# ════════════════════════════════════════════
st.markdown('<div class="section-title"><span class="num">CONFIG</span> Paramètres</div><div class="section-divider"></div>', unsafe_allow_html=True)

with st.container():
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        use_llm = st.toggle("Analyse IA (Claude)", value=True)
    with c2:
        api_key = None
        if use_llm:
            api_key_input = st.text_input("Clé API Anthropic", type="password",
                placeholder="sk-ant-... (ou variable ANTHROPIC_API_KEY)")
            api_key = api_key_input or os.environ.get("ANTHROPIC_API_KEY")

# ════════════════════════════════════════════
# SECTION 1 — BASE DE DONNÉES
# ════════════════════════════════════════════
st.markdown('<div class="scroll-divider">base de données</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title"><span class="num">01</span> Base de données</div><div class="section-divider"></div>', unsafe_allow_html=True)

db_col1, db_col2 = st.columns([3, 2])

with db_col1:
    st.markdown("**Charger une base existante**")
    uploaded_db = st.file_uploader("Importer contrats_base.xlsx", type=["xlsx"], key="db_upload",
        help="Charge ta base depuis ton PC / OneDrive au démarrage")
    if uploaded_db:
        st.session_state.db = load_db(uploaded_db)
        st.success(f"✓ Base chargée — {len(st.session_state.db)} contrat(s)")

    if st.session_state.db is None:
        st.session_state.db = load_db(st.session_state.db_path)

with db_col2:
    db = st.session_state.db
    nb = len(db) if db is not None and not db.empty else 0

    st.markdown(f"""
    <div class="stats-row">
        <div class="stat-badge">
            <div class="val">{nb}</div>
            <div class="lbl">Contrats en base</div>
        </div>
        <div class="stat-badge">
            <div class="val">{len(st.session_state.all_results)}</div>
            <div class="lbl">Analysés (session)</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if db is not None and not db.empty:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            db.to_excel(w, index=False, sheet_name="Contrats analysés")
        st.download_button(
            "⬇ Télécharger la base actuelle",
            data=buf.getvalue(),
            file_name="contrats_base.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ════════════════════════════════════════════
# SECTION 2 — IMPORT & ANALYSE
# ════════════════════════════════════════════
st.markdown('<div class="scroll-divider">import & analyse</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title"><span class="num">02</span> Import & Analyse</div><div class="section-divider"></div>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Déposer les contrats à analyser",
    type=["pdf", "docx", "txt"],
    accept_multiple_files=True,
    key="contracts_upload"
)

if uploaded_files:
    st.markdown(f"<div style='font-size:.75rem;color:#63B3ED;margin:.5rem 0;'>→ {len(uploaded_files)} fichier(s) prêt(s)</div>", unsafe_allow_html=True)
    col_btn1, col_btn2 = st.columns([2, 1])
    with col_btn1:
        analyze_btn = st.button(f"Analyser {len(uploaded_files)} contrat(s)", type="primary", use_container_width=True)
    with col_btn2:
        show_text = st.checkbox("Voir texte extrait")

    if analyze_btn:
        pb = st.progress(0)
        st_txt = st.empty()
        for i, uf in enumerate(uploaded_files):
            st_txt.markdown(f"<div style='font-size:.75rem;color:#6B7A99;font-family:DM Mono,monospace'>→ {uf.name}...</div>", unsafe_allow_html=True)
            pb.progress(i / len(uploaded_files))
            try:
                tmp = save_uploaded_file(uf, TMP_DIR)
                raw, _ = extract_text(tmp)
                clean = clean_text(raw)
                st.session_state.current_text[uf.name] = clean
                if not clean.strip():
                    st.warning(f"Aucun texte extrait — {uf.name}")
                    continue
                regex_r = extract_all_fields(clean)
                llm_r = {}
                if use_llm and (api_key or os.environ.get("ANTHROPIC_API_KEY")):
                    with st.spinner(f"Analyse IA..."):
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
                existing = next((r for r in st.session_state.all_results if r["filename"] == uf.name), None)
                if existing: existing.update(entry)
                else: st.session_state.all_results.append(entry)
                tag = "🔄 MAJ" if already else "✓ NOUVEAU"
                st.success(f"{tag} — {uf.name}  ·  {found}/{total} champs extraits")
            except Exception as e:
                st.error(f"Erreur — {uf.name} : {e}")
                logger.error(e)
        pb.progress(1.0)
        st_txt.empty()
        st.info("→ Faites défiler jusqu'à Résultats pour valider et enregistrer")

    if show_text and st.session_state.current_text:
        sel = st.selectbox("Fichier", list(st.session_state.current_text.keys()))
        if sel:
            st.text_area("Texte extrait", st.session_state.current_text[sel], height=200)

# ════════════════════════════════════════════
# SECTION 3 — RÉSULTATS
# ════════════════════════════════════════════
if st.session_state.all_results:
    st.markdown('<div class="scroll-divider">résultats</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title"><span class="num">03</span> Résultats & Validation</div><div class="section-divider"></div>', unsafe_allow_html=True)

    for result in st.session_state.all_results:
        already = result.get("already_exists", False)
        merged = result["merged"]
        found, total = count_found_fields(merged)
        pct = int(found / total * 100) if total > 0 else 0
        badge_txt = "MISE À JOUR" if already else "NOUVEAU"

        with st.expander(f"{'🔄' if already else '◆'} {result['filename']}  ·  {found}/{total} champs  ·  {badge_txt}", expanded=True):

            # Progress
            st.progress(found / total if total > 0 else 0)

            if already:
                st.warning(f"Ce contrat (N° {result.get('numero_police','')}) existe déjà en base. Enregistrer écrasera la ligne existante.")

            # Fields grid — display only (read mode)
            fields_html = '<div class="field-grid">'
            for field, label in FIELD_LABELS.items():
                fd = merged.get(field, {})
                val = result["edited_values"].get(field, fd.get("value", "") if isinstance(fd, dict) else "")
                if isinstance(val, list): val = ""
                conf = fd.get("confidence", 0.0) if isinstance(fd, dict) else 0.0
                found_class = "found" if val else ""
                empty_class = "empty" if not val else ""
                display_val = val if val else "—"
                fields_html += f'''<div class="field-item {found_class}">
                    <div class="field-lbl">{label}</div>
                    <div class="field-val {empty_class}">{display_val}</div>
                </div>'''
            fields_html += '</div>'
            st.markdown(fields_html, unsafe_allow_html=True)

            # Edit mode
            with st.expander("✏️ Modifier les valeurs", expanded=False):
                cols = st.columns(2)
                for idx, (field, label) in enumerate(FIELD_LABELS.items()):
                    with cols[idx % 2]:
                        fd = merged.get(field, {})
                        cur = result["edited_values"].get(field, fd.get("value", "") if isinstance(fd, dict) else "")
                        if isinstance(cur, list): cur = ""
                        conf = fd.get("confidence", 0.0) if isinstance(fd, dict) else 0.0
                        new_val = st.text_input(label, value=str(cur) if cur else "",
                            key=f"edit_{result['filename']}_{field}")
                        result["edited_values"][field] = new_val

            # Groupe de polices
            gp_d = merged.get("groupe_polices", {})
            polices = gp_d.get("value") if isinstance(gp_d, dict) else None
            if polices and isinstance(polices, list):
                st.markdown(f'<div class="groupe-header">Groupe de polices — {len(polices)} membres</div>', unsafe_allow_html=True)
                df_g = pd.DataFrame(polices).rename(columns={"numero": "N° Police", "devise": "Devise", "assure": "Assuré"})
                df_g["Devise"] = df_g["Devise"].fillna("—")
                st.dataframe(df_g, use_container_width=True, hide_index=True)

            # Save button
            st.markdown("<div style='height:.5rem'></div>", unsafe_allow_html=True)
            save_col1, _ = st.columns([1, 3])
            with save_col1:
                if st.button("💾 Enregistrer en base", key=f"save_{result['filename']}", type="primary", use_container_width=True):
                    merged_edited = dict(merged)
                    for field in FIELD_LABELS:
                        ev = result["edited_values"].get(field, "")
                        orig = merged.get(field, {})
                        merged_edited[field] = {
                            "value": ev,
                            "confidence": orig.get("confidence", 0.0) if isinstance(orig, dict) else 0.0,
                            "method": orig.get("method", "manual") if isinstance(orig, dict) else "manual"
                        }
                    db_row = merged_to_db_row(result["filename"], merged_edited)
                    if st.session_state.db is None:
                        st.session_state.db = load_db(st.session_state.db_path)
                    st.session_state.db, action = upsert_row(st.session_state.db, db_row)
                    verb = "mis à jour" if action == "updated" else "ajouté"
                    st.success(f"✓ Contrat {verb} en base")
                    result["already_exists"] = True

# ════════════════════════════════════════════
# SECTION 4 — BASE / PARCOURIR
# ════════════════════════════════════════════
st.markdown('<div class="scroll-divider">base de données</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title"><span class="num">04</span> Parcourir & Modifier la base</div><div class="section-divider"></div>', unsafe_allow_html=True)

db = st.session_state.db
if db is None or db.empty:
    st.markdown("<div style='color:#3A4560;font-size:.85rem;padding:2rem 0'>Aucune donnée en base. Chargez une base existante ou analysez des contrats.</div>", unsafe_allow_html=True)
else:
    # Stats
    st.markdown(f"<div style='font-size:.7rem;color:#6B7A99;margin-bottom:1.5rem'>{len(db)} contrat(s) · dernière MAJ {datetime.now().strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)

    # Filtres
    fc1, fc2, fc3 = st.columns(3)
    f_police  = fc1.text_input("Filtrer N° Police", "")
    f_assure  = fc2.text_input("Filtrer Assuré", "")
    f_assureur = fc3.text_input("Filtrer Assureur", "")

    mask = pd.Series([True] * len(db), index=db.index)
    if f_police:   mask &= db["N° Police"].str.contains(f_police, case=False, na=False)
    if f_assure:   mask &= db["Assuré"].str.contains(f_assure, case=False, na=False)
    if f_assureur: mask &= db["Assureur"].str.contains(f_assureur, case=False, na=False)
    db_view = db[mask].copy()

    if len(db_view) < len(db):
        st.markdown(f"<div style='font-size:.7rem;color:#63B3ED;margin:.5rem 0'>{len(db_view)} résultat(s) sur {len(db)}</div>", unsafe_allow_html=True)

    # Tableau éditable
    edited_df = st.data_editor(db_view, use_container_width=True, hide_index=False,
        num_rows="fixed", key="db_editor")

    # Actions
    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
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
            st.success("✓ Modifications sauvegardées en mémoire — pensez à télécharger la base")

    with a2:
        to_delete = st.multiselect("Supprimer des polices",
            options=db_view["N° Police"].tolist(), placeholder="Sélectionner N° Police(s)…")

    with a3:
        if to_delete:
            if st.button(f"🗑 Supprimer {len(to_delete)} ligne(s)", type="primary", use_container_width=True):
                indices = db[db["N° Police"].isin(to_delete)].index.tolist()
                st.session_state.db = delete_rows(st.session_state.db, indices)
                st.success(f"✓ {len(indices)} ligne(s) supprimée(s)")
                st.rerun()

# ════════════════════════════════════════════
# SECTION 5 — TÉLÉCHARGEMENT FINAL
# ════════════════════════════════════════════
st.markdown('<div class="scroll-divider">export</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title"><span class="num">05</span> Télécharger la base</div><div class="section-divider"></div>', unsafe_allow_html=True)

dl_col1, dl_col2 = st.columns([2, 1])

with dl_col1:
    st.markdown("""
    <div class="card card-accent">
        <div style="font-size:.7rem;color:#6B7A99;letter-spacing:.1em;text-transform:uppercase;margin-bottom:.5rem">Workflow recommandé</div>
        <div style="font-size:.82rem;color:#A0AEC0;line-height:1.8">
        1 · Charger la base en début de session (section 01)<br>
        2 · Analyser les nouveaux contrats (section 02)<br>
        3 · Valider et enregistrer en base (section 03)<br>
        4 · Télécharger la base mise à jour ci-dessous<br>
        5 · Sauvegarder sur OneDrive / SharePoint
        </div>
    </div>
    """, unsafe_allow_html=True)

with dl_col2:
    db_final = st.session_state.db
    if db_final is not None and not db_final.empty:
        buf = io.BytesIO()
        save_db(db_final, buf if hasattr(buf, 'write') else "/tmp/export_final.xlsx")
        # Use simple pandas export for download
        buf2 = io.BytesIO()
        with pd.ExcelWriter(buf2, engine="openpyxl") as w:
            db_final.to_excel(w, index=False, sheet_name="Contrats analysés")
        st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
        st.download_button(
            "⬇ Télécharger contrats_base.xlsx",
            data=buf2.getvalue(),
            file_name="contrats_base.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        st.markdown(f"<div style='font-size:.65rem;color:#6B7A99;text-align:center;margin-top:.5rem'>{len(db_final)} contrat(s) · {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div style='font-size:.8rem;color:#3A4560;padding:2rem 0;text-align:center'>Aucune donnée à exporter</div>", unsafe_allow_html=True)

# Footer
st.markdown("""
<div style='margin-top:4rem;padding-top:2rem;border-top:1px solid rgba(255,255,255,0.04);
text-align:center;font-size:.65rem;color:#2D3748;letter-spacing:.15em'>
CIA · CREDIT INSURANCE ANALYZER · POWERED BY CLAUDE
</div>
""", unsafe_allow_html=True)
