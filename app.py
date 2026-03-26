import streamlit as st
import pandas as pd
import io
import re
import unicodedata
from collections import defaultdict
from datetime import datetime
import hashlib
import base64
import json
import os
import tempfile
import sys
import platform
from PIL import Image

# ========== CONFIGURACIÓN DE PÁGINA ==========
st.set_page_config(
    page_title="Procesador de Clientes | AR Collect",
    page_icon="https://raw.githubusercontent.com/Iamnotmanolotaco/Finance-Data-structure-and-reporting/main/assets/favicon4.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========== CONFIGURACIÓN DE IMÁGENES FIJAS ==========
BANNER_URL = "https://raw.githubusercontent.com/Iamnotmanolotaco/Finance-Data-structure-and-reporting/main/assets/banner7.png"
LOGO_URL = "https://raw.githubusercontent.com/Iamnotmanolotaco/Finance-Data-structure-and-reporting/main/assets/image5.png"

# ========== CONTRASEÑA PARA EDITOR ==========
EDITOR_PASSWORD = "manolotaco123"
PASSWORD_HASH = hashlib.sha256(EDITOR_PASSWORD.encode()).hexdigest()

# ========== INICIALIZAR VARIABLES DE SESIÓN ==========
if 'modo_editor' not in st.session_state:
    st.session_state.modo_editor = False
if 'password_correcta' not in st.session_state:
    st.session_state.password_correcta = False

# ========== COLORES PRINCIPALES ==========
if 'color_principal' not in st.session_state:
    st.session_state.color_principal = "#f60d2d"
if 'bordes' not in st.session_state:
    st.session_state.bordes = 12
if 'sombra_tarjetas' not in st.session_state:
    st.session_state.sombra_tarjetas = "0 2px 8px rgba(0,0,0,0.08)"
if 'fuente_principal' not in st.session_state:
    st.session_state.fuente_principal = "'Inter', 'Segoe UI', sans-serif"

# ========== FUNCIONES PARA GUARDAR CONFIGURACIÓN ==========
CONFIG_FILE = "app_config.json"

def guardar_configuracion():
    config = {
        'color_principal': st.session_state.color_principal,
        'bordes': st.session_state.bordes
    }
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f)
    except:
        pass

def cargar_configuracion():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                if 'color_principal' in config:
                    st.session_state.color_principal = config['color_principal']
                if 'bordes' in config:
                    st.session_state.bordes = config['bordes']
    except:
        pass

cargar_configuracion()

def mostrar_banner():
    try:
        st.image(BANNER_URL, use_container_width=True)
    except:
        st.markdown(f"""
        <div style="
            background: linear-gradient(135deg, {st.session_state.color_principal} 0%, #b00a22 100%);
            padding: 1.5rem;
            border-radius: {st.session_state.bordes}px;
            margin-bottom: 1rem;
            text-align: center;
        ">
            <h1 style="color: white; margin: 0;">⚖️ Procesador de Clientes</h1>
            <p style="color: rgba(255,255,255,0.9);">AR Collect - Análisis Automático</p>
        </div>
        """, unsafe_allow_html=True)

def mostrar_logo(tamaño=80):
    try:
        st.image(LOGO_URL, width=tamaño)
    except:
        st.markdown(f"<h1 style='font-size: {tamaño//4}px; color: #01bba7;'>⚖️</h1>", unsafe_allow_html=True)

def verificar_password(password):
    return hashlib.sha256(password.encode()).hexdigest() == PASSWORD_HASH

# ========== FUNCIÓN PARA LEER EXCEL CON HASH ==========
def calcular_hash_archivo(file):
    file.seek(0)
    hash_md5 = hashlib.md5()
    for chunk in iter(lambda: file.read(4096), b""):
        hash_md5.update(chunk)
    file.seek(0)
    return hash_md5.hexdigest()

def leer_excel_seguro(file, header=0, nombre="archivo"):
    try:
        hash_archivo = calcular_hash_archivo(file)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(file.read())
            tmp_path = tmp.name
        try:
            df = pd.read_excel(tmp_path, header=header)
        except:
            try:
                df = pd.read_excel(tmp_path, header=header, engine='openpyxl')
            except:
                df = pd.DataFrame()
        try:
            os.unlink(tmp_path)
        except:
            pass
        return df, hash_archivo
    except Exception as e:
        st.error(f"Error al leer {nombre}: {str(e)[:100]}")
        return pd.DataFrame(), None

# ========== FUNCIONES DE NORMALIZACIÓN ==========
SPACE_CHARS = {
    '\u00A0', '\u2000', '\u2001', '\u2002', '\u2003', '\u2004', '\u2005',
    '\u2006', '\u2007', '\u2008', '\u2009', '\u200A', '\u202F', '\u205F', '\u3000'
}

def normalize_spaces(s: str) -> str:
    if not isinstance(s, str):
        return ""
    for ch in SPACE_CHARS:
        s = s.replace(ch, ' ')
    s = re.sub(r'\s+', ' ', s, flags=re.UNICODE).strip()
    return s

def build_case_pattern(keywords):
    alts = []
    for kw in keywords:
        kw = kw.strip()
        if not kw:
            continue
        alt = r'\b' + re.sub(r'\s+', r'\\s+', re.escape(kw)) + r'\b'
        alts.append(alt)
    alts = sorted(alts, key=len, reverse=True)
    if not alts:
        return None
    return re.compile("|".join(alts), flags=re.IGNORECASE | re.UNICODE)

case_keywords = [
    'removal','guardianship','visa','deportation','bond','asylum','affirmative','divorce',
    'u-visa','immigrant','complete','waiver','individual','defense','special','adjustment',
    'u visa','custody','appeal','advanced','foia','immediate','vawa','investigation','work',
    'renewal','ead','lpr','irreconcilable','uscis','bia','601 a waiver','request','daca renewal',
    'consular','nacara applications','representation','job 1','sij application','name change',
    'iaos','citizenship','family petition','lpr replacement','Affirmative','u', 'power','Attorney', 
    'family', 'petition', 'i', 'i-', 'stay', 'motion', 'job', 'uncontested', 'sij', 'pre', 
    'decree', 'applications', 'id', 'mortion', 'def', 'def.', 'job'
]

CASE_PATTERN = build_case_pattern(case_keywords)

def strip_accents(txt: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', txt)
                   if unicodedata.category(c) != 'Mn')

def clean_name(raw) -> str:
    if not isinstance(raw, str):
        return ""
    s = raw.lower()
    s = strip_accents(s)
    s = normalize_spaces(s)
    s = re.sub(r'[^\w\s]', ' ', s, flags=re.UNICODE)
    s = normalize_spaces(s)
    if CASE_PATTERN:
        m = CASE_PATTERN.search(s)
        if m:
            s = s[:m.start()]
            s = normalize_spaces(s)
    return s

def normalize_name(raw) -> str:
    s = clean_name(raw)
    tokens = sorted(s.split())
    return " ".join(tokens)

def token_sets(a: str, b: str):
    s1 = set(a.split())
    s2 = set(b.split())
    return s1, s2, len(s1 & s2)

def classify_match(a: str, b: str, allow_soft: bool = True):
    if not a or not b:
        return (False, "no", 0)
    if a == b:
        return (True, "exact", len(set(a.split())))
    s1, s2, inter = token_sets(a, b)
    n1, n2 = len(s1), len(s2)
    nmin, nmax = min(n1, n2), max(n1, n2)
    if n1 >= 4 or n2 >= 4:
        if inter >= 3:
            return (True, "3+ tokens", inter)
        elif inter == 2:
            return (True, "2/4+ tokens (soft)", inter)
        else:
            return (False, "no", inter)
    if nmin == 3 and nmax >= 3:
        if inter == 3:
            return (True, "3/3 tokens", inter)
        if inter == 2 and allow_soft:
            return (True, "2/3 tokens (soft)", inter)
        return (False, "no", inter)
    if nmin == 2:
        return (inter == 2, "2/2 tokens" if inter == 2 else "no", inter)
    if nmin == 1:
        return (a == b, "1 token exact" if a == b else "no", inter)
    return (False, "no", inter)

def process_data_with_files(AR_file, cl_file, cc_file, allow_soft=True):
    AR, hash_ar = leer_excel_seguro(AR_file, header=0, nombre="ARCollect")
    cl_file_df, hash_cl = leer_excel_seguro(cl_file, header=2, nombre="Case Details")
    cc_data, hash_cc = leer_excel_seguro(cc_file, header=0, nombre="Casos Cerrados")
    
    st.session_state.hash_ar = hash_ar
    st.session_state.hash_cl = hash_cl
    st.session_state.hash_cc = hash_cc
    
    if AR.empty:
        st.error("No se pudo cargar el archivo ARCollect")
        return [], [], []
    
    if 'Customer' not in AR.columns:
        for col in AR.columns:
            if 'customer' in str(col).lower():
                AR.rename(columns={col: 'Customer'}, inplace=True)
                break
    
    if 'Customer' not in AR.columns:
        st.error(f"No se encontró columna 'Customer'. Columnas: {list(AR.columns)[:5]}")
        return [], [], []
    
    AR["normalized_name"] = AR["Customer"].astype(str).apply(normalize_name)
    
    if not cl_file_df.empty and 'Petitioner Name' in cl_file_df.columns:
        cl_file_df["normalized_name"] = cl_file_df["Petitioner Name"].astype(str).apply(normalize_name)
        cl_norms_unique = cl_file_df["normalized_name"].dropna().unique().tolist()
        cl_index = defaultdict(list)
        for i, r in cl_file_df.iterrows():
            cl_index[r["normalized_name"]].append(i)
    else:
        cl_norms_unique = []
        cl_index = defaultdict(list)
    
    if not cc_data.empty:
        cc_data["normalized_name"] = cc_data.iloc[:, 0].astype(str).apply(normalize_name)
        cc_norms = cc_data["normalized_name"].dropna().unique().tolist()
    else:
        cc_norms = []
    
    aging_cols = ["1 - 30 days", "31 - 60 days", "61 - 90 days", "91 - 120 days", "121+ days"]
    for col in aging_cols:
        if col in AR.columns:
            AR[col] = pd.to_numeric(AR[col], errors="coerce").fillna(0)
        else:
            AR[col] = 0
    
    AR["Total_Balance_Calculated"] = AR[aging_cols].sum(axis=1)
    AR_pos_balance = AR[AR["Total_Balance_Calculated"] > 0].copy()
    AR_zero_balance = AR[AR["Total_Balance_Calculated"] == 0].copy()
    
    filtrados_rows = []
    descartados_rows = []
    log_rows = []
    
    for _, row in AR_zero_balance.iterrows():
        row_out = row.to_dict()
        row_out["Estado_final"] = "Balance = 0"
        row_out["Motivo_descartado"] = "Balance calculado = 0"
        row_out["Case_Status"] = ""
        row_out["Case_Number"] = ""
        row_out["Best_Match_cl_name"] = row["Customer"]
        descartados_rows.append(row_out)
    
    for _, ar_row in AR_pos_balance.iterrows():
        cliente = ar_row["Customer"]
        norm_cliente = ar_row["normalized_name"]
        best_match_original = ""
        case_statuses = ""
        case_numbers = ""
        estado = "Sin match"
        accion = "Mantener"
        best_label = ""
        best_inter = 0
        best_match_norm = ""
        
        if norm_cliente and cl_norms_unique:
            matched_normals = []
            match_types = {}
            match_inters = {}
            for cand in cl_norms_unique:
                ok, label, inter = classify_match(norm_cliente, cand, allow_soft)
                if ok:
                    matched_normals.append(cand)
                    match_types[cand] = label
                    match_inters[cand] = inter
            if matched_normals:
                best_match_norm = max(matched_normals, key=lambda c: (match_inters.get(c, 0), len(c.split())))
                best_label = match_types.get(best_match_norm, "")
                best_inter = match_inters.get(best_match_norm, 0)
                matched_indices = cl_index.get(best_match_norm, [])
                if matched_indices:
                    best_df = cl_file_df.loc[matched_indices].copy()
                    case_statuses = "; ".join(sorted(set(best_df["Case Status"].astype(str))))
                    case_numbers = "; ".join(sorted(set(best_df["Case Number"].astype(str))))
                    best_originals = best_df["Petitioner Name"].astype(str).unique().tolist()
                    best_match_original = best_originals[0] if best_originals else ""
                    statuses_upper = []
                    for cs in best_df["Case Status"]:
                        for s in str(cs).split(";"):
                            statuses_upper.append(s.strip().upper())
                    discard_statuses = {"CLOSED", "DELETED", "WITHDRAWING", "WITHDRAWN", "READY_FOR_CLOSING"}
                    all_discardable = all(s in discard_statuses for s in statuses_upper)
                    if all_discardable:
                        if cc_norms:
                            cc_match = any(classify_match(norm_cliente, cc, allow_soft)[0] for cc in cc_norms)
                            if cc_match:
                                estado = "Cerrado confirmado"
                                accion = "Descartar"
                            else:
                                estado = "En proceso de cierre"
                                accion = "Mantener"
                        else:
                            estado = "Sin info CC"
                            accion = "Mantener"
                    else:
                        estado = "Caso abierto"
                        accion = "Mantener"
        
        row_out = ar_row.to_dict()
        
        if best_match_original == "":
            row_out["Best_Match_cl_name"] = cliente
        else:
            row_out["Best_Match_cl_name"] = best_match_original
        
        row_out["Estado_final"] = estado
        row_out["Case_Status"] = case_statuses
        row_out["Case_Number"] = case_numbers
        
        if accion == "Mantener":
            row_out["En_proceso_de_cierre"] = (estado == "En proceso de cierre")
            filtrados_rows.append(row_out)
        else:
            row_out["Motivo_descartado"] = "Cerrado confirmado"
            descartados_rows.append(row_out)
        
        log_rows.append({
            "Cliente_AR": cliente,
            "Nombre_Normalizado_AR": norm_cliente,
            "Best_Match_cl_name": best_match_original if best_match_original else cliente,
            "Best_Match_cl_norm": best_match_norm if best_match_norm else norm_cliente,
            "Best_Match_Type": best_label,
            "Best_Match_OverlapTokens": best_inter,
            "Case_Status": case_statuses,
            "Case_Number": case_numbers,
            "Estado_final": estado,
            "Accion": accion,
            "Revisar": best_label == "2/3 tokens (soft)" or best_label == "2/4+ tokens (soft)"
        })
    
    st.session_state.total_filtrados = len(filtrados_rows)
    st.session_state.total_descatados = len(descartados_rows)
    
    return filtrados_rows, descartados_rows, log_rows

# ========== CSS PERSONALIZADO COMPLETO ==========
st.markdown(f"""
<style>
    /* Importar fuente */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    /* Estilos globales */
    .stApp {{
        font-family: {st.session_state.fuente_principal};
    }}
    
    /* ========== PÁGINA PRINCIPAL ========== */
    .stApp {{
        background-color: #feffff !important;
    }}
    
    /* ========== BARRA LATERAL - COLORES FIJOS ========== */
    
    /* Fondo de la barra lateral */
    [data-testid="stSidebar"] {{
        background-color: #393939 !important;
    }}
    
    /* ========== TÍTULOS EN BARRA LATERAL ========== */
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] h4,
    [data-testid="stSidebar"] .stMarkdown h1,
    [data-testid="stSidebar"] .stMarkdown h2,
    [data-testid="stSidebar"] .stMarkdown h3,
    [data-testid="stSidebar"] .stMarkdown h4 {{
        color: #01bba7 !important;
        font-weight: 700 !important;
    }}
    
    /* ========== TODO EL TEXTO NORMAL EN BARRA LATERAL ========== */
    /* Forzar que todo el texto que no sea título sea #dddcdc */
    [data-testid="stSidebar"] *:not(h1):not(h2):not(h3):not(h4) {{
        color: #dddcdc !important;
    }}
    
    /* ========== FILE UPLOADER EN BARRA LATERAL ========== */
    /* Área de drag and drop */
    [data-testid="stSidebar"] .stFileUploader {{
        background-color: #393939 !important;
        border-radius: {st.session_state.bordes}px !important;
        padding: 0.5rem !important;
        border: 1px solid #5a5a5a !important;
    }}
    
    /* Texto del drag and drop */
    [data-testid="stSidebar"] .stFileUploader div[data-testid="stMarkdownContainer"] p {{
        color: #dddcdc !important;
        font-weight: 500 !important;
    }}
    
    /* Labels de file_uploader */
    [data-testid="stSidebar"] .stFileUploader label {{
        color: #dddcdc !important;
        font-weight: 500;
        font-size: 0.9rem;
    }}
    
    /* Texto de ayuda/información */
    [data-testid="stSidebar"] .stFileUploader p {{
        color: #dddcdc !important;
        font-size: 0.8rem;
        opacity: 0.9;
    }}
    
    /* Botón de "Browse files" */
    [data-testid="stSidebar"] .stFileUploader button {{
        border-radius: {st.session_state.bordes}px !important;
        background-color: #4a4a4a !important;
        color: #dddcdc !important;
        border: 1px solid #5a5a5a !important;
    }}
    
    [data-testid="stSidebar"] .stFileUploader button:hover {{
        background-color: #5a5a5a !important;
        border-color: {st.session_state.color_principal} !important;
    }}
    
    /* Nombre del archivo ya cargado */
    [data-testid="stSidebar"] .stFileUploader div[data-testid="stMarkdownContainer"] {{
        background-color: #393939 !important;
        border-radius: {st.session_state.bordes}px !important;
        padding: 0.5rem !important;
        margin-top: 0.5rem !important;
        border: 1px solid #5a5a5a !important;
    }}
    
    [data-testid="stSidebar"] .stFileUploader div[data-testid="stMarkdownContainer"] p {{
        color: #dddcdc !important;
        font-weight: 500 !important;
    }}
    
    /* Alertas en barra lateral */
    [data-testid="stSidebar"] .stAlert {{
        background-color: #4a4a4a !important;
        color: #dddcdc !important;
        border-left-color: {st.session_state.color_principal} !important;
    }}
    
    /* Expandors en barra lateral */
    [data-testid="stSidebar"] .streamlit-expanderHeader {{
        color: #dddcdc !important;
        background-color: #4a4a4a !important;
        font-weight: 500 !important;
        border-radius: {st.session_state.bordes}px !important;
    }}
    
    [data-testid="stSidebar"] .streamlit-expanderHeader:hover {{
        color: #01bba7 !important;
    }}
    
    [data-testid="stSidebar"] .streamlit-expanderContent {{
        color: #dddcdc !important;
    }}
    
    /* Checkbox en barra lateral */
    [data-testid="stSidebar"] .stCheckbox label {{
        color: #dddcdc !important;
    }}
    
    /* Separadores */
    [data-testid="stSidebar"] hr {{
        border-color: #5a5a5a !important;
    }}
    
    /* ========== TARJETAS EN ÁREA PRINCIPAL ========== */
    .metric-card,
    .file-card {{
        background-color: #dddcdc !important;
        border-radius: {st.session_state.bordes}px;
        padding: 1.2rem;
        text-align: center;
        box-shadow: {st.session_state.sombra_tarjetas};
        transition: all 0.3s ease;
        border: none !important;
    }}
    
    .file-card-success {{
        border-left: none !important;
        background-color: #dddcdc !important;
    }}
    
    .file-card-pending {{
        border-left: none !important;
        background-color: #fafafa !important;
    }}
    
    .metric-card {{
        border-top: none !important;
    }}
    
    .metric-value {{
        font-size: 2.2rem;
        font-weight: 700;
        color: #1a1a1a;
    }}
    
    .metric-label {{
        font-size: 0.85rem;
        color: #555555;
        margin-top: 0.5rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    
    .file-icon {{
        font-size: 2rem;
    }}
    
    .file-title {{
        font-weight: 600;
        color: #1a1a1a;
    }}
    
    .file-status {{
        font-size: 0.8rem;
        color: #555555;
        margin-top: 0.25rem;
    }}
    
    /* ========== TÍTULOS PRINCIPALES ========== */
    h1 {{
        color: #1a1a1a;
        font-size: 2.5rem;
        font-weight: 700;
        font-family: {st.session_state.fuente_principal};
    }}
    
    h2, h3, h4 {{
        color: #1a1a1a;
        font-family: {st.session_state.fuente_principal};
    }}
    
    p, li, .stMarkdown, .stCaption {{
        color: #1a1a1a;
    }}
    
    /* ========== BOTÓN PRINCIPAL ========== */
    .stButton button {{
        background-color: {st.session_state.color_principal};
        color: white;
        font-weight: 600;
        border-radius: {st.session_state.bordes}px;
        border: none;
        transition: all 0.3s ease;
    }}
    
    .stButton button:hover {{
        background-color: {st.session_state.color_principal}cc;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }}
    
    /* ========== BOTÓN DE DESCARGA ========== */
    .stDownloadButton button {{
        background-color: #2c2c2c !important;
        color: white !important;
        border-radius: {st.session_state.bordes}px !important;
        border: 1px solid #5a5a5a !important;
        transition: all 0.3s ease;
    }}
    
    .stDownloadButton button:hover {{
        background-color: #3a3a3a !important;
        border-color: {st.session_state.color_principal} !important;
        transform: translateY(-2px);
    }}
    
    /* ========== TABS ========== */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 0.5rem;
        background-color: #e5e5e5;
        border-radius: {st.session_state.bordes}px;
        padding: 0.5rem;
    }}
    
    .stTabs [data-baseweb="tab"] {{
        border-radius: {st.session_state.bordes - 4}px;
        padding: 0.5rem 1.2rem;
        font-weight: 500;
        color: #555555;
    }}
    
    .stTabs [aria-selected="true"] {{
        background-color: {st.session_state.color_principal};
        color: white;
    }}
    
    /* ========== BANNERS ========== */
    .success-banner {{
        background-color: {st.session_state.color_principal}10;
        border-left: 4px solid {st.session_state.color_principal};
        padding: 1rem;
        border-radius: {st.session_state.bordes}px;
        margin: 1rem 0;
        color: #1a1a1a;
    }}
    
    .info-banner {{
        background-color: #f5f5f5;
        border-left: 4px solid #888888;
        padding: 1rem;
        border-radius: {st.session_state.bordes}px;
        margin: 1rem 0;
        color: #555555;
    }}
    
    /* Expander en área principal */
    .streamlit-expanderHeader {{
        background-color: #f0f0f0;
        border-radius: {st.session_state.bordes}px;
        font-weight: 600;
        color: #1a1a1a;
    }}
    
    /* DataFrames */
    [data-testid="stDataFrame"] {{
        border: 1px solid #eaeaea;
        border-radius: {st.session_state.bordes}px;
    }}
    
    /* Spinner */
    .stSpinner > div {{
        border-color: {st.session_state.color_principal} !important;
    }}
    
    /* Footer */
    .footer {{
        text-align: center;
        padding: 1rem;
        color: #888888;
        font-size: 0.75rem;
        border-top: 1px solid #eaeaea;
        margin-top: 2rem;
    }}
    
    hr {{
        border-color: #eaeaea;
    }}
</style>
""", unsafe_allow_html=True)

# ========== BARRA LATERAL ==========
with st.sidebar:
    mostrar_logo(70)
    st.markdown("### ⚖️ AR Collect")
    st.markdown("---")
    
    st.markdown("#### ⚙️ Configuración")
    allow_soft = st.checkbox(
        "Permitir coincidencias suaves (2/3 tokens)",
        value=True,
        help="Si está activado, permite matches con 2 de 3 tokens coincidentes"
    )
    
    st.markdown("---")
    st.markdown("#### 📁 Subir archivos")
    
    ar_file = st.file_uploader("ARCollect_Age_Analysis.xlsx", type=['xlsx', 'xls'], key="ar")
    case_file = st.file_uploader("Case_Details.xlsx", type=['xlsx', 'xls'], key="case")
    closed_file = st.file_uploader("Casos Cerrados.xlsx", type=['xlsx', 'xls'], key="closed")
    
    st.markdown("---")
    
    with st.expander("🎨 Personalización (Administrador)", expanded=False):
        if not st.session_state.password_correcta:
            password_input = st.text_input("Contraseña", type="password", placeholder="Contraseña de admin")
            if st.button("🔓 Acceder"):
                if verificar_password(password_input):
                    st.session_state.password_correcta = True
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta")
        else:
            st.success("✅ Modo editor activado")
            
            st.markdown("**🎨 Personalización**")
            
            nuevo_color = st.color_picker("Color principal (botones)", st.session_state.color_principal)
            if nuevo_color != st.session_state.color_principal:
                st.session_state.color_principal = nuevo_color
                guardar_configuracion()
                st.rerun()
            
            nuevo_bordes = st.slider("Redondez de bordes", 0, 30, st.session_state.bordes)
            if nuevo_bordes != st.session_state.bordes:
                st.session_state.bordes = nuevo_bordes
                guardar_configuracion()
                st.rerun()
            
            st.markdown("---")
            st.markdown("**🖼️ Imágenes Fijas**")
            st.caption("Logo y banner cargados desde GitHub")
            try:
                st.image(LOGO_URL, width=80, caption="Logo actual")
            except:
                st.caption("No se pudo cargar el logo")
            try:
                st.image(BANNER_URL, use_container_width=True, caption="Banner actual")
            except:
                st.caption("No se pudo cargar el banner")
            st.code(BANNER_URL, language="text")
            
            st.markdown("---")
            
            if st.button("🔄 Resetear colores", use_container_width=True):
                st.session_state.color_principal = "#f60d2d"
                st.session_state.bordes = 12
                guardar_configuracion()
                st.rerun()
            
            if st.button("🚪 Salir modo editor", use_container_width=True):
                st.session_state.password_correcta = False
                st.rerun()
    
    with st.expander("🔧 Diagnóstico", expanded=False):
        st.markdown("**Información del Sistema:**")
        st.caption(f"Python: {sys.version[:40]}")
        st.caption(f"Fecha servidor: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        if 'hash_ar' in st.session_state:
            st.markdown("**Hash de archivos (MD5):**")
            st.caption(f"ARCollect: {st.session_state.hash_ar[:8]}...")
            st.caption(f"Case Details: {st.session_state.hash_cl[:8]}...")
            st.caption(f"Casos Cerrados: {st.session_state.hash_cc[:8]}...")
        if 'total_filtrados' in st.session_state:
            st.markdown("**Resultados última ejecución:**")
            st.caption(f"Mantenidos: {st.session_state.total_filtrados}")
            st.caption(f"Descartados: {st.session_state.total_descatados}")
    
    st.caption("📌 Versión 6.0 | Colores fijos")

# ========== ÁREA PRINCIPAL CON BANNER ==========

# Mostrar banner
mostrar_banner()

# Título y subtítulo
col_logo, col_title = st.columns([1, 5])
with col_logo:
    mostrar_logo(60)
with col_title:
    st.markdown("# Procesador de Clientes")
    st.markdown("### AR Collect - Análisis y Filtrado Automático")
    st.caption(f"🕐 Servidor: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

st.markdown("---")

# Tarjetas de estado - ÁREA PRINCIPAL
col1, col2, col3 = st.columns(3)

with col1:
    if ar_file:
        st.markdown(f"""
        <div class="file-card file-card-success">
            <div class="file-icon">📊</div>
            <div class="file-title">ARCollect</div>
            <div class="file-status">✓ Archivo cargado</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="file-card file-card-pending">
            <div class="file-icon">📄</div>
            <div class="file-title">ARCollect</div>
            <div class="file-status">⏳ Esperando archivo</div>
        </div>
        """, unsafe_allow_html=True)

with col2:
    if case_file:
        st.markdown(f"""
        <div class="file-card file-card-success">
            <div class="file-icon">📋</div>
            <div class="file-title">Case Details</div>
            <div class="file-status">✓ Archivo cargado</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="file-card file-card-pending">
            <div class="file-icon">📄</div>
            <div class="file-title">Case Details</div>
            <div class="file-status">⏳ Esperando archivo</div>
        </div>
        """, unsafe_allow_html=True)

with col3:
    if closed_file:
        st.markdown(f"""
        <div class="file-card file-card-success">
            <div class="file-icon">📁</div>
            <div class="file-title">Casos Cerrados</div>
            <div class="file-status">✓ Archivo cargado</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="file-card file-card-pending">
            <div class="file-icon">📄</div>
            <div class="file-title">Casos Cerrados</div>
            <div class="file-status">⏳ Esperando archivo</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# Botón de procesamiento
if ar_file and case_file and closed_file:
    if st.button("🚀 PROCESAR ARCHIVOS", type="primary", use_container_width=True):
        with st.spinner("Procesando archivos..."):
            try:
                filtrados, descartados, log = process_data_with_files(
                    ar_file, case_file, closed_file, allow_soft
                )
                
                st.markdown(f"""
                <div class="success-banner">
                    ✅ <strong>Procesamiento completado!</strong> Se procesaron {len(filtrados) + len(descartados)} registros.
                </div>
                """, unsafe_allow_html=True)
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{len(filtrados):,}</div>
                        <div class="metric-label">Mantenidos</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{len(descartados):,}</div>
                        <div class="metric-label">Descartados</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    total = len(filtrados) + len(descartados)
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{total:,}</div>
                        <div class="metric-label">Total Procesados</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    porcentaje = (len(filtrados) / total * 100) if total > 0 else 0
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{porcentaje:.1f}%</div>
                        <div class="metric-label">Tasa de retención</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    pd.DataFrame(filtrados).to_excel(writer, sheet_name="AR_filtrada", index=False)
                    pd.DataFrame(descartados).to_excel(writer, sheet_name="AR_descartados", index=False)
                    pd.DataFrame(log).to_excel(writer, sheet_name="Match_log", index=False)
                output.seek(0)
                
                st.download_button(
                    label="📥 DESCARGAR REPORTE EXCEL",
                    data=output,
                    file_name=f"reporte_clientes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.markdown("---")
                st.markdown("### 📋 Vista previa de resultados")
                
                tab1, tab2, tab3 = st.tabs(["📌 MANTENIDOS", "🗑️ DESCARTADOS", "📝 LOG DE MATCHES"])
                
                with tab1:
                    if filtrados:
                        st.dataframe(pd.DataFrame(filtrados).head(20), use_container_width=True)
                        st.caption(f"Mostrando 20 de {len(filtrados)} registros")
                    else:
                        st.info("No hay registros mantenidos")
                
                with tab2:
                    if descartados:
                        st.dataframe(pd.DataFrame(descartados).head(20), use_container_width=True)
                        st.caption(f"Mostrando 20 de {len(descartados)} registros")
                    else:
                        st.info("No hay registros descartados")
                
                with tab3:
                    if log:
                        st.dataframe(pd.DataFrame(log).head(20), use_container_width=True)
                    else:
                        st.info("No hay registros en el log")
                        
            except Exception as e:
                st.error(f"❌ Error: {str(e)[:300]}")
                st.info("💡 Si el error persiste, abre los archivos en Excel y guárdalos nuevamente.")
else:
    st.markdown("""
    <div class="info-banner">
        📂 <strong>Para comenzar</strong>, sube los 3 archivos en la barra lateral izquierda
    </div>
    """, unsafe_allow_html=True)

# Pie de página
st.markdown("---")
st.markdown("""
<div class="footer">
    <span>⚖️ Procesador de Clientes | AR Collect</span>
    <span style="margin: 0 1rem">•</span>
    <span>📊 Versión 6.0</span>
</div>
""", unsafe_allow_html=True)
