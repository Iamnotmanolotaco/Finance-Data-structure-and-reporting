import streamlit as st
import pandas as pd
import io
import re
import unicodedata
from collections import defaultdict
from datetime import datetime
import os
import hashlib
import base64
import json
from PIL import Image

# ========== CONFIGURACIÓN DE PÁGINA ==========
st.set_page_config(
    page_title="Procesador de Clientes | AR Collect",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========== CONTRASEÑA PARA EDITOR ==========
EDITOR_PASSWORD = "admin123"  # CAMBIA ESTA CONTRASEÑA
PASSWORD_HASH = hashlib.sha256(EDITOR_PASSWORD.encode()).hexdigest()

# ========== INICIALIZAR VARIABLES DE SESIÓN ==========
if 'modo_editor' not in st.session_state:
    st.session_state.modo_editor = False
if 'password_correcta' not in st.session_state:
    st.session_state.password_correcta = False
if 'logo_base64' not in st.session_state:
    st.session_state.logo_base64 = None
if 'color_principal' not in st.session_state:
    st.session_state.color_principal = "#f60d2d"
if 'color_fondo' not in st.session_state:
    st.session_state.color_fondo = "#f8f9fa"
if 'color_sidebar' not in st.session_state:
    st.session_state.color_sidebar = "#1e1e1e"
if 'color_card' not in st.session_state:
    st.session_state.color_card = "#ffffff"
if 'bordes' not in st.session_state:
    st.session_state.bordes = 12

# ========== FUNCIONES PARA GUARDAR/CARGAR LOGO ==========
LOGO_CONFIG_FILE = "logo_config.json"

def guardar_logo_en_archivo(logo_base64):
    """Guarda el logo en un archivo JSON"""
    try:
        with open(LOGO_CONFIG_FILE, 'w') as f:
            json.dump({'logo': logo_base64}, f)
        return True
    except:
        return False

def cargar_logo_desde_archivo():
    """Carga el logo desde el archivo JSON"""
    try:
        if os.path.exists(LOGO_CONFIG_FILE):
            with open(LOGO_CONFIG_FILE, 'r') as f:
                data = json.load(f)
                return data.get('logo')
    except:
        pass
    return None

def mostrar_logo(tamaño=80):
    """Muestra el logo desde la sesión o archivo"""
    # Intentar cargar logo guardado
    if st.session_state.logo_base64 is None:
        logo_guardado = cargar_logo_desde_archivo()
        if logo_guardado:
            st.session_state.logo_base64 = logo_guardado
    
    # Mostrar logo si existe
    if st.session_state.logo_base64:
        try:
            st.image(f"data:image/png;base64,{st.session_state.logo_base64}", width=tamaño)
        except:
            st.markdown(f"<h1 style='font-size: {tamaño//4}px;'>⚖️</h1>", unsafe_allow_html=True)
    else:
        st.markdown(f"<h1 style='font-size: {tamaño//4}px;'>⚖️</h1>", unsafe_allow_html=True)

def verificar_password(password):
    return hashlib.sha256(password.encode()).hexdigest() == PASSWORD_HASH

# ========== BARRA LATERAL ==========
with st.sidebar:
    # Logo en la barra lateral
    mostrar_logo(70)
    
    st.markdown("### ⚖️ AR Collect")
    st.markdown("---")
    
    # ========== SECCIÓN DE CONFIGURACIÓN ==========
    st.markdown("#### ⚙️ Configuración")
    
    allow_soft = st.checkbox(
        "Permitir coincidencias suaves (2/3 tokens)",
        value=True,
        help="Si está activado, permite matches con 2 de 3 tokens coincidentes"
    )
    
    st.markdown("---")
    
    # ========== SECCIÓN DE SUBIR ARCHIVOS ==========
    st.markdown("#### 📁 Subir archivos")
    
    ar_file = st.file_uploader("ARCollect_Age_Analysis.xlsx", type=['xlsx'], key="ar")
    case_file = st.file_uploader("Case_Details.xlsx", type=['xlsx'], key="case")
    closed_file = st.file_uploader("Casos Cerrados.xlsx", type=['xlsx'], key="closed")
    
    st.markdown("---")
    
    # ========== SECCIÓN DE PERSONALIZACIÓN (PROTEGIDA) ==========
    with st.expander("🎨 Personalización (Administrador)", expanded=False):
        st.markdown("⚠️ **Acceso restringido**")
        
        if not st.session_state.password_correcta:
            password_input = st.text_input(
                "Contraseña de administrador",
                type="password",
                key="password_input",
                placeholder="Ingresa la contraseña"
            )
            
            if st.button("🔓 Acceder", key="btn_acceder"):
                if verificar_password(password_input):
                    st.session_state.password_correcta = True
                    st.session_state.modo_editor = True
                    st.success("✅ Acceso concedido")
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta")
            
            st.caption("🔒 Solo administradores pueden modificar colores y logos")
            
        else:
            st.success("✅ Modo editor activado")
            
            # Selector de color principal
            nuevo_color = st.color_picker(
                "🎨 Color principal",
                value=st.session_state.color_principal
            )
            if nuevo_color != st.session_state.color_principal:
                st.session_state.color_principal = nuevo_color
                st.rerun()
            
            # Selector de color de fondo
            nuevo_fondo = st.color_picker(
                "📄 Color de fondo",
                value=st.session_state.color_fondo
            )
            if nuevo_fondo != st.session_state.color_fondo:
                st.session_state.color_fondo = nuevo_fondo
                st.rerun()
            
            # Selector de color de sidebar
            nuevo_sidebar = st.color_picker(
                "📁 Color de barra lateral",
                value=st.session_state.color_sidebar
            )
            if nuevo_sidebar != st.session_state.color_sidebar:
                st.session_state.color_sidebar = nuevo_sidebar
                st.rerun()
            
            # Selector de color de tarjetas
            nuevo_card = st.color_picker(
                "💳 Color de tarjetas",
                value=st.session_state.color_card
            )
            if nuevo_card != st.session_state.color_card:
                st.session_state.color_card = nuevo_card
                st.rerun()
            
            # Bordes redondeados
            nuevo_bordes = st.slider(
                "🔘 Redondez de bordes",
                min_value=0,
                max_value=30,
                value=st.session_state.bordes
            )
            if nuevo_bordes != st.session_state.bordes:
                st.session_state.bordes = nuevo_bordes
                st.rerun()
            
            st.markdown("---")
            
            # Subir logo
            st.markdown("### 🖼️ Logo personalizado")
            logo_file = st.file_uploader(
                "Subir logo (PNG, JPG)",
                type=['png', 'jpg', 'jpeg'],
                key="logo_uploader"
            )
            
            if logo_file:
                # Convertir imagen a base64
                logo_bytes = logo_file.read()
                logo_base64 = base64.b64encode(logo_bytes).decode('utf-8')
                st.session_state.logo_base64 = logo_base64
                guardar_logo_en_archivo(logo_base64)
                st.image(logo_file, width=100)
                st.success("✅ Logo guardado permanentemente")
                st.rerun()
            
            # Botón para eliminar logo
            if st.session_state.logo_base64:
                if st.button("🗑️ Eliminar logo", key="btn_eliminar_logo"):
                    st.session_state.logo_base64 = None
                    if os.path.exists(LOGO_CONFIG_FILE):
                        os.remove(LOGO_CONFIG_FILE)
                    st.success("✅ Logo eliminado")
                    st.rerun()
            
            st.markdown("---")
            
            # Botón para salir del modo editor
            if st.button("🚪 Salir modo editor", key="btn_salir"):
                st.session_state.password_correcta = False
                st.session_state.modo_editor = False
                st.rerun()
    
    st.markdown("---")
    st.caption("📌 Versión 2.0 | Protegido")
    st.caption("🔒 Datos procesados localmente")

# ========== CSS DINÁMICO ==========
st.markdown(f"""
<style>
    .stApp {{ background-color: {st.session_state.color_fondo}; }}
    
    [data-testid="stSidebar"] {{ background-color: {st.session_state.color_sidebar}; }}
    [data-testid="stSidebar"] * {{ color: #e0e0e0; }}
    [data-testid="stSidebar"] .stMarkdown h1, [data-testid="stSidebar"] .stMarkdown h2, [data-testid="stSidebar"] .stMarkdown h3 {{ color: #ffffff; }}
    [data-testid="stSidebar"] hr {{ border-color: #3a3a3a; }}
    
    h1 {{ color: #1a1a1a; font-size: 2.5rem; font-weight: 700; }}
    h2, h3 {{ color: #2d2d2d; }}
    
    .stButton button {{
        background-color: {st.session_state.color_principal};
        color: white;
        font-weight: 600;
        border-radius: {st.session_state.bordes}px;
        border: none;
    }}
    .stButton button:hover {{
        background-color: {st.session_state.color_principal}cc;
        transform: translateY(-2px);
    }}
    
    .stDownloadButton button {{
        background-color: #2c2c2c;
        border-radius: {st.session_state.bordes}px;
    }}
    
    .metric-card {{
        background-color: {st.session_state.color_card};
        border-radius: {st.session_state.bordes}px;
        padding: 1.2rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        text-align: center;
        border-top: 4px solid {st.session_state.color_principal};
    }}
    
    .metric-value {{ font-size: 2.2rem; font-weight: 700; color: #1a1a1a; }}
    .metric-label {{ font-size: 0.85rem; color: #666666; margin-top: 0.5rem; }}
    
    .file-card {{
        background-color: {st.session_state.color_card};
        border-radius: {st.session_state.bordes}px;
        padding: 1rem;
        text-align: center;
        border: 1px solid #eaeaea;
    }}
    .file-card-success {{ border-left: 4px solid {st.session_state.color_principal}; }}
    .file-card-pending {{ border-left: 4px solid #cccccc; background-color: #fafafa; }}
    .file-icon {{ font-size: 2rem; }}
    .file-title {{ font-weight: 600; color: #333333; }}
    
    .stTabs [data-baseweb="tab-list"] {{
        gap: 0.5rem;
        background-color: #f0f0f0;
        border-radius: {st.session_state.bordes}px;
        padding: 0.5rem;
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: {st.session_state.bordes - 4}px;
        padding: 0.5rem 1.2rem;
        font-weight: 500;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {st.session_state.color_principal};
        color: white;
    }}
    
    .success-banner {{
        background-color: {st.session_state.color_principal}10;
        border-left: 4px solid {st.session_state.color_principal};
        padding: 1rem;
        border-radius: {st.session_state.bordes}px;
    }}
    .info-banner {{
        background-color: #f5f5f5;
        border-left: 4px solid #888888;
        padding: 1rem;
        border-radius: {st.session_state.bordes}px;
    }}
    
    .footer {{
        text-align: center;
        padding: 1rem;
        color: #888888;
        font-size: 0.75rem;
        border-top: 1px solid #eaeaea;
        margin-top: 2rem;
    }}
</style>
""", unsafe_allow_html=True)

# ========== FUNCIONES DE PROCESAMIENTO ==========
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
    AR = pd.read_excel(AR_file, header=0, engine='openpyxl')
    cl_file_df = pd.read_excel(cl_file, header=2, engine='openpyxl')
    cc_data = pd.read_excel(cc_file, header=0, engine='openpyxl')
    
    if 'Customer' not in AR.columns:
        for col in AR.columns:
            if 'customer' in str(col).lower():
                AR.rename(columns={col: 'Customer'}, inplace=True)
                break
    
    if 'Customer' in AR.columns:
        AR["normalized_name"] = AR["Customer"].apply(normalize_name)
    else:
        return [], [], []
    
    if not cl_file_df.empty and 'Client Name' in cl_file_df.columns:
        cl_file_df["normalized_name"] = cl_file_df["Client Name"].apply(normalize_name)
        cl_norms_unique = cl_file_df["normalized_name"].dropna().unique().tolist()
        cl_index = defaultdict(list)
        for i, r in cl_file_df.iterrows():
            cl_index[r["normalized_name"]].append(i)
    else:
        cl_norms_unique = []
        cl_index = defaultdict(list)
    
    if not cc_data.empty:
        cc_data["normalized_name"] = cc_data.iloc[:, 0].apply(normalize_name)
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
        row_out = row.copy()
        row_out["Estado_final"] = "Balance = 0"
        row_out["Motivo_descartado"] = "Balance calculado = 0"
        row_out["Case_Status"] = ""
        row_out["Case_Number"] = ""
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
                    best_originals = best_df["Client Name"].astype(str).unique().tolist()
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
        
        row_out = ar_row.copy()
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
            "Best_Match_cl_name": best_match_original,
            "Best_Match_cl_norm": best_match_norm if best_match_norm else norm_cliente,
            "Best_Match_Type": best_label,
            "Best_Match_OverlapTokens": best_inter,
            "Case_Status": case_statuses,
            "Case_Number": case_numbers,
            "Estado_final": estado,
            "Accion": accion,
            "Revisar": best_label == "2/3 tokens (soft)" or best_label == "2/4+ tokens (soft)"
        })
    
    return filtrados_rows, descartados_rows, log_rows

# ========== INTERFAZ PRINCIPAL ==========

# Banner superior con logo
col_logo, col_title = st.columns([1, 5])
with col_logo:
    mostrar_logo(80)
with col_title:
    st.markdown("# Procesador de Clientes")
    st.markdown("### AR Collect - Análisis y Filtrado Automático")

st.markdown("---")

# Tarjetas de estado
col1, col2, col3 = st.columns(3)

with col1:
    if ar_file:
        st.markdown(f"""
        <div class="file-card file-card-success">
            <div class="file-icon">📊</div>
            <div class="file-title">ARCollect</div>
            <div class="file-status" style="color: {st.session_state.color_principal};">✓ Archivo cargado</div>
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
            <div class="file-status" style="color: {st.session_state.color_principal};">✓ Archivo cargado</div>
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
            <div class="file-status" style="color: {st.session_state.color_principal};">✓ Archivo cargado</div>
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
                
                # Métricas
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
                
                # Descarga
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
                
                # Tabs con resultados
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
                        df_log = pd.DataFrame(log)
                        st.dataframe(df_log.head(20), use_container_width=True)
                        if 'Best_Match_Type' in df_log.columns:
                            st.markdown("**Resumen de tipos de match:**")
                            st.dataframe(df_log['Best_Match_Type'].value_counts().reset_index(), use_container_width=True, hide_index=True)
                    else:
                        st.info("No hay registros en el log")
                        
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.exception(e)
else:
    st.markdown("""
    <div class="info-banner">
        📂 <strong>Para comenzar</strong>, sube los 3 archivos en la barra lateral izquierda
    </div>
    """, unsafe_allow_html=True)

# Pie de página
st.markdown("---")
st.markdown(f"""
<div class="footer">
    <span>⚖️ Procesador de Clientes | AR Collect</span>
    <span style="margin: 0 1rem">•</span>
    <span>🎨 Colores personalizables</span>
    <span style="margin: 0 1rem">•</span>
    <span>🔒 Datos procesados localmente</span>
    <span style="margin: 0 1rem">•</span>
    <span>📊 Versión 2.0</span>
</div>
""", unsafe_allow_html=True)
