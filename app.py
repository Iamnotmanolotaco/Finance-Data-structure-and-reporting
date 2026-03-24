import streamlit as st
import pandas as pd
import io
import os
import re
import unicodedata
from collections import defaultdict
from datetime import datetime

# Configurar página
st.set_page_config(
    page_title="Procesador de Clientes",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# === TODAS TUS FUNCIONES ORIGINALES (copiadas exactamente) ===
# Configuración
ALLOW_TWO_OF_THREE_SOFT = True

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

def classify_match(a: str, b: str):
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
        if inter == 2 and ALLOW_TWO_OF_THREE_SOFT:
            return (True, "2/3 tokens (soft)", inter)
        return (False, "no", inter)

    if nmin == 2:
        return (inter == 2, "2/2 tokens" if inter == 2 else "no", inter)

    if nmin == 1:
        return (a == b, "1 token exact" if a == b else "no", inter)

    return (False, "no", inter)

def process_data_with_files(AR_file, cl_file, cc_file):
    """Procesa los archivos subidos y devuelve los resultados"""
    
    # Cargar los DataFrames
    AR = pd.read_excel(AR_file, header=0, engine='calamine')
    cl_file_df = pd.read_excel(cl_file, header=2, engine='calamine')
    cc_data = pd.read_excel(cc_file, header=0, engine='calamine')
    
    # Normalizar nombres
    if 'Customer' not in AR.columns:
        for col in AR.columns:
            if 'customer' in str(col).lower():
                AR.rename(columns={col: 'Customer'}, inplace=True)
                break
    
    if 'Customer' in AR.columns:
        AR["normalized_name"] = AR["Customer"].apply(normalize_name)
    else:
        return [], [], []
    
    # Normalizar Case_Details
    if not cl_file_df.empty and 'Client Name' in cl_file_df.columns:
        cl_file_df["normalized_name"] = cl_file_df["Client Name"].apply(normalize_name)
        cl_norms_unique = cl_file_df["normalized_name"].dropna().unique().tolist()
        
        cl_index = defaultdict(list)
        for i, r in cl_file_df.iterrows():
            cl_index[r["normalized_name"]].append(i)
    else:
        cl_norms_unique = []
        cl_index = defaultdict(list)
    
    # Normalizar Casos Cerrados
    if not cc_data.empty:
        cc_data["normalized_name"] = cc_data.iloc[:, 0].apply(normalize_name)
        cc_norms = cc_data["normalized_name"].dropna().unique().tolist()
    else:
        cc_norms = []
    
    # Calcular balances
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
    
    # Procesar balance cero
    for _, row in AR_zero_balance.iterrows():
        row_out = row.copy()
        row_out["Estado_final"] = "Balance = 0"
        row_out["Motivo_descartado"] = "Balance calculado = 0"
        row_out["Case_Status"] = ""
        row_out["Case_Number"] = ""
        descartados_rows.append(row_out)
    
    # Procesar balance positivo
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
                ok, label, inter = classify_match(norm_cliente, cand)
                if ok:
                    matched_normals.append(cand)
                    match_types[cand] = label
                    match_inters[cand] = inter
            
            if matched_normals:
                best_match_norm = max(
                    matched_normals,
                    key=lambda c: (match_inters.get(c, 0), len(c.split()))
                )
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
                            cc_match = any(classify_match(norm_cliente, cc)[0] for cc in cc_norms)
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

# === INTERFAZ STREAMLIT ===
def main():
    st.title("📊 Procesador de Clientes")
    st.markdown("""
    ### Sube los archivos necesarios y genera el reporte automáticamente
    
    **Archivos requeridos:**
    1. **ARCollect_Age_Analysis.xlsx** - Datos de aging de clientes
    2. **Case_Details.xlsx** - Detalles de casos
    3. **Casos Cerrados.xlsx** - Lista de casos cerrados
    """)
    
    st.markdown("---")
    
    # Sidebar para configuraciones
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        allow_soft = st.checkbox(
            "Permitir coincidencias suaves (2/3 tokens)",
            value=True,
            help="Si está activado, permite matches con 2 de 3 tokens coincidentes"
        )
        
        st.markdown("---")
        st.header("📁 Subir archivos")
        
        # Subida de archivos
        ar_file = st.file_uploader(
            "1. ARCollect_Age_Analysis.xlsx",
            type=['xlsx'],
            help="Archivo con el análisis de aging"
        )
        
        case_file = st.file_uploader(
            "2. Case_Details.xlsx",
            type=['xlsx'],
            help="Archivo con detalles de casos"
        )
        
        closed_file = st.file_uploader(
            "3. Casos Cerrados.xlsx",
            type=['xlsx'],
            help="Archivo con casos cerrados"
        )
    
    # Mostrar estado de archivos
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if ar_file:
            st.success("✅ ARCollect cargado")
        else:
            st.info("⏳ Esperando ARCollect")
    
    with col2:
        if case_file:
            st.success("✅ Case Details cargado")
        else:
            st.info("⏳ Esperando Case Details")
    
    with col3:
        if closed_file:
            st.success("✅ Casos Cerrados cargado")
        else:
            st.info("⏳ Esperando Casos Cerrados")
    
    st.markdown("---")
    
    # Botón para procesar
    if ar_file and case_file and closed_file:
        if st.button("🚀 Procesar Archivos", type="primary", use_container_width=True):
            
            # Mostrar spinner mientras procesa
            with st.spinner("Procesando archivos... Esto puede tomar unos segundos"):
                try:
                    # Actualizar configuración global
                    global ALLOW_TWO_OF_THREE_SOFT
                    ALLOW_TWO_OF_THREE_SOFT = allow_soft
                    
                    # Procesar
                    filtrados, descartados, log = process_data_with_files(
                        ar_file, case_file, closed_file
                    )
                    
                    # Mostrar resultados
                    st.success("✅ Procesamiento completado!")
                    
                    # Métricas
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total mantenidos", len(filtrados))
                    with col2:
                        st.metric("Total descartados", len(descartados))
                    with col3:
                        st.metric("Total procesados", len(filtrados) + len(descartados))
                    
                    # Crear archivo Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        pd.DataFrame(filtrados).to_excel(writer, sheet_name="AR_filtrada", index=False)
                        pd.DataFrame(descartados).to_excel(writer, sheet_name="AR_descartados", index=False)
                        pd.DataFrame(log).to_excel(writer, sheet_name="Match_log", index=False)
                    
                    output.seek(0)
                    
                    # Botón de descarga
                    st.download_button(
                        label="📥 Descargar Reporte Excel",
                        data=output,
                        file_name=f"reporte_clientes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    # Mostrar vista previa
                    st.markdown("---")
                    st.subheader("📋 Vista previa de resultados")
                    
                    tab1, tab2, tab3 = st.tabs(["Mantenidos", "Descartados", "Log de Matches"])
                    
                    with tab1:
                        if filtrados:
                            df_filtrados = pd.DataFrame(filtrados)
                            st.dataframe(df_filtrados.head(20), use_container_width=True)
                            st.caption(f"Mostrando 20 de {len(filtrados)} registros")
                        else:
                            st.info("No hay registros mantenidos")
                    
                    with tab2:
                        if descartados:
                            df_descartados = pd.DataFrame(descartados)
                            st.dataframe(df_descartados.head(20), use_container_width=True)
                            st.caption(f"Mostrando 20 de {len(descartados)} registros")
                        else:
                            st.info("No hay registros descartados")
                    
                    with tab3:
                        if log:
                            df_log = pd.DataFrame(log)
                            st.dataframe(df_log.head(20), use_container_width=True)
                            st.caption(f"Mostrando 20 de {len(log)} registros")
                            
                            # Mostrar resumen de tipos de match
                            if 'Best_Match_Type' in df_log.columns:
                                st.markdown("**Resumen de tipos de match:**")
                                st.dataframe(
                                    df_log['Best_Match_Type'].value_counts().reset_index(),
                                    use_container_width=True,
                                    hide_index=True
                                )
                        else:
                            st.info("No hay registros en el log")
                    
                except Exception as e:
                    st.error(f"❌ Error durante el procesamiento: {str(e)}")
                    st.exception(e)
    else:
        st.info("📂 Por favor, sube los 3 archivos para comenzar el procesamiento")
    
    # Instrucciones
    with st.expander("ℹ️ Instrucciones detalladas"):
        st.markdown("""
        ### Cómo usar esta aplicación:
        
        1. **Prepara los archivos** necesarios en tu computadora
        2. **Usa la barra lateral** para subir cada archivo
        3. **Ajusta la configuración** si es necesario (coincidencias suaves)
        4. **Haz clic en "Procesar Archivos"**
        5. **Descarga el resultado** cuando termine
        
        ### Formato esperado de archivos:
        
        - **ARCollect_Age_Analysis.xlsx**: Debe contener columna 'Customer' y columnas de aging
        - **Case_Details.xlsx**: Debe tener headers en fila 3, con columna 'Client Name'
        - **Casos Cerrados.xlsx**: Primera columna con nombres de clientes cerrados
        
        ### Resultados:
        
        - **AR_filtrada**: Clientes con balance positivo y casos activos
        - **AR_descartados**: Clientes descartados (balance cero o casos cerrados)
        - **Match_log**: Registro detallado de todas las coincidencias encontradas
        """)

if __name__ == "__main__":
    main()