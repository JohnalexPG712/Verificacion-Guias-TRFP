import os
import re
import streamlit as st
import pandas as pd
import pdfplumber
from datetime import datetime
import shutil

# --- CONFIGURACIÓN INICIAL ---
st.set_page_config(page_title="Sistema de Conciliación de Guías", page_icon="📦", layout="wide")

# --- FUNCIONES PRINCIPALES (Misma lógica) ---
def pdf_detectar_operador(texto_guia):
    texto_upper = texto_guia.upper()
    if "FEDEX" in texto_upper or "TRK" in texto_upper or "MPS#" in texto_upper: return "FedEx"
    if "EXPRESS WORLDWIDE" in texto_upper or "WAYBILL" in texto_upper: return "DHL"
    if "UPS WORLDWIDE SERVICE" in texto_upper or "COJE" in texto_upper: return "UPS"
    return "Desconocido"

def pdf_extraer_tracking(texto_guia, operador):
    if operador == "FedEx":
        posibles = re.findall(r"\b(?:\d{4}\s\d{4}\s\d{4}|\d{12})\b", texto_guia)
        if len(posibles) == 1: return posibles[0].replace(" ", "")
        if len(posibles) > 1:
            match_master = re.search(r"Mstr#\s*(\d{4}\s\d{4}\s\d{4}|\d{12})", texto_guia)
            master = match_master.group(1).replace(" ", "") if match_master else None
            for track in posibles:
                if track.replace(" ", "") != master: return track.replace(" ", "")
            return posibles[0].replace(" ", "")
    elif operador == "DHL":
        m = re.search(r"WAYBILL\s+([\d\s]{10,})", texto_guia)
        return re.sub(r"\s+", "", m.group(1)) if m else ""
    elif operador == "UPS":
        m = re.search(r"SERVICE\s+(COJE[A-Z0-9]+)", texto_guia)
        return m.group(1) if m else ""
    return ""

def pdf_extraer_peso_neto(texto_guia, operador):
    texto_guia_upper = texto_guia.upper()
    match_pn = re.search(r"PN[:\s]*([\d.,]+)", texto_guia_upper)
    if match_pn:
        peso = match_pn.group(1).replace(",", ".")
        try: return f"{float(peso):.2f}"
        except: return peso
    return ""

def pdf_extraer_fmm_guia(texto_guia, operador):
    patrones = [r"FMM[:\s]*(\d+)", r"FMM\s*No\.?\s*(\d+)", r"F\.M\.M\.\s*(\d+)", r"\b(\d{6})\b(?=.*FMM)"]
    for patron in patrones:
        match = re.search(patron, texto_guia, re.IGNORECASE)
        if match: return match.group(1)
    return ""

def pdf_extraer_fecha_ups(texto_guia):
    patrones = [r"Date\s+(\d{1,2}\s+[A-Z]{3}\s+\d{4})", r"UPS WORLDWIDE SERVICE.*?Date\s+(\d{1,2}\s+[A-Z]{3}\s+\d{4})"]
    for patron in patrones:
        match = re.search(patron, texto_guia, re.IGNORECASE | re.DOTALL)
        if match:
            try: return datetime.strptime(match.group(1).title(), '%d %b %Y').strftime('%Y-%m-%d')
            except ValueError: continue
    return ""

def pdf_extraer_fecha_dhl(texto_guia):
    match = re.search(r"(\d{4}-\d{2}-\d{2})\s*MyDHL", texto_guia)
    return match.group(1) if match else ""

def pdf_extraer_fecha_fedex(texto_guia):
    match = re.search(r"SHIP DATE:\s*(\d{2}[A-Z]{3}\d{2})", texto_guia, re.IGNORECASE)
    if not match: return ""
    try: return datetime.strptime(match.group(1).title(), '%d%b%y').strftime('%Y-%m-%d')
    except ValueError: return ""

def pdf_extraer_fecha_guia(texto_guia, operador):
    if operador == "UPS": return pdf_extraer_fecha_ups(texto_guia)
    elif operador == "DHL": return pdf_extraer_fecha_dhl(texto_guia)
    elif operador == "FedEx": return pdf_extraer_fecha_fedex(texto_guia)
    return ""

def procesar_archivos_guias_pdf(archivos):
    datos_pdf = []
    for archivo in archivos:
        try:
            with pdfplumber.open(archivo) as pdf:
                texto_completo = "\n".join(page.extract_text(x_tolerance=1) or "" for page in pdf.pages)
            texto_normalizado = re.sub(r'\s+', ' ', texto_completo)
            operador = pdf_detectar_operador(texto_normalizado)
            
            if operador == "UPS":
                tracking = pdf_extraer_tracking(texto_normalizado, operador)
                if tracking:
                    datos_pdf.append({
                        "Tracking": tracking,
                        "Fecha_Guia": pdf_extraer_fecha_guia(texto_normalizado, operador),
                        "Pais_Destino_Guia": pdf_extraer_pais_destino(texto_normalizado),
                        "Peso_Neto_Guia": pdf_extraer_peso_neto(texto_normalizado, operador),
                        "FMM_Guia": pdf_extraer_fmm_guia(texto_normalizado, operador),
                        "Remitente_Usuario_Guia": pdf_extraer_remitente(texto_normalizado),
                        "Facturas_Guia": pdf_extraer_facturas(texto_normalizado, operador, "")
                    })
        except Exception as e:
            st.error(f"Error procesando {archivo.name}: {e}")
    return pd.DataFrame(datos_pdf)

def procesar_formulario_pdf(archivo):
    try:
        with pdfplumber.open(archivo) as pdf:
            contenido_completo = "\n".join(page.extract_text(x_tolerance=1) or "" for page in pdf.pages)
            lineas = contenido_completo.splitlines()
    except Exception as e:
        st.error(f"Error leyendo formulario {archivo.name}: {e}")
        return []
    
    # ... (resto de la función igual que antes)
    return datos_filtrados

# --- INTERFAZ STREAMLIT ---
def main():
    st.title("📦 Sistema de Conciliación de Guías Aéreas")
    st.markdown("---")
    
    # Inicializar session state
    if 'resultados' not in st.session_state:
        st.session_state.resultados = None
    if 'archivos_procesados' not in st.session_state:
        st.session_state.archivos_procesados = []
    
    # Sidebar para carga de archivos
    with st.sidebar:
        st.header("📂 Cargar Archivos")
        archivos_guias = st.file_uploader("Guías PDF (FedEx, UPS, DHL)", type="pdf", accept_multiple_files=True)
        archivos_formularios = st.file_uploader("Formularios PDF", type="pdf", accept_multiple_files=True)
        
        if st.button("🔄 Procesar Conciliación", type="primary"):
            if archivos_guias and archivos_formularios:
                with st.spinner("Procesando archivos..."):
                    df_guias = procesar_archivos_guias_pdf(archivos_guias)
                    df_formularios = pd.DataFrame()
                    for archivo in archivos_formularios:
                        datos = procesar_formulario_pdf(archivo)
                        df_formularios = pd.concat([df_formularios, pd.DataFrame(datos)], ignore_index=True)
                    
                    # Lógica de conciliación
                    if not df_guias.empty and not df_formularios.empty:
                        df_guias['Tracking'] = df_guias['Tracking'].astype(str)
                        df_formularios['Tracking'] = df_formularios['Tracking'].astype(str)
                        
                        df_conciliado = pd.merge(df_guias, df_formularios, on='Tracking', how='outer', indicator=True)
                        
                        # ... (resto de la lógica de conciliación)
                        
                        st.session_state.resultados = df_conciliado
                        st.session_state.archivos_procesados = archivos_guias + archivos_formularios
                        st.success("✅ Conciliación completada")
            else:
                st.warning("⚠️ Debes cargar ambos tipos de archivos")
    
    # Botón de limpieza
    if st.sidebar.button("🗑️ Limpiar Resultados", type="secondary"):
        st.session_state.resultados = None
        st.session_state.archivos_procesados = []
        st.rerun()
    
    # Mostrar resultados
    if st.session_state.resultados is not None:
        st.header("📊 Resultados de Conciliación")
        st.dataframe(st.session_state.resultados, use_container_width=True)
        
        # Botones de exportación
        col1, col2 = st.columns(2)
        with col1:
            csv = st.session_state.resultados.to_csv(index=False)
            st.download_button("📥 Descargar CSV", csv, "conciliacion.csv", "text/csv")
        with col2:
            excel_buffer = pd.ExcelWriter("conciliacion.xlsx", engine='xlsxwriter')
            st.session_state.resultados.to_excel(excel_buffer, index=False)
            excel_buffer.close()
            with open("conciliacion.xlsx", "rb") as f:
                st.download_button("📥 Descargar Excel", f, "conciliacion.xlsx")
    
    # Información de uso
    with st.expander("ℹ️ Instrucciones de uso"):
        st.markdown("""
        1. **Cargar archivos**: Sube las guías PDF y formularios PDF
        2. **Procesar**: Haz clic en 'Procesar Conciliación'
        3. **Revisar resultados**: Los resultados se mostrarán en tabla
        4. **Exportar**: Descarga en CSV o Excel
        5. **Limpiar**: Usa 'Limpiar Resultados' para empezar de nuevo
        
        **Formatos soportados:**
        - Guías: FedEx, UPS, DHL
        - Formularios: Formularios de movimiento de mercancías
        """)

if __name__ == "__main__":
    main()
