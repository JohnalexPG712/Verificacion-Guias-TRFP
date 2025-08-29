import os
import re
import streamlit as st
import pandas as pd
import pdfplumber
from datetime import datetime
import shutil

# --- CONFIGURACIÓN INICIAL ---
st.set_page_config(page_title="Sistema de Conciliación de Guías", page_icon="📦", layout="wide")

# --- FUNCIONES PRINCIPALES COMPLETAS ---
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

def pdf_extraer_pais_destino(texto_guia):
    if "UNITED STATES OF AMERICA" in texto_guia: return "UNITED STATES OF AMERICA"
    m_codigo = re.search(r"\(([A-Z]{2})\)", texto_guia)
    if m_codigo: return m_codigo.group(1)
    return ""

def pdf_extraer_facturas(texto_guia, operador, ref_no_maestro_dhl):
    facturas_inv = re.findall(r"INV[:\s]*([A-Z0-9]+)", texto_guia)
    facturas_zffe = re.findall(r"\b(ZFFE\d+|ZFFV\d+)\b", texto_guia)
    todas = set(facturas_inv + facturas_zffe)
    if operador == "DHL" and ref_no_maestro_dhl:
        todas.add(ref_no_maestro_dhl)
    return ", ".join(sorted(list(todas))) if todas else ""

def pdf_extraer_remitente(texto_guia):
    m = re.search(r"(SOLIDEO\s*S\.?A?\.?S\.?)", texto_guia, re.IGNORECASE)
    if m: return "SOLIDEO S.A.S."
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
            else:
                # Para otros operadores
                patrones_inicio = [r"ORIGIN ID:", r"EXPRESS WORLDWIDE", r"UPS WORLDWIDE SERVICE"]
                texto_normalizado = re.sub(r'\s+', ' ', texto_completo)
                indices = [m.start() for m in re.finditer("|".join(patrones_inicio), texto_normalizado)]
                bloques_guias = []
                if not indices:
                    bloques_guias.append(texto_normalizado)
                else:
                    for i in range(len(indices)):
                        inicio = indices[i]
                        fin = indices[i+1] if i + 1 < len(indices) else len(texto_normalizado)
                        bloques_guias.append(texto_normalizado[inicio:fin])
                
                ultimo_ref_dhl = ""
                for bloque in bloques_guias:
                    operador_bloque = pdf_detectar_operador(bloque)
                    if operador_bloque == "DHL":
                        match_ref = re.search(r"#(\d{6,})", bloque)
                        if match_ref: ultimo_ref_dhl = match_ref.group(1)
                    tracking = pdf_extraer_tracking(bloque, operador_bloque)
                    if tracking:
                        datos_pdf.append({
                            "Tracking": tracking,
                            "Fecha_Guia": pdf_extraer_fecha_guia(bloque, operador_bloque),
                            "Pais_Destino_Guia": pdf_extraer_pais_destino(bloque),
                            "Peso_Neto_Guia": pdf_extraer_peso_neto(bloque, operador_bloque),
                            "FMM_Guia": pdf_extraer_fmm_guia(bloque, operador_bloque),
                            "Remitente_Usuario_Guia": pdf_extraer_remitente(bloque),
                            "Facturas_Guia": pdf_extraer_facturas(bloque, operador_bloque, ultimo_ref_dhl)
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
    
    # Extracción de datos básicos
    fmm_formulario, usuario, pais_destino = "", "", ""
    
    for linea in lineas:
        if "FORMULARIO No. No." in linea:
            match = re.search(r'FORMULARIO No\. No\.\s*(\d+)', linea)
            if match: fmm_formulario = match.group(1)
    
    for linea in lineas:
        if "1. USUARIO:" in linea:
            match = re.search(r'1\.\s*USUARIO:\s*(SOLIDEO\s*S\.?A?\.?S\.?)', linea, re.IGNORECASE)
            if match: usuario = "SOLIDEO S.A.S."
    
    for i, linea in enumerate(lineas):
        if "22. País Destino:" in linea:
            texto_busqueda = linea
            for j in range(i, min(i+3, len(lineas))):
                texto_busqueda += " " + lineas[j]
            match = re.search(r'22\.\s*País Destino:.*?(\d+\s+[A-Z\s]+)', texto_busqueda)
            if match:
                pais_destino = re.sub(r'^\d+\s*', '', match.group(1).strip()).strip()
    
    # Búsqueda simplificada de guías en anexos
    datos_finales = []
    en_seccion_anexos = False
    
    for linea in lineas:
        if "DETALLE DE LOS ANEXOS" in linea:
            en_seccion_anexos = True
            continue
        
        if en_seccion_anexos:
            # Buscar tracking numbers
            patrones = [
                r'\b(8837\d{8})\b',
                r'\b(\d{4}\s\d{4}\s\d{4})\b',
                r'\b(\d{12})\b',
                r'\b(COJE[A-Z0-9]{8,})\b',
                r'\b(\d{9,10})\b'
            ]
            
            for patron in patrones:
                matches = re.findall(patron, linea)
                for match in matches:
                    tracking = match.replace(" ", "") if isinstance(match, str) else match
                    datos_finales.append({
                        "Tracking": tracking,
                        "Fecha_FMM": "",
                        "Pais_Destino_FMM": pais_destino,
                        "FMM_Formulario": fmm_formulario,
                        "Remitente_Usuario_FMM": usuario,
                        "Facturas_FMM": ""
                    })
    
    return datos_finales

# --- INTERFAZ STREAMLIT ---
def main():
    st.title("📦 Sistema de Conciliación de Guías Aéreas")
    st.markdown("---")
    
    # Inicializar session state
    if 'resultados' not in st.session_state:
        st.session_state.resultados = None
    if 'archivos_procesados' not in st.session_state:
        st.session_state.archivos_procesados = False
    
    # Sidebar para carga de archivos
    with st.sidebar:
        st.header("📂 Cargar Archivos")
        archivos_guias = st.file_uploader("Guías PDF (FedEx, UPS, DHL)", type="pdf", accept_multiple_files=True)
        archivos_formularios = st.file_uploader("Formularios PDF", type="pdf", accept_multiple_files=True)
        
        if st.button("🔄 Procesar Conciliación", type="primary"):
            if archivos_guias and archivos_formularios:
                with st.spinner("Procesando archivos..."):
                    try:
                        df_guias = procesar_archivos_guias_pdf(archivos_guias)
                        df_formularios = pd.DataFrame()
                        for archivo in archivos_formularios:
                            datos = procesar_formulario_pdf(archivo)
                            if datos:
                                df_formularios = pd.concat([df_formularios, pd.DataFrame(datos)], ignore_index=True)
                        
                        if not df_guias.empty and not df_formularios.empty:
                            df_guias['Tracking'] = df_guias['Tracking'].astype(str)
                            df_formularios['Tracking'] = df_formularios['Tracking'].astype(str)
                            
                            df_conciliado = pd.merge(df_guias, df_formularios, on='Tracking', how='outer', indicator=True)
                            
                            # Simplificar la conciliación para demo
                            def analizar_fila(row):
                                if row['_merge'] == 'left_only': return '❌ SOLO EN GUÍA'
                                if row['_merge'] == 'right_only': return '❌ SOLO EN FMM'
                                return '✅ OK'
                            
                            df_conciliado['Estado_Conciliacion'] = df_conciliado.apply(analizar_fila, axis=1)
                            
                            st.session_state.resultados = df_conciliado
                            st.session_state.archivos_procesados = True
                            st.success("✅ Conciliación completada")
                        else:
                            st.warning("No se pudieron extraer datos suficientes")
                    except Exception as e:
                        st.error(f"Error en procesamiento: {e}")
            else:
                st.warning("⚠️ Debes cargar ambos tipos de archivos")
    
    # Botón de limpieza
    if st.sidebar.button("🗑️ Limpiar Resultados", type="secondary"):
        st.session_state.resultados = None
        st.session_state.archivos_procesados = False
        st.rerun()
    
    # Mostrar resultados
    if st.session_state.resultados is not None:
        st.header("📊 Resultados de Conciliación")
        st.dataframe(st.session_state.resultados, use_container_width=True)
        
        # Estadísticas
        st.subheader("📈 Resumen")
        total_guias = len(st.session_state.resultados)
        solo_guia = len(st.session_state.resultados[st.session_state.resultados['_merge'] == 'left_only'])
        solo_fmm = len(st.session_state.resultados[st.session_state.resultados['_merge'] == 'right_only'])
        coincidencias = len(st.session_state.resultados[st.session_state.resultados['_merge'] == 'both'])
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Guías", total_guias)
        col2.metric("Solo en Guías", solo_guia)
        col3.metric("Solo en FMM", solo_fmm)
        col4.metric("Coincidencias", coincidencias)
        
        # Botones de exportación
        st.subheader("💾 Exportar Resultados")
        col1, col2 = st.columns(2)
        with col1:
            csv = st.session_state.resultados.to_csv(index=False)
            st.download_button("📥 Descargar CSV", csv, "conciliacion.csv", "text/csv")
        
        with col2:
            # Para Excel necesitamos guardar temporalmente
            if st.button("📥 Descargar Excel"):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                    st.session_state.resultados.to_excel(tmp.name, index=False)
                    with open(tmp.name, "rb") as f:
                        st.download_button("Descargar Excel", f, "conciliacion.xlsx")
    
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
