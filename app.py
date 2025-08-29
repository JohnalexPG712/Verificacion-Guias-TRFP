import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime
import streamlit as st
import io

# ==============================================================================
# --- SECCIÓN 1: LÓGICA DE EXTRACCIÓN (Tu código, adaptado para Streamlit) ---
# ==============================================================================

# --- Funciones para PDF de Guías ---
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
        m = re.search(r"\b(COJE[A-Z0-9]{9,})\b", texto_guia)
        return m.group(1).strip() if m else ""
    return ""
def pdf_extraer_pais_destino(texto_guia):
    if "UNITED STATES OF AMERICA" in texto_guia: return "UNITED STATES OF AMERICA"
    m_codigo = re.search(r"\(([A-Z]{2})\)", texto_guia)
    if m_codigo: return m_codigo.group(1)
    return ""
def pdf_extraer_facturas_guia(texto_guia, operador, ref_no_maestro_dhl):
    facturas_inv = re.findall(r"INV[:\s]*([A-Z0-9]+)", texto_guia)
    facturas_zffe = re.findall(r"\b(ZFFE\d+|ZFFV\d+)\b", texto_guia)
    todas = set(facturas_inv + facturas_zffe)
    if operador == "DHL" and ref_no_maestro_dhl:
        todas.add(ref_no_maestro_dhl)
    return ", ".join(sorted(list(todas))) if todas else ""
def pdf_extraer_pn_guia(texto_guia):
    m = re.search(r"PN[:\s]*([\d.,]+)", texto_guia)
    return m.group(1).replace(",", ".") if m else ""
def pdf_extraer_fmm_guia(texto_guia):
    m = re.search(r"(FMM\d+)", texto_guia); return m.group(1) if m else ""
def pdf_extraer_remitente_guia(texto_guia):
    m = re.search(r"(SOLIDEO\s*S\.?A?\.?S\.?)", texto_guia, re.IGNORECASE)
    if m: return "SOLIDEO S.A.S."
    return ""
def pdf_extraer_fecha_guia(texto_guia):
    patron = r"(?i)(?:SHIP DATE:|Date).*?(\d{1,2}\s*[A-Z]{3}\s*\d{4}|\d{4}-\d{2}-\d{2}|\d{2}[A-Z]{3}\d{2})"
    m = re.search(patron, texto_guia, re.DOTALL)
    if not m: return ""
    fecha_str = m.group(1) if m.group(1) else (m.group(2) if m.group(2) else m.group(3))
    if not fecha_str: return ""
    fecha_str_limpia = " ".join(fecha_str.split()).title()
    formatos = [('%d%b%y', '%Y-%m-%d'), ('%Y-%m-%d', '%Y-%m-%d'), ('%d %b %Y', '%Y-%m-%d')]
    for fmt_in, fmt_out in formatos:
        try: return datetime.strptime(fecha_str_limpia, fmt_in).strftime(fmt_out)
        except ValueError: continue
    return fecha_str
@st.cache_data
def procesar_archivos_guias_pdf(lista_archivos_pdf):
    datos_pdf = []
    for archivo in lista_archivos_pdf:
        try:
            with pdfplumber.open(archivo) as pdf:
                texto_completo = "\n".join(page.extract_text(x_tolerance=1) or "" for page in pdf.pages)
            patrones_inicio = [r"ORIGIN ID:", r"EXPRESS WORLDWIDE", r"UPS WORLDWIDE SERVICE"]
            indices = [m.start() for m in re.finditer("|".join(patrones_inicio), texto_completo)]
            bloques_guias = []
            if not indices:
                bloques_guias.append(texto_completo)
            else:
                for i in range(len(indices)):
                    inicio = indices[i]
                    fin = indices[i+1] if i + 1 < len(indices) else len(texto_completo)
                    bloques_guias.append(texto_completo[inicio:fin])
            ultimo_ref_dhl = ""
            for bloque in bloques_guias:
                operador = pdf_detectar_operador(bloque)
                if operador == "DHL":
                    match_ref = re.search(r"#(\d{6,})", bloque)
                    if match_ref: ultimo_ref_dhl = match_ref.group(1)
                tracking = pdf_extraer_tracking(bloque, operador)
                if tracking:
                    datos_pdf.append({
                        "Tracking": tracking,
                        "Fecha_Guia": pdf_extraer_fecha_guia(bloque),
                        "Pais_Destino_Guia": pdf_extraer_pais_destino(bloque),
                        "FMM_Guia": pdf_extraer_fmm_guia(bloque),
                        "Peso_Neto_Guia": pdf_extraer_pn_guia(bloque),
                        "Remitente_Usuario_Guia": pdf_extraer_remitente_guia(bloque),
                        "Facturas_Guia": pdf_extraer_facturas_guia(bloque, operador, ultimo_ref_dhl)
                    })
        except Exception as e:
            st.error(f"Error procesando PDF de Guía '{archivo.name}': {e}")
    return pd.DataFrame(datos_pdf)

# --- Funciones para PDF de Formularios ---
@st.cache_data
def procesar_archivos_formularios_pdf(lista_archivos_pdf):
    todos_los_datos = []
    for archivo in lista_archivos_pdf:
        try:
            with pdfplumber.open(archivo) as pdf:
                contenido_completo = "\n".join(page.extract_text(x_tolerance=1) or "" for page in pdf.pages)
                lineas = contenido_completo.splitlines()
        except Exception as e:
            st.error(f"Error leyendo PDF de Formulario '{archivo.name}': {e}")
            continue
        formulario, usuario, pais_destino = "", "", ""
        for linea in lineas:
            if not formulario and "FORMULARIO No. No." in linea:
                match = re.search(r'FORMULARIO No\. No\.\s*(\d+)', linea)
                if match: formulario = match.group(1)
            if not usuario and "1. USUARIO:" in linea:
                match = re.search(r'1\.\s*USUARIO:\s*(SOLIDEO\s*S\.?A?\.?S\.?)', linea)
                if match: usuario = "SOLIDEO S.A.S."
            if not pais_destino and "22. País Destino:" in linea:
                match = re.search(r"22\.\s*País Destino:.*?(\d+\s+[A-Z\s]+)", linea)
                if match: pais_destino = re.sub(r'^\d+\s*', '', match.group(1).strip()).strip()
            if formulario and usuario and pais_destino:
                break
        factura_a_asignar = ""
        en_anexos = False
        for linea in lineas:
            if "DETALLE DE LOS ANEXOS" in linea: en_anexos = True
            if en_anexos and re.search(r'^\s*6\s', linea) and "FACTURA COMERCIAL" in linea:
                if 'servicio' not in linea.lower():
                    match_factura = re.search(r'\b(ZFFE\d+|ZFFV\d+)\b', linea)
                    if match_factura: factura_a_asignar = match_factura.group(0)
        en_anexos = False
        for linea in lineas:
            if "DETALLE DE LOS ANEXOS" in linea: en_anexos = True
            if en_anexos and re.search(r'^\s*127\s', linea) and "GUIAS DE TRAFICO POSTAL" in linea:
                guia = re.search(r'\b(8837\d{8})\b', linea)
                fecha = re.search(r'(\d{4}/\d{2}/\d{2})', linea)
                if guia:
                    todos_los_datos.append({
                        "Tracking": guia.group(0),
                        "Fecha_FMM": datetime.strptime(fecha.group(0), '%Y/%m/%d').strftime('%Y-%m-%d') if fecha else "",
                        "Pais_Destino_FMM": pais_destino,
                        "FMM_FMM": formulario,
                        "Remitente_Usuario_FMM": usuario,
                        "Facturas_FMM": factura_a_asignar
                    })
    return pd.DataFrame(todos_los_datos)

# ==============================================================================
# --- SECCIÓN 2: INTERFAZ DE STREAMLIT ---
# ==============================================================================
st.set_page_config(layout="wide", page_title="Conciliador de Guías")
st.title("🚀 Herramienta de Conciliación de Guías y Formularios")
st.markdown("Carga los archivos PDF de guías y los archivos PDF de formularios para compararlos.")

if 'df_resultado' not in st.session_state:
    st.session_state.df_resultado = None
if 'file_uploader_key' not in st.session_state:
    st.session_state.file_uploader_key = 0

uploaded_files = st.file_uploader(
    "Selecciona todos tus archivos PDF",
    type=['pdf'],
    accept_multiple_files=True,
    key=f"file_uploader_{st.session_state.file_uploader_key}"
)

col1, col2, _ = st.columns([1.5, 1.5, 3])
with col1:
    if st.button("📊 Conciliar Archivos", type="primary"):
        if not uploaded_files:
            st.warning("Por favor, carga al menos un PDF de guías y un PDF de formulario.")
        else:
            files_guias_pdf, files_formularios_pdf = [], []
            with st.spinner("Clasificando archivos PDF..."):
                for archivo in uploaded_files:
                    try:
                        with pdfplumber.open(archivo) as pdf:
                            primer_pagina_texto = pdf.pages[0].extract_text(x_tolerance=1) or ""
                            if "FORMULARIO DE MOVIMIENTO DE MERCANCÍAS" in primer_pagina_texto:
                                files_formularios_pdf.append(archivo)
                            else:
                                files_guias_pdf.append(archivo)
                    except Exception:
                        st.warning(f"No se pudo leer '{archivo.name}' para clasificarlo. Se asumirá que es una guía.")
                        files_guias_pdf.append(archivo)

            if not files_guias_pdf or not files_formularios_pdf:
                st.error("Error: Debes cargar al menos un PDF de guías Y un PDF de formulario.")
            else:
                with st.spinner("Procesando y conciliando datos..."):
                    df_guias = procesar_archivos_guias_pdf(files_guias_pdf)
                    df_formulario = procesar_archivos_formularios_pdf(files_formularios_pdf)

                    if df_guias.empty or df_formulario.empty:
                        st.error("No se pudo extraer información de una de las fuentes. No es posible comparar.")
                        st.session_state.df_resultado = None
                    else:
                        df_guias['Tracking'] = df_guias['Tracking'].astype(str)
                        df_formulario['Tracking'] = df_formulario['Tracking'].astype(str)
                        
                        df_conciliado = pd.merge(df_guias, df_formulario, on='Tracking', how='outer', indicator=True)
                        
                        mapa_paises = {"US": "UNITED STATES OF AMERICA", "ESTADOS UNIDOS": "UNITED STATES OF AMERICA", "JP": "JAPAN"}
                        df_conciliado['Pais_Normalizado_GUÍA'] = df_conciliado['Pais_Destino_GUÍA'].str.upper().map(mapa_paises).fillna(df_conciliado['Pais_Destino_GUÍA'].str.upper())
                        df_conciliado['Pais_Normalizado_FMM'] = df_conciliado['Pais_Destino_FMM'].str.upper().map(mapa_paises).fillna(df_conciliado['Pais_Destino_FMM'].str.upper())
                        df_conciliado['FMM_Normalizado_GUÍA'] = df_conciliado['FMM_GUÍA'].str.replace('FMM', '', regex=False)
                        
                        def analizar_fila(row):
                            origen = row['_merge']
                            if origen == 'left_only': return '❌ SOLO EN GUÍA'
                            if origen == 'right_only': return '❌ SOLO EN FMM'
                            diferencias = []
                            if str(row['Fecha_GUÍA']) != str(row['Fecha_FMM']): diferencias.append("Fecha")
                            if str(row['Pais_Normalizado_GUÍA']) != str(row['Pais_Normalizado_FMM']): diferencias.append("País")
                            if str(row['FMM_Normalizado_GUÍA']) != str(row['FMM_FMM']): diferencias.append("FMM")
                            if str(row['Remitente_Usuario_GUÍA']) != str(row['Remitente_Usuario_FMM']): diferencias.append("Remitente")
                            if str(row['Facturas_GUÍA']) != str(row['Facturas_FMM']): diferencias.append("Facturas")
                            if not diferencias: return '✅ OK'
                            return '❌ Diferencia en: ' + ", ".join(diferencias)
                        
                        df_conciliado['Estado_Conciliacion'] = df_conciliado.apply(analizar_fila, axis=1)
                        st.session_state.df_resultado = df_conciliado
with col2:
    if st.session_state.df_resultado is not None:
        if st.button("🧹 Limpiar Resultados"):
            st.session_state.df_resultado = None
            st.session_state.file_uploader_key += 1
            st.rerun()

if st.session_state.df_resultado is not None:
    st.success("¡Conciliación completada!")
    df_final = st.session_state.df_resultado.drop(columns=['_merge', 'Pais_Normalizado_GUÍA', 'Pais_Normalizado_FMM', 'FMM_Normalizado_GUÍA'], errors='ignore')
    
    st.header("Resumen Ejecutivo")
    conteo_estados = df_final['Estado_Conciliacion'].value_counts()
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("✅ Coincidencias OK", conteo_estados.get('✅ OK', 0))
    col2.metric("❌ Con Diferencias", sum(1 for estado in conteo_estados.index if estado.startswith('❌ Diferencia')))
    col3.metric("❓ Solo en Guías (PDF)", conteo_estados.get('❌ SOLO EN GUÍA', 0))
    col4.metric("❓ Solo en Formularios (FMM)", conteo_estados.get('❌ SOLO EN FMM', 0))

    st.header("Reporte Detallado de Conciliación")
    columnas_ordenadas = [
        'Estado_Conciliacion', 'Tracking',
        'Fecha_GUÍA', 'Fecha_FMM',
        'Pais_Destino_GUÍA', 'Pais_Destino_FMM',
        'FMM_GUÍA', 'FMM_FMM',
        'Peso_Neto_GUÍA',
        'Remitente_Usuario_GUÍA', 'Remitente_Usuario_FMM',
        'Facturas_GUÍA', 'Facturas_FMM',
    ]
    columnas_existentes = [col for col in columnas_ordenadas if col in df_final.columns]
    df_display = df_final[columnas_existentes].copy()
    df_display.index = range(1, len(df_display) + 1)
    st.dataframe(df_display)
    
    @st.cache_data
    def convertir_df_a_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Conciliacion')
        return output.getvalue()

    excel_file = convertir_df_a_excel(df_final[columnas_existentes])
    
    st.download_button(
        label="📥 Descargar Reporte en Excel",
        data=excel_file,
        file_name="Resultado_Verificación_Guias.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

