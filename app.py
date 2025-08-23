import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime
import streamlit as st
import io

# ==============================================================================
# --- SECCIÓN 1: LÓGICA DE EXTRACCIÓN (Sin cambios) ---
# ==============================================================================
# --- Funciones para PDF ---
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
def pdf_extraer_facturas(texto_guia, operador, ref_no_maestro_dhl):
    facturas_inv = re.findall(r"INV[:\s]*([A-Z0-9]+)", texto_guia)
    facturas_zffe = re.findall(r"\b(ZFFE\d+|ZFFV\d+)\b", texto_guia)
    todas = set(facturas_inv + facturas_zffe)
    if operador == "DHL" and ref_no_maestro_dhl:
        todas.add(ref_no_maestro_dhl)
    return ", ".join(sorted(list(todas))) if todas else ""
@st.cache_data
def procesar_archivos_pdf(lista_archivos_pdf):
    datos_pdf = []
    for archivo in lista_archivos_pdf:
        try:
            with pdfplumber.open(archivo) as pdf:
                texto_completo = "\n".join(page.extract_text(x_tolerance=1) or "" for page in pdf.pages)
            patrones_inicio = [r"ORIGIN ID:", r"EXPRESS WORLDWIDE", r"UPS WORLDWIDE SERVICE"]
            indices = [m.start() for m in re.finditer("|".join(patrones_inicio), texto_completo)]
            if not indices:
                bloques_guias = [texto_completo]
            else:
                bloques_guias = []
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
                        "País destino guía pdf": pdf_extraer_pais_destino(bloque),
                        "Facturas comercial guía pdf": pdf_extraer_facturas(bloque, operador, ultimo_ref_dhl),
                        "Archivo guía pdf": archivo.name
                    })
        except Exception as e:
            st.error(f"Error procesando el PDF '{archivo.name}': {e}")
    return pd.DataFrame(datos_pdf)

# --- Funciones para CSV ---
@st.cache_data
def procesar_reporte_csv(lista_archivos_csv):
    todos_los_datos = []
    for archivo in lista_archivos_csv:
        try:
            contenido_completo = archivo.getvalue().decode('latin-1')
            lineas = contenido_completo.splitlines()
        except Exception as e:
            st.error(f"Error leyendo el CSV '{archivo.name}': {e}")
            continue
        formulario, usuario, pais_destino = "", "", ""
        for linea in lineas:
            if not formulario and "FORMULARIO No. No." in linea:
                try:
                    parte_posterior = linea.split("FORMULARIO No. No.")[1]
                    match = re.search(r'(\d+)', parte_posterior)
                    if match: formulario = match.group(1).strip()
                except IndexError: pass
            if not usuario and "1. USUARIO:" in linea:
                try:
                    usuario = linea.split(":", 1)[1].split(",")[0].strip()
                except IndexError: pass
            if not pais_destino and "22." in linea and "Destino:" in linea:
                try:
                    match_etiqueta = re.search(r"22\..*Pa..s Destino:", linea)
                    if match_etiqueta:
                        parte_posterior = linea[match_etiqueta.end():]
                        partes = parte_posterior.split(',')
                        for parte in partes:
                            valor_completo = parte.strip()
                            if valor_completo:
                                pais_destino = re.sub(r'^\d+\s*', '', valor_completo).strip()
                                break
                except IndexError: pass
            if formulario and usuario and pais_destino:
                break
        factura_a_asignar = ""
        en_anexos = False
        for linea in lineas:
            linea_strip = linea.strip()
            if "DETALLE DE LOS ANEXOS" in linea_strip: en_anexos = True
            if en_anexos and linea_strip.startswith('6,'):
                if 'servicio' not in linea.lower():
                    match_factura = re.search(r'\b(ZFFE\d+|ZFFV\d+)\b', linea_strip)
                    if match_factura: factura_a_asignar = match_factura.group(0)
        en_anexos = False
        for linea in lineas:
            linea_strip = linea.strip()
            if "DETALLE DE LOS ANEXOS" in linea_strip: en_anexos = True
            if en_anexos and linea_strip.startswith('127,'):
                guia = re.search(r'\b(8837\d{8})\b', linea_strip)
                if guia:
                    todos_los_datos.append({
                        "Tracking": guia.group(0),
                        "País Destino csv": pais_destino,
                        "Factura Comercial csv": factura_a_asignar
                    })
    return pd.DataFrame(todos_los_datos)

# ==============================================================================
# --- SECCIÓN 2: INTERFAZ DE STREAMLIT ---
# ==============================================================================

st.set_page_config(layout="wide", page_title="Conciliador de Guías")
st.title("🚀 Herramienta de Conciliación de Guías Aéreas")
st.markdown("Esta aplicación procesa guías desde archivos **PDF** y las compara con los datos de un formulario **CSV** para identificar diferencias.")

if 'df_resultado' not in st.session_state:
    st.session_state.df_resultado = None

st.header("1. Carga tus archivos")
st.info("Puedes arrastrar y soltar múltiples archivos PDF y CSV a la vez.")
uploaded_files = st.file_uploader(
    "Selecciona los archivos PDF de guías y los CSV de formularios",
    type=['pdf', 'csv'],
    accept_multiple_files=True
)

# --- Botones de Acción ---
col1, col2 = st.columns(2)
with col1:
    if st.button("📊 Conciliar Archivos", type="primary"):
        if not uploaded_files:
            st.warning("Por favor, carga al menos un archivo PDF y un archivo CSV.")
        else:
            files_pdf = [f for f in uploaded_files if f.name.lower().endswith('.pdf')]
            files_csv = [f for f in uploaded_files if f.name.lower().endswith('.csv')]
            if not files_pdf or not files_csv:
                st.error("Error: Debes cargar al menos un archivo PDF Y al menos un archivo CSV.")
            else:
                with st.spinner("Procesando y conciliando datos..."):
                    df_guias = procesar_archivos_pdf(files_pdf)
                    df_formulario = procesar_reporte_csv(files_csv)
                    if df_guias.empty or df_formulario.empty:
                        st.error("No se pudo extraer información de una de las fuentes. No es posible comparar.")
                        st.session_state.df_resultado = None
                    else:
                        df_guias['Tracking'] = df_guias['Tracking'].astype(str)
                        df_formulario['Tracking'] = df_formulario['Tracking'].astype(str)
                        
                        mapa_paises = {"US": "UNITED STATES OF AMERICA", "ESTADOS UNIDOS": "UNITED STATES OF AMERICA", "JP": "JAPAN"}
                        df_guias['Pais_Normalizado'] = df_guias['País destino guía pdf'].str.upper().map(mapa_paises).fillna(df_guias['País destino guía pdf'].str.upper())
                        df_formulario['Pais_Normalizado'] = df_formulario['País Destino csv'].str.upper().map(mapa_paises).fillna(df_formulario['País Destino csv'].str.upper())

                        df_conciliado = pd.merge(df_guias, df_formulario, on='Tracking', how='outer', indicator=True)
                        
                        def analizar_fila(row):
                            origen = row['_merge']
                            if origen == 'left_only': return '❌ SOLO EN PDF'
                            if origen == 'right_only': return '❌ SOLO EN CSV'
                            diferencias = []
                            if str(row['Pais_Normalizado_x']) != str(row['Pais_Normalizado_y']): diferencias.append("País")
                            if str(row['Facturas comercial guía pdf']) != str(row['Factura Comercial csv']): diferencias.append("Facturas")
                            if not diferencias: return '✅ OK'
                            return '❌ Diferencia en: ' + ", ".join(diferencias)
                        
                        df_conciliado['Estado_Conciliacion'] = df_conciliado.apply(analizar_fila, axis=1)
                        st.session_state.df_resultado = df_conciliado

with col2:
    # AJUSTE: El botón de limpiar solo aparece si hay resultados
    if st.session_state.df_resultado is not None:
        if st.button("🧹 Limpiar Resultados"):
            st.session_state.df_resultado = None
            st.experimental_rerun() # Refresca la página para que los resultados desaparezcan

# --- Visualización de Resultados ---
if st.session_state.df_resultado is not None:
    st.success("¡Conciliación completada!")
    df_final = st.session_state.df_resultado
    
    st.header("2. Resumen Ejecutivo")
    conteo_estados = df_final['Estado_Conciliacion'].value_counts()
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("✅ Coincidencias OK", conteo_estados.get('✅ OK', 0))
    col2.metric("❌ Con Diferencias", sum(1 for estado in conteo_estados.index if estado.startswith('❌ Diferencia')))
    col3.metric("❓ Solo en PDF", conteo_estados.get('❌ SOLO EN PDF', 0))
    col4.metric("❓ Solo en CSV", conteo_estados.get('❌ SOLO EN CSV', 0))

    st.header("3. Reporte Detallado de Conciliación")
    columnas_ordenadas = [
        'Estado_Conciliacion', 'Tracking',
        'País destino guía pdf', 'País Destino csv',
        'Facturas comercial guía pdf', 'Factura Comercial csv',
        'Archivo guía pdf'
    ]
    columnas_existentes = [col for col in columnas_ordenadas if col in df_final.columns]
    
    # AJUSTE: El índice de la tabla empieza en 1
    df_display = df_final[columnas_existentes].copy()
    df_display.index = df_display.index + 1
    st.dataframe(df_display)

    @st.cache_data
    def convertir_df_a_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Se guarda sin el índice de pandas
            df.to_excel(writer, index=False, sheet_name='Conciliacion')
        return output.getvalue()

    excel_file = convertir_df_a_excel(df_final[columnas_existentes])
    
    st.download_button(
        label="📥 Descargar Reporte en Excel",
        data=excel_file,
        file_name="reporte_conciliacion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
