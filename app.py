import os
import re
import streamlit as st
import pandas as pd
import pdfplumber
from datetime import datetime
import tempfile

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
                patrones_inicio = [r"ORIGIN ID:", r"EXPRESS WORLDWIDE", r"UPS WORLDWIDE SERVICE"]
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
    
    # Eliminar duplicados
    df = pd.DataFrame(datos_pdf)
    if not df.empty:
        df = df.drop_duplicates(subset=['Tracking'], keep='first')
    return df

def procesar_formulario_pdf(archivo):
    try:
        with pdfplumber.open(archivo) as pdf:
            contenido_completo = "\n".join(page.extract_text(x_tolerance=1) or "" for page in pdf.pages)
            lineas = contenido_completo.splitlines()
    except Exception as e:
        st.error(f"Error leyendo formulario {archivo.name}: {e}")
        return []
    
    # Extracción de datos básicos
    fmm_formulario, usuario, pais_destino, factura_a_asignar = "", "", "", ""
    
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
            match = re.search(r'22\.\s*País Destino:\s*(\d+\s+[A-Z\s]+)', linea)
            if match:
                pais_destino = re.sub(r'^\d+\s*', '', match.group(1).strip()).strip()
    
    # Extraer facturas que NO mencionen servicio
    en_anexos = False
    facturas_validas = []
    
    for i, linea in enumerate(lineas):
        if "DETALLE DE LOS ANEXOS" in linea:
            en_anexos = True
        
        if en_anexos and "FACTURA COMERCIAL" in linea:
            tiene_servicio = any('servicio' in lineas[j].lower() for j in range(max(0, i-2), min(len(lineas), i+3)))
            if not tiene_servicio:
                match = re.search(r'\b(ZFFE\d+|ZFFV\d+)\b', linea)
                if match:
                    facturas_validas.append(match.group(0))
    
    factura_a_asignar = ", ".join(facturas_validas) if facturas_validas else ""
    
    # Extraer guías de la sección de anexos
    datos_finales = []
    en_seccion_anexos = False
    
    for linea in lineas:
        if "DETALLE DE LOS ANEXOS" in linea:
            en_seccion_anexos = True
            continue
        
        if en_seccion_anexos:
            # Buscar tracking numbers específicamente en líneas de guías
            if re.search(r'127\s+GUIAS DE TRAFICO POSTAL', linea) or re.search(r'127\s+GUAS DE TRAFICO POSTAL', linea):
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
                        # Buscar fecha en la misma línea
                        fecha_match = re.search(r'(\d{4}/\d{2}/\d{2})', linea)
                        fecha = ""
                        if fecha_match:
                            try:
                                fecha = datetime.strptime(fecha_match.group(1), '%Y/%m/%d').strftime('%Y-%m-%d')
                            except:
                                fecha = ""
                        
                        datos_finales.append({
                            "Tracking": tracking,
                            "Fecha_FMM": fecha,
                            "Pais_Destino_FMM": pais_destino,
                            "FMM_Formulario": fmm_formulario,
                            "Remitente_Usuario_FMM": usuario,
                            "Facturas_FMM": factura_a_asignar
                        })
                        break
    
    # Eliminar duplicados
    unique_data = []
    seen_trackings = set()
    for item in datos_finales:
        if item['Tracking'] not in seen_trackings:
            unique_data.append(item)
            seen_trackings.add(item['Tracking'])
    
    return unique_data

# --- INTERFAZ STREAMLIT MEJORADA ---
def main():
    st.title("📦 Sistema de Conciliación de Guías Aéreas")
    st.markdown("---")
    
    # Inicializar session state
    if 'resultados' not in st.session_state:
        st.session_state.resultados = None
    
    # Sidebar para carga de archivos
    with st.sidebar:
        st.header("📂 Cargar Archivos")
        archivos_guias = st.file_uploader("Guías PDF (FedEx, UPS, DHL)", type="pdf", accept_multiple_files=True)
        archivos_formularios = st.file_uploader("Formularios PDF", type="pdf", accept_multiple_files=True)
        
        if st.button("🔄 Procesar Conciliación", type="primary"):
            if archivos_guias and archivos_formularios:
                with st.spinner("Procesando archivos..."):
                    try:
                        # Procesar guías
                        df_guias = procesar_archivos_guias_pdf(archivos_guias)
                        st.info(f"Guías procesadas: {len(df_guias)}")
                        
                        # Procesar formularios
                        todos_datos_formularios = []
                        for archivo in archivos_formularios:
                            datos = procesar_formulario_pdf(archivo)
                            todos_datos_formularios.extend(datos)
                        
                        df_formularios = pd.DataFrame(todos_datos_formularios)
                        st.info(f"Formularios procesados: {len(df_formularios)}")
                        
                        if not df_guias.empty and not df_formularios.empty:
                            # Conciliación con verificación real de datos
                            df_conciliado = pd.merge(
                                df_guias, 
                                df_formularios, 
                                on='Tracking', 
                                how='outer', 
                                indicator=True,
                                suffixes=('_Guia', '_FMM')
                            )
                            
                            # Función de análisis REAL
                            def analizar_fila(row):
                                if row['_merge'] == 'left_only': 
                                    return '❌ SOLO EN GUÍA'
                                if row['_merge'] == 'right_only': 
                                    return '❌ SOLO EN FMM'
                                
                                # Verificar coincidencias reales
                                diferencias = []
                                if str(row.get('Fecha_Guia', '')) != str(row.get('Fecha_FMM', '')):
                                    diferencias.append("Fecha")
                                if str(row.get('Pais_Destino_Guia', '')) != str(row.get('Pais_Destino_FMM', '')):
                                    diferencias.append("País")
                                if str(row.get('FMM_Guia', '')) != str(row.get('FMM_Formulario', '')):
                                    diferencias.append("FMM")
                                if str(row.get('Facturas_Guia', '')) != str(row.get('Facturas_FMM', '')):
                                    diferencias.append("Facturas")
                                
                                if not diferencias:
                                    return '✅ OK'
                                else:
                                    return f'⚠️ Diferencias: {", ".join(diferencias)}'
                            
                            df_conciliado['Estado_Conciliacion'] = df_conciliado.apply(analizar_fila, axis=1)
                            
                            st.session_state.resultados = df_conciliado
                            st.success("✅ Conciliación completada")
                            
                        else:
                            st.warning("No se pudieron extraer datos suficientes para comparar")
                            
                    except Exception as e:
                        st.error(f"Error en procesamiento: {str(e)}")
            else:
                st.warning("⚠️ Debes cargar ambos tipos de archivos")
    
    # Botón de limpieza
    if st.sidebar.button("🗑️ Limpiar Resultados", type="secondary"):
        st.session_state.resultados = None
        st.rerun()
    
    # Mostrar resultados
    if st.session_state.resultados is not None:
        st.header("📊 Resultados de Conciliación")
        
        # Filtrar columnas para mejor visualización
        columnas_a_mostrar = [
            'Tracking', 'Fecha_Guia', 'Fecha_FMM', 'Pais_Destino_Guia', 
            'Pais_Destino_FMM', 'Peso_Neto_Guia', 'FMM_Guia', 'FMM_Formulario',
            'Facturas_Guia', 'Facturas_FMM', 'Estado_Conciliacion'
        ]
        
        columnas_existentes = [col for col in columnas_a_mostrar if col in st.session_state.resultados.columns]
        df_mostrar = st.session_state.resultados[columnas_existentes]
        
        st.dataframe(df_mostrar, use_container_width=True)
        
        # Estadísticas reales
        st.subheader("📈 Resumen de Conciliación")
        if 'Estado_Conciliacion' in st.session_state.resultados.columns:
            conteo_estados = st.session_state.resultados['Estado_Conciliacion'].value_counts()
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total", len(st.session_state.resultados))
            col2.metric("✅ OK", conteo_estados.get('✅ OK', 0))
            col3.metric("❌ Solo Guía", conteo_estados.get('❌ SOLO EN GUÍA', 0))
            col4.metric("❌ Solo FMM", conteo_estados.get('❌ SOLO EN FMM', 0))
            
            if '⚠️ Diferencias:' in conteo_estados:
                st.metric("⚠️ Con Diferencias", conteo_estados['⚠️ Diferencias:'])
        
        # Botones de exportación
        st.subheader("💾 Exportar Resultados")
        csv = st.session_state.resultados.to_csv(index=False)
        st.download_button("📥 Descargar CSV", csv, "conciliacion.csv", "text/csv")
    
    # Información de uso
    with st.expander("ℹ️ Instrucciones de uso"):
        st.markdown("""
        **⚠️ IMPORTANTE:** Esta versión incluye verificaciones reales de datos:
        - Compara fechas, países, FMM y facturas
        - Muestra diferencias reales, no solo el merge
        - Elimina duplicados automáticamente
        """)

if __name__ == "__main__":
    main()
