import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime

# --- 📂 CONFIGURACIÓN PRINCIPAL ---
CARPETA_DATOS = "E:/Users/Lenovo/Desktop/PYTHON/Guis Aereas Solideo/FMM844217"
ARCHIVO_SALIDA_COMPARACION = "Resultado_Verificación_Guias.xlsx"

# ==============================================================================
# --- SECCIÓN 1: PROCESAMIENTO DE GUÍAS PDF (Tu lógica) ---
# ==============================================================================
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
def pdf_extraer_pn(texto_guia):
    m = re.search(r"PN[:\s]*([\d.,]+)", texto_guia)
    return m.group(1).replace(",", ".") if m else ""
def pdf_extraer_fmm(texto_guia):
    m = re.search(r"(FMM\d+)", texto_guia); return m.group(1) if m else ""
def pdf_extraer_remitente(texto_guia):
    m = re.search(r"(SOLIDEO\s*S\.?A?\.?S\.?)", texto_guia, re.IGNORECASE)
    if m: return "SOLIDEO S.A.S."
    return ""
def pdf_extraer_fecha(texto_guia):
    patron = r"(?i)(?:SHIP DATE:\s*|Date\s)(\d{2}[A-Z]{3}\d{2})|(\d{4}-\d{2}-\d{2})|Date\s*(\d{2}\s[A-Z]{3}\s\d{4})"
    m = re.search(patron, texto_guia)
    if not m: return ""
    fecha_str = next((g for g in m.groups() if g is not None), None)
    if not fecha_str: return ""
    fecha_str_limpia = " ".join(fecha_str.split()).title()
    formatos = [('%d%b%y', '%Y-%m-%d'), ('%Y-%m-%d', '%Y-%m-%d'), ('%d %b %Y', '%Y-%m-%d')]
    for fmt_in, fmt_out in formatos:
        try: return datetime.strptime(fecha_str_limpia, fmt_in).strftime(fmt_out)
        except ValueError: continue
    return fecha_str
def procesar_archivos_pdf(carpeta):
    print("\n--- INICIANDO PROCESO DE GUÍAS PDF ---")
    archivos_pdf = [f for f in os.listdir(carpeta) if f.lower().endswith(".pdf")]
    if not archivos_pdf:
        print("   -> No se encontraron archivos PDF.")
        return pd.DataFrame()
    print(f"   -> Encontrados {len(archivos_pdf)} archivos PDF. Procesando...")
    datos_pdf = []
    for archivo in archivos_pdf:
        try:
            ruta_completa = os.path.join(carpeta, archivo)
            with pdfplumber.open(ruta_completa) as pdf:
                texto_completo = "\n".join(page.extract_text(x_tolerance=1) or "" for page in pdf.pages)
            patrones_inicio = [r"ORIGIN ID:", r"EXPRESS WORLDWIDE", r"UPS WORLDWIDE SERVICE"]
            indices = [m.start() for m in re.finditer("|".join(patrones_inicio), texto_completo)]
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
                        "Fecha_GUÍA": pdf_extraer_fecha(bloque),
                        "Pais_Destino_GUÍA": pdf_extraer_pais_destino(bloque),
                        "FMM_GUÍA": pdf_extraer_fmm(bloque),
                        "Peso_Neto_GUÍA": pdf_extraer_pn(bloque),
                        "Remitente_Usuario_GUÍA": pdf_extraer_remitente(bloque),
                        "Facturas_GUÍA": pdf_extraer_facturas(bloque, operador, ultimo_ref_dhl)
                    })
        except Exception as e:
            print(f"   -> ⚠️  Error procesando PDF '{archivo}': {e}")
    print(f"   -> Se extrajeron datos de {len(datos_pdf)} guías desde los PDFs.")
    return pd.DataFrame(datos_pdf)

# ==============================================================================
# --- SECCIÓN 2: PROCESAMIENTO DE FORMULARIOS CSV (Tu lógica) ---
# ==============================================================================
def procesar_reporte_csv(ruta_csv):
    try:
        with open(ruta_csv, 'r', encoding='latin-1') as f:
            lineas = f.readlines()
    except Exception: return []
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
    
    datos_finales = []
    en_anexos = False
    for linea in lineas:
        linea_strip = linea.strip()
        if "DETALLE DE LOS ANEXOS" in linea_strip: en_anexos = True
        if en_anexos and linea_strip.startswith('127,'):
            guia = re.search(r'\b(8837\d{8})\b', linea_strip)
            fecha = re.search(r'(\d{4}/\d{2}/\d{2})', linea_strip)
            if guia:
                datos_finales.append({
                    "Tracking": guia.group(0),
                    "Fecha_FMM": datetime.strptime(fecha.group(0), '%Y/%m/%d').strftime('%Y-%m-%d') if fecha else "",
                    "Pais_Destino_FMM": pais_destino,
                    "FMM_FMM": formulario,
                    "Remitente_Usuario_FMM": usuario,
                    "Facturas_FMM": factura_a_asignar
                })
    return datos_finales
def procesar_archivos_csv(carpeta):
    print("\n--- INICIANDO PROCESO DE FORMULARIOS CSV ---")
    archivos_csv = [f for f in os.listdir(carpeta) if f.lower().endswith(".csv")]
    if not archivos_csv:
        print("   -> No se encontraron archivos CSV.")
        return pd.DataFrame()
    print(f"   -> Encontrados {len(archivos_csv)} archivos CSV. Procesando...")
    todos_los_datos = []
    for nombre_archivo in archivos_csv:
        todos_los_datos.extend(procesar_reporte_csv(os.path.join(carpeta, nombre_archivo)))
    print(f"   -> Se extrajeron datos de {len(todos_los_datos)} guías desde los CSVs.")
    return pd.DataFrame(todos_los_datos)

# ==============================================================================
# --- SECCIÓN 3: LÓGICA DE CONCILIACIÓN FINAL (Con tus ajustes) ---
# ==============================================================================
def main():
    df_guias = procesar_archivos_pdf(CARPETA_DATOS)
    df_formulario = procesar_archivos_csv(CARPETA_DATOS)

    if df_guias.empty or df_formulario.empty:
        print("\n❌ No se pueden comparar los datos."); return

    df_guias['Tracking'] = df_guias['Tracking'].astype(str)
    df_formulario['Tracking'] = df_formulario['Tracking'].astype(str)
    
    print("\n--- INICIANDO CONCILIACIÓN DE DATOS ---")
    df_conciliado = pd.merge(
        df_guias, df_formulario, on='Tracking', how='outer', indicator=True
    )
    
    # --- Etapa de Normalización ---
    mapa_paises = {"US": "UNITED STATES OF AMERICA", "ESTADOS UNIDOS": "UNITED STATES OF AMERICA", "JP": "JAPAN"}
    df_conciliado['Pais_Normalizado_GUÍA'] = df_conciliado['Pais_Destino_GUÍA'].str.upper().map(mapa_paises).fillna(df_conciliado['Pais_Destino_GUÍA'].str.upper())
    df_conciliado['Pais_Normalizado_FMM'] = df_conciliado['Pais_Destino_FMM'].str.upper().map(mapa_paises).fillna(df_conciliado['Pais_Destino_FMM'].str.upper())
    
    # AJUSTE: Normalización de FMM para quitar el prefijo y comparar solo números
    df_conciliado['FMM_Normalizado_GUÍA'] = df_conciliado['FMM_GUÍA'].str.replace('FMM', '', regex=False)
    
    def analizar_fila(row):
        origen = row['_merge']
        if origen == 'left_only':
            return '❌ SOLO EN GUÍA'
        if origen == 'right_only':
            return '❌ SOLO EN FMM'
        
        diferencias = []
        if str(row['Fecha_GUÍA']) != str(row['Fecha_FMM']):
            diferencias.append("Fecha")
        if str(row['Pais_Normalizado_GUÍA']) != str(row['Pais_Normalizado_FMM']):
            diferencias.append("País")
        # AJUSTE: Comparar las columnas normalizadas de FMM
        if str(row['FMM_Normalizado_GUÍA']) != str(row['FMM_FMM']):
            diferencias.append("FMM")
        if str(row['Remitente_Usuario_GUÍA']) != str(row['Remitente_Usuario_FMM']):
            diferencias.append("Remitente")
        if str(row['Facturas_GUÍA']) != str(row['Facturas_FMM']):
            diferencias.append("Facturas")
        
        if not diferencias:
            return '✅ OK'
        else:
            return '❌ Diferencia en: ' + ", ".join(diferencias)
    
    df_conciliado['Estado_Conciliacion'] = df_conciliado.apply(analizar_fila, axis=1)
    
    # AJUSTE: Eliminamos la columna _merge y las columnas de normalización temporales
    df_conciliado.drop(columns=['_merge', 'Pais_Normalizado_GUÍA', 'Pais_Normalizado_FMM', 'FMM_Normalizado_GUÍA'], inplace=True, errors='ignore')
    
    # AJUSTE: El orden de las columnas se mantiene como lo tenías, eliminando "Origen"
    columnas_ordenadas = [
        'Estado_Conciliacion', 'Tracking',
        'Fecha_GUÍA', 'Fecha_FMM',
        'Pais_Destino_GUÍA', 'Pais_Destino_FMM',
        'FMM_GUÍA', 'FMM_FMM',
        'Peso_Neto_GUÍA',
        'Remitente_Usuario_GUÍA', 'Remitente_Usuario_FMM',
        'Facturas_GUÍA', 'Facturas_FMM',
    ]
    columnas_existentes = [col for col in columnas_ordenadas if col in df_conciliado.columns]
    df_conciliado = df_conciliado[columnas_existentes]
    
    print("   -> Conciliación completada.")
    
    # --- Resumen en Consola ---
    print("\n--- 📊 RESUMEN DE CONCILIACIÓN ---")
    total_pdf = len(df_guias)
    total_csv = len(df_formulario)
    conteo_estados = df_conciliado['Estado_Conciliacion'].value_counts()
    
    print(f"Total de guías en PDFs: {total_pdf}")
    print(f"Total de guías en FMMs: {total_csv}")
    print("-" * 35)
    print(f"✅ Coincidencias perfectas (OK): {conteo_estados.get('✅ OK', 0)}")
    print(f"❌ Con diferencias: {sum(1 for estado in conteo_estados.index if estado.startswith('❌ Diferencia'))}")
    # AJUSTE: Se usan los nombres de estado que definiste
    print(f"   - SOLO EN GUIAS: {conteo_estados.get('❌ SOLO EN GUÍA', 0)}")
    print(f"   - SOLO EN FMM: {conteo_estados.get('❌ SOLO EN FMM', 0)}")
    
    try:
        df_conciliado.to_excel(ARCHIVO_SALIDA_COMPARACION, index=False)
        print(f"\n✅ ¡Proceso finalizado! El reporte ejecutivo se ha guardado en:")
        print(f"   '{ARCHIVO_SALIDA_COMPARACION}'")
    except Exception as e:
        print(f"\n❌ ERROR AL GUARDAR EL ARCHIVO EXCEL: {e}")
        print(f"   Causa probable: ¿Tienes el archivo '{ARCHIVO_SALIDA_COMPARACION}' abierto?")

if __name__ == "__main__":
    main()
