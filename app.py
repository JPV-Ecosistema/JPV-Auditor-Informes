# ==============================================================================
# ### --- BLOQUE 0: IMPORTACIONES Y CONFIGURACIÓN BASE --- ###
# ==============================================================================
import streamlit as st
import pandas as pd
import pdfplumber
import docx
from docx.shared import Pt, RGBColor
import re
import io
from datetime import datetime

st.set_page_config(page_title="Auditor JPV - Revisión Individual", layout="wide")

if 'df_maestro' not in st.session_state:
    st.session_state.df_maestro = None

st.title("🔎 Auditor Individual de Informes")
st.markdown("Revisión técnica exhaustiva de informes de liquidación contra el Reporte de Acciones.")

# ==============================================================================
# ### --- BLOQUE 1: FUNCIONES DE EXTRACCIÓN DE TEXTO BASE --- ###
# ==============================================================================
def extraer_texto_pdf(archivo):
    texto = ""
    try:
        with pdfplumber.open(archivo) as pdf:
            for pagina in pdf.pages:
                txt = pagina.extract_text(layout=True)
                if txt: texto += txt + "\n"
    except Exception as e:
        texto = f"Error al leer PDF: {e}"
    return texto

def extraer_texto_docx(archivo):
    texto = ""
    try:
        doc = docx.Document(archivo)
        for para in doc.paragraphs:
            texto += para.text + "\n"
        for table in doc.tables:
            for row in table.rows:
                texto += " | ".join([cell.text.replace("\n", " ").strip() for cell in row.cells]) + "\n"
    except Exception as e:
        texto = f"Error al leer DOCX: {e}"
    return texto

# ==============================================================================
# ### --- BLOQUE 2: MOTOR LÓGICO Y EXTRACCIÓN DE DATOS (REGEX) --- ###
# ==============================================================================
def limpiar_monto(texto_monto):
    if not texto_monto or pd.isna(texto_monto): return 0.0
    if isinstance(texto_monto, (int, float)): return float(texto_monto)
    
    # Extrae solo la parte numérica con puntos y comas
    match = re.search(r'[\d\.,]+', str(texto_monto))
    if not match: return 0.0
    limpio = match.group(0)
    
    # Formato Chileno a Float computacional: 1.500.000,50 -> 1500000.50
    limpio = limpio.replace('.', '').replace(',', '.')
    try:
        return float(limpio)
    except:
        return 0.0

def extraer_datos_informe(texto):
    datos = {
        "Liquidacion": None, "Poliza": None, "Compania": None, "Ajustador": None,
        "Fecha_Siniestro": None, "Fecha_Denuncia": None,
        "Divisa": None, "Perdida_Bruta": 0.0, "Texto_Monto_Crudo": ""
    }
    
    match_liq = re.search(r'(?:LIQUIDACI[OÓ]N Nº|Ref\. JPV\s*:)\s*(\d+)', texto, re.IGNORECASE)
    if match_liq: datos["Liquidacion"] = match_liq.group(1).strip()
        
    match_pol = re.search(r'(?:Nº Póliza|Póliza Nº|Póliza número)[\s:]*([A-Za-z0-9-]+)', texto, re.IGNORECASE)
    if match_pol: datos["Poliza"] = match_pol.group(1).strip()
    
    match_comp = re.search(r'(?:ASEGURADOR|COMPAÑ[IÍ]A DE SEGUROS|ASEGURADORA|COMPAÑ[IÍ]A)[\s:]*([^\n|]+)', texto, re.IGNORECASE)
    if match_comp: datos["Compania"] = match_comp.group(1).strip().split('|')[0].strip()

    match_ajust = re.search(r'(?:Ajustador a cargo|Ajustador senior|Ajustador|Firmante)[\s:]*([A-Za-z\s\.]+)', texto, re.IGNORECASE)
    if match_ajust: datos["Ajustador"] = match_ajust.group(1).strip()

    match_fsin = re.search(r'Fecha de Siniestro[\s:]*([\d\-]+)', texto, re.IGNORECASE)
    if match_fsin: datos["Fecha_Siniestro"] = match_fsin.group(1).strip()
        
    match_fden = re.search(r'Fecha Denuncia[\s:]*([\d\-]+)', texto, re.IGNORECASE)
    if match_fden: datos["Fecha_Denuncia"] = match_fden.group(1).strip()

    # Corrección de Montos: Buscar la última aparición explícita para evitar sumas duplicadas
    matches_bruta = re.findall(r'(?:Pérdida Bruta|Pérdida Probable|Reserva Determinada|Pérdida Estimada)[^\d]*([\d\.,]+)', texto, re.IGNORECASE)
    if matches_bruta:
        datos["Texto_Monto_Crudo"] = matches_bruta[-1] # Guardamos el texto exacto capturado para diagnóstico
        datos["Perdida_Bruta"] = limpiar_monto(matches_bruta[-1])
    
    if re.search(r'\bUF\b', texto, re.IGNORECASE): datos["Divisa"] = "UF"
    elif re.search(r'\b(US\$|USD|D[OÓ]LARES|US \$)\b', texto, re.IGNORECASE): datos["Divisa"] = "US$"
    elif re.search(r'\b(PESOS|CLP|\$)\b', texto, re.IGNORECASE): datos["Divisa"] = "PESOS"
        
    return datos

# ==============================================================================
# ### --- BLOQUE 3: GESTIÓN DE LA BASE MAESTRA (SIDEBAR) --- ###
# ==============================================================================
with st.sidebar:
    st.header("⚙️ Configuración")
    archivo_excel = st.file_uploader("Cargar/Actualizar Reporte de Acciones", type=["xlsx"])
    
    if archivo_excel:
        for i in range(15): 
            try:
                df_temp = pd.read_excel(archivo_excel, skiprows=i)
                df_temp.columns = [str(c).strip() for c in df_temp.columns]
                if 'Número de caso' in df_temp.columns:
                    df_temp['Número de caso'] = df_temp['Número de caso'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                    st.session_state.df_maestro = df_temp.dropna(how='all', axis=0)
                    st.success("✅ Reporte de Acciones cargado.")
                    break
            except: continue

    if st.session_state.df_maestro is not None:
        st.write(f"📊 **Casos en Base:** {len(st.session_state.df_maestro)}")
        if st.button("Limpiar Memoria"):
            st.session_state.df_maestro = None
            st.rerun()

# ==============================================================================
# ### --- BLOQUE 4: GENERADOR DE CERTIFICADO Y REPORTES (WORD) --- ###
# ==============================================================================
def generar_word_individual(resultado):
    doc = docx.Document()
    doc.add_heading('CERTIFICADO DE AUDITORÍA JPV', 0)
    doc.add_paragraph(f"Fecha: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
    
    doc.add_heading('1. DATOS DEL INFORME', level=1)
    doc.add_paragraph(f"Liquidación Nº: {resultado['N° Caso']}\nArchivo: {resultado['Documento']}")

    estado = "❌ OBSERVADO" if resultado['Detalles_Criticos'] else "✅ APROBADO"
    doc.add_heading(f'2. RESULTADO DETALLADO: {estado}', level=1)

    sections = [
        ("Póliza de Seguros", resultado['Detalles_Poliza']),
        ("Compañía de Seguros", resultado['Detalles_Compania']),
        ("Firma / Ajustador Senior", resultado['Detalles_Firmas']),
        ("Fechas (Siniestro y Denuncia)", resultado['Detalles_Fechas']),
        ("Financiera (Pérdida Bruta y Divisa)", resultado['Detalles_Monto'])
    ]

    for titulo, errores in sections:
        doc.add_heading(titulo, level=2)
        if not errores:
            doc.add_paragraph("✅ Sin observaciones detectadas. Coincide con sistema.")
        else:
            for err in errores:
                p = doc.add_paragraph(style='List Bullet')
                run = p.add_run(err)
                run.font.color.rgb = RGBColor(200, 0, 0)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer

# ==============================================================================
# ### --- BLOQUE 5: LÓGICA DE CRUCE Y FLUJO PRINCIPAL DE INTERFAZ --- ###
# ==============================================================================
if st.session_state.df_maestro is not None:
    st.subheader("📄 Subir Informe para Revisión Individual")
    archivo_informe = st.file_uploader("Sube el Informe (PDF o DOCX)", type=["pdf", "docx"])

    if archivo_informe:
        texto = extraer_texto_pdf(archivo_informe) if archivo_informe.name.lower().endswith('.pdf') else extraer_texto_docx(archivo_informe)
        datos_doc = extraer_datos_informe(texto)
        
        if not datos_doc["Liquidacion"]:
            st.error("❌ No se detectó el N° de Liquidación en el documento.")
        else:
            df = st.session_state.df_maestro
            fila = df[df['Número de caso'] == datos_doc["Liquidacion"]]
            
            if fila.empty:
                st.warning(f"⚠️ El caso {datos_doc['Liquidacion']} no figura en el Reporte de Acciones.")
            else:
                fila = fila.iloc[0]
                
                # Arrays separados rigurosamente para cada análisis
                detalles_poliza = []
                detalles_compania = []
                detalles_firmas = []
                detalles_fechas = []
                detalles_monto = []

                # 1. Cruce de Póliza Separado
                poliza_sist = str(fila.get('Póliza de seguros', '')).strip()
                if poliza_sist.upper() not in str(datos_doc["Poliza"]).upper() and poliza_sist != "nan" and poliza_sist != "":
                    detalles_poliza.append(f"Póliza: Sist({poliza_sist}) vs Doc({datos_doc['Poliza']})")
                
                # 2. Cruce de Asegurador Separado
                comp_sist = str(fila.get('Compañía de seguros', '')).strip().upper()
                if datos_doc["Compania"] and comp_sist not in datos_doc["Compania"].upper():
                    detalles_compania.append(f"Asegurador: Sist({comp_sist}) vs Doc({datos_doc['Compania']})")

                # 3. Cruce de Firmas / Ajustador Separado
                ajust_sist = str(fila.get('Ajustador senior', '')).strip().upper()
                if datos_doc["Ajustador"] and ajust_sist not in datos_doc["Ajustador"].upper() and ajust_sist != "NAN":
                    detalles_firmas.append(f"Ajustador: Sist({ajust_sist}) vs Doc({datos_doc['Ajustador']})")

                # 4. Cruce de Fechas Exactas Separado
                f_ocurr = str(fila.get('Fecha de ocurrencia', '')).strip()
                if datos_doc["Fecha_Siniestro"] and datos_doc["Fecha_Siniestro"][:10] not in f_ocurr:
                    detalles_fechas.append(f"Fecha Ocurrencia: Sist({f_ocurr}) vs Doc({datos_doc['Fecha_Siniestro']})")

                f_denun = str(fila.get('Fecha de denuncio', '')).strip()
                if datos_doc["Fecha_Denuncia"] and datos_doc["Fecha_Denuncia"][:10] not in f_denun:
                    detalles_fechas.append(f"Fecha Denuncio: Sist({f_denun}) vs Doc({datos_doc['Fecha_Denuncia']})")

                # 5. Cruce Financiero y Monto Separado
                bruta_sist = limpiar_monto(fila.get('Perdida bruta (en moneda del caso)', 0))
                if abs(datos_doc["Perdida_Bruta"] - bruta_sist) > 1.0:
                    detalles_monto.append(f"Monto Bruto: Sist({bruta_sist:,.2f}) vs Doc({datos_doc['Perdida_Bruta']:,.2f}) [Texto extraído: '{datos_doc['Texto_Monto_Crudo']}']")

                div_sist = str(fila.get('Divisa', '')).strip().upper()
                if datos_doc["Divisa"] and div_sist != datos_doc["Divisa"] and div_sist != "NAN":
                    detalles_monto.append(f"Divisa: Sist({div_sist}) vs Doc({datos_doc['Divisa']})")

                # Consolidación de Resultados
                todos_los_errores = detalles_poliza + detalles_compania + detalles_firmas + detalles_fechas + detalles_monto
                
                res_final = {
                    "Documento": archivo_informe.name,
                    "N° Caso": datos_doc["Liquidacion"],
                    "Detalles_Poliza": detalles_poliza,
                    "Detalles_Compania": detalles_compania,
                    "Detalles_Firmas": detalles_firmas,
                    "Detalles_Fechas": detalles_fechas,
                    "Detalles_Monto": detalles_monto,
                    "Detalles_Criticos": todos_los_errores
                }

                estado_visual = "❌ Con Observaciones" if todos_los_errores else "✅ Aprobado"
                st.success(f"Auditoría Finalizada. Estado: {estado_visual}")
                
                # Métricas visuales desglosadas al 100%
                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Póliza", "❌" if detalles_poliza else "✅")
                c2.metric("Compañía", "❌" if detalles_compania else "✅")
                c3.metric("Firma", "❌" if detalles_firmas else "✅")
                c4.metric("Fechas", "❌" if detalles_fechas else "✅")
                c5.metric("Monto/Divisa", "❌" if detalles_monto else "✅")

                if todos_los_errores:
                    st.error("Observaciones encontradas:")
                    for d in todos_los_errores: st.write(f"• {d}")

                # Botón Word
                word_buf = generar_word_individual(res_final)
                st.download_button("📄 Descargar Certificado Word", word_buf.getvalue(), f"Auditoria_{datos_doc['Liquidacion']}.docx")
else:
    st.info("👈 Sube primero el Reporte de Acciones en la barra lateral.")
