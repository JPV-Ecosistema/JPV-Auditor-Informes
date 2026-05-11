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

# Estado de sesión (Persistencia de la Base Maestra)
if 'df_maestro' not in st.session_state:
    st.session_state.df_maestro = None

st.title("🔎 Auditor Individual de Informes")
st.markdown("Revisión técnica de informes de liquidación contra el Reporte de Acciones.")

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
    if not texto_monto: return 0.0
    limpio = re.sub(r'[^\d,\.-]', '', str(texto_monto))
    limpio = limpio.replace('.', '').replace(',', '.')
    try:
        return float(limpio)
    except:
        return 0.0

def extraer_datos_informe(texto):
    datos = {
        "Liquidacion": None, "Poliza": None, "Compania": None, "Ajustador": None,
        "Fecha_Siniestro": None, "Fecha_Denuncia": None,
        "Divisa": None, "Perdida_Bruta": 0.0
    }
    
    # 1. Identificación del Caso (Llave)
    match_liq = re.search(r'(?:LIQUIDACI[OÓ]N Nº|Ref\. JPV\s*:)\s*(\d+)', texto, re.IGNORECASE)
    if match_liq: datos["Liquidacion"] = match_liq.group(1).strip()
        
    # 2. Póliza
    match_pol = re.search(r'(?:Nº Póliza|Póliza Nº|Póliza número)[\s:]*([A-Za-z0-9-]+)', texto, re.IGNORECASE)
    if match_pol: datos["Poliza"] = match_pol.group(1).strip()
    
    # 3. Asegurador
    match_comp = re.search(r'(?:ASEGURADOR|COMPAÑ[IÍ]A DE SEGUROS|ASEGURADORA|COMPAÑ[IÍ]A)[\s:]*([^\n|]+)', texto, re.IGNORECASE)
    if match_comp: 
        datos["Compania"] = match_comp.group(1).strip().split('|')[0].strip()

    # 4. Ajustador / Firmante
    match_ajust = re.search(r'(?:Ajustador a cargo|Ajustador senior)[\s:]*([^\n|]+)', texto, re.IGNORECASE)
    if match_ajust: datos["Ajustador"] = match_ajust.group(1).strip()

    # 5. Fechas
    match_fsin = re.search(r'Fecha de Siniestro[\s:]*([\d\-]+)', texto, re.IGNORECASE)
    if match_fsin: datos["Fecha_Siniestro"] = match_fsin.group(1).strip()
        
    match_fden = re.search(r'Fecha Denuncia[\s:]*([\d\-]+)', texto, re.IGNORECASE)
    if match_fden: datos["Fecha_Denuncia"] = match_fden.group(1).strip()

    # 6. Pérdida Bruta (Equivalencia con Pérdida Probable)
    patrones_bruta = [
        r'(?:Pérdida Bruta|Pérdida Probable|Reserva Determinada|Pérdida Estimada)[^\d]*([\d\.,]+)',
        r'(?:BF|BI)[^\d]*([\d\.,]+)'
    ]
    
    total_bruta = 0.0
    match_total_bruta = re.findall(r'(?:Total Pérdida Bruta|Total Pérdida Probable)[^\d]*([\d\.,]+)', texto, re.IGNORECASE)
    if match_total_bruta:
        total_bruta = limpiar_monto(match_total_bruta[-1])
    else:
        for patron in patrones_bruta:
            matches = re.findall(patron, texto, re.IGNORECASE)
            if matches:
                for m in matches:
                    val = limpiar_monto(m)
                    if val > total_bruta: 
                        total_bruta += val
    
    datos["Perdida_Bruta"] = total_bruta
    
    # 7. Divisa (Detección Robusta)
    if re.search(r'\bUF\b', texto, re.IGNORECASE):
        datos["Divisa"] = "UF"
    elif re.search(r'\b(US\$|USD|D[OÓ]LARES|US \$)\b', texto, re.IGNORECASE):
        datos["Divisa"] = "US$"
    elif re.search(r'\b(PESOS|CLP|\$)\b', texto, re.IGNORECASE):
        datos["Divisa"] = "PESOS"
        
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
    doc.add_heading(f'2. RESULTADO: {estado}', level=1)

    sections = [
        ("Póliza y Compañía", resultado['Detalles_Forma']),
        ("Firmas / Ajustador", resultado['Detalles_Firmas']),
        ("Fechas (Siniestro y Denuncia)", resultado['Detalles_Fechas']),
        ("Financiera y Moneda", resultado['Detalles_Bruta'])
    ]

    for titulo, errores in sections:
        doc.add_heading(titulo, level=2)
        if not errores:
            doc.add_paragraph("✅ Sin observaciones detectadas.")
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
                detalles_forma, detalles_firmas, detalles_fechas, detalles_bruta = [], [], [], []

                # Cruce de Póliza
                poliza_sist = str(fila.get('Póliza de seguros', '')).strip()
                if poliza_sist.upper() not in str(datos_doc["Poliza"]).upper() and poliza_sist != "nan" and poliza_sist != "":
                    detalles_forma.append(f"Póliza: Sist({poliza_sist}) vs Doc({datos_doc['Poliza']})")
                
                # Cruce de Asegurador
                comp_sist = str(fila.get('Compañía de seguros', '')).strip().upper()
                if datos_doc["Compania"] and comp_sist not in datos_doc["Compania"].upper():
                    detalles_forma.append(f"Asegurador: Sist({comp_sist}) vs Doc({datos_doc['Compania']})")

                # Cruce de Firmas / Ajustador
                ajust_sist = str(fila.get('Ajustador senior', '')).strip().upper()
                if datos_doc["Ajustador"] and ajust_sist not in datos_doc["Ajustador"].upper() and ajust_sist != "NAN":
                    detalles_firmas.append(f"Firma Ajustador: Sist({ajust_sist}) vs Doc({datos_doc['Ajustador']})")

                # Cruce de Fechas Exactas
                f_ocurr = str(fila.get('Fecha de ocurrencia', '')).strip()
                if datos_doc["Fecha_Siniestro"] and datos_doc["Fecha_Siniestro"][:10] not in f_ocurr:
                    detalles_fechas.append(f"Ocurrencia: Sist({f_ocurr}) vs Doc({datos_doc['Fecha_Siniestro']})")

                f_denun = str(fila.get('Fecha de denuncio', '')).strip()
                if datos_doc["Fecha_Denuncia"] and datos_doc["Fecha_Denuncia"][:10] not in f_denun:
                    detalles_fechas.append(f"Denuncio: Sist({f_denun}) vs Doc({datos_doc['Fecha_Denuncia']})")

                # Cruce Financiero (Pérdida Bruta)
                bruta_sist = limpiar_monto(fila.get('Perdida bruta (en moneda del caso)', 0))
                if abs(datos_doc["Perdida_Bruta"] - bruta_sist) > 1.0:
                    detalles_bruta.append(f"Monto Bruto: Sist({bruta_sist:,.2f}) vs Doc({datos_doc['Perdida_Bruta']:,.2f})")

                div_sist = str(fila.get('Divisa', '')).strip().upper()
                if datos_doc["Divisa"] and div_sist != datos_doc["Divisa"]:
                    detalles_bruta.append(f"Divisa: Sist({div_sist}) vs Doc({datos_doc['Divisa']})")

                # Resultados
                res_final = {
                    "Documento": archivo_informe.name,
                    "N° Caso": datos_doc["Liquidacion"],
                    "Detalles_Forma": detalles_forma,
                    "Detalles_Firmas": detalles_firmas,
                    "Detalles_Fechas": detalles_fechas,
                    "Detalles_Bruta": detalles_bruta,
                    "Detalles_Criticos": detalles_forma + detalles_fechas + detalles_firmas + detalles_bruta,
                }

                estado_visual = "❌ Con Observaciones" if res_final["Detalles_Criticos"] else "✅ Aprobado"
                st.success(f"Auditoría Finalizada. Estado: {estado_visual}")
                
                # Métricas visuales explícitas
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Póliza/Cía", "❌" if detalles_forma else "✅")
                c2.metric("Ajustador", "❌" if detalles_firmas else "✅")
                c3.metric("Fechas", "❌" if detalles_fechas else "✅")
                c4.metric("Monto/Divisa", "⚠️" if detalles_bruta else "✅")

                if res_final["Detalles_Criticos"]:
                    st.error("Observaciones encontradas:")
                    for d in res_final["Detalles_Criticos"]: st.write(f"• {d}")

                # Botón Word
                word_buf = generar_word_individual(res_final)
                st.download_button("📄 Descargar Certificado Word", word_buf.getvalue(), f"Auditoria_{datos_doc['Liquidacion']}.docx")
else:
    st.info("👈 Sube primero el Reporte de Acciones en la barra lateral.")
