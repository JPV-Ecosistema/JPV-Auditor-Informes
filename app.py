import streamlit as st
import pandas as pd
import pdfplumber
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX
import re
import io
from datetime import datetime

st.set_page_config(page_title="Auditor de Informes JPV", layout="wide")
st.title("🔎 Auditor Automático de Informes")
st.markdown("Revisión avanzada con lectura dinámica de cuadros, exclusión de IVA, cruce de Pérdida Bruta y generación de reportes en Word.")

# --- 1. FUNCIONES DE EXTRACCIÓN Y LIMPIEZA ---
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

def limpiar_monto(texto_monto):
    if not texto_monto: return 0.0
    limpio = re.sub(r'[^\d,\.-]', '', str(texto_monto))
    limpio = limpio.replace('.', '').replace(',', '.')
    try:
        return float(limpio)
    except:
        return 0.0

def buscar_ultimo_valor(patron, texto):
    matches = re.findall(patron, texto, re.IGNORECASE)
    return limpiar_monto(matches[-1]) if matches else 0.0

# --- 2. MOTOR ARITMÉTICO DINÁMICO ---
def extraer_datos_informe(texto):
    datos = {
        "Liquidacion": None, "Poliza": None, "Fecha_Siniestro": None, "Fecha_Denuncia": None,
        "Total_Bruto_Dinámico": 0.0, "Total_Neto_Dinámico": 0.0,
        "Honorarios": 0.0, "Gastos": 0.0, "Total_Reserva_Declarado": 0.0,
        "Errores_Cascada": []
    }
    
    # Llaves y Forma
    match_liq = re.search(r'(?:LIQUIDACI[OÓ]N Nº|Ref\. JPV\s*:)\s*(\d+)', texto, re.IGNORECASE)
    if match_liq: datos["Liquidacion"] = match_liq.group(1).strip()
        
    match_pol = re.search(r'(?:Nº Póliza|Póliza Nº|Póliza número)[\s:]*([A-Za-z0-9-]+)', texto, re.IGNORECASE)
    if match_pol: datos["Poliza"] = match_pol.group(1).strip()
        
    match_fsin = re.search(r'Fecha de Siniestro[\s:]*([\d\-]+)', texto, re.IGNORECASE)
    if match_fsin: datos["Fecha_Siniestro"] = match_fsin.group(1).strip()
        
    match_fden = re.search(r'Fecha Denuncia[\s:]*([\d\-]+)', texto, re.IGNORECASE)
    if match_fden: datos["Fecha_Denuncia"] = match_fden.group(1).strip()

    # Extracción de montos fijos
    datos["Honorarios"] = buscar_ultimo_valor(r'Honorarios[^\d]*([\d\.,]+)', texto)
    datos["Gastos"] = buscar_ultimo_valor(r'Gastos[^\d]*([\d\.,]+)', texto)
    datos["Total_Reserva_Declarado"] = buscar_ultimo_valor(r'(?:Total reserva recomendada|Total Reserva del siniestro|Total Reserva)[^\d]*([\d\.,]+)', texto)

    # Lógica Dinámica en Cascada (Busca patrones Base - Deducible = Neto)
    lineas = texto.split('\n')
    patron_numeros = r'(?<![\w])\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?(?![\w])'
    
    for linea in lineas:
        # Evitar líneas de totales consolidados para no duplicar sumas
        if "total reserva" in linea.lower(): 
            continue
            
        matches = re.findall(patron_numeros, linea)
        montos = [limpio for m in matches if (limpio := limpiar_monto(m)) > 0]
        
        # Si la línea tiene al menos 3 montos, comprobamos si es una fila de cálculo (Base, Deducible, Neto)
        if len(montos) >= 3:
            # Tomamos los últimos 3 números por si la línea incluye "Monto Asegurado" al principio
            base, deducible, neto = montos[-3], montos[-2], montos[-1]
            
            # Verificamos si matemáticamente tiene sentido (Base - Deducible = Neto) con tolerancia
            if abs((base - deducible) - neto) <= 1.0:
                datos["Total_Bruto_Dinámico"] += base
                datos["Total_Neto_Dinámico"] += neto
            elif abs((base - deducible) - neto) <= (base * 0.5): # Si parece fila de reserva pero está mal restada
                # Solo alertamos si la palabra pérdida o deducible está cerca para no tomar fechas al azar
                if "pérdida" in linea.lower() or "deducible" in linea.lower():
                    datos["Errores_Cascada"].append(f"Error en fila: Base {base:,.2f} - Ded {deducible:,.2f} != {neto:,.2f}")

    return datos

# --- 3. CARGA DE BASE MAESTRA INTELIGENTE ---
def cargar_reporte_acciones(archivo):
    if archivo is None: return None
    for i in range(10):
        try:
            df_temp = pd.read_excel(archivo, skiprows=i)
            df_temp.columns = [str(c).strip() for c in df_temp.columns]
            posibles_nombres = ['Número de caso', 'Numero de caso', 'N° caso', 'Caso']
            col_found = next((c for c in df_temp.columns if c in posibles_nombres), None)
            
            if col_found:
                df_temp[col_found] = df_temp[col_found].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                df_temp.rename(columns={col_found: 'Llave_Caso'}, inplace=True)
                return df_temp.dropna(how='all', axis=0)
        except:
            continue
    return None

# --- 4. GENERADOR DE REPORTE WORD ---
def generar_word_auditoria(resultados):
    doc = docx.Document()
    doc.add_heading('Reporte de Auditoría de Informes - JPV', 0)
    doc.add_paragraph(f"Fecha de revisión: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
    
    for res in resultados:
        # Título del Caso
        estado_gral = "⚠️ CON HALLAZGOS" if "❌" in res["Detalle"] or "⚠️" in res["Detalle"] else "✅ APROBADO"
        heading = doc.add_heading(f"Liquidación: {res['N° Caso']} | {estado_gral}", level=2)
        doc.add_paragraph(f"Archivo: {res['Documento']}", style='Subtitle')
        
        # Resultados Semáforo
        p = doc.add_paragraph()
        p.add_run(f"Forma y Póliza: {res['Validación Forma']}\n")
        p.add_run(f"Fechas (Sin/Den): {res['Validación Fechas']}\n")
        p.add_run(f"Aritmética Integral: {res['Aritmética']}\n")
        p.add_run(f"Cruce Pérdida Bruta: {res['Desviación Reserva']}\n")
        
        # Detalle de errores
        if res['Detalle'] and "Auditoría completada" not in res['Detalle']:
            p_det = doc.add_paragraph("Detalle de las Inconsistencias detectadas:\n", style='Intense Quote')
            for error in res['Detalle'].split(' | '):
                r = p_det.add_run(f"• {error}\n")
                r.font.color.rgb = RGBColor(200, 0, 0) # Texto rojo para errores
                
        doc.add_paragraph("_" * 50)
        
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer

# --- INTERFAZ PRINCIPAL ---
st.sidebar.header("1. Base Maestra")
archivo_reporte = st.sidebar.file_uploader("Sube el Reporte de Acciones (Excel)", type=["xlsx"])

st.sidebar.header("2. Documentos a Revisar")
archivos_informes = st.sidebar.file_uploader("Sube los informes (PDF o DOCX)", type=["pdf", "docx"], accept_multiple_files=True)

if archivo_reporte and archivos_informes:
    df_acciones = cargar_reporte_acciones(archivo_reporte)
    
    if df_acciones is None or 'Llave_Caso' not in df_acciones.columns:
        st.error("❌ No se encontró la columna 'Número de caso' en el Excel.")
    else:
        st.info(f"Base Maestra cargada. Auditando {len(archivos_informes)} informe(s)...")
        resultados = []

        for archivo in archivos_informes:
            nombre_arch = archivo.name
            texto_doc = extraer_texto_pdf(archivo) if nombre_arch.lower().endswith('.pdf') else extraer_texto_docx(archivo)
            datos = extraer_datos_informe(texto_doc)
            caso_informe = datos["Liquidacion"]

            if not caso_informe:
                resultados.append({
                    "Documento": nombre_arch, "N° Caso": "N/D",
                    "Validación Fechas": "❌", "Validación Forma": "❌",
                    "Aritmética": "❌", "Desviación Reserva": "❌",
                    "Detalle": "❌ No se encontró el N° de Liquidación en el documento."
                })
                continue

            filas_match = df_acciones[df_acciones['Llave_Caso'] == caso_informe]
            
            if filas_match.empty:
                resultados.append({
                    "Documento": nombre_arch, "N° Caso": caso_informe,
                    "Validación Fechas": "❌", "Validación Forma": "❌",
                    "Aritmética": "❌", "Desviación Reserva": "❌",
                    "Detalle": "❌ El caso no existe en el Reporte de Acciones."
                })
                continue
                
            fila_caso = filas_match.iloc[0]
            poliza_sistema = str(fila_caso.get('Póliza de seguros', '')).strip()
            
            col_fsin = next((c for c in df_acciones.columns if 'siniestro' in c.lower()), None)
            fsin_sistema = str(fila_caso.get(col_fsin, '')).strip() if col_fsin else ""
            
            col_bruta = next((c for c in df_acciones.columns if 'bruta' in c.lower()), None)
            bruta_sistema_val = limpiar_monto(str(fila_caso.get(col_bruta, '0'))) if col_bruta else 0.0

            alerta_fechas, alerta_forma, alerta_aritmetica, alerta_reserva = "✅ OK", "✅ OK", "✅ OK", "✅ OK"
            detalles_errores = datos["Errores_Cascada"]

            # 1. FECHAS
            if datos["Fecha_Siniestro"]:
                fsin_doc_corta = datos["Fecha_Siniestro"][:10]
                if fsin_doc_corta not in fsin_sistema:
                    alerta_fechas = "❌ Error"
                    detalles_errores.append(f"❌ Fecha Sin: Sist({fsin_sistema}) vs Doc({fsin_doc_corta})")

            # 2. PÓLIZA
            if poliza_sistema != "nan" and poliza_sistema != "":
                if datos["Poliza"] and poliza_sistema.upper() not in datos["Poliza"].upper():
                    alerta_forma = "❌ Error"
                    detalles_errores.append(f"❌ Póliza: Sist({poliza_sistema}) vs Doc({datos['Poliza']})")
            
            # 3. ARITMÉTICA FINAL (Neto + Honorarios + Gastos. SIN IVA)
            if datos["Total_Neto_Dinámico"] > 0:
                suma_auditor = datos["Total_Neto_Dinámico"] + datos["Honorarios"] + datos["Gastos"]
                if abs(suma_auditor - datos["Total_Reserva_Declarado"]) > 1.0:
                    alerta_aritmetica = "❌ Error"
                    detalles_errores.append(f"❌ Aritmética final descuadrada (¿IVA incluido?): Suma correcta {suma_auditor:,.2f} vs Tipeado {datos['Total_Reserva_Declarado']:,.2f}")

            # 4. PÉRDIDA BRUTA (Cruce Sistema vs Documento)
            if datos["Total_Bruto_Dinámico"] > 0:
                if abs(datos["Total_Bruto_Dinámico"] - bruta_sistema_val) > 1.0 and bruta_sistema_val > 0:
                    alerta_reserva = "⚠️ Warning"
                    detalles_errores.append(f"⚠️ Pérdida Bruta: Sist({bruta_sistema_val:,.2f}) difiere de Suma Doc({datos['Total_Bruto_Dinámico']:,.2f})")

            resultados.append({
                "Documento": nombre_arch, "N° Caso": caso_informe,
                "Validación Fechas": alerta_fechas, "Validación Forma": alerta_forma,
                "Aritmética": alerta_aritmetica, "Desviación Reserva": alerta_reserva,
                "Detalle": " | ".join(detalles_errores) if detalles_errores else "✅ Auditoría completada sin hallazgos."
            })

        st.subheader("📊 Dashboard de Auditoría Integral")
        df_resultados = pd.DataFrame(resultados)
        
        def colorear_estados(val):
            color = ''
            if '✅' in str(val): color = 'background-color: #d4edda; color: #155724;'
            elif '❌' in str(val): color = 'background-color: #f8d7da; color: #721c24;'
            elif '⚠️' in str(val): color = 'background-color: #fff3cd; color: #856404;'
            return color

        st.dataframe(
            df_resultados.style.map(
                colorear_estados, 
                subset=['Validación Fechas', 'Validación Forma', 'Aritmética', 'Desviación Reserva']
            ), 
            use_container_width=True, hide_index=True
        )

        st.divider()
        
        # Generación y Descarga del Word
        word_buffer = generar_word_auditoria(resultados)
        st.download_button(
            label="📄 Descargar Reporte en Word",
            data=word_buffer.getvalue(),
            file_name=f"Auditoria_JPV_{datetime.now().strftime('%d-%m-%y')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.info("👈 Sube tu Reporte de Acciones y los informes a auditar.")
