import streamlit as st
import pandas as pd
import pdfplumber
import docx
from docx.shared import Pt, RGBColor
import re
import io
from datetime import datetime

st.set_page_config(page_title="Auditor JPV - Comparativa de Montos", layout="wide")
st.title("🔎 Auditor de Informes: Comparativa vs. Sistema")
st.markdown("Este módulo compara los montos del Reporte de Acciones contra los declarados en los informes Word/PDF.")

# --- 1. FUNCIONES DE EXTRACCIÓN ---
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

def extraer_montos_informe(texto):
    """Extrae los montos finales buscando patrones en tablas y texto"""
    datos = {
        "Liquidacion": None, "Poliza": None,
        "Bruta": 0.0, "Deducible": 0.0, "Neta": 0.0, "Total_Reserva": 0.0
    }
    
    # Identificación
    match_liq = re.search(r'(?:LIQUIDACI[OÓ]N Nº|Ref\. JPV\s*:)\s*(\d+)', texto, re.IGNORECASE)
    if match_liq: datos["Liquidacion"] = match_liq.group(1).strip()
    
    match_pol = re.search(r'(?:Nº Póliza|Póliza Nº|Póliza número)[\s:]*([A-Za-z0-9-]+)', texto, re.IGNORECASE)
    if match_pol: datos["Poliza"] = match_pol.group(1).strip()

    # Buscamos totales en tablas (patrones comunes en tus informes)
    # Sumamos todas las 'Pérdidas Estimadas' para la Bruta
    brutas = re.findall(r'(?:Pérdida Estimada|Reserva Determinada)[^\d]*([\d\.,]+)', texto, re.IGNORECASE)
    datos["Bruta"] = sum([limpiar_monto(m) for m in brutas]) if brutas else 0.0

    # Deducibles
    deducibles = re.findall(r'(?:Deducible|Deducible Contratado)[^\d]*([\d\.,]+)', texto, re.IGNORECASE)
    datos["Deducible"] = sum([limpiar_monto(m) for m in deducibles]) if deducibles else 0.0

    # Neta
    netas = re.findall(r'(?:Reserva Neta|Total Reserva Neta)[^\d]*([\d\.,]+)', texto, re.IGNORECASE)
    datos["Neta"] = sum([limpiar_monto(m) for m in netas]) if netas else 0.0

    # Total Reserva
    total = re.findall(r'(?:Total reserva recomendada|Total Reserva del siniestro|Total Reserva)[^\d]*([\d\.,]+)', texto, re.IGNORECASE)
    datos["Total_Reserva"] = limpiar_monto(total[-1]) if total else 0.0
    
    return datos

# --- 2. CARGA DE EXCEL ---
def cargar_reporte_acciones(archivo):
    if archivo is None: return None
    for i in range(10):
        try:
            df = pd.read_excel(archivo, skiprows=i)
            df.columns = [str(c).strip() for c in df.columns]
            col_liq = next((c for c in df.columns if 'número de caso' in c.lower() or 'caso' in c.lower()), None)
            if col_liq:
                df[col_liq] = df[col_liq].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                df.rename(columns={col_liq: 'Llave'}, inplace=True)
                return df
        except: continue
    return None

# --- 3. GENERADOR DE WORD ---
def generar_word_comparativo(resultados):
    doc = docx.Document()
    doc.add_heading('Auditoría de Montos: Sistema vs. Informe', 0)
    
    for res in resultados:
        doc.add_heading(f"Liquidación N° {res['Caso']}", level=1)
        doc.add_paragraph(f"Documento: {res['Archivo']}")
        
        # Tabla comparativa
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Concepto'
        hdr_cells[1].text = 'Sistema (Excel)'
        hdr_cells[2].text = 'Informe (Doc)'
        hdr_cells[3].text = 'Diferencia'
        
        conceptos = [
            ('Pérdida Bruta', res['Bruta_Sist'], res['Bruta_Inf']),
            ('Deducible', res['Ded_Sist'], res['Ded_Inf']),
            ('Pérdida Neta', res['Neta_Sist'], res['Neta_Inf']),
            ('Total Reserva', res['Total_Sist'], res['Total_Inf'])
        ]
        
        for nombre, sist, inf in conceptos:
            row_cells = table.add_row().cells
            row_cells[0].text = nombre
            row_cells[1].text = f"{sist:,.2f}"
            row_cells[2].text = f"{inf:,.2f}"
            diff = abs(sist - inf)
            row_cells[3].text = f"{diff:,.2f}"
            if diff > 1.0:
                run = row_cells[3].paragraphs[0].runs[0]
                run.font.color.rgb = RGBColor(255, 0, 0)
                run.bold = True

        doc.add_paragraph("\n")
        if res['Alerta_Poliza']:
            p = doc.add_paragraph()
            r = p.add_run(f"⚠️ {res['Alerta_Poliza']}")
            r.font.color.rgb = RGBColor(200, 0, 0)
            
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer

# --- INTERFAZ ---
st.sidebar.header("1. Carga de Datos")
archivo_excel = st.sidebar.file_uploader("Reporte de Acciones (Excel)", type=["xlsx"])
archivos_docs = st.sidebar.file_uploader("Informes (Word/PDF)", type=["docx", "pdf"], accept_multiple_files=True)

if archivo_excel and archivos_docs:
    df_sist = cargar_reporte_acciones(archivo_excel)
    
    if df_sist is not None:
        resultados_finales = []
        
        for archivo in archivos_docs:
            texto = extraer_texto_pdf(archivo) if archivo.name.endswith('.pdf') else extraer_texto_docx(archivo)
            inf = extraer_montos_informe(texto)
            
            if inf["Liquidacion"]:
                fila = df_sist[df_sist['Llave'] == inf["Liquidacion"]]
                if not fila.empty:
                    f = fila.iloc[0]
                    # Mapeo dinámico de columnas del Excel
                    c_bruta = next((c for c in df_sist.columns if 'bruta' in c.lower()), None)
                    c_neta = next((c for c in df_sist.columns if 'neta' in c.lower()), None)
                    c_pol = next((c for c in df_sist.columns if 'póliza' in c.lower()), None)
                    
                    bruta_s = limpiar_monto(f.get(c_bruta, 0))
                    neta_s = limpiar_monto(f.get(c_neta, 0))
                    pol_s = str(f.get(c_pol, '')).strip()
                    
                    # El deducible en el sistema suele ser la diferencia
                    ded_s = bruta_s - neta_s
                    
                    alerta_pol = ""
                    if pol_s and inf["Poliza"] and pol_s not in inf["Poliza"]:
                        alerta_pol = f"Póliza inconsistente: Sistema dice {pol_s} e informe dice {inf['Poliza']}"

                    resultados_finales.append({
                        "Caso": inf["Liquidacion"], "Archivo": archivo.name,
                        "Bruta_Sist": bruta_s, "Bruta_Inf": inf["Bruta"],
                        "Ded_Sist": ded_s, "Ded_Inf": inf["Deducible"],
                        "Neta_Sist": neta_s, "Neta_Inf": inf["Neta"],
                        "Total_Sist": neta_s, "Total_Inf": inf["Total_Reserva"], # Ajustar según necesidad de honorarios
                        "Alerta_Poliza": alerta_pol
                    })

        if resultados_finales:
            st.success(f"Se procesaron {len(resultados_finales)} informes con éxito.")
            word_buf = generar_word_comparativo(resultados_finales)
            st.download_button("📥 Descargar Reporte Comparativo (Word)", data=word_buf.getvalue(), file_name="Auditoria_Montos_JPV.docx")
            
            # Vista previa en tabla
            st.dataframe(pd.DataFrame(resultados_finales)[["Caso", "Bruta_Sist", "Bruta_Inf", "Neta_Sist", "Neta_Inf"]])
