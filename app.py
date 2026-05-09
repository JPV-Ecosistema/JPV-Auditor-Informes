import streamlit as st
import pandas as pd
import pdfplumber
import docx
import re

st.set_page_config(page_title="Auditor de Informes", layout="wide")
st.title("🔎 Auditor Automático de Informes")
st.markdown("Revisión en lote de documentos (Word/PDF) contra el Reporte de Acciones.")

# --- FUNCIONES DE LECTURA DE TEXTO ---
def extraer_texto_pdf(archivo):
    texto = ""
    try:
        with pdfplumber.open(archivo) as pdf:
            for pagina in pdf.pages:
                txt = pagina.extract_text()
                if txt:
                    texto += txt + "\n"
    except Exception as e:
        texto = f"Error al leer PDF: {e}"
    return texto

def extraer_texto_docx(archivo):
    texto = ""
    try:
        doc = docx.Document(archivo)
        texto = "\n".join([para.text for para in doc.paragraphs])
        # También leer tablas dentro del Word
        for table in doc.tables:
            for row in table.rows:
                texto += " ".join([cell.text if cell.text else "" for cell in row.cells]) + "\n"
    except Exception as e:
        texto = f"Error al leer DOCX: {e}"
    return texto

def cargar_reporte_acciones(archivo):
    if archivo is None: return None
    # El reporte de acciones trae 5 filas de encabezado institucional
    df = pd.read_excel(archivo, skiprows=5)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how='all', axis=0)
    
    # Asegurar que el número de caso sea texto limpio para la búsqueda exacta
    if 'Número de caso' in df.columns:
        df['Número de caso'] = df['Número de caso'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    return df

# --- INTERFAZ DE CARGA ---
st.sidebar.header("1. Carga de Base de Datos")
archivo_reporte = st.sidebar.file_uploader("Sube el Reporte de Acciones (Excel)", type=["xlsx"])

st.sidebar.header("2. Carga de Informes a Auditar")
archivos_informes = st.sidebar.file_uploader("Sube los informes (PDF o DOCX)", type=["pdf", "docx"], accept_multiple_files=True)

if archivo_reporte and archivos_informes:
    df_acciones = cargar_reporte_acciones(archivo_reporte)
    
    if df_acciones is not None and 'Número de caso' not in df_acciones.columns:
        st.error("El Reporte de Acciones no tiene la columna 'Número de caso'.")
    elif df_acciones is not None:
        st.success("Reporte de acciones cargado correctamente. Iniciando auditoría...")
        
        resultados = []

        # --- MOTOR DE AUDITORÍA ---
        for archivo in archivos_informes:
            nombre_arch = archivo.name
            
            # 1. Extracción de texto según formato
            if nombre_arch.endswith('.pdf'):
                texto_doc = extraer_texto_pdf(archivo)
            elif nombre_arch.endswith('.docx'):
                texto_doc = extraer_texto_docx(archivo)
            else:
                continue
                
            texto_limpio = texto_doc.upper().replace("\n", " ") # Normalizar para búsqueda
            
            # 2. Búsqueda del Número de Caso (Nuestra llave)
            casos_posibles = df_acciones['Número de caso'].unique()
            caso_encontrado = None
            
            for caso in casos_posibles:
                if caso != "NAN" and caso != "" and caso in texto_limpio:
                    caso_encontrado = caso
                    break
            
            # 3. Validaciones si encontramos el caso
            if caso_encontrado:
                fila_caso = df_acciones[df_acciones['Número de caso'] == caso_encontrado].iloc[0]
                
                poliza_sistema = str(fila_caso.get('Póliza de seguros', '')).upper().strip()
                reserva_sistema = str(fila_caso.get('Perdida bruta (en moneda del caso)', '0')).strip()
                
                # a. Póliza
                alerta_poliza = "✅ OK"
                if poliza_sistema != "NAN" and poliza_sistema != "--" and poliza_sistema != "":
                    if poliza_sistema not in texto_limpio:
                        alerta_poliza = f"❌ Error: Póliza {poliza_sistema} no encontrada en texto."
                else:
                    alerta_poliza = "⚠️ Sin Póliza en Sistema"

                # b. Reserva (Monto)
                alerta_reserva = "✅ OK"
                if reserva_sistema != "NAN" and reserva_sistema != "0" and reserva_sistema != "0.0":
                    reserva_limpia = reserva_sistema.replace('.0', '')
                    # Buscamos la reserva limpia, ignorando comas de miles para mayor flexibilidad
                    reserva_limpia_sin_puntos = reserva_limpia.replace('.', '').replace(',', '')
                    texto_sin_puntos = texto_limpio.replace('.', '').replace(',', '')
                    
                    if reserva_limpia_sin_puntos not in texto_sin_puntos:
                        alerta_reserva = f"⚠️ Warning: Monto de reserva ({reserva_sistema}) difiere del documento."

                resultados.append({
                    "Documento": nombre_arch,
                    "N° Caso Detectado": caso_encontrado,
                    "Estado Póliza": alerta_poliza,
                    "Estado Reserva": alerta_reserva,
                    "Detalle": "Auditoría completada."
                })
            else:
                resultados.append({
                    "Documento": nombre_arch,
                    "N° Caso Detectado": "No Encontrado",
                    "Estado Póliza": "N/A",
                    "Estado Reserva": "N/A",
                    "Detalle": "❌ No se pudo vincular el documento con un caso del sistema."
                })

        # --- VISUALIZACIÓN DE RESULTADOS ---
        st.subheader("Resultados de la Auditoría")
        df_resultados = pd.DataFrame(resultados)
        
        # Aplicar colores a la tabla
        def colorear_estados(val):
            color = ''
            if '✅' in str(val): color = 'background-color: #d4edda; color: #155724;'
            elif '❌' in str(val): color = 'background-color: #f8d7da; color: #721c24;'
            elif '⚠️' in str(val): color = 'background-color: #fff3cd; color: #856404;'
            return color

        st.dataframe(df_resultados.style.map(colorear_estados, subset=['Estado Póliza', 'Estado Reserva', 'Detalle']), use_container_width=True)

else:
    st.info("Sube el Reporte de Acciones y al menos un documento para comenzar la revisión.")
