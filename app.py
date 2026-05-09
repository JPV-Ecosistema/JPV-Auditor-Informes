import streamlit as st
import pandas as pd
import pdfplumber
import docx
import re
import io
from datetime import datetime

st.set_page_config(page_title="Auditor de Informes JPV", layout="wide")
st.title("🔎 Auditor Automático de Informes")
st.markdown("Revisión en lote de documentos (Word/PDF) contra el Reporte de Acciones. Valida forma, aritmética y cambios de reserva.")

# --- 1. FUNCIONES DE EXTRACCIÓN DE TEXTO ---
def extraer_texto_pdf(archivo):
    texto = ""
    try:
        with pdfplumber.open(archivo) as pdf:
            for pagina in pdf.pages:
                txt = pagina.extract_text(layout=True)
                if txt:
                    texto += txt + "\n"
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
                texto += " | ".join([cell.text.replace("\n", " ") for cell in row.cells]) + "\n"
    except Exception as e:
        texto = f"Error al leer DOCX: {e}"
    return texto

# --- 2. FUNCIONES DE LIMPIEZA Y BÚSQUEDA ---
def limpiar_monto(texto_monto):
    if not texto_monto: return 0.0
    # Limpia el string dejando solo números, puntos y comas
    limpio = re.sub(r'[^\d,\.-]', '', texto_monto)
    # Convierte formato chileno "1.600,00" a float "1600.00"
    limpio = limpio.replace('.', '').replace(',', '.')
    try:
        return float(limpio)
    except:
        return 0.0

def extraer_datos_informe(texto):
    datos = {
        "Liquidacion": None,
        "Poliza": None,
        "Fecha_Siniestro": None,
        "Fecha_Denuncia": None,
        "Reserva_Neta": 0.0,
        "Honorarios": 0.0,
        "Gastos": 0.0,
        "IVA": 0.0,
        "Total_Reserva": 0.0
    }
    
    # 1. Buscar Número de Liquidación o Caso (La Llave)
    match_liq = re.search(r'(?:LIQUIDACI[OÓ]N Nº|Ref\. JPV\s*:)\s*(\d+)', texto, re.IGNORECASE)
    if match_liq:
        datos["Liquidacion"] = match_liq.group(1).strip()
        
    # 2. Buscar Póliza
    match_pol = re.search(r'(?:Nº Póliza|Póliza Nº|Póliza número)[\s:]*([A-Za-z0-9-]+)', texto, re.IGNORECASE)
    if match_pol:
        datos["Poliza"] = match_pol.group(1).strip()
        
    # 3. Buscar Fechas
    match_fsin = re.search(r'Fecha de Siniestro[\s:]*([\d\-]+)', texto, re.IGNORECASE)
    if match_fsin: datos["Fecha_Siniestro"] = match_fsin.group(1).strip()
        
    match_fden = re.search(r'Fecha Denuncia[\s:]*([\d\-]+)', texto, re.IGNORECASE)
    if match_fden: datos["Fecha_Denuncia"] = match_fden.group(1).strip()

    # 4. Búsqueda de variables para Aritmética (Buscamos todas las coincidencias y tomamos la última que suele ser la tabla final)
    matches_neta = re.findall(r'(?:Reserva Neta|Pérdida Probable Neta)[^\d]*([\d\.,]+)', texto, re.IGNORECASE)
    if matches_neta: datos["Reserva_Neta"] = limpiar_monto(matches_neta[-1])
        
    matches_hon = re.findall(r'Honorarios[^\d]*([\d\.,]+)', texto, re.IGNORECASE)
    if matches_hon: datos["Honorarios"] = limpiar_monto(matches_hon[-1])

    matches_gas = re.findall(r'Gastos[^\d]*([\d\.,]+)', texto, re.IGNORECASE)
    if matches_gas: datos["Gastos"] = limpiar_monto(matches_gas[-1])

    matches_iva = re.findall(r'IVA[^\d]*([\d\.,]+)', texto, re.IGNORECASE)
    if matches_iva: datos["IVA"] = limpiar_monto(matches_iva[-1])
        
    # Total Reserva
    matches_tot = re.findall(r'(?:Total reserva recomendada|Total Reserva del siniestro|Total Reserva)[^\d]*([\d\.,]+)', texto, re.IGNORECASE)
    if matches_tot: datos["Total_Reserva"] = limpiar_monto(matches_tot[-1])
        
    return datos

# --- 3. CARGA DE BASE MAESTRA ---
def cargar_reporte_acciones(archivo):
    if archivo is None: return None
    # Asumimos el formato estándar: saltamos 5 filas de encabezado institucional
    df = pd.read_excel(archivo, skiprows=5)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how='all', axis=0)
    
    posibles_nombres = ['Número de caso', 'Numero de caso', 'N° caso', 'Caso']
    col_llave = next((c for c in df.columns if c in posibles_nombres), None)
    
    if col_llave:
        df[col_llave] = df[col_llave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        # Renombramos internamente para facilitar el cruce sin perder la columna
        df.rename(columns={col_llave: 'Llave_Caso'}, inplace=True)
    return df

# --- INTERFAZ PRINCIPAL ---
st.sidebar.header("1. Base Maestra")
archivo_reporte = st.sidebar.file_uploader("Sube el Reporte de Acciones (Excel)", type=["xlsx"])

st.sidebar.header("2. Documentos a Revisar")
archivos_informes = st.sidebar.file_uploader("Sube los informes (PDF o DOCX)", type=["pdf", "docx"], accept_multiple_files=True)

if archivo_reporte and archivos_informes:
    df_acciones = cargar_reporte_acciones(archivo_reporte)
    
    if df_acciones is None or 'Llave_Caso' not in df_acciones.columns:
        st.error("❌ El Reporte de Acciones no pudo ser procesado o no tiene la columna 'Número de caso'.")
    else:
        st.info(f"Reporte de acciones cargado. Procesando y auditando {len(archivos_informes)} informe(s)...")
        resultados = []

        # --- MOTOR DE AUDITORÍA ---
        for archivo in archivos_informes:
            nombre_arch = archivo.name
            
            if nombre_arch.lower().endswith('.pdf'):
                texto_doc = extraer_texto_pdf(archivo)
            else:
                texto_doc = extraer_texto_docx(archivo)
                
            datos = extraer_datos_informe(texto_doc)
            caso_informe = datos["Liquidacion"]

            if not caso_informe:
                resultados.append({
                    "Documento": nombre_arch,
                    "N° Caso": "No detectado",
                    "Validación Fechas": "❌ Fallo",
                    "Validación Forma": "❌ Fallo",
                    "Aritmética": "❌ Fallo",
                    "Desviación Reserva": "❌ Fallo",
                    "Detalle": "No se encontró el Número de Liquidación/Caso en el documento."
                })
                continue

            # Buscar el caso en el Reporte de Acciones
            filas_match = df_acciones[df_acciones['Llave_Caso'] == caso_informe]
            
            if filas_match.empty:
                resultados.append({
                    "Documento": nombre_arch,
                    "N° Caso": caso_informe,
                    "Validación Fechas": "❌ Fallo",
                    "Validación Forma": "❌ Fallo",
                    "Aritmética": "❌ Fallo",
                    "Desviación Reserva": "❌ Fallo",
                    "Detalle": "El caso existe en el informe pero NO se encontró en el Reporte de Acciones."
                })
                continue
                
            # Extraer datos reales del sistema
            fila_caso = filas_match.iloc[0]
            poliza_sistema = str(fila_caso.get('Póliza de seguros', '')).strip()
            
            # Buscar fechas en sistema
            col_fsin = next((c for c in df_acciones.columns if 'siniestro' in c.lower()), None)
            fsin_sistema = str(fila_caso.get(col_fsin, '')).strip() if col_fsin else ""
            
            col_reserva = next((c for c in df_acciones.columns if 'Perdida' in c or 'Reserva' in c), None)
            reserva_sistema_val = limpiar_monto(str(fila_caso.get(col_reserva, '0'))) if col_reserva else 0.0

            # --- VALIDACIONES ESTRICTAS ---
            alerta_fechas = "✅ OK"
            alerta_forma = "✅ OK"
            alerta_aritmetica = "✅ OK"
            alerta_reserva = "✅ OK"
            detalle = "Auditoría completada sin hallazgos."
            detalles_errores = []

            # 1. Fechas
            if datos["Fecha_Siniestro"]:
                # Comparamos solo los primeros 10 caracteres (dd-mm-aaaa)
                fsin_doc_corta = datos["Fecha_Siniestro"][:10]
                if fsin_doc_corta not in fsin_sistema:
                    alerta_fechas = "❌ Error"
                    detalles_errores.append(f"Fecha Sin: Sist({fsin_sistema}) vs Doc({fsin_doc_corta})")

            # 2. Forma (Póliza)
            if poliza_sistema != "nan" and poliza_sistema != "":
                if datos["Poliza"] and poliza_sistema.upper() not in datos["Poliza"].upper():
                    alerta_forma = "❌ Error"
                    detalles_errores.append(f"Póliza: Sist({poliza_sistema}) vs Doc({datos['Poliza']})")
            
            # 3. Aritmética (Recálculo Interno)
            # Suma de componentes (Asumiendo que si la reserva neta existe, la suma debe cuadrar)
            if datos["Reserva_Neta"] > 0 or datos["Total_Reserva"] > 0:
                suma_calculada = datos["Reserva_Neta"] + datos["Honorarios"] + datos["Gastos"] + datos["IVA"]
                diferencia_aritmetica = abs(suma_calculada - datos["Total_Reserva"])
                
                # Tolerancia de 0.02 para posibles errores de redondeo de céntimos en Excel vs Texto
                if diferencia_aritmetica > 0.02:
                    alerta_aritmetica = "❌ Error"
                    detalles_errores.append(f"Aritmética descuadrada: Suma real {suma_calculada:,.2f} vs Tipeado {datos['Total_Reserva']:,.2f}")

            # 4. Desviación de Reserva vs Sistema
            diferencia_sistema = abs(datos["Total_Reserva"] - reserva_sistema_val)
            if diferencia_sistema > 1.0 and reserva_sistema_val > 0:
                alerta_reserva = "⚠️ Warning"
                detalles_errores.append(f"Reserva Sist({reserva_sistema_val:,.2f}) difiere de Doc({datos['Total_Reserva']:,.2f})")

            if detalles_errores:
                detalle = " | ".join(detalles_errores)

            # --- REGISTRO DEL RESULTADO ---
            resultados.append({
                "Documento": nombre_arch,
                "N° Caso": caso_informe,
                "Validación Fechas": alerta_fechas,
                "Validación Forma": alerta_forma,
                "Aritmética": alerta_aritmetica,
                "Desviación Reserva": alerta_reserva,
                "Detalle": detalle
            })

        # --- TABLA DE RESULTADOS (DASHBOARD) ---
        st.subheader("📊 Dashboard de Auditoría Integral")
        df_resultados = pd.DataFrame(resultados)
        
        # Estilos visuales (Utilizando .map() para evitar el AttributeError de versiones nuevas de Pandas)
        def colorear_estados(val):
            color = ''
            if '✅' in str(val): color = 'background-color: #d4edda; color: #155724;'
            elif '❌' in str(val): color = 'background-color: #f8d7da; color: #721c24;'
            elif '⚠️' in str(val): color = 'background-color: #fff3cd; color: #856404;'
            return color

        # Aplicamos .map() en lugar de .applymap()
        st.dataframe(
            df_resultados.style.map(
                colorear_estados, 
                subset=['Validación Fechas', 'Validación Forma', 'Aritmética', 'Desviación Reserva']
            ), 
            use_container_width=True,
            hide_index=True
        )

        # Botón de Descarga de Resultados
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_resultados.to_excel(writer, sheet_name="Auditoria", index=False)
        
        st.divider()
        st.download_button(
            label="📥 Descargar Reporte de Auditoría",
            data=buffer.getvalue(),
            file_name=f"Resultado_Auditoria_{datetime.now().strftime('%d-%m-%y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👈 Sube tu Reporte de Acciones (Excel) y los informes (Word/PDF) para comenzar la revisión masiva.")
