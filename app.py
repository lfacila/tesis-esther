import streamlit as st
import openpyxl
import json
import google.generativeai as genai
import os
from datetime import datetime

# --- CONFIGURACIÓN ---

try:
    # Si la app está en la nube, coge la clave de los secretos protegidos
    API_KEY = st.secrets["GEMINI_API_KEY"]
except:
    # Si la app está en tu Mac, usa esta (¡Asegúrate de borrarla antes de subir a GitHub!)
    API_KEY = "LA REAL AQUI"

genai.configure(api_key=API_KEY)



# Esto fuerza al programa a buscar el Excel exactamente en la misma carpeta donde está guardado app.py
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "CRD_Tesis_Cardiorrenal.xlsx")

model = genai.GenerativeModel('gemini-1.5-flash')

st.set_page_config(page_title="CRD Tesis - Ingreso de Datos", layout="wide")
st.title("🩺 Extracción Inteligente y Formulario Completo")

# --- DICCIONARIO DE VARIABLES Y CLAVES ---
# Se inicializan todas las variables del CRD en el session_state
claves_por_defecto = {
    # Clínicos
    "id_pac": "", "fecha_inc": datetime.today().strftime('%d/%m/%Y'), "edad": "", "sexo": "H", 
    "peso": "", "talla": "", "eti_erc": "HTA", "eti_ic": "HFpEF", 
    "dm2": "No", "hta": "No", "fa": "No", "epoc": "No", "sd_metab": "No",
    # Laboratorio
    "creat": "", "cist_c": "", "fge": "", "urea": "", "ac_urico": "", "prot_creat": "",
    "ast": "", "alt": "", "plaq": "", "bili_t": "", "bili_d": "", "albumina": "", "hba1c": "", "colest": "",
    "nt_probnp": "", "ca125": "", "gal3": "", "sst2": "", "gdf15": "",
    # Eco y Elastografía
    "fevi": "", "gls": "", "masa_vi": "", "tapse": "", "vai": "", "vci": "",
    "lsm": "", "cat_fibro": "F0-F1 (<7)", "cap": "", "med_val": "", "iqr_med": "Sí", 
    "dias_desc": "", "nt_prueba": "", "edemas_prueba": "No",
    # Tratamiento
    "ieca": "No", "bb": "No", "sglt2i": "No", "sac_val": "No", 
    "diur_tipo": "Ninguno", "diur_dosis": "", "estatinas": "No", "epo": "No"
}

for clave, valor in claves_por_defecto.items():
    if clave not in st.session_state:
        st.session_state[clave] = valor

# --- FUNCIÓN DE EXTRACCIÓN IA ---
def extraer_datos(texto):
    prompt = f"""
    Eres un asistente médico experto. Analiza el siguiente texto clínico y extrae todas las variables posibles para un estudio de insuficiencia cardíaca y renal.
    Devuelve ÚNICAMENTE un objeto JSON válido con las siguientes claves (si no encuentras el dato, pon null o dejalo vacío):
    - id_pac (ID o SIP, texto)
    - edad (número)
    - sexo (H o M)
    - peso (número)
    - talla (número)
    - eti_erc (DM2, HTA, Glomerulonefritis, Poliquistosis, Otras)
    - eti_ic (Isquémica, Hipertensiva, MCD, HFpEF, Valvular)
    - dm2 (Sí o No)
    - hta (Sí o No)
    - fa (Sí o No)
    - epoc (Sí o No)
    - sd_metab (Sí o No)
    - creat (número, creatinina)
    - fge (número, filtrado glomerular)
    - ast (número)
    - alt (número)
    - plaq (número, plaquetas)
    - nt_probnp (número)
    - ca125 (número)
    - fevi (número)
    - tapse (número)
    - lsm (número, rigidez hepática en kPa)
    - ieca (Sí o No, incluye ARA-II)
    - bb (Sí o No, betabloqueantes)
    - sglt2i (Sí o No, gliflozinas)
    - diur_tipo (Furosemida, Torasemida, Tiazida, MRA, Ninguno)
    - diur_dosis (número)

    Texto clínico a analizar:
    {texto}
    """
    try:
        res = model.generate_content(prompt)
        # Limpieza básica para extraer solo el JSON
        texto_json = res.text.strip().replace('```json', '').replace('```', '')
        return json.loads(texto_json)
    except Exception as e:
        return None

# --- MAQUETACIÓN DE LA INTERFAZ ---
col_izq, col_der = st.columns([1, 2])

# COLUMNA IZQUIERDA: Área de copiado y procesado
with col_izq:
    st.subheader("1. Pegar Evolución / Historia")
    texto_input = st.text_area("Pega aquí los datos en texto plano:", height=300)
    
    if st.button("🧠 Auto-completar Formulario", type="primary", use_container_width=True):
        if texto_input and API_KEY != "TU_API_KEY_AQUI":
            with st.spinner("Procesando historia clínica..."):
                datos_ia = extraer_datos(texto_input)
                if datos_ia:
                    # Actualizar session_state con lo encontrado
                    for k, v in datos_ia.items():
                        if v is not None and str(v).strip() != "":
                            # Para evitar errores con selectboxes, normalizamos strings si es necesario
                            valor_str = str(v).strip()
                            if valor_str in ["Si", "SI", "sí", "SÍ"]: valor_str = "Sí"
                            if valor_str in ["No", "NO", "no"]: valor_str = "No"
                            if k in st.session_state:
                                st.session_state[k] = valor_str
                    st.success("Extracción completada. Revisa los campos a la derecha.")
                    st.rerun() # Refrescar para mostrar los datos en los campos
                else:
                    st.error("No se pudo extraer la información.")
        else:
            st.warning("Pega un texto y asegúrate de configurar tu API Key.")

# COLUMNA DERECHA: Pestañas (Tabs) tipo Excel
with col_der:
    st.subheader("2. Formulario de Datos (Revisar y Guardar)")
    
    # Crear pestañas idénticas al Excel
    tab1, tab2, tab3, tab4 = st.tabs(["01. Clínicos", "02. Analítica", "03. Eco/Elastografía", "04. Tratamiento"])
    
    with tab1:
        c1, c2, c3, c4 = st.columns(4)
        c1.text_input("ID Paciente", key="id_pac")
        c2.text_input("Fecha Inclusión", key="fecha_inc")
        c3.text_input("Edad", key="edad")
        c4.selectbox("Sexo", ["H", "M"], key="sexo")
        
        c5, c6, c7, c8 = st.columns(4)
        c5.text_input("Peso (kg)", key="peso")
        c6.text_input("Talla (cm)", key="talla")
        c7.selectbox("Etiología ERC", ["DM2", "HTA", "Glomerulonefritis", "Poliquistosis", "Otras"], key="eti_erc")
        c8.selectbox("Etiología IC", ["Isquémica", "Hipertensiva", "MCD", "HFpEF", "Valvular"], key="eti_ic")
        
        st.markdown("**Comorbilidades**")
        c9, c10, c11, c12, c13 = st.columns(5)
        c9.selectbox("DM2", ["Sí", "No"], key="dm2")
        c10.selectbox("HTA", ["Sí", "No"], key="hta")
        c11.selectbox("FA", ["Sí", "No"], key="fa")
        c12.selectbox("EPOC", ["Sí", "No"], key="epoc")
        c13.selectbox("Sd. Metab", ["Sí", "No"], key="sd_metab")

    with tab2:
        st.markdown("**Función Renal y Hepática**")
        c1, c2, c3, c4 = st.columns(4)
        c1.text_input("Creatinina", key="creat")
        c2.text_input("Cistatina C", key="cist_c")
        c3.text_input("FGe CKD-EPI", key="fge")
        c4.text_input("Prot/Creat", key="prot_creat")
        
        c5, c6, c7, c8 = st.columns(4)
        c5.text_input("AST", key="ast")
        c6.text_input("ALT", key="alt")
        c7.text_input("Plaquetas", key="plaq")
        c8.text_input("Bilirrubina Tot", key="bili_t")
        
        st.markdown("**Biomarcadores Tesis**")
        c9, c10, c11, c12, c13 = st.columns(5)
        c9.text_input("NT-proBNP", key="nt_probnp")
        c10.text_input("CA125", key="ca125")
        c11.text_input("Galectina-3", key="gal3")
        c12.text_input("sST2", key="sst2")
        c13.text_input("GDF-15", key="gdf15")

    with tab3:
        st.markdown("**Ecocardiografía**")
        c1, c2, c3, c4 = st.columns(4)
        c1.text_input("FEVI (%)", key="fevi")
        c2.text_input("TAPSE", key="tapse")
        c3.text_input("VAI", key="vai")
        c4.text_input("VCI (cm)", key="vci")
        
        st.markdown("**Elastografía (FibroScan)**")
        c5, c6, c7, c8 = st.columns(4)
        c5.text_input("LSM (kPa)", key="lsm")
        c6.selectbox("Cat. Fibrosis", ["F0-F1 (<7)", "F2 (7-9.4)", "F3 (9.5-12.9)", "F4 (≥13)"], key="cat_fibro")
        c7.text_input("CAP (dB/m)", key="cap")
        c8.selectbox("IQR/Med ≤0.30", ["Sí", "No"], key="iqr_med")
        
        st.markdown("**Estado Volémico (día de prueba)**")
        c9, c10, c11 = st.columns(3)
        c9.text_input("Días desde descomp.", key="dias_desc")
        c10.text_input("NT-proBNP día prueba", key="nt_prueba")
        c11.selectbox("Edemas activos", ["Sí", "No"], key="edemas_prueba")

    with tab4:
        st.markdown("**Tratamiento Médico**")
        c1, c2, c3, c4 = st.columns(4)
        c1.selectbox("IECAs/ARAs", ["Sí", "No"], key="ieca")
        c2.selectbox("Betabloqueantes", ["Sí", "No"], key="bb")
        c3.selectbox("SGLT2i", ["Sí", "No"], key="sglt2i")
        c4.selectbox("Sacubitrilo-Val", ["Sí", "No"], key="sac_val")
        
        c5, c6, c7, c8 = st.columns(4)
        c5.selectbox("Diurético", ["Furosemida", "Torasemida", "Tiazida", "MRA", "Ninguno"], key="diur_tipo")
        c6.text_input("Dosis Diurético (mg)", key="diur_dosis")
        c7.selectbox("Estatinas", ["Sí", "No"], key="estatinas")
        c8.selectbox("Eritropoyetina", ["Sí", "No"], key="epo")

    st.markdown("---")
    # Botón para guardar todos los datos
    if st.button("💾 Guardar Registro en Excel", use_container_width=True):
        if not st.session_state.id_pac:
            st.error("⚠️ El ID de Paciente es obligatorio para guardar.")
        elif os.path.exists(EXCEL_FILE):
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                s = st.session_state # Atajo
                
                # Función para encontrar la primera fila donde el ID está vacío
                def fila_vacia(ws):
                    for row in range(2, ws.max_row + 100):
                        if ws.cell(row=row, column=1).value is None or str(ws.cell(row=row, column=1).value).strip() == "":
                            return row
                    return 2

                # 01_Datos_Clinicos
                ws_clin = wb["01_Datos_Clinicos"]
                fila = fila_vacia(ws_clin)
                datos_clin = [s.id_pac, s.fecha_inc, s.edad, s.sexo, s.peso, s.talla, "", s.eti_erc, s.eti_ic, s.dm2, s.hta, s.fa, s.epoc, s.sd_metab]
                for col, val in enumerate(datos_clin, start=1): ws_clin.cell(row=fila, column=col, value=val)
                
                # 02_Analitica_Biomarcadores
                ws_lab = wb["02_Analitica_Biomarcadores"]
                fila = fila_vacia(ws_lab)
                datos_lab = [s.id_pac, s.creat, s.cist_c, s.fge, s.urea, s.ac_urico, s.prot_creat, s.ast, s.alt, s.plaq, "", s.bili_t, s.bili_d, s.albumina, s.hba1c, s.colest, s.nt_probnp, s.ca125, s.gal3, s.sst2, s.gdf15]
                for col, val in enumerate(datos_lab, start=1): ws_lab.cell(row=fila, column=col, value=val)
                
                # 03_Eco_Elastografia
                ws_eco = wb["03_Eco_Elastografia"]
                fila = fila_vacia(ws_eco)
                datos_eco = [s.id_pac, s.fevi, s.gls, s.masa_vi, s.tapse, s.vai, s.vci, s.lsm, s.cat_fibro, s.cap, s.med_val, s.iqr_med, s.dias_desc, s.nt_prueba, s.edemas_prueba]
                for col, val in enumerate(datos_eco, start=1): ws_eco.cell(row=fila, column=col, value=val)
                
                # 04_Tratamiento
                ws_tx = wb["04_Tratamiento"]
                fila = fila_vacia(ws_tx)
                datos_tx = [s.id_pac, s.ieca, s.bb, s.sglt2i, s.sac_val, s.diur_tipo, s.diur_dosis, s.estatinas, s.epo]
                for col, val in enumerate(datos_tx, start=1): ws_tx.cell(row=fila, column=col, value=val)
                
                wb.save(EXCEL_FILE)
                st.success(f"✅ Paciente {s.id_pac} guardado correctamente en la fila {fila}.")
                
            except PermissionError:
                st.error("❌ Error: Cierra el archivo Excel antes de guardar.")
            except Exception as e:
                st.error(f"Error inesperado al guardar: {e}")
        else:
            st.error(f"No se encuentra {EXCEL_FILE}. Asegúrate de que está en la misma carpeta.")

st.markdown("---")
    # Botón para descargar la base de datos
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="⬇️ Descargar Excel Actualizado",
                data=f,
                file_name="CRD_Tesis_Cardiorrenal_Actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )