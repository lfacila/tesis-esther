import streamlit as st
import openpyxl
import json
import google.generativeai as genai
import os
from datetime import datetime

# --- CONFIGURACIÓN ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except:
    API_KEY = "TU_CLAVE_LOCAL_AQUI" 

genai.configure(api_key=API_KEY)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "CRD_Tesis_Cardiorrenal.xlsx")

model = genai.GenerativeModel('gemini-1.5-flash')

st.set_page_config(page_title="CRD Tesis - Ingreso de Datos", layout="wide")
st.title("Tesis Esther Tamarit")

# --- DICCIONARIO DE VARIABLES ---
claves_por_defecto = {
    # 01 Datos Clínicos
    "id_pac": "", "fecha_inc": datetime.today().strftime('%d/%m/%Y'), "edad": "", "sexo": "H", 
    "peso": "", "talla": "", "eti_erc": "HTA", "eti_ic": "HFpEF", 
    "dm2": "No", "hta": "No", "fa": "No", "epoc": "No", "sd_metab": "No",
    "tabaco": "No", "enolismo": "No", "hepato": "No",
    # 02 Analítica
    "hb": "", "creat": "", "cist_c": "", "fge": "", "urea": "", "ac_urico": "", "prot_creat": "",
    "ast": "", "alt": "", "plaq": "", "bili_t": "", "bili_d": "", "albumina": "", "hba1c": "", "colest": "",
    "nt_probnp": "", "ca125": "", "gal3": "", "sst2": "", "gdf15": "", "biobanco": "No",
    # 03 Eco Elastografía
    "fevi": "", "gls": "", "masa_vi": "", "tapse": "", "vai": "", "vci": "",
    "lsm": "", "cat_fibro": "F0-F1 (<7)", "cap": "", "med_val": "", "iqr_med": "Sí", 
    "dias_desc": "", "nt_prueba": "", "edemas_prueba": "No",
    # 04 Tratamiento
    "ieca": "No", "ara2": "No", "bb": "No", "amr": "No", "sac_val": "No", "sglt2i": "No",
    "diur_asa": "No", "hctz": "No", "acetazolamida": "No", "estatinas": "No", "epo": "No",
    # 05 Seguimiento
    "meses_seg": "", "m_cv": "No", "f_m_cv": "", "hosp_ic": "No", "f_hosp_ic": "", 
    "iam": "No", "f_iam": "", "acv": "No", "f_acv": "", "m_tot": "No", "f_m_tot": "",
    "trs": "No", "f_trs": "", "caida_fge": "No", "f_caida_fge": "", "sd_cr": "No", "f_sd_cr": ""
}

for clave, valor in claves_por_defecto.items():
    if clave not in st.session_state:
        st.session_state[clave] = valor

# --- FUNCIÓN DE EXTRACCIÓN IA ---
def extraer_datos(texto):
    prompt = f"""
    Eres un asistente médico experto. Analiza el siguiente texto clínico y extrae todas las variables posibles para un estudio de insuficiencia cardíaca y renal.
    Devuelve ÚNICAMENTE un objeto JSON válido con las siguientes claves (si no encuentras el dato, pon null o dejalo vacío):
    - id_pac, edad, sexo, peso, talla, eti_erc, eti_ic
    - dm2, hta, fa, epoc, sd_metab, tabaco, enolismo, hepato (valores: Sí o No)
    - hb, creat, fge, ast, alt, plaq, nt_probnp, ca125, fevi, tapse, lsm
    - ieca, ara2, bb, amr, sac_val, sglt2i, diur_asa, hctz, acetazolamida (valores: Sí o No)

    Texto clínico a analizar:
    {texto}
    """
    try:
        res = model.generate_content(prompt)
        texto_json = res.text.strip().replace('```json', '').replace('```', '')
        return json.loads(texto_json)
    except Exception as e:
        return None

# --- MAQUETACIÓN DE LA INTERFAZ ---
col_izq, col_der = st.columns([1, 2])

with col_izq:
    st.subheader("1. Pegar Evolución / Historia")
    texto_input = st.text_area("Pega aquí los datos en texto plano:", height=300)
    
    if st.button("🧠 Auto-completar Formulario", type="primary", use_container_width=True):
        if texto_input:
            with st.spinner("Procesando historia clínica..."):
                datos_ia = extraer_datos(texto_input)
                if datos_ia:
                    for k, v in datos_ia.items():
                        if v is not None and str(v).strip() != "":
                            valor_str = str(v).strip()
                            if valor_str.lower() in ["si", "sí"]: valor_str = "Sí"
                            if valor_str.lower() in ["no"]: valor_str = "No"
                            if k in st.session_state:
                                st.session_state[k] = valor_str
                    st.success("Extracción completada. Revisa los campos.")
                    st.rerun()
                else:
                    st.error("No se pudo extraer la información.")
        else:
            st.warning("Pega un texto primero.")

with col_der:
    st.subheader("2. Formulario de Datos")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["01. Clínicos", "02. Analítica", "03. Eco/Fibro", "04. Tratamiento", "05. Seguimiento"])
    
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
        
        st.markdown("**Comorbilidades Adicionales**")
        c9, c10, c11, c12 = st.columns(4)
        c9.selectbox("DM2", ["Sí", "No"], key="dm2")
        c10.selectbox("HTA", ["Sí", "No"], key="hta")
        c11.selectbox("FA", ["Sí", "No"], key="fa")
        c12.selectbox("EPOC", ["Sí", "No"], key="epoc")
        
        c13, c14, c15, c16 = st.columns(4)
        c13.selectbox("Sd. Metab", ["Sí", "No"], key="sd_metab")
        c14.selectbox("Tabaquismo", ["Sí", "No"], key="tabaco")
        c15.selectbox("Enolismo", ["Sí", "No"], key="enolismo")
        c16.selectbox("Hepatopatía", ["Sí", "No"], key="hepato")

    with tab2:
        c1, c2, c3, c4 = st.columns(4)
        c1.text_input("Hemoglobina", key="hb")
        c2.text_input("Creatinina", key="creat")
        c3.text_input("Cistatina C", key="cist_c")
        c4.text_input("FGe CKD-EPI", key="fge")
        
        c5, c6, c7, c8 = st.columns(4)
        c5.text_input("Prot/Creat", key="prot_creat")
        c6.text_input("AST", key="ast")
        c7.text_input("ALT", key="alt")
        c8.text_input("Plaquetas", key="plaq")
        
        st.markdown("**Biomarcadores y Muestras**")
        c9, c10, c11, c12 = st.columns(4)
        c9.text_input("NT-proBNP", key="nt_probnp")
        c10.text_input("CA125", key="ca125")
        c11.text_input("Galectina-3", key="gal3")
        c12.selectbox("Muestra a Biobanco", ["Sí", "No"], key="biobanco")

    with tab3:
        c1, c2, c3, c4 = st.columns(4)
        c1.text_input("FEVI (%)", key="fevi")
        c2.text_input("TAPSE", key="tapse")
        c3.text_input("LSM (kPa)", key="lsm")
        c4.text_input("CAP (dB/m)", key="cap")
        c5, c6 = st.columns(2)
        c5.text_input("NT-proBNP (día prueba)", key="nt_prueba")
        c6.selectbox("Edemas activos", ["Sí", "No"], key="edemas_prueba")

    with tab4:
        st.markdown("**Tratamiento Farmacológico Optimizado**")
        c1, c2, c3, c4 = st.columns(4)
        c1.selectbox("IECAS", ["Sí", "No"], key="ieca")
        c2.selectbox("ARA2", ["Sí", "No"], key="ara2")
        c3.selectbox("Betabloqueantes", ["Sí", "No"], key="bb")
        c4.selectbox("AMR", ["Sí", "No"], key="amr")
        
        c5, c6, c7, c8 = st.columns(4)
        c5.selectbox("SAC/VAL", ["Sí", "No"], key="sac_val")
        c6.selectbox("SGLT2i", ["Sí", "No"], key="sglt2i")
        c7.selectbox("Diurético de Asa", ["Sí", "No"], key="diur_asa")
        c8.selectbox("HCTZ", ["Sí", "No"], key="hctz")
        
        c9, c10, c11, _ = st.columns(4)
        c9.selectbox("Acetazolamida", ["Sí", "No"], key="acetazolamida")
        c10.selectbox("Estatinas", ["Sí", "No"], key="estatinas")
        c11.selectbox("Eritropoyetina", ["Sí", "No"], key="epo")

    with tab5:
        st.markdown("**Seguimiento a 24 Meses (Endpoints y Fechas)**")
        st.text_input("Meses de Seguimiento Total", key="meses_seg")
        st.markdown("---")
        
        e1, e2 = st.columns(2)
        e1.selectbox("Muerte Cardiovascular", ["Sí", "No"], key="m_cv")
        e2.text_input("Fecha Muerte CV", key="f_m_cv", placeholder="DD/MM/AAAA")
        
        e3, e4 = st.columns(2)
        e3.selectbox("Hosp. IC descompensada", ["Sí", "No"], key="hosp_ic")
        e4.text_input("Fecha Hosp. IC", key="f_hosp_ic", placeholder="DD/MM/AAAA")
        
        e5, e6 = st.columns(2)
        e5.selectbox("IAM no fatal", ["Sí", "No"], key="iam")
        e6.text_input("Fecha IAM", key="f_iam", placeholder="DD/MM/AAAA")
        
        e7, e8 = st.columns(2)
        e7.selectbox("ACV no fatal", ["Sí", "No"], key="acv")
        e8.text_input("Fecha ACV", key="f_acv", placeholder="DD/MM/AAAA")
        
        st.markdown("---")
        e9, e10 = st.columns(2)
        e9.selectbox("Muerte Total", ["Sí", "No"], key="m_tot")
        e10.text_input("Fecha Muerte Total", key="f_m_tot", placeholder="DD/MM/AAAA")
        
        e11, e12 = st.columns(2)
        e11.selectbox("Inicio TRS", ["Sí", "No"], key="trs")
        e12.text_input("Fecha TRS", key="f_trs", placeholder="DD/MM/AAAA")
        
        e13, e14 = st.columns(2)
        e13.selectbox("Caída FGe ≥25%", ["Sí", "No"], key="caida_fge")
        e14.text_input("Fecha Caída FGe", key="f_caida_fge", placeholder="DD/MM/AAAA")
        
        e15, e16 = st.columns(2)
        e15.selectbox("Sd. CR Agudo", ["Sí", "No"], key="sd_cr")
        e16.text_input("Fecha Sd. CR Agudo", key="f_sd_cr", placeholder="DD/MM/AAAA")

# --- BOTONES DE GUARDADO Y DESCARGA ---
st.markdown("---")

col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    if st.button("💾 Guardar Registro en Excel", use_container_width=True):
        if not st.session_state.id_pac:
            st.error("⚠️ El ID de Paciente es obligatorio.")
        elif os.path.exists(EXCEL_FILE):
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                s = st.session_state 
                
                def fila_vacia(ws):
                    for row in range(2, ws.max_row + 100):
                        if ws.cell(row=row, column=1).value is None or str(ws.cell(row=row, column=1).value).strip() == "":
                            return row
                    return 2

                # 01_Datos_Clinicos
                ws_clin = wb["01_Datos_Clinicos"]
                fila = fila_vacia(ws_clin)
                datos_clin = [s.id_pac, s.fecha_inc, s.edad, s.sexo, s.peso, s.talla, "", s.eti_erc, s.eti_ic, 
                              s.dm2, s.hta, s.fa, s.epoc, s.sd_metab, s.tabaco, s.enolismo, s.hepato]
                for col, val in enumerate(datos_clin, start=1): ws_clin.cell(row=fila, column=col, value=val)
                
                # 02_Analitica_Biomarcadores
                ws_lab = wb["02_Analitica_Biomarcadores"]
                datos_lab = [s.id_pac, s.hb, s.creat, s.cist_c, s.fge, s.urea, s.ac_urico, s.prot_creat, 
                             s.ast, s.alt, s.plaq, "", s.bili_t, s.bili_d, s.albumina, s.hba1c, s.colest, 
                             s.nt_probnp, s.ca125, s.gal3, s.sst2, s.gdf15, s.biobanco]
                for col, val in enumerate(datos_lab, start=1): ws_lab.cell(row=fila, column=col, value=val)
                
                # 03_Eco_Elastografia
                ws_eco = wb["03_Eco_Elastografia"]
                datos_eco = [s.id_pac, s.fevi, s.gls, s.masa_vi, s.tapse, s.vai, s.vci, 
                             s.lsm, s.cat_fibro, s.cap, s.med_val, s.iqr_med, s.dias_desc, s.nt_prueba, s.edemas_prueba]
                for col, val in enumerate(datos_eco, start=1): ws_eco.cell(row=fila, column=col, value=val)
                
                # 04_Tratamiento
                ws_tx = wb["04_Tratamiento"]
                datos_tx = [s.id_pac, s.ieca, s.ara2, s.bb, s.amr, s.sac_val, s.sglt2i, 
                            s.diur_asa, s.hctz, s.acetazolamida, s.estatinas, s.epo]
                for col, val in enumerate(datos_tx, start=1): ws_tx.cell(row=fila, column=col, value=val)

                # 05_Seguimiento_24m
                ws_end = wb["05_Seguimiento_24m"]
                datos_end = [s.id_pac, s.meses_seg, s.m_cv, s.f_m_cv, s.hosp_ic, s.f_hosp_ic, 
                             s.iam, s.f_iam, s.acv, s.f_acv, "", s.m_tot, s.f_m_tot, 
                             s.trs, s.f_trs, s.caida_fge, s.f_caida_fge, s.sd_cr, s.f_sd_cr]
                for col, val in enumerate(datos_end, start=1): ws_end.cell(row=fila, column=col, value=val)
                
                wb.save(EXCEL_FILE)
                st.success(f"✅ Paciente {s.id_pac} guardado correctamente (Variables optimizadas).")
                
            except Exception as e:
                st.error(f"Error inesperado al guardar: {e}")
        else:
            st.error(f"No se encuentra {EXCEL_FILE}.")

with col_btn2:
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="⬇️ Descargar Excel Actualizado",
                data=f,
                file_name="CRD_Tesis_Cardiorrenal_Actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
