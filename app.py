import streamlit as st
import openpyxl
import json
import os
import re
from datetime import datetime

# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="CRD Tesis - Ingreso de Datos",
    layout="wide"
)

try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
    API_KEY = ""

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "CRD_Tesis_Cardiorrenal.xlsx")

# Gemini is optional: the manual form still works without it.
MODEL_NAME = "gemini-1.5-flash"

if API_KEY:
    try:
        import google.generativeai as genai
        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel(MODEL_NAME)
    except Exception:
        model = None
else:
    model = None

st.title("Tesis Esther Tamarit")
st.caption("Registro cardiorrenal · extracción asistida por IA · revisión humana obligatoria")


# ============================================================
# VARIABLES
# ============================================================

claves_por_defecto = {
    "id_pac": "",
    "fecha_inc": datetime.today().strftime("%d/%m/%Y"),
    "edad": "",
    "sexo": "",
    "peso": "",
    "talla": "",
    "eti_erc": "",
    "eti_ic": "",
    "dm2": "No",
    "hta": "No",
    "fa": "No",
    "epoc": "No",
    "sd_metab": "No",
    "tabaco": "No",
    "enolismo": "No",
    "hepato": "No",

    "hb": "",
    "creat": "",
    "cist_c": "",
    "fge": "",
    "urea": "",
    "ac_urico": "",
    "prot_creat": "",
    "ast": "",
    "alt": "",
    "plaq": "",
    "bili_t": "",
    "bili_d": "",
    "albumina": "",
    "hba1c": "",
    "colest": "",
    "nt_probnp": "",
    "ca125": "",
    "gal3": "",
    "sst2": "",
    "gdf15": "",
    "biobanco": "No",

    "fevi": "",
    "gls": "",
    "masa_vi": "",
    "tapse": "",
    "vai": "",
    "vci": "",
    "lsm": "",
    "cat_fibro": "",
    "cap": "",
    "med_val": "",
    "iqr_med": "",
    "dias_desc": "",
    "nt_prueba": "",
    "edemas_prueba": "No",

    "ieca": "No",
    "ara2": "No",
    "bb": "No",
    "amr": "No",
    "sac_val": "No",
    "sglt2i": "No",
    "diur_asa": "No",
    "hctz": "No",
    "acetazolamida": "No",
    "estatinas": "No",
    "epo": "No",

    "meses_seg": "",
    "m_cv": "No",
    "f_m_cv": "",
    "hosp_ic": "No",
    "f_hosp_ic": "",
    "iam": "No",
    "f_iam": "",
    "acv": "No",
    "f_acv": "",
    "m_tot": "No",
    "f_m_tot": "",
    "trs": "No",
    "f_trs": "",
    "caida_fge": "No",
    "f_caida_fge": "",
    "sd_cr": "No",
    "f_sd_cr": "",
}

# ============================================================
# UTILIDADES
# ============================================================

BOOLS = ["Sí", "No"]

def init_state():
    for key, value in claves_por_defecto.items():
        if key not in st.session_state:
            st.session_state[key] = value

def normalizar_si_no(v):
    if v is None:
        return v
    s = str(v).strip().lower()
    if s in {"si", "sí", "yes", "true"}:
        return "Sí"
    if s in {"no", "false"}:
        return "No"
    return str(v).strip()

def limpiar_numero(v):
    """Convierte valores numéricos habituales a string limpio.
    No intenta interpretar unidades ni hacer conversiones clínicas.
    """
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    s = s.replace(",", ".")
    # Elimina espacios, pero no fuerza conversiones clínicas.
    s = re.sub(r"\s+", "", s)
    return s

def es_numero(v):
    try:
        float(str(v).replace(",", "."))
        return True
    except Exception:
        return False

def validar_rangos():
    """Devuelve advertencias. No modifica datos."""
    warnings = []

    rangos = {
        "edad": (0, 120),
        "peso": (1, 500),
        "talla": (30, 250),
        "hb": (1, 30),
        "creat": (0.1, 30),
        "fge": (0, 200),
        "fevi": (0, 100),
        "tapse": (0, 50),
        "lsm": (0, 100),
        "cap": (0, 1000),
        "meses_seg": (0, 120),
        "plaq": (1, 2000),
    }

    for key, (lo, hi) in rangos.items():
        value = st.session_state.get(key, "")
        if value not in ("", None) and es_numero(value):
            x = float(str(value).replace(",", "."))
            if x < lo or x > hi:
                warnings.append(
                    f"{key}: valor {value} fuera del rango de comprobación ({lo}–{hi})."
                )

    return warnings


# ============================================================
# ESTADO
# ============================================================

init_state()

if "last_ai_data" not in st.session_state:
    st.session_state.last_ai_data = {}

if "audit_log" not in st.session_state:
    st.session_state.audit_log = []


# ============================================================
# EXCEL
# ============================================================

def encontrar_fila(ws, id_buscado):
    id_buscado = str(id_buscado).strip()

    # Buscar primero un ID existente.
    for r in range(2, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val is not None and str(val).strip() == id_buscado:
            return r, True

    # Después buscar primera fila vacía.
    for r in range(2, ws.max_row + 100):
        val = ws.cell(row=r, column=1).value
        if val is None or str(val).strip() == "":
            return r, False

    return ws.max_row + 1, False


def cargar_paciente(id_buscado):
    if not os.path.exists(EXCEL_FILE):
        return False

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)

    if "01_Datos_Clinicos" not in wb.sheetnames:
        return False

    ws_clin = wb["01_Datos_Clinicos"]
    fila_obj = None

    for r in range(2, ws_clin.max_row + 1):
        if str(ws_clin.cell(row=r, column=1).value).strip() == str(id_buscado).strip():
            fila_obj = r
            break

    if not fila_obj:
        return False

    def get_val(ws_name, col):
        v = wb[ws_name].cell(row=fila_obj, column=col).value
        return "" if v is None else str(v).strip()

    mapping = {
        "fecha_inc": ("01_Datos_Clinicos", 2),
        "edad": ("01_Datos_Clinicos", 3),
        "sexo": ("01_Datos_Clinicos", 4),
        "peso": ("01_Datos_Clinicos", 5),
        "talla": ("01_Datos_Clinicos", 6),
        "eti_erc": ("01_Datos_Clinicos", 8),
        "eti_ic": ("01_Datos_Clinicos", 9),
        "dm2": ("01_Datos_Clinicos", 10),
        "hta": ("01_Datos_Clinicos", 11),
        "fa": ("01_Datos_Clinicos", 12),
        "epoc": ("01_Datos_Clinicos", 13),
        "sd_metab": ("01_Datos_Clinicos", 14),
        "tabaco": ("01_Datos_Clinicos", 15),
        "enolismo": ("01_Datos_Clinicos", 16),
        "hepato": ("01_Datos_Clinicos", 17),

        "hb": ("02_Analitica_Biomarcadores", 2),
        "creat": ("02_Analitica_Biomarcadores", 3),
        "cist_c": ("02_Analitica_Biomarcadores", 4),
        "fge": ("02_Analitica_Biomarcadores", 5),
        "urea": ("02_Analitica_Biomarcadores", 6),
        "ac_urico": ("02_Analitica_Biomarcadores", 7),
        "prot_creat": ("02_Analitica_Biomarcadores", 8),
        "ast": ("02_Analitica_Biomarcadores", 9),
        "alt": ("02_Analitica_Biomarcadores", 10),
        "plaq": ("02_Analitica_Biomarcadores", 11),
        "bili_t": ("02_Analitica_Biomarcadores", 13),
        "bili_d": ("02_Analitica_Biomarcadores", 14),
        "albumina": ("02_Analitica_Biomarcadores", 15),
        "hba1c": ("02_Analitica_Biomarcadores", 16),
        "colest": ("02_Analitica_Biomarcadores", 17),
        "nt_probnp": ("02_Analitica_Biomarcadores", 18),
        "ca125": ("02_Analitica_Biomarcadores", 19),
        "gal3": ("02_Analitica_Biomarcadores", 20),
        "sst2": ("02_Analitica_Biomarcadores", 21),
        "gdf15": ("02_Analitica_Biomarcadores", 22),
        "biobanco": ("02_Analitica_Biomarcadores", 23),

        "fevi": ("03_Eco_Elastografia", 2),
        "gls": ("03_Eco_Elastografia", 3),
        "masa_vi": ("03_Eco_Elastografia", 4),
        "tapse": ("03_Eco_Elastografia", 5),
        "vai": ("03_Eco_Elastografia", 6),
        "vci": ("03_Eco_Elastografia", 7),
        "lsm": ("03_Eco_Elastografia", 8),
        "cat_fibro": ("03_Eco_Elastografia", 9),
        "cap": ("03_Eco_Elastografia", 10),
        "med_val": ("03_Eco_Elastografia", 11),
        "iqr_med": ("03_Eco_Elastografia", 12),
        "dias_desc": ("03_Eco_Elastografia", 13),
        "nt_prueba": ("03_Eco_Elastografia", 14),
        "edemas_prueba": ("03_Eco_Elastografia", 15),

        "ieca": ("04_Tratamiento", 2),
        "ara2": ("04_Tratamiento", 3),
        "bb": ("04_Tratamiento", 4),
        "amr": ("04_Tratamiento", 5),
        "sac_val": ("04_Tratamiento", 6),
        "sglt2i": ("04_Tratamiento", 7),
        "diur_asa": ("04_Tratamiento", 8),
        "hctz": ("04_Tratamiento", 9),
        "acetazolamida": ("04_Tratamiento", 10),
        "estatinas": ("04_Tratamiento", 11),
        "epo": ("04_Tratamiento", 12),

        "meses_seg": ("05_Seguimiento_24m", 2),
        "m_cv": ("05_Seguimiento_24m", 3),
        "f_m_cv": ("05_Seguimiento_24m", 4),
        "hosp_ic": ("05_Seguimiento_24m", 5),
        "f_hosp_ic": ("05_Seguimiento_24m", 6),
        "iam": ("05_Seguimiento_24m", 7),
        "f_iam": ("05_Seguimiento_24m", 8),
        "acv": ("05_Seguimiento_24m", 9),
        "f_acv": ("05_Seguimiento_24m", 10),
        "m_tot": ("05_Seguimiento_24m", 12),
        "f_m_tot": ("05_Seguimiento_24m", 13),
        "trs": ("05_Seguimiento_24m", 14),
        "f_trs": ("05_Seguimiento_24m", 15),
        "caida_fge": ("05_Seguimiento_24m", 16),
        "f_caida_fge": ("05_Seguimiento_24m", 17),
        "sd_cr": ("05_Seguimiento_24m", 18),
        "f_sd_cr": ("05_Seguimiento_24m", 19),
    }

    for key, (sheet, col) in mapping.items():
        st.session_state[key] = get_val(sheet, col)

    st.session_state.id_pac = str(id_buscado).strip()
    return True


def guardar_registro():
    if not st.session_state.id_pac:
        return False, "El ID de Paciente es obligatorio."

    if not os.path.exists(EXCEL_FILE):
        return False, f"No se encuentra {EXCEL_FILE}."

    warnings = validar_rangos()
    if warnings:
        return False, "Hay valores que deben revisarse antes de guardar."

    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        s = st.session_state

        # Guardado específico por columnas: preserva las columnas no tocadas,
        # incluyendo posibles fórmulas/formato del libro.
        sheets_data = {
            "01_Datos_Clinicos": [
                (1, s.id_pac), (2, s.fecha_inc), (3, s.edad), (4, s.sexo),
                (5, s.peso), (6, s.talla), (8, s.eti_erc), (9, s.eti_ic),
                (10, s.dm2), (11, s.hta), (12, s.fa), (13, s.epoc),
                (14, s.sd_metab), (15, s.tabaco), (16, s.enolismo), (17, s.hepato)
            ],
            "02_Analitica_Biomarcadores": [
                (1, s.id_pac), (2, s.hb), (3, s.creat), (4, s.cist_c),
                (5, s.fge), (6, s.urea), (7, s.ac_urico), (8, s.prot_creat),
                (9, s.ast), (10, s.alt), (11, s.plaq), (13, s.bili_t),
                (14, s.bili_d), (15, s.albumina), (16, s.hba1c),
                (17, s.colest), (18, s.nt_probnp), (19, s.ca125),
                (20, s.gal3), (21, s.sst2), (22, s.gdf15), (23, s.biobanco)
            ],
            "03_Eco_Elastografia": [
                (1, s.id_pac), (2, s.fevi), (3, s.gls), (4, s.masa_vi),
                (5, s.tapse), (6, s.vai), (7, s.vci), (8, s.lsm),
                (9, s.cat_fibro), (10, s.cap), (11, s.med_val),
                (12, s.iqr_med), (13, s.dias_desc), (14, s.nt_prueba),
                (15, s.edemas_prueba)
            ],
            "04_Tratamiento": [
                (1, s.id_pac), (2, s.ieca), (3, s.ara2), (4, s.bb),
                (5, s.amr), (6, s.sac_val), (7, s.sglt2i), (8, s.diur_asa),
                (9, s.hctz), (10, s.acetazolamida), (11, s.estatinas), (12, s.epo)
            ],
            "05_Seguimiento_24m": [
                (1, s.id_pac), (2, s.meses_seg), (3, s.m_cv), (4, s.f_m_cv),
                (5, s.hosp_ic), (6, s.f_hosp_ic), (7, s.iam), (8, s.f_iam),
                (9, s.acv), (10, s.f_acv), (12, s.m_tot), (13, s.f_m_tot),
                (14, s.trs), (15, s.f_trs), (16, s.caida_fge),
                (17, s.f_caida_fge), (18, s.sd_cr), (19, s.f_sd_cr)
            ]
        }

        existence = None
        for sheet_name, data in sheets_data.items():
            ws = wb[sheet_name]
            fila, existe = encontrar_fila(ws, s.id_pac)
            if existence is None:
                existence = existe

            for col, val in data:
                ws.cell(row=fila, column=col, value=val)

        wb.save(EXCEL_FILE)

        return True, (
            f"Paciente {s.id_pac} "
            + ("ACTUALIZADO correctamente." if existence else "NUEVO guardado correctamente.")
        )

    except PermissionError:
        return False, "No se puede guardar el Excel. Comprueba que no esté abierto en Excel."
    except Exception as e:
        return False, f"Error inesperado al guardar: {e}"


# ============================================================
# IA
# ============================================================

AI_SCHEMA = {
    "id_pac": "string",
    "edad": "number",
    "sexo": "H|M",
    "peso": "number",
    "talla": "number",
    "eti_erc": "DM2|HTA|Glomerulonefritis|Poliquistosis|Otras",
    "eti_ic": "Isquémica|Hipertensiva|MCD|HFpEF|Valvular",

    "dm2": "Sí|No",
    "hta": "Sí|No",
    "fa": "Sí|No",
    "epoc": "Sí|No",
    "sd_metab": "Sí|No",
    "tabaco": "Sí|No",
    "enolismo": "Sí|No",
    "hepato": "Sí|No",

    "hb": "number",
    "creat": "number",
    "cist_c": "number",
    "fge": "number",
    "urea": "number",
    "ac_urico": "number",
    "prot_creat": "number",
    "ast": "number",
    "alt": "number",
    "plaq": "number",
    "bili_t": "number",
    "bili_d": "number",
    "albumina": "number",
    "hba1c": "number",
    "colest": "number",
    "nt_probnp": "number",
    "ca125": "number",
    "gal3": "number",
    "sst2": "number",
    "gdf15": "number",
    "biobanco": "Sí|No",

    "fevi": "number",
    "gls": "number",
    "masa_vi": "number",
    "tapse": "number",
    "vai": "number",
    "vci": "number",
    "lsm": "number",
    "cat_fibro": "string",
    "cap": "number",
    "med_val": "number",
    "iqr_med": "number|string",
    "dias_desc": "number",
    "nt_prueba": "number",
    "edemas_prueba": "Sí|No",

    "ieca": "Sí|No",
    "ara2": "Sí|No",
    "bb": "Sí|No",
    "amr": "Sí|No",
    "sac_val": "Sí|No",
    "sglt2i": "Sí|No",
    "diur_asa": "Sí|No",
    "hctz": "Sí|No",
    "acetazolamida": "Sí|No",
    "estatinas": "Sí|No",
    "epo": "Sí|No",

    "meses_seg": "number",
    "m_cv": "Sí|No",
    "f_m_cv": "DD/MM/AAAA",
    "hosp_ic": "Sí|No",
    "f_hosp_ic": "DD/MM/AAAA",
    "iam": "Sí|No",
    "f_iam": "DD/MM/AAAA",
    "acv": "Sí|No",
    "f_acv": "DD/MM/AAAA",
    "m_tot": "Sí|No",
    "f_m_tot": "DD/MM/AAAA",
    "trs": "Sí|No",
    "f_trs": "DD/MM/AAAA",
    "caida_fge": "Sí|No",
    "f_caida_fge": "DD/MM/AAAA",
    "sd_cr": "Sí|No",
    "f_sd_cr": "DD/MM/AAAA",
}


def extraer_datos(texto):
    if model is None:
        st.error("Gemini no está configurado. Añade GEMINI_API_KEY en st.secrets.")
        return None

    prompt = f"""
Eres un extractor de datos clínicos para una base de investigación cardiorrenal.

OBJETIVO:
Extraer exclusivamente información explícitamente presente en el texto clínico.
NO inventes datos.
NO completes campos por conocimiento médico.
NO infieras una variable porque otra variable la haga probable.
Si un dato no aparece, devuelve null.
Si existe duda, devuelve null.

REGLAS:
1. Mantén exactamente las unidades y el valor numérico encontrado.
2. No conviertas unidades salvo que la conversión esté explícitamente indicada.
3. Para Sí/No, sólo usa Sí cuando el texto lo documente claramente y No cuando esté documentado como ausente/negado.
4. Si aparece un tratamiento, marca el correspondiente Sí/No sólo si está explícitamente documentado.
5. Las fechas deben conservarse como DD/MM/AAAA cuando sea posible.
6. MUY IMPORTANTE: NO conviertas LSM en fibrosis.
7. NO asignes cat_fibro a partir de LSM. Si el texto proporciona una categoría explícita, extráela; si no, null.
8. No confundas antecedentes con eventos ocurridos durante el seguimiento.
9. Devuelve ÚNICAMENTE JSON válido. Sin markdown. Sin explicaciones.

ESQUEMA:
{json.dumps(AI_SCHEMA, ensure_ascii=False, indent=2)}

TEXTO CLÍNICO:
{texto}
"""

    try:
        res = model.generate_content(prompt)
        raw = res.text.strip()
        raw = re.sub(r"^```json\s*", "", raw, flags=re.I)
        raw = re.sub(r"\s*```$", "", raw)

        data = json.loads(raw)

        # Sólo aceptamos claves conocidas.
        data = {k: v for k, v in data.items() if k in claves_por_defecto}

        return data

    except Exception as e:
        st.error(f"No se pudo interpretar la respuesta de la IA: {e}")
        return None


def aplicar_datos_ia(datos):
    cambios = []

    for k, v in datos.items():
        if k not in claves_por_defecto or v is None:
            continue

        if isinstance(v, str):
            v = v.strip()

        if v == "":
            continue

        if k in {
            "dm2", "hta", "fa", "epoc", "sd_metab", "tabaco", "enolismo",
            "hepato", "biobanco", "edemas_prueba", "ieca", "ara2", "bb",
            "amr", "sac_val", "sglt2i", "diur_asa", "hctz", "acetazolamida",
            "estatinas", "epo", "m_cv", "hosp_ic", "iam", "acv", "m_tot",
            "trs", "caida_fge", "sd_cr"
        }:
            v = normalizar_si_no(v)

        elif isinstance(v, (int, float)):
            v = str(v)

        else:
            v = limpiar_numero(v) if k not in {
                "id_pac", "sexo", "eti_erc", "eti_ic", "cat_fibro",
                "f_m_cv", "f_hosp_ic", "f_iam", "f_acv", "f_m_tot",
                "f_trs", "f_caida_fge", "f_sd_cr"
            } else str(v).strip()

        old = st.session_state.get(k, "")
        if str(old) != str(v):
            cambios.append((k, old, v))
            st.session_state[k] = v

    st.session_state.last_ai_data = datos
    st.session_state.audit_log.append({
        "fecha": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "accion": "IA",
        "cambios": cambios
    })

    return cambios


# ============================================================
# INTERFAZ
# ============================================================

col_izq, col_der = st.columns([1, 2])

with col_izq:
    st.subheader("1. Pegar Evolución / Historia")
    texto_input = st.text_area(
        "Pega aquí los datos en texto plano:",
        height=350,
        key="texto_clinico"
    )

    if st.button(
        "🧠 Auto-completar formulario",
        type="primary",
        use_container_width=True
    ):
        if not texto_input.strip():
            st.warning("Pega un texto primero.")
        else:
            with st.spinner("Extrayendo datos clínicos..."):
                datos_ia = extraer_datos(texto_input)

            if datos_ia:
                cambios = aplicar_datos_ia(datos_ia)
                st.success(
                    f"Extracción completada. Se han propuesto {len(cambios)} cambios. "
                    "REVISA EL FORMULARIO antes de guardar."
                )
                st.rerun()

    if st.session_state.last_ai_data:
        st.markdown("### Última extracción IA")
        st.caption("La IA propone valores; el investigador decide el valor final.")
        st.json(st.session_state.last_ai_data)


with col_der:
    c_head1, c_head2 = st.columns([2, 1])

    with c_head1:
        st.subheader("2. Formulario de Datos")

    with c_head2:
        if st.button("🔄 Nuevo paciente / limpiar", use_container_width=True):
            for clave, valor in claves_por_defecto.items():
                st.session_state[clave] = valor
            st.session_state.last_ai_data = {}
            st.rerun()

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "01. Clínicos",
        "02. Analítica",
        "03. Eco/Fibro",
        "04. Tratamiento",
        "05. Seguimiento"
    ])

    with tab1:
        c1, c2, c3, c4 = st.columns([1.5, 0.7, 1, 1])

        c1.text_input("ID Paciente", key="id_pac")

        with c2:
            st.write("")
            if st.button("🔍 Buscar", use_container_width=True):
                if st.session_state.id_pac:
                    if cargar_paciente(st.session_state.id_pac):
                        st.success("Paciente cargado. Modo edición.")
                        st.rerun()
                    else:
                        st.warning("ID no existe. Se guardará como nuevo.")
                else:
                    st.warning("Introduce un ID.")

        c3.text_input("Fecha Inclusión", key="fecha_inc")
        c4.selectbox("Sexo", ["", "H", "M"], key="sexo")

        c5, c6, c7, c8 = st.columns(4)
        c5.text_input("Edad", key="edad")
        c6.text_input("Peso (kg)", key="peso")
        c7.text_input("Talla (cm)", key="talla")

        c8.selectbox(
            "Etiología ERC",
            ["", "DM2", "HTA", "Glomerulonefritis", "Poliquistosis", "Otras"],
            key="eti_erc"
        )

        c_ic, _ = st.columns([1, 3])
        c_ic.selectbox(
            "Etiología IC",
            ["", "Isquémica", "Hipertensiva", "MCD", "HFpEF", "Valvular"],
            key="eti_ic"
        )

        st.markdown("**Comorbilidades adicionales**")
        c9, c10, c11, c12 = st.columns(4)
        c9.selectbox("DM2", BOOLS, key="dm2")
        c10.selectbox("HTA", BOOLS, key="hta")
        c11.selectbox("FA", BOOLS, key="fa")
        c12.selectbox("EPOC", BOOLS, key="epoc")

        c13, c14, c15, c16 = st.columns(4)
        c13.selectbox("Sd. Metab", BOOLS, key="sd_metab")
        c14.selectbox("Tabaquismo", BOOLS, key="tabaco")
        c15.selectbox("Enolismo", BOOLS, key="enolismo")
        c16.selectbox("Hepatopatía", BOOLS, key="hepato")

    with tab2:
        c1, c2, c3, c4 = st.columns(4)
        c1.text_input("Hemoglobina", key="hb")
        c2.text_input("Creatinina", key="creat")
        c3.text_input("Cistatina C", key="cist_c")
        c4.text_input("FGe CKD-EPI", key="fge")

        c5, c6, c7, c8 = st.columns(4)
        c5.text_input("Urea", key="urea")
        c6.text_input("Ácido úrico", key="ac_urico")
        c7.text_input("Prot/Creat", key="prot_creat")
        c8.text_input("Plaquetas", key="plaq")

        c9, c10, c11, c12 = st.columns(4)
        c9.text_input("AST", key="ast")
        c10.text_input("ALT", key="alt")
        c11.text_input("Bilirrubina total", key="bili_t")
        c12.text_input("Bilirrubina directa", key="bili_d")

        c13, c14, c15, c16 = st.columns(4)
        c13.text_input("Albúmina", key="albumina")
        c14.text_input("HbA1c", key="hba1c")
        c15.text_input("Colesterol", key="colest")
        c16.selectbox("Muestra a Biobanco", BOOLS, key="biobanco")

        st.markdown("**Biomarcadores**")
        c17, c18, c19, c20 = st.columns(4)
        c17.text_input("NT-proBNP", key="nt_probnp")
        c18.text_input("CA125", key="ca125")
        c19.text_input("Galectina-3", key="gal3")
        c20.text_input("sST2", key="sst2")

        c21, c22 = st.columns(2)
        c21.text_input("GDF-15", key="gdf15")

    with tab3:
        c1, c2, c3, c4 = st.columns(4)
        c1.text_input("FEVI (%)", key="fevi")
        c2.text_input("GLS (%)", key="gls")
        c3.text_input("Masa VI", key="masa_vi")
        c4.text_input("TAPSE", key="tapse")

        c5, c6, c7, c8 = st.columns(4)
        c5.text_input("VAI", key="vai")
        c6.text_input("VCI", key="vci")
        c7.text_input("LSM (kPa)", key="lsm")
        c8.text_input("CAP (dB/m)", key="cap")

        c9, c10, c11, c12 = st.columns(4)
        c9.text_input("Mediana elastografía", key="med_val")
        c10.text_input("IQR / IQR-mediana", key="iqr_med")
        c11.text_input("Días descompensación", key="dias_desc")
        c12.text_input("NT-proBNP (día prueba)", key="nt_prueba")

        st.markdown("**Interpretación de elastografía**")
        c13, c14 = st.columns(2)
        c13.text_input(
            "Categoría fibrosis (si está explícitamente documentada)",
            key="cat_fibro"
        )
        c14.selectbox("Edemas activos", BOOLS, key="edemas_prueba")

        st.info(
            "La IA no asigna automáticamente una categoría de fibrosis a partir de la LSM. "
            "La categoría sólo se conserva si está documentada explícitamente."
        )

    with tab4:
        st.markdown("**Tratamiento farmacológico**")
        c1, c2, c3, c4 = st.columns(4)
        c1.selectbox("IECA", BOOLS, key="ieca")
        c2.selectbox("ARA2", BOOLS, key="ara2")
        c3.selectbox("Betabloqueantes", BOOLS, key="bb")
        c4.selectbox("AMR", BOOLS, key="amr")

        c5, c6, c7, c8 = st.columns(4)
        c5.selectbox("SAC/VAL", BOOLS, key="sac_val")
        c6.selectbox("SGLT2i", BOOLS, key="sglt2i")
        c7.selectbox("Diurético de asa", BOOLS, key="diur_asa")
        c8.selectbox("HCTZ", BOOLS, key="hctz")

        c9, c10, c11, _ = st.columns(4)
        c9.selectbox("Acetazolamida", BOOLS, key="acetazolamida")
        c10.selectbox("Estatinas", BOOLS, key="estatinas")
        c11.selectbox("Eritropoyetina", BOOLS, key="epo")

    with tab5:
        st.markdown("**Seguimiento a 24 meses (endpoints y fechas)**")
        st.text_input("Meses de seguimiento total", key="meses_seg")
        st.markdown("---")

        e1, e2 = st.columns(2)
        e1.selectbox("Muerte cardiovascular", BOOLS, key="m_cv")
        e2.text_input("Fecha muerte CV", key="f_m_cv", placeholder="DD/MM/AAAA")

        e3, e4 = st.columns(2)
        e3.selectbox("Hosp. IC descompensada", BOOLS, key="hosp_ic")
        e4.text_input("Fecha hosp. IC", key="f_hosp_ic", placeholder="DD/MM/AAAA")

        e5, e6 = st.columns(2)
        e5.selectbox("IAM no fatal", BOOLS, key="iam")
        e6.text_input("Fecha IAM", key="f_iam", placeholder="DD/MM/AAAA")

        e7, e8 = st.columns(2)
        e7.selectbox("ACV no fatal", BOOLS, key="acv")
        e8.text_input("Fecha ACV", key="f_acv", placeholder="DD/MM/AAAA")

        st.markdown("---")

        e9, e10 = st.columns(2)
        e9.selectbox("Muerte total", BOOLS, key="m_tot")
        e10.text_input("Fecha muerte total", key="f_m_tot", placeholder="DD/MM/AAAA")

        e11, e12 = st.columns(2)
        e11.selectbox("Inicio TRS", BOOLS, key="trs")
        e12.text_input("Fecha TRS", key="f_trs", placeholder="DD/MM/AAAA")

        e13, e14 = st.columns(2)
        e13.selectbox("Caída FGe ≥25%", BOOLS, key="caida_fge")
        e14.text_input("Fecha caída FGe", key="f_caida_fge", placeholder="DD/MM/AAAA")

        e15, e16 = st.columns(2)
        e15.selectbox("Sd. CR agudo", BOOLS, key="sd_cr")
        e16.text_input("Fecha Sd. CR agudo", key="f_sd_cr", placeholder="DD/MM/AAAA")


# ============================================================
# VALIDACIÓN + GUARDADO
# ============================================================

st.markdown("---")

warnings = validar_rangos()

if warnings:
    st.warning("Revisa estos valores antes de guardar:")
    for w in warnings:
        st.write("• " + w)

col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    if st.button("💾 Guardar registro en Excel", use_container_width=True):
        if warnings:
            st.error("Corrige primero los valores fuera de rango.")
        else:
            ok, msg = guardar_registro()
            if ok:
                st.success(msg)
            else:
                st.error(msg)

with col_btn2:
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="⬇️ Descargar Excel actualizado",
                data=f,
                file_name="CRD_Tesis_Cardiorrenal_Actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ============================================================
# AUDITORÍA DE LA SESIÓN
# ============================================================

with st.expander("🔎 Auditoría de esta sesión"):
    if st.session_state.audit_log:
        for item in reversed(st.session_state.audit_log):
            st.write(item)
    else:
        st.caption("No hay cambios IA registrados en esta sesión.")
