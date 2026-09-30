import json
import math
import re
from datetime import datetime
from io import BytesIO

import openpyxl
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# CRD TESIS CARDIORRENAL — V9
# Supabase central + Gemini (nuevo SDK) + interfaz clínica optimizada
# ============================================================

st.set_page_config(page_title="CRD Tesis Cardiorrenal", page_icon="❤️", layout="wide")

# ---------- Estilo ----------
st.markdown(
    """
    <style>
    .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
    div[data-testid="stMetric"] { padding: .35rem .25rem; }
    .small-muted { color: #6b7280; font-size: 0.86rem; }
    .ai-note { font-size: 0.88rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- Secrets ----------
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "")
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")

if not SUPABASE_URL or not SUPABASE_KEY:
    st.error("Faltan SUPABASE_URL y/o SUPABASE_KEY en los Secrets de Streamlit.")
    st.stop()

if APP_PASSWORD:
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        st.title("CRD Tesis Cardiorrenal")
        st.subheader("Acceso")
        password = st.text_input("Contraseña", type="password")
        if st.button("Entrar", type="primary", use_container_width=True):
            if password == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        st.stop()

# ---------- Cliente Supabase ----------
try:
    from supabase import create_client

    @st.cache_resource(show_spinner=False)
    def get_supabase(url, key):
        return create_client(url, key)

    supabase = get_supabase(SUPABASE_URL, SUPABASE_KEY)
except Exception as e:
    st.error(f"No se pudo conectar con Supabase: {e}")
    st.stop()

# ============================================================
# VARIABLES / ESTADO
# ============================================================

YES_NO_FIELDS = {
    "dm2", "hta", "fa", "epoc", "sd_metab", "tabaco", "enolismo", "hepato",
    "biobanco", "edemas_prueba", "ieca", "ara2", "bb", "amr", "sac_val",
    "sglt2i", "diur_asa", "hctz", "acetazolamida", "estatinas", "epo",
    "m_cv", "hosp_ic", "iam", "acv", "m_tot", "trs", "caida_fge", "sd_cr"
}

DEFAULTS = {
    "id_pac": "",
    "fecha_inc": "",
    "edad": "",
    "sexo": "",
    "peso": "",
    "talla": "",
    "imc": "",
    "eti_erc": "",
    "eti_ic": "",
    "dm2": "",
    "hta": "",
    "fa": "",
    "epoc": "",
    "sd_metab": "",
    "tabaco": "",
    "enolismo": "",
    "hepato": "",
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
    "fib4": "",
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
    "biobanco": "",
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
    "edemas_prueba": "",
    "ieca": "",
    "ara2": "",
    "bb": "",
    "amr": "",
    "sac_val": "",
    "sglt2i": "",
    "diur_asa": "",
    "hctz": "",
    "acetazolamida": "",
    "estatinas": "",
    "epo": "",
    "meses_seg": "",
    "m_cv": "",
    "f_m_cv": "",
    "hosp_ic": "",
    "f_hosp_ic": "",
    "iam": "",
    "f_iam": "",
    "acv": "",
    "f_acv": "",
    "mace_plus": "",
    "m_tot": "",
    "f_m_tot": "",
    "trs": "",
    "f_trs": "",
    "caida_fge": "",
    "f_caida_fge": "",
    "sd_cr": "",
    "f_sd_cr": "",
}

SECTION_FIELDS = {
    "01. Datos clínicos": [
        "fecha_inc", "edad", "sexo", "peso", "talla", "imc", "eti_erc", "eti_ic",
        "dm2", "hta", "fa", "epoc", "sd_metab", "tabaco", "enolismo", "hepato"
    ],
    "02. Analítica / biomarcadores": [
        "hb", "creat", "cist_c", "fge", "urea", "ac_urico", "prot_creat", "ast", "alt", "plaq",
        "fib4", "bili_t", "bili_d", "albumina", "hba1c", "colest", "nt_probnp", "ca125", "gal3",
        "sst2", "gdf15", "biobanco"
    ],
    "03. Ecocardiografía / elastografía": [
        "fevi", "gls", "masa_vi", "tapse", "vai", "vci", "lsm", "cat_fibro", "cap", "med_val",
        "iqr_med", "dias_desc", "nt_prueba", "edemas_prueba"
    ],
    "04. Tratamiento": [
        "ieca", "ara2", "bb", "amr", "sac_val", "sglt2i", "diur_asa", "hctz", "acetazolamida",
        "estatinas", "epo"
    ],
    "05. Seguimiento 24 m": [
        "meses_seg", "m_cv", "f_m_cv", "hosp_ic", "f_hosp_ic", "iam", "f_iam", "acv", "f_acv",
        "mace_plus", "m_tot", "f_m_tot", "trs", "f_trs", "caida_fge", "f_caida_fge", "sd_cr", "f_sd_cr"
    ],
}

DERIVED_FIELDS = {"imc", "fib4", "mace_plus"}

if "form_mode" not in st.session_state:
    st.session_state.form_mode = "new"  # new | edit
if "loaded_patient_id" not in st.session_state:
    st.session_state.loaded_patient_id = ""
if "loaded_updated_at" not in st.session_state:
    st.session_state.loaded_updated_at = ""
if "id_exists" not in st.session_state:
    st.session_state.id_exists = False
if "id_checked" not in st.session_state:
    st.session_state.id_checked = ""
if "ui_message" not in st.session_state:
    st.session_state.ui_message = ""
if "ui_message_type" not in st.session_state:
    st.session_state.ui_message_type = "info"
if "last_ai_data" not in st.session_state:
    st.session_state.last_ai_data = {}
if "ai_changed_keys" not in st.session_state:
    st.session_state.ai_changed_keys = []
if "clinical_text" not in st.session_state:
    st.session_state.clinical_text = ""

for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ============================================================
# DATABASE
# ============================================================

def normalize_patient_id(value):
    return str(value or "").strip().upper()


def db_get_patient(patient_id):
    pid = normalize_patient_id(patient_id)
    if not pid:
        return None
    res = supabase.table("patients").select("*").eq("id_pac", pid).limit(1).execute()
    return res.data[0] if res.data else None


def db_patient_exists(patient_id):
    return db_get_patient(patient_id) is not None


def db_count():
    res = supabase.table("patients").select("id_pac", count="exact").limit(1).execute()
    return res.count if res.count is not None else 0


def db_recent(limit=20):
    return supabase.table("patients").select(
        "id_pac,updated_at,updated_by"
    ).order("updated_at", desc=True).limit(limit).execute().data or []


def db_audit(patient_id, limit=20):
    return supabase.table("audit_log").select(
        "id_pac,action,changed_at,changed_by"
    ).eq("id_pac", normalize_patient_id(patient_id)).order(
        "changed_at", desc=True
    ).limit(limit).execute().data or []


def db_save_patient(data):
    patient_id = normalize_patient_id(data.get("id_pac", ""))
    if not patient_id:
        raise ValueError("El ID de paciente es obligatorio.")

    data = dict(data)
    data["id_pac"] = patient_id
    now = datetime.now().isoformat()
    old = db_get_patient(patient_id)
    mode = st.session_state.form_mode
    loaded_id = normalize_patient_id(st.session_state.loaded_patient_id)

    payload = {
        "id_pac": patient_id,
        "data": data,
        "updated_at": now,
        "updated_by": st.session_state.get("user_label", "usuario"),
    }

    if mode == "new":
        if old:
            raise ValueError(f"DUPLICADO:{patient_id}")
        res = supabase.table("patients").insert(payload).select("*").execute()
        action = "CREATE"
        result_label = "creado"
    else:
        if not old:
            raise ValueError("El paciente cargado ya no existe en la base de datos.")
        if loaded_id != patient_id:
            raise ValueError("La ID del paciente cargado no coincide.")
        loaded_at = st.session_state.get("loaded_updated_at", "")
        current_at = old.get("updated_at", "")
        if loaded_at and current_at and loaded_at != current_at:
            raise ValueError(
                "Este paciente ha sido modificado por otro usuario desde que lo cargaste. "
                "Vuelve a cargarlo antes de guardar para evitar sobrescribir cambios."
            )
        res = (
            supabase.table("patients")
            .update({
                "data": data,
                "updated_at": now,
                "updated_by": st.session_state.get("user_label", "usuario"),
            })
            .eq("id_pac", patient_id)
            .select("*")
            .execute()
        )
        action = "UPDATE"
        result_label = "actualizado"

    if not res.data:
        raise ValueError("Supabase no devolvió el registro guardado.")

    supabase.table("audit_log").insert({
        "id_pac": patient_id,
        "action": action,
        "snapshot": data,
        "changed_at": now,
        "changed_by": st.session_state.get("user_label", "usuario"),
    }).execute()

    st.session_state.form_mode = "edit"
    st.session_state.loaded_patient_id = patient_id
    st.session_state.loaded_updated_at = now
    st.session_state.id_exists = True
    st.session_state.id_checked = patient_id
    return result_label


def db_load_patient(patient_id):
    row = db_get_patient(patient_id)
    if not row:
        return False
    data = row.get("data") or {}
    for k in DEFAULTS:
        st.session_state[k] = data.get(k, DEFAULTS[k])
    st.session_state.id_pac = normalize_patient_id(patient_id)
    st.session_state.form_mode = "edit"
    st.session_state.loaded_patient_id = normalize_patient_id(patient_id)
    st.session_state.loaded_updated_at = row.get("updated_at", "")
    st.session_state.id_exists = True
    st.session_state.id_checked = normalize_patient_id(patient_id)
    st.session_state.ai_changed_keys = []
    recalculate_derived_fields()
    return True


def set_ui_message(message, kind="info"):
    st.session_state.ui_message = message
    st.session_state.ui_message_type = kind


def reset_patient_state():
    for k, v in DEFAULTS.items():
        st.session_state[k] = v
    st.session_state.form_mode = "new"
    st.session_state.loaded_patient_id = ""
    st.session_state.loaded_updated_at = ""
    st.session_state.id_exists = False
    st.session_state.id_checked = ""
    st.session_state.last_ai_data = {}
    st.session_state.ai_changed_keys = []
    st.session_state.clinical_text = ""


def handle_new_patient():
    reset_patient_state()
    set_ui_message("Formulario preparado para un paciente nuevo.", "info")


def handle_id_change():
    pid = normalize_patient_id(st.session_state.get("id_pac", ""))
    st.session_state.id_checked = pid
    st.session_state.id_exists = bool(pid and db_patient_exists(pid))


def handle_load_patient():
    pid = normalize_patient_id(st.session_state.get("id_pac", ""))
    if not pid:
        set_ui_message("Introduce primero una ID de paciente.", "warning")
        return
    if db_load_patient(pid):
        set_ui_message(f"Paciente {pid} cargado. Puedes editarlo.", "success")
    else:
        st.session_state.id_exists = False
        st.session_state.id_checked = pid
        set_ui_message(
            f"El paciente {pid} no existe en la base. Puedes crear una ficha nueva.",
            "warning",
        )


def handle_save_patient():
    pid = normalize_patient_id(st.session_state.get("id_pac", ""))
    if not pid:
        set_ui_message("Introduce primero una ID de paciente.", "error")
        return
    if st.session_state.form_mode == "new" and st.session_state.id_exists:
        set_ui_message(
            f"⚠️ El paciente {pid} ya existe en la base de datos. No se ha modificado nada. "
            "Pulsa 'Cargar paciente' para editarlo.",
            "warning",
        )
        return

    recalculate_derived_fields()
    warnings = validation_warnings()
    if warnings:
        set_ui_message("Hay valores fuera de rango. Revísalos antes de guardar.", "error")
        return

    data = {k: st.session_state.get(k, DEFAULTS[k]) for k in DEFAULTS}
    data["id_pac"] = pid
    try:
        label = db_save_patient(data)
        set_ui_message(f"Paciente {pid} {label} correctamente en la base central.", "success")
    except Exception as e:
        msg = str(e)
        if msg.startswith("DUPLICADO:"):
            set_ui_message(
                f"⚠️ El paciente {pid} ya existe en la base de datos. No se ha modificado nada. "
                "Pulsa 'Cargar paciente' para editarlo.",
                "warning",
            )
        else:
            set_ui_message(f"No se pudo guardar en la base central: {e}", "error")


# ============================================================
# DERIVADAS / VALIDACIÓN
# ============================================================

def _to_float(value):
    try:
        s = str(value).strip().replace(",", ".")
        if s in ("", "None", "nan"):
            return None
        return float(s)
    except Exception:
        return None


def calculate_imc():
    peso = _to_float(st.session_state.get("peso", ""))
    talla_cm = _to_float(st.session_state.get("talla", ""))
    if peso is None or talla_cm is None or peso <= 0 or talla_cm <= 0:
        return ""
    return f"{peso / ((talla_cm / 100.0) ** 2):.2f}"


def calculate_fib4():
    edad = _to_float(st.session_state.get("edad", ""))
    ast = _to_float(st.session_state.get("ast", ""))
    alt = _to_float(st.session_state.get("alt", ""))
    plaq = _to_float(st.session_state.get("plaq", ""))
    if any(v is None for v in (edad, ast, alt, plaq)):
        return ""
    if edad < 0 or ast <= 0 or alt <= 0 or plaq <= 0:
        return ""
    return f"{(edad * ast) / (plaq * math.sqrt(alt)):.2f}"


def calculate_mace_plus():
    events = [st.session_state.get(k, "") for k in ("m_cv", "hosp_ic", "iam", "acv")]
    if any(v == "Sí" for v in events):
        return "Sí"
    if all(v == "No" for v in events):
        return "No"
    return ""


def recalculate_derived_fields():
    st.session_state["imc"] = calculate_imc()
    st.session_state["fib4"] = calculate_fib4()
    st.session_state["mace_plus"] = calculate_mace_plus()


def is_number(v):
    try:
        float(str(v).replace(",", "."))
        return True
    except Exception:
        return False


def validation_warnings():
    warnings = []
    ranges = {
        "edad": (0, 120), "peso": (1, 500), "talla": (30, 250), "hb": (1, 30),
        "creat": (0.1, 30), "fge": (0, 200), "fevi": (0, 100), "tapse": (0, 50),
        "lsm": (0, 100), "cap": (0, 1000), "plaq": (1, 2000), "meses_seg": (0, 120),
    }
    for key, (lo, hi) in ranges.items():
        value = st.session_state.get(key, "")
        if value not in ("", None) and is_number(value):
            number = float(str(value).replace(",", "."))
            if number < lo or number > hi:
                warnings.append(f"{key}: {value} fuera del rango de comprobación ({lo}–{hi}).")
    return warnings


# ============================================================
# COMPLETITUD / PRESENTACIÓN
# ============================================================

def is_filled(value):
    return value not in (None, "", "—")


def field_display_name(key):
    labels = {
        "fecha_inc": "Fecha inclusión", "edad": "Edad", "sexo": "Sexo", "peso": "Peso",
        "talla": "Talla", "imc": "IMC", "eti_erc": "Etiología ERC", "eti_ic": "Etiología IC",
        "dm2": "DM2", "hta": "HTA", "fa": "FA", "epoc": "EPOC", "sd_metab": "Sd. metabólico",
        "tabaco": "Tabaquismo", "enolismo": "Enolismo", "hepato": "Hepatopatía",
        "hb": "Hemoglobina", "creat": "Creatinina", "cist_c": "Cistatina C", "fge": "FGe",
        "urea": "Urea", "ac_urico": "Ácido úrico", "prot_creat": "Prot/Creat", "ast": "AST",
        "alt": "ALT", "plaq": "Plaquetas", "fib4": "FIB-4", "bili_t": "Bilirrubina total",
        "bili_d": "Bilirrubina directa", "albumina": "Albúmina", "hba1c": "HbA1c", "colest": "Colesterol total",
        "nt_probnp": "NT-proBNP", "ca125": "CA125", "gal3": "Galectina-3", "sst2": "sST2", "gdf15": "GDF-15",
        "biobanco": "Biobanco", "fevi": "FEVI", "gls": "GLS", "masa_vi": "Masa VI", "tapse": "TAPSE",
        "vai": "VAI", "vci": "VCI", "lsm": "LSM", "cat_fibro": "Categoría fibrosis", "cap": "CAP",
        "med_val": "Mediciones válidas", "iqr_med": "IQR/mediana", "dias_desc": "Días desde descompensación",
        "nt_prueba": "NT-proBNP día prueba", "edemas_prueba": "Edemas día prueba", "ieca": "IECA", "ara2": "ARA2",
        "bb": "Betabloqueante", "amr": "AMR", "sac_val": "SAC/VAL", "sglt2i": "SGLT2i", "diur_asa": "Diurético de asa",
        "hctz": "HCTZ", "acetazolamida": "Acetazolamida", "estatinas": "Estatinas", "epo": "Eritropoyetina",
        "meses_seg": "Meses seguimiento", "m_cv": "Muerte CV", "f_m_cv": "Fecha muerte CV", "hosp_ic": "Hosp. IC",
        "f_hosp_ic": "Fecha hosp. IC", "iam": "IAM no fatal", "f_iam": "Fecha IAM", "acv": "ACV no fatal", "f_acv": "Fecha ACV",
        "mace_plus": "MACE+", "m_tot": "Muerte total", "f_m_tot": "Fecha muerte total", "trs": "Inicio TRS",
        "f_trs": "Fecha inicio TRS", "caida_fge": "Caída FGe ≥25%", "f_caida_fge": "Fecha caída FGe",
        "sd_cr": "Sd. cardiorrenal agudo", "f_sd_cr": "Fecha Sd. CR agudo",
    }
    return labels.get(key, key)


def section_completion(field_list):
    relevant = []
    for key in field_list:
        # Dates of events count only when event is explicitly Yes.
        if key.startswith("f_"):
            parent = {
                "f_m_cv": "m_cv", "f_hosp_ic": "hosp_ic", "f_iam": "iam", "f_acv": "acv",
                "f_m_tot": "m_tot", "f_trs": "trs", "f_caida_fge": "caida_fge", "f_sd_cr": "sd_cr",
            }.get(key)
            if parent and st.session_state.get(parent, "") != "Sí":
                continue
        # Category fibrosis is optional when not documented.
        if key == "cat_fibro":
            continue
        relevant.append(key)
    completed = sum(1 for k in relevant if is_filled(st.session_state.get(k, "")))
    return completed, len(relevant)


def overall_completion():
    counts = [section_completion(fields) for fields in SECTION_FIELDS.values()]
    total_done = sum(c for c, _ in counts)
    total = sum(t for _, t in counts)
    return total_done, total, (total_done / total if total else 0)


def widget_label(text, key):
    return f"🤖 {text}" if key in st.session_state.ai_changed_keys else text


def tri_state_select(label, key, disabled=False, container=None):
    options = ["", "Sí", "No"]
    target = container if container is not None else st
    return target.selectbox(
        widget_label(label, key),
        options,
        key=key,
        disabled=disabled,
        format_func=lambda x: "— No recogido —" if x == "" else x,
    )


# ============================================================
# IA — GEMINI NUEVO SDK
# ============================================================

AI_KEYS = [k for k in DEFAULTS if k not in DERIVED_FIELDS and k != "id_pac"]
AI_SCHEMA_TEXT = {k: "valor explícito o null" for k in AI_KEYS}
for k in YES_NO_FIELDS:
    AI_SCHEMA_TEXT[k] = "Sí|No|null"


def get_gemini_client():
    if not GEMINI_API_KEY:
        raise RuntimeError("Falta GEMINI_API_KEY en los Secrets de Streamlit.")
    try:
        from google import genai
    except Exception:
        try:
            import importlib
            genai = importlib.import_module("google.genai")
        except Exception as e:
            raise RuntimeError(
                "No está instalado el SDK nuevo de Gemini. Instala 'google-genai' en requirements.txt "
                "y elimina la dependencia antigua 'google-generativeai'."
            ) from e
    return genai.Client(api_key=GEMINI_API_KEY)


@st.cache_resource(show_spinner=False)
def cached_gemini_client(api_key):
    return get_gemini_client()


def normalize_yes_no(value):
    if value is None:
        return None
    s = str(value).strip().lower()
    if s in ("si", "sí", "yes", "true", "presente", "positivo"):
        return "Sí"
    if s in ("no", "false", "ausente", "negativo"):
        return "No"
    return str(value).strip()


def extract_ai(texto):
    client = cached_gemini_client(GEMINI_API_KEY)
    prompt = f"""
Eres un extractor de datos clínicos para una base de investigación cardiorrenal.

OBJETIVO
Extraer exclusivamente datos que estén explícitos en el texto clínico.

REGLAS OBLIGATORIAS
1. No inventes datos.
2. No infieras datos clínicos que no estén explícitos.
3. Si un dato no aparece, devuelve null.
4. Si hay ambigüedad o contradicción, devuelve null para ese campo.
5. No calcules IMC, FIB-4 ni MACE+; esos campos los calcula la aplicación.
6. No conviertas LSM en fibrosis. LSM y categoría de fibrosis son variables independientes.
7. No asignar categoría de fibrosis por umbrales ni por conocimiento médico.
8. No convertir unidades.
9. No interpretar "posible", "sugestivo", "a valorar" como diagnóstico confirmado.
10. En variables Sí/No, solo usar Sí o No si existe evidencia textual explícita.
11. "Nunca fumador" -> No; "fumador" o "exfumador" -> Sí.
12. Para tratamiento, marcar Sí solo si el texto indica que está tomando/recibiendo el tratamiento; si dice explícitamente que no lo toma, No.
13. No extraer la ID de paciente del texto. El ID se controla fuera de la IA.
14. Distingue antecedentes/comorbilidades de eventos del seguimiento.
15. Devuelve SOLO un objeto JSON válido, sin markdown ni comentarios.

CLAVES PERMITIDAS
{json.dumps(AI_SCHEMA_TEXT, ensure_ascii=False, indent=2)}

TEXTO CLÍNICO
{text}
"""

    try:
        from google.genai import types
        response = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0,
                max_output_tokens=6000,
            ),
        )
        raw = (response.text or "").strip()
        if not raw:
            raise ValueError("Gemini devolvió una respuesta vacía.")
        raw = re.sub(r"^```json\s*", "", raw, flags=re.I)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)
        if not isinstance(data, dict):
            raise ValueError("La respuesta de Gemini no es un objeto JSON.")
        return {k: v for k, v in data.items() if k in DEFAULTS and k not in DERIVED_FIELDS and k != "id_pac"}
    except Exception as e:
        st.error(f"Error al extraer datos con Gemini: {e}")
        return None


def apply_ai(data):
    changes = []
    changed_keys = []
    for key, value in data.items():
        if key not in DEFAULTS or key in DERIVED_FIELDS or key == "id_pac" or value in (None, ""):
            continue
        if key in YES_NO_FIELDS:
            value = normalize_yes_no(value)
            if value not in ("Sí", "No"):
                continue
        elif isinstance(value, (int, float)):
            value = str(value)
        else:
            value = str(value).strip()

        old = st.session_state.get(key, "")
        if str(old) != str(value):
            changes.append((key, old, value))
            changed_keys.append(key)
            st.session_state[key] = value

    st.session_state.last_ai_data = data
    st.session_state.ai_changed_keys = changed_keys
    recalculate_derived_fields()
    return changes


# ============================================================
# EXPORTACIÓN EXCEL
# ============================================================

EXPORT_SHEETS = {
    "01_Datos_Clinicos": [
        ("ID Paciente", "id_pac"), ("Fecha Inclusión", "fecha_inc"), ("Edad", "edad"), ("Sexo", "sexo"),
        ("Peso (kg)", "peso"), ("Talla (cm)", "talla"), ("IMC (calc)", "imc"), ("Etiología ERC", "eti_erc"),
        ("Etiología IC", "eti_ic"), ("Comorbilidad: DM2", "dm2"), ("Comorbilidad: HTA", "hta"),
        ("Comorbilidad: FA", "fa"), ("Comorbilidad: EPOC", "epoc"), ("Sd. Metabólico", "sd_metab"),
        ("Tabaquismo", "tabaco"), ("Enolismo", "enolismo"), ("Hepatopatía", "hepato")
    ],
    "02_Analitica_Biomarcadores": [
        ("ID Paciente", "id_pac"), ("Hemoglobina", "hb"), ("Creatinina", "creat"), ("Cistatina C", "cist_c"),
        ("FGe CKD-EPI", "fge"), ("Urea", "urea"), ("Ácido Úrico", "ac_urico"), ("Prot/Creat", "prot_creat"),
        ("AST", "ast"), ("ALT", "alt"), ("Plaquetas (x10^9/L)", "plaq"), ("FIB-4 (calc)", "fib4"),
        ("Bilirrubina Total", "bili_t"), ("Bilirrubina Directa", "bili_d"), ("Albúmina", "albumina"),
        ("HbA1c", "hba1c"), ("Colesterol Total", "colest"), ("NT-proBNP (pg/mL)", "nt_probnp"),
        ("CA125 (U/mL)", "ca125"), ("Galectina-3 (ng/mL)", "gal3"), ("sST2 (ng/mL)", "sst2"),
        ("GDF-15 (pg/mL)", "gdf15"), ("Muestra Biobanco", "biobanco")
    ],
    "03_Eco_Elastografia": [
        ("ID Paciente", "id_pac"), ("FEVI (%)", "fevi"), ("GLS", "gls"), ("Masa VI", "masa_vi"),
        ("TAPSE", "tapse"), ("VAI", "vai"), ("VCI (cm)", "vci"), ("LSM (kPa)", "lsm"),
        ("Categoría Fibrosis", "cat_fibro"), ("CAP (dB/m)", "cap"), ("Mediciones Válidas", "med_val"),
        ("IQR/Mediana ≤0.30", "iqr_med"), ("Días desde última descompensación", "dias_desc"),
        ("NT-proBNP día prueba", "nt_prueba"), ("Edemas (día prueba)", "edemas_prueba")
    ],
    "04_Tratamiento": [
        ("ID Paciente", "id_pac"), ("IECAS", "ieca"), ("ARA2", "ara2"), ("Betabloqueantes", "bb"),
        ("AMR", "amr"), ("SAC/VAL", "sac_val"), ("SGLT2i", "sglt2i"), ("Diurético de asa", "diur_asa"),
        ("HCTZ", "hctz"), ("Acetazolamida", "acetazolamida"), ("Estatinas", "estatinas"), ("Eritropoyetina", "epo")
    ],
    "05_Seguimiento_24m": [
        ("ID Paciente", "id_pac"), ("Meses Seguimiento", "meses_seg"), ("Muerte Cardiovascular", "m_cv"),
        ("Fecha Muerte CV", "f_m_cv"), ("Hosp. IC descompensada", "hosp_ic"), ("Fecha Hosp. IC", "f_hosp_ic"),
        ("IAM no fatal", "iam"), ("Fecha IAM", "f_iam"), ("ACV no fatal", "acv"), ("Fecha ACV", "f_acv"),
        ("MACE+ (Compuesto)", "mace_plus"), ("Muerte Total", "m_tot"), ("Fecha Muerte Total", "f_m_tot"),
        ("Inicio TRS (Diálisis/Tx)", "trs"), ("Fecha Inicio TRS", "f_trs"), ("Caída FGe ≥25%", "caida_fge"),
        ("Fecha Caída FGe", "f_caida_fge"), ("Sd. Cardiorrenal Agudo", "sd_cr"), ("Fecha Sd. CR Agudo", "f_sd_cr")
    ],
}


def export_all_excel():
    rows = supabase.table("patients").select("id_pac,data").order("id_pac").execute().data or []
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "00_Instrucciones"
    ws["A1"] = "CRD Tesis Cardiorrenal"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "La fuente maestra de datos es Supabase. Este archivo es una exportación de trabajo."
    ws["A3"] = "Los campos calculados (IMC, FIB-4, MACE+) se generan en la aplicación."
    ws["A4"] = "Las celdas vacías indican que el dato no está recogido; no equivalen a 'No'."
    ws.column_dimensions["A"].width = 100

    for sheet_name, mapping in EXPORT_SHEETS.items():
        ws = wb.create_sheet(sheet_name)
        for col_idx, (header, _) in enumerate(mapping, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(mapping))}1"

        for r_idx, row in enumerate(rows, start=2):
            data = row.get("data") or {}
            for col_idx, (_, key) in enumerate(mapping, start=1):
                ws.cell(row=r_idx, column=col_idx, value=data.get(key, ""))
        for col_idx, (header, _) in enumerate(mapping, start=1):
            width = min(max(len(header) + 2, 12), 30)
            ws.column_dimensions[get_column_letter(col_idx)].width = width

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), len(rows)


# ============================================================
# CABECERA
# ============================================================

st.title("Tesis Esther Tamarit · Base Cardiorrenal")

recent = db_recent(1)
count = db_count()
last_update = recent[0]["updated_at"] if recent else "—"
done, total, pct = overall_completion()

m1, m2, m3 = st.columns(3)
m1.metric("Pacientes", count)
m2.metric("Recogida actual", f"{done}/{total} · {pct:.0%}")
m3.metric("Estado", "🟢 Base central")

st.progress(pct, text=f"Completitud de la ficha actual: {pct:.0%}")

# ---------- Control de paciente ----------
control1, control2, control3 = st.columns([2, 1, 1])
with control1:
    st.text_input(
        "ID paciente",
        key="id_pac",
        disabled=(st.session_state.form_mode == "edit"),
        on_change=handle_id_change,
        help="Introduce primero la ID. Al salir del campo se comprueba automáticamente si ya existe.",
    )
with control2:
    st.button(
        "🔍 Cargar paciente",
        use_container_width=True,
        on_click=handle_load_patient,
        disabled=(st.session_state.form_mode == "edit" or not normalize_patient_id(st.session_state.get("id_pac", ""))),
    )
with control3:
    st.button("＋ Nuevo paciente", use_container_width=True, on_click=handle_new_patient)

current_id = normalize_patient_id(st.session_state.get("id_pac", ""))

if st.session_state.form_mode == "edit":
    st.success(f"🟢 EDITANDO PACIENTE · ID {st.session_state.loaded_patient_id}")
elif current_id and st.session_state.id_checked == current_id and st.session_state.id_exists:
    st.warning(
        f"⚠️ Este paciente ya existe (ID {current_id}). Cárgalo para editarlo. "
        "El formulario queda bloqueado hasta cargarlo."
    )
    if st.button(f"🔍 Cargar paciente {current_id}", type="primary", use_container_width=True):
        handle_load_patient()
        st.rerun()
    
elif current_id and st.session_state.id_checked == current_id and not st.session_state.id_exists:
    st.success(f"✅ ID {current_id} disponible para un paciente nuevo.")
elif current_id:
    st.info("Pulsa Tab o haz clic fuera del campo ID para comprobar si ya existe.")

if st.session_state.ui_message:
    kind = st.session_state.ui_message_type
    message = st.session_state.ui_message
    st.session_state.ui_message = ""
    st.session_state.ui_message_type = "info"
    if kind == "success":
        st.success(message)
    elif kind == "warning":
        st.warning(message)
    elif kind == "error":
        st.error(message)
    else:
        st.info(message)

form_locked = (
    st.session_state.form_mode == "new"
    and (not current_id or st.session_state.id_exists or st.session_state.id_checked != current_id)
)

# ============================================================
# IA + HISTORIA
# ============================================================

st.markdown("### 🧠 Texto clínico e IA")
left, right = st.columns([1.05, 1.95])

with left:
    st.text_area(
        "Pega aquí la evolución / historia clínica",
        height=270,
        key="clinical_text",
        disabled=form_locked,
        placeholder="Pega el texto clínico anonimizado…",
    )
    st.caption("El texto clínico no se guarda en Supabase.")
    if st.button("🧠 Extraer datos con IA", type="primary", use_container_width=True, disabled=form_locked):
        if not st.session_state.clinical_text.strip():
            st.warning("Pega primero el texto clínico.")
        else:
            with st.spinner("Gemini está extrayendo solo los datos explícitos…"):
                result = extract_ai(st.session_state.clinical_text)
            if result is not None:
                changes = apply_ai(result)
                if changes:
                    st.success(f"{len(changes)} campos propuestos por IA. Revísalos antes de guardar.")
                else:
                    st.info("La IA no encontró nuevos datos explícitos que añadir.")
                st.rerun()

with right:
    if st.session_state.ai_changed_keys:
        st.info(
            "🤖 **Campos propuestos por IA:** "
            + ", ".join(field_display_name(k) for k in st.session_state.ai_changed_keys)
            + ". Revísalos manualmente antes de guardar."
        )
    else:
        st.caption("Los campos propuestos por IA aparecen marcados con 🤖.")

    if st.session_state.last_ai_data:
        with st.expander("Ver respuesta estructurada de IA"):
            st.json(st.session_state.last_ai_data)

# ============================================================
# FORMULARIO CLÍNICO
# ============================================================

section_titles = []
for title, fields in SECTION_FIELDS.items():
    c, t = section_completion(fields)
    section_titles.append(f"{title} · {c}/{t}")

tab1, tab2, tab3, tab4, tab5 = st.tabs(section_titles)

with tab1:
    st.subheader("Datos basales")
    a, b, c = st.columns(3)
    a.text_input(widget_label("Fecha inclusión", "fecha_inc"), key="fecha_inc", disabled=form_locked)
    b.selectbox(widget_label("Sexo", "sexo"), ["", "H", "M"], key="sexo", disabled=form_locked,
                format_func=lambda x: "— No recogido —" if x == "" else x)
    c.text_input(widget_label("Edad", "edad"), key="edad", disabled=form_locked)

    a, b, c, d = st.columns(4)
    a.text_input(widget_label("Peso (kg)", "peso"), key="peso", disabled=form_locked)
    b.text_input(widget_label("Talla (cm)", "talla"), key="talla", disabled=form_locked)
    c.text_input("IMC (calculado)", value=calculate_imc(), disabled=True)
    d.selectbox(widget_label("Etiología ERC", "eti_erc"), ["", "DM2", "HTA", "Glomerulonefritis", "Poliquistosis", "Otras"], key="eti_erc", disabled=form_locked,
                format_func=lambda x: "— No recogida —" if x == "" else x)

    a, b = st.columns(2)
    a.selectbox(widget_label("Etiología IC", "eti_ic"), ["", "Isquémica", "Hipertensiva", "MCD", "HFpEF", "Valvular", "Otras"], key="eti_ic", disabled=form_locked,
                format_func=lambda x: "— No recogida —" if x == "" else x)
    b.caption("No se infieren etiologías a partir de otros datos.")

    st.markdown("**Comorbilidades**")
    a, b, c, d = st.columns(4)
    tri_state_select("DM2", "dm2", form_locked, a)
    tri_state_select("HTA", "hta", form_locked, b)
    tri_state_select("FA", "fa", form_locked, c)
    tri_state_select("EPOC", "epoc", form_locked, d)
    a, b, c, d = st.columns(4)
    with a: tri_state_select("Sd. Metabólico", "sd_metab", form_locked)
    with b: tri_state_select("Tabaquismo", "tabaco", form_locked)
    with c: tri_state_select("Enolismo", "enolismo", form_locked)
    with d: tri_state_select("Hepatopatía", "hepato", form_locked)

with tab2:
    st.subheader("Analítica y biomarcadores")
    a, b, c, d = st.columns(4)
    a.text_input(widget_label("Hemoglobina", "hb"), key="hb", disabled=form_locked)
    b.text_input(widget_label("Creatinina", "creat"), key="creat", disabled=form_locked)
    c.text_input(widget_label("Cistatina C", "cist_c"), key="cist_c", disabled=form_locked)
    d.text_input(widget_label("FGe CKD-EPI", "fge"), key="fge", disabled=form_locked)
    a, b, c, d = st.columns(4)
    a.text_input(widget_label("Urea", "urea"), key="urea", disabled=form_locked)
    b.text_input(widget_label("Ácido úrico", "ac_urico"), key="ac_urico", disabled=form_locked)
    c.text_input(widget_label("Prot/Creat", "prot_creat"), key="prot_creat", disabled=form_locked)
    d.text_input(widget_label("Plaquetas", "plaq"), key="plaq", disabled=form_locked)
    a, b, c, d = st.columns(4)
    a.text_input(widget_label("AST", "ast"), key="ast", disabled=form_locked)
    b.text_input(widget_label("ALT", "alt"), key="alt", disabled=form_locked)
    c.text_input("FIB-4 (calculado)", value=calculate_fib4(), disabled=True)
    d.text_input(widget_label("Bilirrubina total", "bili_t"), key="bili_t", disabled=form_locked)
    a, b, c, d = st.columns(4)
    a.text_input(widget_label("Bilirrubina directa", "bili_d"), key="bili_d", disabled=form_locked)
    b.text_input(widget_label("Albúmina", "albumina"), key="albumina", disabled=form_locked)
    c.text_input(widget_label("HbA1c", "hba1c"), key="hba1c", disabled=form_locked)
    d.text_input(widget_label("Colesterol total", "colest"), key="colest", disabled=form_locked)

    st.markdown("**Biomarcadores**")
    a, b, c, d = st.columns(4)
    a.text_input(widget_label("NT-proBNP (pg/mL)", "nt_probnp"), key="nt_probnp", disabled=form_locked)
    b.text_input(widget_label("CA125 (U/mL)", "ca125"), key="ca125", disabled=form_locked)
    c.text_input(widget_label("Galectina-3 (ng/mL)", "gal3"), key="gal3", disabled=form_locked)
    d.text_input(widget_label("sST2 (ng/mL)", "sst2"), key="sst2", disabled=form_locked)
    a, b = st.columns(2)
    a.text_input(widget_label("GDF-15 (pg/mL)", "gdf15"), key="gdf15", disabled=form_locked)
    with b:
        tri_state_select("Muestra biobanco", "biobanco", form_locked)

with tab3:
    st.subheader("Ecocardiografía / elastografía")
    a, b, c, d = st.columns(4)
    a.text_input(widget_label("FEVI (%)", "fevi"), key="fevi", disabled=form_locked)
    b.text_input(widget_label("GLS", "gls"), key="gls", disabled=form_locked)
    c.text_input(widget_label("Masa VI", "masa_vi"), key="masa_vi", disabled=form_locked)
    d.text_input(widget_label("TAPSE", "tapse"), key="tapse", disabled=form_locked)
    a, b, c, d = st.columns(4)
    a.text_input(widget_label("VAI", "vai"), key="vai", disabled=form_locked)
    b.text_input(widget_label("VCI (cm)", "vci"), key="vci", disabled=form_locked)
    c.text_input(widget_label("LSM (kPa)", "lsm"), key="lsm", disabled=form_locked)
    d.text_input(widget_label("CAP (dB/m)", "cap"), key="cap", disabled=form_locked)
    a, b, c, d = st.columns(4)
    a.text_input(widget_label("Mediciones válidas", "med_val"), key="med_val", disabled=form_locked)
    b.text_input(widget_label("IQR/Mediana ≤0.30", "iqr_med"), key="iqr_med", disabled=form_locked)
    c.text_input(widget_label("Días desde última descompensación", "dias_desc"), key="dias_desc", disabled=form_locked)
    d.text_input(widget_label("NT-proBNP día prueba", "nt_prueba"), key="nt_prueba", disabled=form_locked)
    a, b = st.columns(2)
    a.text_input(widget_label("Categoría fibrosis (solo si documentada)", "cat_fibro"), key="cat_fibro", disabled=form_locked)
    with b:
        tri_state_select("Edemas (día prueba)", "edemas_prueba", form_locked)
    st.info("LSM y categoría de fibrosis se mantienen separadas. La IA no infiere fibrosis a partir de LSM.")

with tab4:
    st.subheader("Tratamiento")
    a, b, c, d = st.columns(4)
    tri_state_select("IECA", "ieca", form_locked, a)
    tri_state_select("ARA2", "ara2", form_locked, b)
    tri_state_select("Betabloqueante", "bb", form_locked, c)
    tri_state_select("AMR", "amr", form_locked, d)
    a, b, c, d = st.columns(4)
    tri_state_select("SAC/VAL", "sac_val", form_locked, a)
    tri_state_select("SGLT2i", "sglt2i", form_locked, b)
    tri_state_select("Diurético de asa", "diur_asa", form_locked, c)
    tri_state_select("HCTZ", "hctz", form_locked, d)
    a, b, c = st.columns(3)
    with a: tri_state_select("Acetazolamida", "acetazolamida", form_locked)
    with b: tri_state_select("Estatinas", "estatinas", form_locked)
    with c: tri_state_select("Eritropoyetina", "epo", form_locked)

with tab5:
    st.subheader("Seguimiento 24 meses")
    st.text_input(widget_label("Meses de seguimiento", "meses_seg"), key="meses_seg", disabled=form_locked)

    st.markdown("**Eventos cardiovasculares**")
    a, b = st.columns(2)
    with a:
        tri_state_select("Muerte cardiovascular", "m_cv", form_locked, a)
        if st.session_state.m_cv == "Sí":
            st.text_input("Fecha muerte CV", key="f_m_cv", disabled=form_locked)
    with b:
        tri_state_select("Hospitalización por IC descompensada", "hosp_ic", form_locked, b)
        if st.session_state.hosp_ic == "Sí":
            st.text_input("Fecha hospitalización IC", key="f_hosp_ic", disabled=form_locked)

    a, b = st.columns(2)
    with a:
        tri_state_select("IAM no fatal", "iam", form_locked, a)
        if st.session_state.iam == "Sí":
            st.text_input("Fecha IAM", key="f_iam", disabled=form_locked)
    with b:
        tri_state_select("ACV no fatal", "acv", form_locked, b)
        if st.session_state.acv == "Sí":
            st.text_input("Fecha ACV", key="f_acv", disabled=form_locked)

    st.info(f"**MACE+ (calculado):** {calculate_mace_plus() or '— no determinable con los datos actuales —'}")

    st.markdown("**Otros desenlaces**")
    a, b = st.columns(2)
    with a:
        tri_state_select("Muerte total", "m_tot", form_locked, a)
        if st.session_state.m_tot == "Sí":
            st.text_input("Fecha muerte total", key="f_m_tot", disabled=form_locked)
    with b:
        tri_state_select("Inicio TRS (diálisis/trasplante)", "trs", form_locked, b)
        if st.session_state.trs == "Sí":
            st.text_input("Fecha inicio TRS", key="f_trs", disabled=form_locked)
    a, b = st.columns(2)
    with a:
        tri_state_select("Caída FGe ≥25%", "caida_fge", form_locked, a)
        if st.session_state.caida_fge == "Sí":
            st.text_input("Fecha caída FGe", key="f_caida_fge", disabled=form_locked)
    with b:
        tri_state_select("Síndrome cardiorrenal agudo", "sd_cr", form_locked, b)
        if st.session_state.sd_cr == "Sí":
            st.text_input("Fecha Sd. CR agudo", key="f_sd_cr", disabled=form_locked)

# Recalcular después de renderizar valores modificables.
recalculate_derived_fields()

# ============================================================
# REVISIÓN / GUARDADO / EXPORTACIÓN
# ============================================================

st.markdown("---")
done, total, pct = overall_completion()
missing = []
for title, fields in SECTION_FIELDS.items():
    c, t = section_completion(fields)
    if c < t:
        missing.append((title, t - c))

st.markdown("### Revisión antes de guardar")
r1, r2, r3 = st.columns(3)
r1.metric("Variables recogidas", f"{done}/{total}")
r2.metric("Completitud", f"{pct:.0%}")
r3.metric("ID", current_id or "—")

if st.session_state.form_mode == "new" and current_id and st.session_state.id_exists:
    st.warning("Este paciente ya existe. Cárgalo antes de introducir o guardar datos.")
elif missing:
    st.caption("Pendientes de recoger: " + " · ".join(f"{title} ({n})" for title, n in missing))
else:
    st.success("Ficha completa según las variables monitorizadas por la aplicación.")

save_disabled = form_locked or not current_id or (st.session_state.form_mode == "new" and st.session_state.id_exists)

b1, b2, b3 = st.columns(3)
with b1:
    st.button("💾 GUARDAR PACIENTE", type="primary", use_container_width=True, on_click=handle_save_patient, disabled=save_disabled)
with b2:
    try:
        excel_bytes, n = export_all_excel()
        st.download_button(
            "⬇️ EXPORTAR EXCEL",
            data=excel_bytes,
            file_name=f"CRD_Tesis_Cardiorrenal_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"No se pudo preparar el Excel: {e}")
with b3:
    st.button("🔄 Refrescar", use_container_width=True, on_click=lambda: st.rerun())

# ============================================================
# HISTORIAL / ÚLTIMOS PACIENTES
# ============================================================

if current_id:
    with st.expander("🕒 Historial de cambios del paciente"):
        try:
            history = db_audit(current_id)
            if history:
                st.dataframe(history, use_container_width=True, hide_index=True)
            else:
                st.caption("Sin historial.")
        except Exception as e:
            st.caption(f"No se pudo cargar el historial: {e}")

with st.expander("👥 Últimos pacientes modificados"):
    try:
        st.dataframe(db_recent(25), use_container_width=True, hide_index=True)
    except Exception as e:
        st.caption(f"No se pudo cargar la lista: {e}")

st.caption("La base central es la fuente maestra. El Excel es una exportación para análisis/copia.")
