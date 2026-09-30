import json
import math
import re
import time
from datetime import datetime, date
from io import BytesIO

import openpyxl
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# CRD TESIS CARDIORRENAL — V10.1
# Interfaz clínica + Supabase central + Gemini + Dashboard
# ============================================================

st.set_page_config(
    page_title="CRD Tesis Cardiorrenal",
    page_icon="❤️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.0rem; padding-bottom: 2.5rem; max-width: 1500px; }
    .hero { padding: 0.2rem 0 0.8rem 0; }
    .hero h1 { margin-bottom: .15rem; font-size: 2rem; }
    .hero p { color:#6b7280; margin-top:0; }
    .section-card { border:1px solid #e5e7eb; border-radius:12px; padding:14px 16px; margin-bottom:12px; background:#ffffff; }
    .mini-label { color:#6b7280; font-size:.78rem; text-transform:uppercase; letter-spacing:.04em; }
    .status-ok { color:#166534; font-weight:600; }
    .status-warn { color:#92400e; font-weight:600; }
    .status-blue { color:#1d4ed8; font-weight:600; }
    .ai-note { font-size:.88rem; }
    div[data-testid="stMetric"] { background:#f8fafc; border:1px solid #e5e7eb; border-radius:10px; padding:.55rem .7rem; }
    .stTabs [data-baseweb="tab-list"] { gap: 0.35rem; }
    .stTabs [data-baseweb="tab"] { padding: .55rem .9rem; }
    div[role="radiogroup"] { gap: .35rem; }
    div[role="radiogroup"] > label {
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: .45rem .8rem;
        background: #ffffff;
    }
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
        st.markdown('<div class="hero"><h1>CRD Tesis Cardiorrenal</h1><p>Acceso a la base de recogida clínica</p></div>', unsafe_allow_html=True)
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
    "id_pac": "", "fecha_inc": "", "edad": "", "sexo": "", "peso": "", "talla": "", "imc": "",
    "eti_erc": "", "eti_ic": "", "dm2": "", "hta": "", "fa": "", "epoc": "", "sd_metab": "",
    "tabaco": "", "enolismo": "", "hepato": "", "hb": "", "creat": "", "cist_c": "", "fge": "",
    "urea": "", "ac_urico": "", "prot_creat": "", "ast": "", "alt": "", "plaq": "", "fib4": "",
    "bili_t": "", "bili_d": "", "albumina": "", "hba1c": "", "colest": "", "nt_probnp": "",
    "ca125": "", "gal3": "", "sst2": "", "gdf15": "", "biobanco": "", "fevi": "", "gls": "",
    "masa_vi": "", "tapse": "", "vai": "", "vci": "", "lsm": "", "cat_fibro": "", "cap": "",
    "med_val": "", "iqr_med": "", "dias_desc": "", "nt_prueba": "", "edemas_prueba": "", "ieca": "",
    "ara2": "", "bb": "", "amr": "", "sac_val": "", "sglt2i": "", "diur_asa": "", "hctz": "",
    "acetazolamida": "", "estatinas": "", "epo": "", "meses_seg": "", "m_cv": "", "f_m_cv": "",
    "hosp_ic": "", "f_hosp_ic": "", "iam": "", "f_iam": "", "acv": "", "f_acv": "", "mace_plus": "",
    "m_tot": "", "f_m_tot": "", "trs": "", "f_trs": "", "caida_fge": "", "f_caida_fge": "", "sd_cr": "", "f_sd_cr": "",
}

SECTION_FIELDS = {
    "01. Datos clínicos": [
        "fecha_inc", "edad", "sexo", "peso", "talla", "imc", "eti_erc", "eti_ic", "dm2", "hta", "fa", "epoc", "sd_metab", "tabaco", "enolismo", "hepato"
    ],
    "02. Analítica / biomarcadores": [
        "hb", "creat", "cist_c", "fge", "urea", "ac_urico", "prot_creat", "ast", "alt", "plaq", "fib4", "bili_t", "bili_d", "albumina", "hba1c", "colest", "nt_probnp", "ca125", "gal3", "sst2", "gdf15", "biobanco"
    ],
    "03. Ecocardiografía / elastografía": [
        "fevi", "gls", "masa_vi", "tapse", "vai", "vci", "lsm", "cat_fibro", "cap", "med_val", "iqr_med", "dias_desc", "nt_prueba", "edemas_prueba"
    ],
    "04. Tratamiento": [
        "ieca", "ara2", "bb", "amr", "sac_val", "sglt2i", "diur_asa", "hctz", "acetazolamida", "estatinas", "epo"
    ],
    "05. Seguimiento 24 m": [
        "meses_seg", "m_cv", "f_m_cv", "hosp_ic", "f_hosp_ic", "iam", "f_iam", "acv", "f_acv", "mace_plus", "m_tot", "f_m_tot", "trs", "f_trs", "caida_fge", "f_caida_fge", "sd_cr", "f_sd_cr"
    ],
}

DERIVED_FIELDS = {"imc", "fib4", "mace_plus"}
DATE_FIELDS = {"fecha_inc", "f_m_cv", "f_hosp_ic", "f_iam", "f_acv", "f_m_tot", "f_trs", "f_caida_fge", "f_sd_cr"}
AI_MODELS = ["gemini-3.8-flash", "gemini-3.7-flash", "gemini-3.6-flash"]

FIELD_LABELS = {
    "fecha_inc": "Fecha inclusión", "edad": "Edad", "sexo": "Sexo", "peso": "Peso", "talla": "Talla", "imc": "IMC",
    "eti_erc": "Etiología ERC", "eti_ic": "Etiología IC", "dm2": "DM2", "hta": "HTA", "fa": "FA", "epoc": "EPOC",
    "sd_metab": "Sd. metabólico", "tabaco": "Tabaquismo", "enolismo": "Enolismo", "hepato": "Hepatopatía",
    "hb": "Hemoglobina", "creat": "Creatinina", "cist_c": "Cistatina C", "fge": "FGe", "urea": "Urea", "ac_urico": "Ácido úrico",
    "prot_creat": "Prot/Creat", "ast": "AST", "alt": "ALT", "plaq": "Plaquetas", "fib4": "FIB-4", "bili_t": "Bilirrubina total",
    "bili_d": "Bilirrubina directa", "albumina": "Albúmina", "hba1c": "HbA1c", "colest": "Colesterol total",
    "nt_probnp": "NT-proBNP", "ca125": "CA125", "gal3": "Galectina-3", "sst2": "sST2", "gdf15": "GDF-15", "biobanco": "Biobanco",
    "fevi": "FEVI", "gls": "GLS", "masa_vi": "Masa VI", "tapse": "TAPSE", "vai": "VAI", "vci": "VCI", "lsm": "LSM",
    "cat_fibro": "Categoría fibrosis", "cap": "CAP", "med_val": "Mediciones válidas", "iqr_med": "IQR/mediana", "dias_desc": "Días desde descompensación",
    "nt_prueba": "NT-proBNP día prueba", "edemas_prueba": "Edemas día prueba", "ieca": "IECA", "ara2": "ARA2", "bb": "Betabloqueante",
    "amr": "AMR", "sac_val": "SAC/VAL", "sglt2i": "SGLT2i", "diur_asa": "Diurético de asa", "hctz": "HCTZ", "acetazolamida": "Acetazolamida",
    "estatinas": "Estatinas", "epo": "Eritropoyetina", "meses_seg": "Meses seguimiento", "m_cv": "Muerte CV", "f_m_cv": "Fecha muerte CV",
    "hosp_ic": "Hosp. IC", "f_hosp_ic": "Fecha hosp. IC", "iam": "IAM no fatal", "f_iam": "Fecha IAM", "acv": "ACV no fatal", "f_acv": "Fecha ACV",
    "mace_plus": "MACE+", "m_tot": "Muerte total", "f_m_tot": "Fecha muerte total", "trs": "Inicio TRS", "f_trs": "Fecha inicio TRS",
    "caida_fge": "Caída FGe ≥25%", "f_caida_fge": "Fecha caída FGe", "sd_cr": "Sd. cardiorrenal agudo", "f_sd_cr": "Fecha Sd. CR agudo",
}

for state_key, initial in {
    "form_mode": "new", "loaded_patient_id": "", "loaded_updated_at": "", "id_exists": False, "id_checked": "",
    "ui_message": "", "ui_message_type": "info", "last_ai_data": {}, "ai_changed_keys": [], "clinical_text": "",
    "ai_model_used": "", "dashboard_cache": [], "patient_search": "", "selected_patient_row": "",
}.items():
    if state_key not in st.session_state:
        st.session_state[state_key] = initial

for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v
for date_key in DATE_FIELDS:
    if f"_date_{date_key}" not in st.session_state:
        st.session_state[f"_date_{date_key}"] = None

# ============================================================
# HELPERS / DATABASE
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
    return supabase.table("patients").select("id_pac,updated_at,updated_by").order("updated_at", desc=True).limit(limit).execute().data or []


def db_all_patients():
    # For the thesis scale this is deliberately simple and transparent.
    return supabase.table("patients").select("id_pac,data,updated_at,updated_by").order("id_pac").execute().data or []


def db_audit(patient_id, limit=50):
    return supabase.table("audit_log").select("id_pac,action,changed_at,changed_by").eq("id_pac", normalize_patient_id(patient_id)).order("changed_at", desc=True).limit(limit).execute().data or []


def parse_date_value(value):
    if value in (None, ""):
        return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def set_date_widget_state(key, value):
    st.session_state[f"_date_{key}"] = parse_date_value(value)


def sync_date_field(key, value):
    st.session_state[key] = value.isoformat() if value else ""


def date_input_field(label, key, disabled=False, container=None):
    target = container if container is not None else st
    widget_key = f"_date_{key}"
    if widget_key not in st.session_state:
        st.session_state[widget_key] = parse_date_value(st.session_state.get(key, ""))
    selected = target.date_input(widget_label(label, key), key=widget_key, format="DD/MM/YYYY", disabled=disabled)
    sync_date_field(key, selected)
    return selected


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


def is_filled(value):
    return value not in (None, "", "—")


def section_completion(field_list):
    relevant = []
    for key in field_list:
        if key.startswith("f_"):
            parent = {"f_m_cv":"m_cv","f_hosp_ic":"hosp_ic","f_iam":"iam","f_acv":"acv","f_m_tot":"m_tot","f_trs":"trs","f_caida_fge":"caida_fge","f_sd_cr":"sd_cr"}.get(key)
            if parent and st.session_state.get(parent, "") != "Sí":
                continue
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
    target = container if container is not None else st
    return target.selectbox(
        widget_label(label, key), ["", "Sí", "No"], key=key, disabled=disabled,
        format_func=lambda x: "— No recogido —" if x == "" else x,
    )


def validation_warnings():
    warnings = []
    ranges = {
        "edad": (0, 120), "peso": (1, 500), "talla": (30, 250), "hb": (1, 30), "creat": (0.1, 30), "fge": (0, 200),
        "fevi": (0, 100), "tapse": (0, 50), "lsm": (0, 100), "cap": (0, 1000), "plaq": (1, 2000), "meses_seg": (0, 120),
    }
    for key, (lo, hi) in ranges.items():
        value = st.session_state.get(key, "")
        if value not in ("", None):
            number = _to_float(value)
            if number is not None and (number < lo or number > hi):
                warnings.append(f"{FIELD_LABELS.get(key, key)}: {value} fuera del rango de comprobación ({lo}–{hi}).")
    return warnings

# ============================================================
# CRUD CALLBACKS
# ============================================================

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
    payload = {"id_pac": patient_id, "data": data, "updated_at": now, "updated_by": st.session_state.get("user_label", "usuario")}
    if mode == "new":
        if old:
            raise ValueError(f"DUPLICADO:{patient_id}")
        res = supabase.table("patients").insert(payload).select("*").execute()
        action, result_label = "CREATE", "creado"
    else:
        if not old:
            raise ValueError("El paciente cargado ya no existe en la base de datos.")
        if loaded_id != patient_id:
            raise ValueError("La ID del paciente cargado no coincide.")
        loaded_at = st.session_state.get("loaded_updated_at", "")
        current_at = old.get("updated_at", "")
        if loaded_at and current_at and loaded_at != current_at:
            raise ValueError("Este paciente ha sido modificado por otro usuario desde que lo cargaste. Vuelve a cargarlo antes de guardar.")
        res = supabase.table("patients").update({"data": data, "updated_at": now, "updated_by": st.session_state.get("user_label", "usuario")}).eq("id_pac", patient_id).select("*").execute()
        action, result_label = "UPDATE", "actualizado"
    if not res.data:
        raise ValueError("Supabase no devolvió el registro guardado.")
    supabase.table("audit_log").insert({"id_pac": patient_id, "action": action, "snapshot": data, "changed_at": now, "changed_by": st.session_state.get("user_label", "usuario")}).execute()
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
    for key in DATE_FIELDS:
        set_date_widget_state(key, st.session_state.get(key, ""))
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
    for key in DATE_FIELDS:
        set_date_widget_state(key, "")
    st.session_state.form_mode = "new"
    st.session_state.loaded_patient_id = ""
    st.session_state.loaded_updated_at = ""
    st.session_state.id_exists = False
    st.session_state.id_checked = ""
    st.session_state.last_ai_data = {}
    st.session_state.ai_changed_keys = []
    st.session_state.clinical_text = ""
    st.session_state.ai_model_used = ""


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
        set_ui_message(f"El paciente {pid} no existe en la base. Puedes crear una ficha nueva.", "warning")


def handle_save_patient():
    pid = normalize_patient_id(st.session_state.get("id_pac", ""))
    if not pid:
        set_ui_message("Introduce primero una ID de paciente.", "error")
        return
    if st.session_state.form_mode == "new" and st.session_state.id_exists:
        set_ui_message(f"⚠️ El paciente {pid} ya existe en la base de datos. No se ha modificado nada. Pulsa 'Cargar paciente' para editarlo.", "warning")
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
            set_ui_message(f"⚠️ El paciente {pid} ya existe en la base de datos. No se ha modificado nada. Pulsa 'Cargar paciente' para editarlo.", "warning")
        else:
            set_ui_message(f"No se pudo guardar en la base central: {e}", "error")

# ============================================================
# GEMINI
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
            raise RuntimeError("No está instalado el SDK nuevo de Gemini. Instala 'google-genai' en requirements.txt.") from e
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


def is_transient_gemini_error(exc):
    msg = str(exc).upper()
    return any(token in msg for token in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "500", "INTERNAL"))


def generate_gemini_json(client, prompt):
    from google.genai import types
    last_error = None
    for idx, model in enumerate(AI_MODELS):
        try:
            response = client.models.generate_content(
                model=model,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    max_output_tokens=6000,
                    thinking_config=types.ThinkingConfig(thinking_level="low"),
                ),
            )
            return response, model
        except Exception as exc:
            last_error = exc
            if not is_transient_gemini_error(exc) or idx == len(AI_MODELS) - 1:
                raise
            time.sleep(1.5)
    raise last_error


def extract_ai(texto):
    client = cached_gemini_client(GEMINI_API_KEY)
    prompt = f"""
Eres un extractor de datos clínicos para una base de investigación cardiorrenal.
Extrae exclusivamente datos explícitos. No inventes, no infieras, no calcules IMC/FIB-4/MACE+, no conviertas LSM en fibrosis, no conviertas unidades y no trates posibilidades como diagnósticos confirmados. En Sí/No solo usar Sí/No si existe evidencia textual explícita. Devuelve SOLO JSON válido.

CLAVES PERMITIDAS:
{json.dumps(AI_SCHEMA_TEXT, ensure_ascii=False, indent=2)}

TEXTO CLÍNICO:
{texto}
"""
    try:
        response, model_used = generate_gemini_json(client, prompt)
        raw = (response.text or "").strip()
        if not raw:
            raise ValueError("Gemini devolvió una respuesta vacía.")
        raw = re.sub(r"^```json\s*", "", raw, flags=re.I)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)
        if not isinstance(data, dict):
            raise ValueError("La respuesta de Gemini no es un objeto JSON.")
        st.session_state.ai_model_used = model_used
        return {k: v for k, v in data.items() if k in DEFAULTS and k not in DERIVED_FIELDS and k != "id_pac"}
    except Exception as e:
        if is_transient_gemini_error(e):
            st.error("Gemini está temporalmente saturado. Se han probado los modelos configurados. Vuelve a intentarlo en unos segundos.")
        else:
            st.error(f"Error al extraer datos con Gemini: {e}")
        return None


def apply_ai(data):
    changes, changed_keys = [], []
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
# DASHBOARD HELPERS
# ============================================================

def patient_rows_flat():
    rows = db_all_patients()
    flat = []
    for row in rows:
        d = row.get("data") or {}
        flat.append({"id_pac": row.get("id_pac", ""), "updated_at": row.get("updated_at", ""), "updated_by": row.get("updated_by", ""), **d})
    return flat


def numeric_values(rows, key):
    vals = []
    for r in rows:
        x = _to_float(r.get(key, ""))
        if x is not None:
            vals.append(x)
    return vals


def median(values):
    values = sorted(values)
    if not values:
        return None
    n = len(values)
    mid = n // 2
    return values[mid] if n % 2 else (values[mid - 1] + values[mid]) / 2


def count_yes(rows, key):
    return sum(1 for r in rows if r.get(key) == "Sí")


def count_value(rows, key):
    out = {}
    for r in rows:
        value = r.get(key, "")
        if value not in (None, ""):
            out[value] = out.get(value, 0) + 1
    return dict(sorted(out.items(), key=lambda x: (-x[1], str(x[0]))))


def dataframe_from_records(records):
    try:
        import pandas as pd
        return pd.DataFrame(records)
    except Exception:
        return records


def cohort_completeness(rows):
    out = []
    for r in rows:
        total_done = 0
        total = 0
        for fields in SECTION_FIELDS.values():
            relevant = []
            for key in fields:
                if key.startswith("f_"):
                    parent = {"f_m_cv":"m_cv","f_hosp_ic":"hosp_ic","f_iam":"iam","f_acv":"acv","f_m_tot":"m_tot","f_trs":"trs","f_caida_fge":"caida_fge","f_sd_cr":"sd_cr"}.get(key)
                    if parent and r.get(parent, "") != "Sí":
                        continue
                if key == "cat_fibro":
                    continue
                relevant.append(key)
            total_done += sum(1 for k in relevant if is_filled(r.get(k, "")))
            total += len(relevant)
        out.append((r.get("id_pac", ""), total_done, total, total_done / total if total else 0))
    return out

# ============================================================
# EXPORT EXCEL
# ============================================================

EXPORT_SHEETS = {
    "01_Datos_Clinicos": [("ID Paciente", "id_pac"), ("Fecha Inclusión", "fecha_inc"), ("Edad", "edad"), ("Sexo", "sexo"), ("Peso (kg)", "peso"), ("Talla (cm)", "talla"), ("IMC (calc)", "imc"), ("Etiología ERC", "eti_erc"), ("Etiología IC", "eti_ic"), ("Comorbilidad: DM2", "dm2"), ("Comorbilidad: HTA", "hta"), ("Comorbilidad: FA", "fa"), ("Comorbilidad: EPOC", "epoc"), ("Sd. Metabólico", "sd_metab"), ("Tabaquismo", "tabaco"), ("Enolismo", "enolismo"), ("Hepatopatía", "hepato")],
    "02_Analitica_Biomarcadores": [("ID Paciente", "id_pac"), ("Hemoglobina", "hb"), ("Creatinina", "creat"), ("Cistatina C", "cist_c"), ("FGe CKD-EPI", "fge"), ("Urea", "urea"), ("Ácido Úrico", "ac_urico"), ("Prot/Creat", "prot_creat"), ("AST", "ast"), ("ALT", "alt"), ("Plaquetas (x10^9/L)", "plaq"), ("FIB-4 (calc)", "fib4"), ("Bilirrubina Total", "bili_t"), ("Bilirrubina Directa", "bili_d"), ("Albúmina", "albumina"), ("HbA1c", "hba1c"), ("Colesterol Total", "colest"), ("NT-proBNP (pg/mL)", "nt_probnp"), ("CA125 (U/mL)", "ca125"), ("Galectina-3 (ng/mL)", "gal3"), ("sST2 (ng/mL)", "sst2"), ("GDF-15 (pg/mL)", "gdf15"), ("Muestra Biobanco", "biobanco")],
    "03_Eco_Elastografia": [("ID Paciente", "id_pac"), ("FEVI (%)", "fevi"), ("GLS", "gls"), ("Masa VI", "masa_vi"), ("TAPSE", "tapse"), ("VAI", "vai"), ("VCI (cm)", "vci"), ("LSM (kPa)", "lsm"), ("Categoría Fibrosis", "cat_fibro"), ("CAP (dB/m)", "cap"), ("Mediciones Válidas", "med_val"), ("IQR/Mediana ≤0.30", "iqr_med"), ("Días desde última descompensación", "dias_desc"), ("NT-proBNP día prueba", "nt_prueba"), ("Edemas (día prueba)", "edemas_prueba")],
    "04_Tratamiento": [("ID Paciente", "id_pac"), ("IECAS", "ieca"), ("ARA2", "ara2"), ("Betabloqueantes", "bb"), ("AMR", "amr"), ("SAC/VAL", "sac_val"), ("SGLT2i", "sglt2i"), ("Diurético de asa", "diur_asa"), ("HCTZ", "hctz"), ("Acetazolamida", "acetazolamida"), ("Estatinas", "estatinas"), ("Eritropoyetina", "epo")],
    "05_Seguimiento_24m": [("ID Paciente", "id_pac"), ("Meses Seguimiento", "meses_seg"), ("Muerte Cardiovascular", "m_cv"), ("Fecha Muerte CV", "f_m_cv"), ("Hosp. IC descompensada", "hosp_ic"), ("Fecha Hosp. IC", "f_hosp_ic"), ("IAM no fatal", "iam"), ("Fecha IAM", "f_iam"), ("ACV no fatal", "acv"), ("Fecha ACV", "f_acv"), ("MACE+ (Compuesto)", "mace_plus"), ("Muerte Total", "m_tot"), ("Fecha Muerte Total", "f_m_tot"), ("Inicio TRS (Diálisis/Tx)", "trs"), ("Fecha Inicio TRS", "f_trs"), ("Caída FGe ≥25%", "caida_fge"), ("Fecha Caída FGe", "f_caida_fge"), ("Sd. Cardiorrenal Agudo", "sd_cr"), ("Fecha Sd. CR Agudo", "f_sd_cr")],
}


def export_all_excel():
    rows = supabase.table("patients").select("id_pac,data").order("id_pac").execute().data or []
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "00_Instrucciones"
    ws["A1"] = "CRD Tesis Cardiorrenal"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Fuente maestra: Supabase. Este archivo es una exportación de trabajo."
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
        for r_idx, row in enumerate(rows, start=2):
            data = row.get("data") or {}
            for col_idx, (_, key) in enumerate(mapping, start=1):
                ws.cell(row=r_idx, column=col_idx, value=data.get(key, ""))
        ws.auto_filter.ref = f"A1:{get_column_letter(len(mapping))}{max(1, len(rows)+1)}"
        for col_idx, (header, _) in enumerate(mapping, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(len(header)+2, 12), 30)
    output = BytesIO(); wb.save(output); output.seek(0)
    return output.getvalue(), len(rows)

# ============================================================
# TOP HEADER + MAIN NAVIGATION
# ============================================================

st.markdown('<div class="hero"><h1>❤️ CRD Tesis Cardiorrenal</h1><p>Recogida clínica · base central · control de calidad · dashboard</p></div>', unsafe_allow_html=True)

count = db_count()
recent = db_recent(1)
last_update = recent[0]["updated_at"] if recent else "—"

top1, top2, top3, top4 = st.columns(4)
top1.metric("Pacientes", count)
top2.metric("Última actualización", last_update[:16].replace("T", " ") if isinstance(last_update, str) else "—")
top3.metric("Base", "🟢 Conectada")
active_patient_display = normalize_patient_id(st.session_state.get("loaded_patient_id", "")) or normalize_patient_id(st.session_state.get("id_pac", ""))
top4.metric("Paciente activo", active_patient_display or "—")

# Navegación persistente: a diferencia de st.tabs(), el valor queda guardado
# en session_state y no vuelve a la primera sección cuando hay un rerun.
MAIN_NAV = ["recogida", "dashboard", "pacientes", "admin"]
MAIN_NAV_LABELS = {
    "recogida": "📝 Recogida clínica",
    "dashboard": "📊 Dashboard",
    "pacientes": "👥 Pacientes",
    "admin": "⚙️ Administración",
}
if "main_nav" not in st.session_state:
    st.session_state.main_nav = "recogida"
main_nav = st.radio(
    "Sección principal",
    MAIN_NAV,
    key="main_nav",
    horizontal=True,
    label_visibility="collapsed",
    format_func=lambda x: MAIN_NAV_LABELS[x],
)

# ============================================================
# TAB 1 — RECOGIDA CLÍNICA
# ============================================================
if main_nav == "recogida":
    rec1, rec2, rec3 = st.columns([2.2, 1, 1])
    with rec1:
        st.text_input(
            "ID paciente",
            key="id_pac",
            disabled=(st.session_state.form_mode == "edit"),
            on_change=handle_id_change,
            help="La comprobación se hace automáticamente al salir del campo.",
        )
    with rec2:
        st.button("🔍 Cargar", use_container_width=True, on_click=handle_load_patient, disabled=(st.session_state.form_mode == "edit" or not normalize_patient_id(st.session_state.get("id_pac", ""))))
    with rec3:
        st.button("＋ Nuevo", use_container_width=True, on_click=handle_new_patient)

    current_id = normalize_patient_id(st.session_state.get("id_pac", ""))
    if st.session_state.form_mode == "edit":
        st.success(f"🟢 **Editando paciente {st.session_state.loaded_patient_id}**")
    elif current_id and st.session_state.id_checked == current_id and st.session_state.id_exists:
        st.warning(f"⚠️ **El paciente {current_id} ya existe.** Cárgalo para editarlo.")
        st.button(f"🔍 Cargar {current_id}", type="primary", use_container_width=True, on_click=handle_load_patient)
    elif current_id and st.session_state.id_checked == current_id:
        st.success(f"✅ ID **{current_id}** disponible para un paciente nuevo.")
    elif current_id:
        st.info("Pulsa Tab o haz clic fuera del campo ID para comprobar si ya existe.")

    if st.session_state.ui_message:
        msg = st.session_state.ui_message; kind = st.session_state.ui_message_type
        st.session_state.ui_message = ""; st.session_state.ui_message_type = "info"
        {"success": st.success, "warning": st.warning, "error": st.error, "info": st.info}.get(kind, st.info)(msg)

    form_locked = st.session_state.form_mode == "new" and (not current_id or st.session_state.id_exists or st.session_state.id_checked != current_id)
    done, total, pct = overall_completion()

    st.progress(pct, text=f"Completitud de la ficha actual: {done}/{total} · {pct:.0%}")

    # IA compacta
    with st.expander("🧠 **Texto clínico e IA**", expanded=False):
        left, right = st.columns([1.15, 1])
        with left:
            st.text_area("Pega aquí la evolución / historia clínica", height=220, key="clinical_text", disabled=form_locked, placeholder="Pega el texto clínico anonimizado…")
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
                            model_note = f" · {st.session_state.ai_model_used}" if st.session_state.ai_model_used else ""
                            st.success(f"{len(changes)} campos propuestos por IA{model_note}. Revísalos antes de guardar.")
                        else:
                            st.info("La IA no encontró nuevos datos explícitos que añadir.")
                        st.rerun()
        with right:
            if st.session_state.ai_changed_keys:
                st.info("🤖 **Revisar campos propuestos:** " + ", ".join(FIELD_LABELS.get(k, k) for k in st.session_state.ai_changed_keys))
            else:
                st.caption("Los campos propuestos por IA se marcan con 🤖 para su revisión.")
            if st.session_state.ai_model_used:
                st.caption(f"Última extracción: {st.session_state.ai_model_used}")
            if st.session_state.last_ai_data:
                with st.expander("Ver respuesta estructurada", expanded=False):
                    st.json(st.session_state.last_ai_data)

    section_options = list(SECTION_FIELDS.keys())
    section_labels = {}
    for title, fields in SECTION_FIELDS.items():
        c, t = section_completion(fields)
        section_labels[title] = f"{title} · {c}/{t}"
    if "clinical_section" not in st.session_state or st.session_state.clinical_section not in section_options:
        st.session_state.clinical_section = section_options[0]
    clinical_section = st.radio(
        "Bloque clínico",
        section_options,
        key="clinical_section",
        horizontal=True,
        label_visibility="collapsed",
        format_func=lambda x: section_labels[x],
    )

    if clinical_section == section_options[0]:
        st.subheader("Datos basales")
        a,b,c = st.columns(3)
        date_input_field("Fecha inclusión", "fecha_inc", form_locked, a)
        b.selectbox(widget_label("Sexo", "sexo"), ["", "H", "M"], key="sexo", disabled=form_locked, format_func=lambda x: "— No recogido —" if x == "" else x)
        c.text_input(widget_label("Edad", "edad"), key="edad", disabled=form_locked)
        a,b,c,d = st.columns(4)
        a.text_input(widget_label("Peso (kg)", "peso"), key="peso", disabled=form_locked)
        b.text_input(widget_label("Talla (cm)", "talla"), key="talla", disabled=form_locked)
        c.text_input("IMC · calculado", value=calculate_imc(), disabled=True)
        d.selectbox(widget_label("Etiología ERC", "eti_erc"), ["", "DM2", "HTA", "Glomerulonefritis", "Poliquistosis", "Otras"], key="eti_erc", disabled=form_locked, format_func=lambda x: "— No recogida —" if x == "" else x)
        a,b = st.columns(2)
        a.selectbox(widget_label("Etiología IC", "eti_ic"), ["", "Isquémica", "Hipertensiva", "MCD", "HFpEF", "Valvular", "Otras"], key="eti_ic", disabled=form_locked, format_func=lambda x: "— No recogida —" if x == "" else x)
        b.caption("Las etiologías no se infieren a partir de otros datos.")
        st.markdown("**Comorbilidades**")
        a,b,c,d = st.columns(4)
        tri_state_select("DM2", "dm2", form_locked, a); tri_state_select("HTA", "hta", form_locked, b); tri_state_select("FA", "fa", form_locked, c); tri_state_select("EPOC", "epoc", form_locked, d)
        a,b,c,d = st.columns(4)
        tri_state_select("Sd. Metabólico", "sd_metab", form_locked, a); tri_state_select("Tabaquismo", "tabaco", form_locked, b); tri_state_select("Enolismo", "enolismo", form_locked, c); tri_state_select("Hepatopatía", "hepato", form_locked, d)

    if clinical_section == section_options[1]:
        st.subheader("Analítica y biomarcadores")
        a,b,c,d = st.columns(4)
        a.text_input(widget_label("Hemoglobina", "hb"), key="hb", disabled=form_locked); b.text_input(widget_label("Creatinina", "creat"), key="creat", disabled=form_locked); c.text_input(widget_label("Cistatina C", "cist_c"), key="cist_c", disabled=form_locked); d.text_input(widget_label("FGe CKD-EPI", "fge"), key="fge", disabled=form_locked)
        a,b,c,d = st.columns(4)
        a.text_input(widget_label("Urea", "urea"), key="urea", disabled=form_locked); b.text_input(widget_label("Ácido úrico", "ac_urico"), key="ac_urico", disabled=form_locked); c.text_input(widget_label("Prot/Creat", "prot_creat"), key="prot_creat", disabled=form_locked); d.text_input(widget_label("Plaquetas", "plaq"), key="plaq", disabled=form_locked)
        a,b,c,d = st.columns(4)
        a.text_input(widget_label("AST", "ast"), key="ast", disabled=form_locked); b.text_input(widget_label("ALT", "alt"), key="alt", disabled=form_locked); c.text_input("FIB-4 · calculado", value=calculate_fib4(), disabled=True); d.text_input(widget_label("Bilirrubina total", "bili_t"), key="bili_t", disabled=form_locked)
        a,b,c,d = st.columns(4)
        a.text_input(widget_label("Bilirrubina directa", "bili_d"), key="bili_d", disabled=form_locked); b.text_input(widget_label("Albúmina", "albumina"), key="albumina", disabled=form_locked); c.text_input(widget_label("HbA1c", "hba1c"), key="hba1c", disabled=form_locked); d.text_input(widget_label("Colesterol total", "colest"), key="colest", disabled=form_locked)
        st.markdown("**Biomarcadores**")
        a,b,c,d = st.columns(4)
        a.text_input(widget_label("NT-proBNP (pg/mL)", "nt_probnp"), key="nt_probnp", disabled=form_locked); b.text_input(widget_label("CA125 (U/mL)", "ca125"), key="ca125", disabled=form_locked); c.text_input(widget_label("Galectina-3 (ng/mL)", "gal3"), key="gal3", disabled=form_locked); d.text_input(widget_label("sST2 (ng/mL)", "sst2"), key="sst2", disabled=form_locked)
        a,b = st.columns(2)
        a.text_input(widget_label("GDF-15 (pg/mL)", "gdf15"), key="gdf15", disabled=form_locked); tri_state_select("Muestra biobanco", "biobanco", form_locked, b)

    if clinical_section == section_options[2]:
        st.subheader("Ecocardiografía / elastografía")
        a,b,c,d = st.columns(4)
        a.text_input(widget_label("FEVI (%)", "fevi"), key="fevi", disabled=form_locked); b.text_input(widget_label("GLS", "gls"), key="gls", disabled=form_locked); c.text_input(widget_label("Masa VI", "masa_vi"), key="masa_vi", disabled=form_locked); d.text_input(widget_label("TAPSE", "tapse"), key="tapse", disabled=form_locked)
        a,b,c,d = st.columns(4)
        a.text_input(widget_label("VAI", "vai"), key="vai", disabled=form_locked); b.text_input(widget_label("VCI (cm)", "vci"), key="vci", disabled=form_locked); c.text_input(widget_label("LSM (kPa)", "lsm"), key="lsm", disabled=form_locked); d.text_input(widget_label("CAP (dB/m)", "cap"), key="cap", disabled=form_locked)
        a,b,c,d = st.columns(4)
        a.text_input(widget_label("Mediciones válidas", "med_val"), key="med_val", disabled=form_locked); b.text_input(widget_label("IQR/Mediana ≤0.30", "iqr_med"), key="iqr_med", disabled=form_locked); c.text_input(widget_label("Días desde descompensación", "dias_desc"), key="dias_desc", disabled=form_locked); d.text_input(widget_label("NT-proBNP día prueba", "nt_prueba"), key="nt_prueba", disabled=form_locked)
        a,b = st.columns(2)
        a.text_input(widget_label("Categoría fibrosis · solo si documentada", "cat_fibro"), key="cat_fibro", disabled=form_locked); tri_state_select("Edemas día prueba", "edemas_prueba", form_locked, b)
        st.info("LSM y categoría de fibrosis se mantienen separadas. La IA no infiere fibrosis a partir de LSM.")

    if clinical_section == section_options[3]:
        st.subheader("Tratamiento")
        a,b,c,d = st.columns(4)
        tri_state_select("IECA", "ieca", form_locked, a); tri_state_select("ARA2", "ara2", form_locked, b); tri_state_select("Betabloqueante", "bb", form_locked, c); tri_state_select("AMR", "amr", form_locked, d)
        a,b,c,d = st.columns(4)
        tri_state_select("SAC/VAL", "sac_val", form_locked, a); tri_state_select("SGLT2i", "sglt2i", form_locked, b); tri_state_select("Diurético de asa", "diur_asa", form_locked, c); tri_state_select("HCTZ", "hctz", form_locked, d)
        a,b,c = st.columns(3)
        tri_state_select("Acetazolamida", "acetazolamida", form_locked, a); tri_state_select("Estatinas", "estatinas", form_locked, b); tri_state_select("Eritropoyetina", "epo", form_locked, c)

    if clinical_section == section_options[4]:
        st.subheader("Seguimiento 24 meses")
        st.text_input(widget_label("Meses de seguimiento", "meses_seg"), key="meses_seg", disabled=form_locked)
        st.markdown("**Eventos cardiovasculares**")
        a,b = st.columns(2)
        tri_state_select("Muerte cardiovascular", "m_cv", form_locked, a)
        if st.session_state.m_cv == "Sí": date_input_field("Fecha muerte CV", "f_m_cv", form_locked, a)
        tri_state_select("Hospitalización por IC descompensada", "hosp_ic", form_locked, b)
        if st.session_state.hosp_ic == "Sí": date_input_field("Fecha hospitalización IC", "f_hosp_ic", form_locked, b)
        a,b = st.columns(2)
        tri_state_select("IAM no fatal", "iam", form_locked, a)
        if st.session_state.iam == "Sí": date_input_field("Fecha IAM", "f_iam", form_locked, a)
        tri_state_select("ACV no fatal", "acv", form_locked, b)
        if st.session_state.acv == "Sí": date_input_field("Fecha ACV", "f_acv", form_locked, b)
        st.metric("MACE+ · calculado", calculate_mace_plus() or "—")
        st.markdown("**Otros desenlaces**")
        a,b = st.columns(2)
        tri_state_select("Muerte total", "m_tot", form_locked, a)
        if st.session_state.m_tot == "Sí": date_input_field("Fecha muerte total", "f_m_tot", form_locked, a)
        tri_state_select("Inicio TRS (diálisis/trasplante)", "trs", form_locked, b)
        if st.session_state.trs == "Sí": date_input_field("Fecha inicio TRS", "f_trs", form_locked, b)
        a,b = st.columns(2)
        tri_state_select("Caída FGe ≥25%", "caida_fge", form_locked, a)
        if st.session_state.caida_fge == "Sí": date_input_field("Fecha caída FGe", "f_caida_fge", form_locked, a)
        tri_state_select("Síndrome cardiorrenal agudo", "sd_cr", form_locked, b)
        if st.session_state.sd_cr == "Sí": date_input_field("Fecha Sd. CR agudo", "f_sd_cr", form_locked, b)

    recalculate_derived_fields()
    st.markdown("### Revisión antes de guardar")
    missing = []
    for title, fields in SECTION_FIELDS.items():
        c,t = section_completion(fields)
        if c < t: missing.append((title, t-c))
    r1,r2,r3 = st.columns(3)
    r1.metric("Variables recogidas", f"{done}/{total}"); r2.metric("Completitud", f"{pct:.0%}"); r3.metric("ID", current_id or "—")
    if st.session_state.form_mode == "new" and current_id and st.session_state.id_exists:
        st.warning("Este paciente ya existe. Cárgalo antes de editarlo.")
    elif missing:
        st.caption("Pendientes: " + " · ".join(f"{x} ({n})" for x,n in missing))
    else:
        st.success("Ficha completa según las variables monitorizadas.")
    save_disabled = form_locked or not current_id or (st.session_state.form_mode == "new" and st.session_state.id_exists)
    s1,s2 = st.columns([2,1])
    with s1: st.button("💾 GUARDAR PACIENTE", type="primary", use_container_width=True, on_click=handle_save_patient, disabled=save_disabled)
    with s2:
        try:
            excel_bytes, n = export_all_excel()
            st.download_button("⬇️ Exportar Excel", data=excel_bytes, file_name=f"CRD_Tesis_Cardiorrenal_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        except Exception as e: st.error(f"No se pudo preparar el Excel: {e}")

    if current_id:
        with st.expander("🕒 Historial de cambios del paciente", expanded=False):
            try:
                history = db_audit(current_id)
                st.dataframe(dataframe_from_records(history), use_container_width=True, hide_index=True)
            except Exception as e: st.caption(f"No se pudo cargar el historial: {e}")

# ============================================================
# TAB 2 — DASHBOARD
# ============================================================
if main_nav == "dashboard":
    st.subheader("Dashboard de la cohorte")
    st.caption("Panel descriptivo de la base actual. Los indicadores no sustituyen el análisis estadístico de la tesis.")
    refresh_dash = st.button("↻ Actualizar dashboard", key="refresh_dashboard")
    try:
        rows = patient_rows_flat()
    except Exception as e:
        st.error(f"No se pudieron cargar los datos de la cohorte: {e}")
        rows = []
    n = len(rows)
    ages = numeric_values(rows, "edad")
    fges = numeric_values(rows, "fge")
    creats = numeric_values(rows, "creat")
    ntprobnp = numeric_values(rows, "nt_probnp")
    imcs = numeric_values(rows, "imc")
    completeness = cohort_completeness(rows)
    avg_comp = (sum(x[3] for x in completeness) / len(completeness)) if completeness else 0

    a,b,c,d,e = st.columns(5)
    a.metric("N pacientes", n)
    b.metric("Edad mediana", f"{median(ages):.0f}" if median(ages) is not None else "—")
    c.metric("IMC mediano", f"{median(imcs):.1f}" if median(imcs) is not None else "—")
    d.metric("FGe mediano", f"{median(fges):.1f}" if median(fges) is not None else "—")
    e.metric("Completitud media", f"{avg_comp:.0%}")

    st.markdown("### Perfil de la cohorte")
    p1,p2,p3 = st.columns(3)
    with p1:
        sex = count_value(rows, "sexo")
        st.markdown("**Sexo**")
        st.bar_chart(dataframe_from_records([{"Categoría": k, "N": v} for k,v in sex.items()]).set_index("Categoría") if sex else {"N": []})
    with p2:
        erc = count_value(rows, "eti_erc")
        erc = {k:v for k,v in erc.items() if k}
        st.markdown("**Etiología ERC**")
        st.bar_chart(dataframe_from_records([{"Categoría": k, "N": v} for k,v in erc.items()]).set_index("Categoría") if erc else {"N": []})
    with p3:
        ic = count_value(rows, "eti_ic")
        ic = {k:v for k,v in ic.items() if k}
        st.markdown("**Etiología IC**")
        st.bar_chart(dataframe_from_records([{"Categoría": k, "N": v} for k,v in ic.items()]).set_index("Categoría") if ic else {"N": []})

    st.markdown("### Comorbilidades y tratamientos")
    q1,q2 = st.columns(2)
    comorb_keys = ["dm2","hta","fa","epoc","sd_metab","tabaco","enolismo","hepato"]
    treatment_keys = ["ieca","ara2","bb","amr","sac_val","sglt2i","diur_asa","estatinas"]
    with q1:
        data = [{"Variable": FIELD_LABELS[k], "N": count_yes(rows,k)} for k in comorb_keys]
        st.bar_chart(dataframe_from_records(data).set_index("Variable") if data else [])
    with q2:
        data = [{"Tratamiento": FIELD_LABELS[k], "N": count_yes(rows,k)} for k in treatment_keys]
        st.bar_chart(dataframe_from_records(data).set_index("Tratamiento") if data else [])

    st.markdown("### Eventos de seguimiento")
    event_keys = ["m_cv","hosp_ic","iam","acv","mace_plus","m_tot","trs","caida_fge","sd_cr"]
    event_table = [{"Evento": FIELD_LABELS[k], "N": count_yes(rows,k)} for k in event_keys]
    st.dataframe(dataframe_from_records(event_table), use_container_width=True, hide_index=True)

    st.markdown("### Calidad de datos")
    c1,c2 = st.columns([1,2])
    with c1:
        bins = {"0–49%":0,"50–74%":0,"75–89%":0,"90–99%":0,"100%":0}
        for _,_,_,p in completeness:
            if p < .5: bins["0–49%"] += 1
            elif p < .75: bins["50–74%"] += 1
            elif p < .9: bins["75–89%"] += 1
            elif p < 1: bins["90–99%"] += 1
            else: bins["100%"] += 1
        st.bar_chart(dataframe_from_records([{"Completitud":k,"N":v} for k,v in bins.items()]).set_index("Completitud"))
    with c2:
        missingness = []
        for key in [k for fields in SECTION_FIELDS.values() for k in fields if k not in DERIVED_FIELDS and not k.startswith("f_")]:
            present = sum(1 for r in rows if is_filled(r.get(key,"")))
            missingness.append({"Variable": FIELD_LABELS.get(key,key), "Sin dato": n-present, "% sin dato": round((n-present)/n*100,1) if n else 0})
        missingness.sort(key=lambda x: x["% sin dato"], reverse=True)
        st.dataframe(dataframe_from_records(missingness[:12]), use_container_width=True, hide_index=True)

    with st.expander("📋 Resumen analítico rápido", expanded=False):
        summary = {
            "Pacientes": n,
            "Edad mediana": median(ages),
            "Creatinina mediana": median(creats),
            "FGe mediano": median(fges),
            "NT-proBNP mediano": median(ntprobnp),
            "MACE+": count_yes(rows, "mace_plus"),
            "Muerte CV": count_yes(rows, "m_cv"),
            "Hospitalización IC": count_yes(rows, "hosp_ic"),
        }
        st.dataframe(dataframe_from_records([{"Indicador":k,"Valor":v} for k,v in summary.items()]), use_container_width=True, hide_index=True)

# ============================================================
# TAB 3 — PACIENTES
# ============================================================
if main_nav == "pacientes":
    st.subheader("Pacientes")
    st.caption("Búsqueda rápida y control de pacientes de la base central.")
    try:
        all_rows = patient_rows_flat()
    except Exception as e:
        all_rows = []
        st.error(f"No se pudieron cargar los pacientes: {e}")
    search = st.text_input("Buscar por ID", key="patient_search", placeholder="Ej.: 3444")
    filtered = [r for r in all_rows if not search.strip() or search.strip().upper() in str(r.get("id_pac", "")).upper()]
    table = [{"ID": r.get("id_pac",""), "Última actualización": r.get("updated_at",""), "Usuario": r.get("updated_by","")} for r in filtered]
    st.dataframe(dataframe_from_records(table), use_container_width=True, hide_index=True)
    ids = [r.get("id_pac","") for r in filtered if r.get("id_pac")]
    if ids:
        pick_col, button_col = st.columns([2,1])
        with pick_col:
            selected = st.selectbox("Paciente", ids, key="selected_patient_row")
        with button_col:
            st.write("")
            st.write("")
            if st.button("🔍 Cargar para editar", type="primary", use_container_width=True):
                st.session_state.id_pac = selected
                if db_load_patient(selected):
                    set_ui_message(f"Paciente {selected} cargado. Ve a 'Recogida clínica' para editarlo.", "success")
                st.rerun()
    else:
        st.info("No hay pacientes que coincidan con la búsqueda.")

# ============================================================
# TAB 4 — ADMINISTRACIÓN
# ============================================================
if main_nav == "admin":
    st.subheader("Administración y exportación")
    st.warning("Supabase es la fuente maestra. El Excel es una exportación para análisis, copia y trabajo estadístico.")
    try:
        excel_bytes, n_export = export_all_excel()
        st.download_button("⬇️ Descargar base completa en Excel", data=excel_bytes, file_name=f"CRD_Tesis_Cardiorrenal_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.caption(f"Exportación preparada: {n_export} pacientes.")
    except Exception as e:
        st.error(f"No se pudo preparar la exportación: {e}")

    current_id = normalize_patient_id(st.session_state.get("id_pac", ""))
    if current_id:
        st.markdown("### Historial del paciente activo")
        try:
            history = db_audit(current_id)
            st.dataframe(dataframe_from_records(history), use_container_width=True, hide_index=True)
        except Exception as e:
            st.caption(f"No se pudo cargar el historial: {e}")

    st.markdown("### Estado de la aplicación")
    a,b,c = st.columns(3)
    a.metric("Supabase", "🟢 Conectado")
    b.metric("Gemini", "🟢 Configurado" if GEMINI_API_KEY else "🔴 Falta clave")
    c.metric("Pacientes", db_count())
    st.caption("La exportación no modifica la base central.")
