import streamlit as st
import json
import os
import re
from datetime import datetime
from io import BytesIO
import openpyxl

# ============================================================
# CRD TESIS CARDIORRENAL V6
# Base central Supabase + exportación Excel + control robusto de estado
# ============================================================

st.set_page_config(page_title="CRD Tesis Cardiorrenal", layout="wide")

# ---------- Secrets ----------
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "")

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
        if st.button("Entrar", type="primary"):
            if password == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        st.stop()

try:
    from supabase import create_client
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
except Exception as e:
    st.error(f"No se pudo conectar con Supabase: {e}")
    st.stop()

# ---------- Variables ----------
DEFAULTS = {
    "id_pac": "", "fecha_inc": datetime.today().strftime("%d/%m/%Y"),
    "edad": "", "sexo": "", "peso": "", "talla": "", "imc": "",
    "eti_erc": "", "eti_ic": "",
    "dm2": "No", "hta": "No", "fa": "No", "epoc": "No", "sd_metab": "No",
    "tabaco": "No", "enolismo": "No", "hepato": "No",
    "hb": "", "creat": "", "cist_c": "", "fge": "", "urea": "", "ac_urico": "",
    "prot_creat": "", "ast": "", "alt": "", "plaq": "", "bili_t": "", "bili_d": "",
    "albumina": "", "hba1c": "", "colest": "", "fib4": "", "nt_probnp": "", "ca125": "",
    "gal3": "", "sst2": "", "gdf15": "", "biobanco": "No",
    "fevi": "", "gls": "", "masa_vi": "", "tapse": "", "vai": "", "vci": "",
    "lsm": "", "cat_fibro": "", "cap": "", "med_val": "", "iqr_med": "",
    "dias_desc": "", "nt_prueba": "", "edemas_prueba": "No",
    "ieca": "No", "ara2": "No", "bb": "No", "amr": "No", "sac_val": "No",
    "sglt2i": "No", "diur_asa": "No", "hctz": "No", "acetazolamida": "No",
    "estatinas": "No", "epo": "No",
    "meses_seg": "", "m_cv": "No", "f_m_cv": "", "hosp_ic": "No",
    "f_hosp_ic": "", "iam": "No", "f_iam": "", "acv": "No", "f_acv": "",
    "m_tot": "No", "f_m_tot": "", "trs": "No", "f_trs": "",
    "caida_fge": "No", "f_caida_fge": "", "sd_cr": "No", "f_sd_cr": ""
}

BOOLS = ["Sí", "No"]

for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

if "last_ai_data" not in st.session_state:
    st.session_state.last_ai_data = {}
if "audit_log" not in st.session_state:
    st.session_state.audit_log = []
if "form_mode" not in st.session_state:
    st.session_state.form_mode = "new"  # new | edit
if "loaded_patient_id" not in st.session_state:
    st.session_state.loaded_patient_id = ""
if "loaded_updated_at" not in st.session_state:
    st.session_state.loaded_updated_at = ""
if "ui_message" not in st.session_state:
    st.session_state.ui_message = ""
if "ui_message_type" not in st.session_state:
    st.session_state.ui_message_type = ""
if "id_exists" not in st.session_state:
    st.session_state.id_exists = False
if "id_checked" not in st.session_state:
    st.session_state.id_checked = ""


# ============================================================
# DATABASE
# ============================================================

def db_get_patient(patient_id):
    res = supabase.table("patients").select("*").eq("id_pac", patient_id).limit(1).execute()
    if not res.data:
        return None
    return res.data[0]

def db_count():
    res = supabase.table("patients").select("id_pac", count="exact").execute()
    return res.count if res.count is not None else len(res.data or [])

def normalize_patient_id(value):
    return str(value or "").strip().upper()

def check_id_exists():
    """Comprueba la ID en cuanto el usuario sale del campo de ID."""
    patient_id = normalize_patient_id(st.session_state.get("id_pac", ""))
    st.session_state.id_pac = patient_id
    st.session_state.id_checked = patient_id
    if not patient_id or st.session_state.get("form_mode", "new") != "new":
        st.session_state.id_exists = False
        return
    try:
        st.session_state.id_exists = db_get_patient(patient_id) is not None
    except Exception:
        # No bloqueamos la interfaz por un error transitorio de red.
        st.session_state.id_exists = False

def db_save_patient(data):
    patient_id = normalize_patient_id(data.get("id_pac", ""))
    if not patient_id:
        raise ValueError("El ID de paciente es obligatorio.")

    data = dict(data)
    data["id_pac"] = patient_id
    now = datetime.now().isoformat()
    old = db_get_patient(patient_id)
    mode = st.session_state.get("form_mode", "new")
    loaded_id = normalize_patient_id(st.session_state.get("loaded_patient_id", ""))

    payload = {
        "id_pac": patient_id,
        "data": data,
        "updated_at": now,
        "updated_by": st.session_state.get("user_label", "usuario")
    }

    if mode == "new":
        if old:
            raise ValueError(
                f"DUPLICADO: la ID {patient_id} ya existe en la base. "
                "Pulsa 'Cargar paciente' para editarlo; no se creará una segunda ficha."
            )
        res = supabase.table("patients").insert(payload).select("*").single().execute()
        action = "CREATE"
        result_label = "creado"
    else:
        if not old:
            raise ValueError(
                f"La ID {patient_id} ya no existe en la base. Vuelve a 'Nuevo paciente' para crearla."
            )
        if loaded_id != patient_id:
            raise ValueError(
                "Has cambiado la ID después de cargar un paciente. "
                "Para crear otra ficha pulsa 'Nuevo paciente'."
            )
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
                "updated_by": st.session_state.get("user_label", "usuario")
            })
            .eq("id_pac", patient_id)
            .select("*")
            .single()
            .execute()
        )
        action = "UPDATE"
        result_label = "actualizado"

    # Audit record: full snapshot + timestamp.
    supabase.table("audit_log").insert({
        "id_pac": patient_id,
        "action": action,
        "snapshot": data,
        "changed_at": now,
        "changed_by": st.session_state.get("user_label", "usuario")
    }).execute()

    st.session_state.form_mode = "edit"
    st.session_state.loaded_patient_id = patient_id
    st.session_state.loaded_updated_at = now
    return bool(res.data), result_label

def db_load_patient(patient_id):
    row = db_get_patient(patient_id)
    if not row:
        return False

    data = row.get("data") or {}
    for k in DEFAULTS:
        st.session_state[k] = data.get(k, DEFAULTS[k])

    st.session_state.id_pac = patient_id
    st.session_state.form_mode = "edit"
    st.session_state.loaded_patient_id = patient_id
    st.session_state.loaded_updated_at = row.get("updated_at", "")
    st.session_state.id_exists = False
    st.session_state.id_checked = patient_id
    recalculate_derived_fields()
    return True

def set_ui_message(message, kind="info"):
    st.session_state.ui_message = message
    st.session_state.ui_message_type = kind

def handle_load_patient():
    normalized_search = normalize_patient_id(st.session_state.get("id_pac", ""))
    if not normalized_search:
        set_ui_message("Introduce primero una ID de paciente.", "warning")
        return
    if db_load_patient(normalized_search):
        set_ui_message(f"Paciente {normalized_search} cargado desde la base central.", "success")
    else:
        set_ui_message(
            f"El paciente {normalized_search} no existe en la base de datos. "
            "Pulsa 'Nuevo paciente' para crear una ficha nueva.",
            "warning"
        )

def handle_new_patient():
    for k, v in DEFAULTS.items():
        st.session_state[k] = v
    st.session_state.last_ai_data = {}
    st.session_state.form_mode = "new"
    st.session_state.loaded_patient_id = ""
    st.session_state.loaded_updated_at = ""
    st.session_state.id_exists = False
    st.session_state.id_checked = ""
    set_ui_message("Formulario preparado para un paciente nuevo.", "info")

def handle_save_patient():
    normalized_id = normalize_patient_id(st.session_state.get("id_pac", ""))
    if not normalized_id:
        set_ui_message("El ID de paciente es obligatorio.", "error")
        return

    warnings_now = validation_warnings()
    if warnings_now:
        set_ui_message("Corrige primero los valores fuera de rango.", "error")
        return

    recalculate_derived_fields()
    data = {k: st.session_state.get(k, DEFAULTS[k]) for k in DEFAULTS}
    data["id_pac"] = normalized_id

    try:
        ok, mode = db_save_patient(data)
        if ok:
            set_ui_message(
                f"Paciente {data['id_pac']} {mode} correctamente en la BASE CENTRAL.",
                "success"
            )
    except Exception as e:
        msg = str(e)
        if "DUPLICADO" in msg:
            st.session_state.id_exists = True
            st.session_state.id_checked = normalized_id
            set_ui_message(
                f"⚠️ El paciente {normalized_id} ya existe en la base de datos. "
                "No se ha modificado nada. Cárgalo para poder editarlo.",
                "warning"
            )
        else:
            set_ui_message(f"No se pudo guardar en la base central: {e}", "error")

def db_recent(limit=20):
    return supabase.table("patients").select(
        "id_pac,updated_at,updated_by"
    ).order("updated_at", desc=True).limit(limit).execute().data or []

def db_audit(patient_id, limit=20):
    return supabase.table("audit_log").select(
        "id_pac,action,changed_at,changed_by"
    ).eq("id_pac", patient_id).order(
        "changed_at", desc=True
    ).limit(limit).execute().data or []


# ============================================================
# VARIABLES DERIVADAS
# ============================================================

DERIVED_FIELDS = {"imc", "fib4"}


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
    talla_m = talla_cm / 100.0
    return f"{peso / (talla_m ** 2):.1f}"


def calculate_fib4():
    edad = _to_float(st.session_state.get("edad", ""))
    ast = _to_float(st.session_state.get("ast", ""))
    alt = _to_float(st.session_state.get("alt", ""))
    plaq = _to_float(st.session_state.get("plaq", ""))
    if any(v is None for v in (edad, ast, alt, plaq)):
        return ""
    if edad < 0 or ast <= 0 or alt <= 0 or plaq <= 0:
        return ""
    # FIB-4 = (edad x AST) / (plaquetas x sqrt(ALT))
    return f"{(edad * ast) / (plaq * (alt ** 0.5)):.2f}"


def recalculate_derived_fields():
    st.session_state["imc"] = calculate_imc()
    st.session_state["fib4"] = calculate_fib4()


# ============================================================
# VALIDACIÓN
# ============================================================

def is_number(v):
    try:
        float(str(v).replace(",", "."))
        return True
    except Exception:
        return False

def validation_warnings():
    warnings = []
    ranges = {
        "edad": (0, 120), "peso": (1, 500), "talla": (30, 250),
        "hb": (1, 30), "creat": (0.1, 30), "fge": (0, 200),
        "fevi": (0, 100), "tapse": (0, 50), "lsm": (0, 100),
        "cap": (0, 1000), "plaq": (1, 2000), "meses_seg": (0, 120)
    }
    for key, (lo, hi) in ranges.items():
        v = st.session_state.get(key, "")
        if v not in ("", None) and is_number(v):
            x = float(str(v).replace(",", "."))
            if x < lo or x > hi:
                warnings.append(f"{key}: {v} fuera del rango de comprobación ({lo}–{hi}).")
    return warnings


# ============================================================
# IA
# ============================================================

AI_KEYS = [k for k in DEFAULTS.keys() if k not in DERIVED_FIELDS]

AI_SCHEMA = {k: "valor explícito o null" for k in AI_KEYS}
for k in ["dm2","hta","fa","epoc","sd_metab","tabaco","enolismo","hepato",
          "biobanco","edemas_prueba","ieca","ara2","bb","amr","sac_val",
          "sglt2i","diur_asa","hctz","acetazolamida","estatinas","epo",
          "m_cv","hosp_ic","iam","acv","m_tot","trs","caida_fge","sd_cr"]:
    AI_SCHEMA[k] = "Sí|No|null"

def normalize_yes_no(v):
    if v is None:
        return None
    s = str(v).strip().lower()
    if s in ("si", "sí", "yes", "true"):
        return "Sí"
    if s in ("no", "false"):
        return "No"
    return str(v).strip()

def extract_ai(texto):
    try:
        import google.generativeai as genai
        api_key = st.secrets.get("GEMINI_API_KEY", "")
        if not api_key:
            st.error("Falta GEMINI_API_KEY en los Secrets.")
            return None
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-1.5-flash")
    except Exception as e:
        st.error(f"No se pudo iniciar Gemini: {e}")
        return None

    prompt = f"""
Eres un extractor de datos clínicos para una base de investigación cardiorrenal.

Extrae EXCLUSIVAMENTE información explícitamente presente en el texto.
NO inventes datos.
NO infieras datos.
NO completes campos por conocimiento médico.
Si no aparece un dato, devuelve null.
Si hay duda, devuelve null.

Reglas:
- Mantén los valores numéricos tal como aparecen.
- No hagas conversiones de unidades.
- Para Sí/No exige evidencia textual.
- No confundas antecedentes con eventos de seguimiento.
- NO conviertas LSM en fibrosis.
- NO asignes cat_fibro a partir de LSM. Sólo extrae una categoría si aparece explícitamente.
- Devuelve únicamente JSON válido, sin markdown.

Claves permitidas:
{json.dumps(AI_SCHEMA, ensure_ascii=False)}

Texto clínico:
{texto}
"""
    try:
        response = model.generate_content(prompt)
        raw = response.text.strip()
        raw = re.sub(r"^```json\s*", "", raw, flags=re.I)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)
        return {k: v for k, v in data.items() if k in DEFAULTS}
    except Exception as e:
        st.error(f"Error interpretando la respuesta de Gemini: {e}")
        return None

def apply_ai(data):
    changes = []
    for k, v in data.items():
        if k not in DEFAULTS or k in DERIVED_FIELDS or k == "id_pac" or v is None or str(v).strip() == "":
            continue
        if k in {
            "dm2","hta","fa","epoc","sd_metab","tabaco","enolismo","hepato",
            "biobanco","edemas_prueba","ieca","ara2","bb","amr","sac_val",
            "sglt2i","diur_asa","hctz","acetazolamida","estatinas","epo",
            "m_cv","hosp_ic","iam","acv","m_tot","trs","caida_fge","sd_cr"
        }:
            v = normalize_yes_no(v)
        elif isinstance(v, (int, float)):
            v = str(v)
        else:
            v = str(v).strip()

        old = st.session_state.get(k, "")
        if str(old) != str(v):
            changes.append((k, old, v))
            st.session_state[k] = v

    st.session_state.last_ai_data = data
    return changes


# ============================================================
# EXPORTACIÓN EXCEL DESDE LA BASE CENTRAL
# ============================================================

SHEETS = {
    "01_Datos_Clinicos": [
        (1,"id_pac"),(2,"fecha_inc"),(3,"edad"),(4,"sexo"),(5,"peso"),(6,"talla"),(7,"imc"),
        (8,"eti_erc"),(9,"eti_ic"),(10,"dm2"),(11,"hta"),(12,"fa"),(13,"epoc"),
        (14,"sd_metab"),(15,"tabaco"),(16,"enolismo"),(17,"hepato")
    ],
    "02_Analitica_Biomarcadores": [
        (1,"id_pac"),(2,"hb"),(3,"creat"),(4,"cist_c"),(5,"fge"),(6,"urea"),
        (7,"ac_urico"),(8,"prot_creat"),(9,"ast"),(10,"alt"),(11,"plaq"),(12,"fib4"),
        (13,"bili_t"),(14,"bili_d"),(15,"albumina"),(16,"hba1c"),(17,"colest"),
        (18,"nt_probnp"),(19,"ca125"),(20,"gal3"),(21,"sst2"),(22,"gdf15"),
        (23,"biobanco")
    ],
    "03_Eco_Elastografia": [
        (1,"id_pac"),(2,"fevi"),(3,"gls"),(4,"masa_vi"),(5,"tapse"),(6,"vai"),
        (7,"vci"),(8,"lsm"),(9,"cat_fibro"),(10,"cap"),(11,"med_val"),
        (12,"iqr_med"),(13,"dias_desc"),(14,"nt_prueba"),(15,"edemas_prueba")
    ],
    "04_Tratamiento": [
        (1,"id_pac"),(2,"ieca"),(3,"ara2"),(4,"bb"),(5,"amr"),(6,"sac_val"),
        (7,"sglt2i"),(8,"diur_asa"),(9,"hctz"),(10,"acetazolamida"),
        (11,"estatinas"),(12,"epo")
    ],
    "05_Seguimiento_24m": [
        (1,"id_pac"),(2,"meses_seg"),(3,"m_cv"),(4,"f_m_cv"),(5,"hosp_ic"),
        (6,"f_hosp_ic"),(7,"iam"),(8,"f_iam"),(9,"acv"),(10,"f_acv"),
        (12,"m_tot"),(13,"f_m_tot"),(14,"trs"),(15,"f_trs"),(16,"caida_fge"),
        (17,"f_caida_fge"),(18,"sd_cr"),(19,"f_sd_cr")
    ]
}

def export_all_excel():
    rows = supabase.table("patients").select("id_pac,data").order("id_pac").execute().data or []
    wb = openpyxl.Workbook()
    first = True

    for sheet_name, mapping in SHEETS.items():
        if first:
            ws = wb.active
            ws.title = sheet_name
            first = False
        else:
            ws = wb.create_sheet(sheet_name)

        max_col = max(c for c,_ in mapping)
        for col, key in mapping:
            ws.cell(row=1, column=col, value=key)

        for r_idx, row in enumerate(rows, start=2):
            data = row.get("data") or {}
            for col, key in mapping:
                ws.cell(row=r_idx, column=col, value=data.get(key, ""))

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), len(rows)


# ============================================================
# CABECERA
# ============================================================

st.title("Tesis Esther Tamarit · Base Cardiorrenal")

count = db_count()
recent = db_recent(1)
last_update = recent[0]["updated_at"] if recent else "—"

m1, m2, m3, m4 = st.columns(4)
m1.metric("Pacientes en la base", count)
m2.metric("Base central", "SUPABASE")
m3.metric("Última actualización", last_update[:19].replace("T", " ") if last_update != "—" else "—")
m4.metric("Estado", "🟢 ONLINE")


# ============================================================
# BUSCADOR
# ============================================================

top1, top2, top3 = st.columns([1.5, 1, 1])

with top1:
    st.text_input(
        "ID paciente",
        key="id_pac",
        on_change=check_id_exists,
        help="Escribe la ID. Al salir del campo se comprueba inmediatamente si ese paciente ya existe."
    )

with top2:
    st.button("🔍 Cargar paciente", use_container_width=True, on_click=handle_load_patient)

with top3:
    st.button("🔄 Nuevo paciente", use_container_width=True, on_click=handle_new_patient)

# Aviso inmediato: se muestra antes de que el usuario tenga que rellenar el formulario.
if (
    st.session_state.form_mode == "new"
    and st.session_state.id_exists
    and normalize_patient_id(st.session_state.get("id_pac", "")) == st.session_state.get("id_checked", "")
):
    existing_id = st.session_state.id_checked
    st.warning(
        f"⚠️ El paciente **{existing_id} ya existe** en la base de datos. "
        "Si quieres modificarlo, pulsa **🔍 Cargar paciente**. "
        "Si quieres crear uno nuevo, utiliza otra ID."
    )

if st.session_state.ui_message:
    kind = st.session_state.ui_message_type
    message = st.session_state.ui_message
    st.session_state.ui_message = ""
    st.session_state.ui_message_type = ""
    if kind == "success":
        st.success(message)
    elif kind == "warning":
        st.warning(message)
    elif kind == "error":
        st.error(message)
    else:
        st.info(message)


mode_label = "EDICIÓN: paciente cargado" if st.session_state.form_mode == "edit" else "NUEVO PACIENTE"
if st.session_state.form_mode == "edit":
    st.success(f"🟢 {mode_label} · ID {st.session_state.loaded_patient_id}")


# ============================================================
# HISTORIA + IA
# ============================================================

col_izq, col_der = st.columns([1, 2])

with col_izq:
    st.subheader("1. Evolución / Historia")
    texto = st.text_area("Pega aquí el texto clínico", height=350)

    if st.button("🧠 Auto-completar con IA", type="primary", use_container_width=True):
        if not texto.strip():
            st.warning("Pega primero el texto clínico.")
        else:
            with st.spinner("Extrayendo datos..."):
                result = extract_ai(texto)
            if result:
                changes = apply_ai(result)
                st.success(f"Extracción completada: {len(changes)} campos propuestos.")
                st.rerun()

    if st.session_state.last_ai_data:
        with st.expander("Ver extracción IA"):
            st.json(st.session_state.last_ai_data)


# ============================================================
# FORMULARIO
# ============================================================

with col_der:
    st.subheader("2. Formulario")
    st.caption("La ID se introduce únicamente en la parte superior. Aquí aparecen solo los datos clínicos.")

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "01. Clínicos","02. Analítica","03. Eco/Fibro","04. Tratamiento","05. Seguimiento"
    ])

    with tab1:
        a,b,c = st.columns(3)
        a.text_input("Fecha inclusión", key="fecha_inc")
        b.selectbox("Sexo", ["","H","M"], key="sexo")
        c.text_input("Edad", key="edad")

        a,b,c,d = st.columns(4)
        a.text_input("Peso (kg)", key="peso")
        b.text_input("Talla (cm)", key="talla")
        c.text_input("IMC (calculado)", value=calculate_imc(), disabled=True)
        d.selectbox("Etiología ERC", ["","DM2","HTA","Glomerulonefritis","Poliquistosis","Otras"], key="eti_erc")

        a,b = st.columns(2)
        a.selectbox("Etiología IC", ["","Isquémica","Hipertensiva","MCD","HFpEF","Valvular"], key="eti_ic")

        st.markdown("**Comorbilidades**")
        a,b,c,d = st.columns(4)
        a.selectbox("DM2", BOOLS, key="dm2")
        b.selectbox("HTA", BOOLS, key="hta")
        c.selectbox("FA", BOOLS, key="fa")
        d.selectbox("EPOC", BOOLS, key="epoc")
        a,b,c,d = st.columns(4)
        a.selectbox("Sd. Metab", BOOLS, key="sd_metab")
        b.selectbox("Tabaquismo", BOOLS, key="tabaco")
        c.selectbox("Enolismo", BOOLS, key="enolismo")
        d.selectbox("Hepatopatía", BOOLS, key="hepato")

    with tab2:
        a,b,c,d = st.columns(4)
        a.text_input("Hemoglobina", key="hb")
        b.text_input("Creatinina", key="creat")
        c.text_input("Cistatina C", key="cist_c")
        d.text_input("FGe", key="fge")
        a,b,c,d = st.columns(4)
        a.text_input("Urea", key="urea")
        b.text_input("Ácido úrico", key="ac_urico")
        c.text_input("Prot/Creat", key="prot_creat")
        d.text_input("Plaquetas", key="plaq")
        a,b,c,d = st.columns(4)
        a.text_input("AST", key="ast")
        b.text_input("ALT", key="alt")
        c.text_input("FIB-4 (calculado)", value=calculate_fib4(), disabled=True)
        d.text_input("Bilirrubina total", key="bili_t")
        a,b,c,d = st.columns(4)
        a.text_input("Bilirrubina directa", key="bili_d")
        b.text_input("Albúmina", key="albumina")
        c.text_input("HbA1c", key="hba1c")
        d.text_input("Colesterol", key="colest")
        st.selectbox("Biobanco", BOOLS, key="biobanco")
        st.markdown("**Biomarcadores**")
        a,b,c,d = st.columns(4)
        a.text_input("NT-proBNP", key="nt_probnp")
        b.text_input("CA125", key="ca125")
        c.text_input("Galectina-3", key="gal3")
        d.text_input("sST2", key="sst2")
        st.text_input("GDF-15", key="gdf15")

    with tab3:
        a,b,c,d = st.columns(4)
        a.text_input("FEVI (%)", key="fevi")
        b.text_input("GLS (%)", key="gls")
        c.text_input("Masa VI", key="masa_vi")
        d.text_input("TAPSE", key="tapse")
        a,b,c,d = st.columns(4)
        a.text_input("VAI", key="vai")
        b.text_input("VCI", key="vci")
        c.text_input("LSM (kPa)", key="lsm")
        d.text_input("CAP (dB/m)", key="cap")
        a,b,c,d = st.columns(4)
        a.text_input("Mediana elastografía", key="med_val")
        b.text_input("IQR / IQR-mediana", key="iqr_med")
        c.text_input("Días descompensación", key="dias_desc")
        d.text_input("NT-proBNP día prueba", key="nt_prueba")
        a,b = st.columns(2)
        a.text_input("Categoría fibrosis (sólo si documentada)", key="cat_fibro")
        b.selectbox("Edemas activos", BOOLS, key="edemas_prueba")
        st.info("LSM y categoría de fibrosis se mantienen separadas: la IA no infiere fibrosis a partir de LSM.")

    with tab4:
        a,b,c,d = st.columns(4)
        a.selectbox("IECA", BOOLS, key="ieca")
        b.selectbox("ARA2", BOOLS, key="ara2")
        c.selectbox("Betabloqueante", BOOLS, key="bb")
        d.selectbox("AMR", BOOLS, key="amr")
        a,b,c,d = st.columns(4)
        a.selectbox("SAC/VAL", BOOLS, key="sac_val")
        b.selectbox("SGLT2i", BOOLS, key="sglt2i")
        c.selectbox("Diurético asa", BOOLS, key="diur_asa")
        d.selectbox("HCTZ", BOOLS, key="hctz")
        a,b,c = st.columns(3)
        a.selectbox("Acetazolamida", BOOLS, key="acetazolamida")
        b.selectbox("Estatinas", BOOLS, key="estatinas")
        c.selectbox("Eritropoyetina", BOOLS, key="epo")

    with tab5:
        st.text_input("Meses seguimiento", key="meses_seg")
        a,b = st.columns(2)
        a.selectbox("Muerte CV", BOOLS, key="m_cv")
        b.text_input("Fecha muerte CV", key="f_m_cv")
        a,b = st.columns(2)
        a.selectbox("Hospitalización IC", BOOLS, key="hosp_ic")
        b.text_input("Fecha hosp. IC", key="f_hosp_ic")
        a,b = st.columns(2)
        a.selectbox("IAM no fatal", BOOLS, key="iam")
        b.text_input("Fecha IAM", key="f_iam")
        a,b = st.columns(2)
        a.selectbox("ACV no fatal", BOOLS, key="acv")
        b.text_input("Fecha ACV", key="f_acv")
        st.markdown("---")
        a,b = st.columns(2)
        a.selectbox("Muerte total", BOOLS, key="m_tot")
        b.text_input("Fecha muerte total", key="f_m_tot")
        a,b = st.columns(2)
        a.selectbox("Inicio TRS", BOOLS, key="trs")
        b.text_input("Fecha TRS", key="f_trs")
        a,b = st.columns(2)
        a.selectbox("Caída FGe ≥25%", BOOLS, key="caida_fge")
        b.text_input("Fecha caída FGe", key="f_caida_fge")
        a,b = st.columns(2)
        a.selectbox("Sd. CR agudo", BOOLS, key="sd_cr")
        b.text_input("Fecha Sd. CR agudo", key="f_sd_cr")


# ============================================================
# GUARDAR / EXPORTAR
# ============================================================

st.markdown("---")

warnings = validation_warnings()
if warnings:
    st.warning("Revisa antes de guardar:")
    for w in warnings:
        st.write("• " + w)

c1,c2,c3 = st.columns(3)

with c1:
    duplicate_new = (
        st.session_state.form_mode == "new"
        and st.session_state.id_exists
        and normalize_patient_id(st.session_state.get("id_pac", "")) == st.session_state.get("id_checked", "")
    )
    st.button(
        "💾 GUARDAR EN BASE CENTRAL",
        type="primary",
        use_container_width=True,
        on_click=handle_save_patient,
        disabled=duplicate_new
    )

with c2:
    try:
        excel_bytes, n = export_all_excel()
        st.download_button(
            "⬇️ EXPORTAR EXCEL ACTUAL",
            data=excel_bytes,
            file_name=f"CRD_Tesis_Cardiorrenal_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"No se pudo preparar el Excel: {e}")

with c3:
    if st.button("🔄 Actualizar datos", use_container_width=True):
        st.rerun()


# ============================================================
# HISTORIAL DEL PACIENTE
# ============================================================

if st.session_state.get("id_pac"):
    with st.expander("🕒 Historial de cambios del paciente"):
        try:
            history = db_audit(st.session_state.id_pac)
            if history:
                st.dataframe(history, use_container_width=True, hide_index=True)
            else:
                st.caption("Sin historial.")
        except Exception as e:
            st.caption(f"No se pudo cargar el historial: {e}")


# ============================================================
# ÚLTIMOS PACIENTES
# ============================================================

with st.expander("👥 Últimos pacientes modificados"):
    try:
        st.dataframe(db_recent(25), use_container_width=True, hide_index=True)
    except Exception as e:
        st.caption(f"No se pudo cargar la lista: {e}")

st.caption(
    "CRD Tesis Cardiorrenal V4 · ID única protegida y la base central es la fuente maestra. "
    "El Excel es una exportación para análisis/copia."
)
