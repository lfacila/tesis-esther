import streamlit as st
import json
import os
import re
from datetime import datetime
from io import BytesIO
import openpyxl

# ============================================================
# CRD TESIS CARDIORRENAL V3
# Base central Supabase + exportación Excel
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
    "edad": "", "sexo": "", "peso": "", "talla": "", "eti_erc": "", "eti_ic": "",
    "dm2": "No", "hta": "No", "fa": "No", "epoc": "No", "sd_metab": "No",
    "tabaco": "No", "enolismo": "No", "hepato": "No",
    "hb": "", "creat": "", "cist_c": "", "fge": "", "urea": "", "ac_urico": "",
    "prot_creat": "", "ast": "", "alt": "", "plaq": "", "bili_t": "", "bili_d": "",
    "albumina": "", "hba1c": "", "colest": "", "nt_probnp": "", "ca125": "",
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

def db_upsert_patient(data):
    patient_id = str(data["id_pac"]).strip()
    now = datetime.now().isoformat()

    old = db_get_patient(patient_id)
    payload = {
        "id_pac": patient_id,
        "data": data,
        "updated_at": now,
        "updated_by": st.session_state.get("user_label", "usuario")
    }

    res = supabase.table("patients").upsert(payload, on_conflict="id_pac").execute()

    # Audit record: store complete snapshot + timestamp.
    supabase.table("audit_log").insert({
        "id_pac": patient_id,
        "action": "UPDATE" if old else "CREATE",
        "snapshot": data,
        "changed_at": now,
        "changed_by": st.session_state.get("user_label", "usuario")
    }).execute()

    return bool(res.data), ("actualizado" if old else "nuevo")

def db_load_patient(patient_id):
    row = db_get_patient(patient_id)
    if not row:
        return False

    data = row.get("data") or {}
    for k in DEFAULTS:
        st.session_state[k] = data.get(k, DEFAULTS[k])

    st.session_state.id_pac = patient_id
    st.session_state.loaded_updated_at = row.get("updated_at", "")
    return True

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

AI_KEYS = list(DEFAULTS.keys())

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
        if k not in DEFAULTS or v is None or str(v).strip() == "":
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
        (1,"id_pac"),(2,"fecha_inc"),(3,"edad"),(4,"sexo"),(5,"peso"),(6,"talla"),
        (8,"eti_erc"),(9,"eti_ic"),(10,"dm2"),(11,"hta"),(12,"fa"),(13,"epoc"),
        (14,"sd_metab"),(15,"tabaco"),(16,"enolismo"),(17,"hepato")
    ],
    "02_Analitica_Biomarcadores": [
        (1,"id_pac"),(2,"hb"),(3,"creat"),(4,"cist_c"),(5,"fge"),(6,"urea"),
        (7,"ac_urico"),(8,"prot_creat"),(9,"ast"),(10,"alt"),(11,"plaq"),
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

st.info(
    "Esta aplicación utiliza una base central. Cualquier Mac, Windows o navegador "
    "que entre en esta misma aplicación trabaja sobre los mismos datos."
)


# ============================================================
# BUSCADOR
# ============================================================

top1, top2, top3 = st.columns([1.5, 1, 1])

with top1:
    search_id = st.text_input("ID paciente", value=st.session_state.get("id_pac",""))

with top2:
    if st.button("🔍 Cargar paciente", use_container_width=True):
        if search_id.strip():
            if db_load_patient(search_id.strip()):
                st.success("Paciente cargado desde la base central.")
                st.rerun()
            else:
                st.warning("Ese ID no existe. Puedes crear un paciente nuevo.")

with top3:
    if st.button("🔄 Nuevo paciente", use_container_width=True):
        for k, v in DEFAULTS.items():
            st.session_state[k] = v
        st.session_state.last_ai_data = {}
        st.rerun()


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

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "01. Clínicos","02. Analítica","03. Eco/Fibro","04. Tratamiento","05. Seguimiento"
    ])

    with tab1:
        a,b,c,d = st.columns(4)
        a.text_input("ID Paciente", key="id_pac")
        b.text_input("Fecha inclusión", key="fecha_inc")
        c.selectbox("Sexo", ["","H","M"], key="sexo")
        d.text_input("Edad", key="edad")

        a,b,c,d = st.columns(4)
        a.text_input("Peso (kg)", key="peso")
        b.text_input("Talla (cm)", key="talla")
        c.selectbox("Etiología ERC", ["","DM2","HTA","Glomerulonefritis","Poliquistosis","Otras"], key="eti_erc")
        d.selectbox("Etiología IC", ["","Isquémica","Hipertensiva","MCD","HFpEF","Valvular"], key="eti_ic")

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
        c.text_input("Bilirrubina total", key="bili_t")
        d.text_input("Bilirrubina directa", key="bili_d")
        a,b,c,d = st.columns(4)
        a.text_input("Albúmina", key="albumina")
        b.text_input("HbA1c", key="hba1c")
        c.text_input("Colesterol", key="colest")
        d.selectbox("Biobanco", BOOLS, key="biobanco")
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
    if st.button("💾 GUARDAR EN BASE CENTRAL", type="primary", use_container_width=True):
        if not st.session_state.id_pac.strip():
            st.error("El ID de paciente es obligatorio.")
        elif warnings:
            st.error("Corrige primero los valores fuera de rango.")
        else:
            data = {k: st.session_state.get(k, DEFAULTS[k]) for k in DEFAULTS}
            try:
                ok, mode = db_upsert_patient(data)
                if ok:
                    st.success(
                        f"Paciente {data['id_pac']} {mode} en la BASE CENTRAL."
                    )
                    st.rerun()
            except Exception as e:
                st.error(f"No se pudo guardar en la base central: {e}")

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
    "CRD Tesis Cardiorrenal V3 · La base central es la fuente maestra. "
    "El Excel es una exportación para análisis/copia."
)
