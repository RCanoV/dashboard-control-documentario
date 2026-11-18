import streamlit as st
import requests
import base64
import json

GITHUB_REPO = "rcanov/dashboard-control-documentario"
RUTA_JSON = "config_reportes.json"

PROYECTOS_ACC = {
    "Muelle": "b.3fe40740-4483-4ed9-895b-9cb579cd7b1d",
    "Línea 3": "b.24914611-716e-4e2b-a8a2-bf28757efbe9",
    "Subestación Mina 02": "b.18ec0b1a-6335-472d-95bb-854a9259704e",
    "Otro (Ingresar Project ID)": ""
}

def _get_headers():
    return {
        "Authorization": f"token {st.secrets['GITHUB_TOKEN']}",
        "Accept": "application/vnd.github.v3+json"
    }

def cargar_config_github():
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{RUTA_JSON}"
    r = requests.get(url, headers=_get_headers())
    if r.status_code == 200:
        data = r.json()
        contenido = base64.b64decode(data["content"]).decode("utf-8")
        return json.loads(contenido), data["sha"]
    return {}, None

def guardar_config_github(config_dict, sha=None):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{RUTA_JSON}"
    contenido_b64 = base64.b64encode(
        json.dumps(config_dict, indent=4, ensure_ascii=False).encode("utf-8")
    ).decode("utf-8")
    payload = {
        "message": "Actualización de reportes diarios desde Streamlit",
        "content": contenido_b64
    }
    if sha:
        payload["sha"] = sha
    return requests.put(url, headers=_get_headers(), json=payload)

def mostrar_formulario_reportes():
    st.header("📝 Creador de Reportes Diarios Automáticos")
    
    if "GITHUB_TOKEN" not in st.secrets:
        st.error("Falta configurar `GITHUB_TOKEN` en los Secrets de Streamlit.")
        return

    config_actual, sha_actual = cargar_config_github()
    opciones = ["-- Crear Nuevo Reporte --"] + list(config_actual.keys())
    seleccion = st.selectbox("Selecciona un reporte para editar o crea uno nuevo:", opciones)

    datos = config_actual.get(seleccion, {}) if seleccion != "-- Crear Nuevo Reporte --" else {}
    cfg_mail = datos.get("correo", {})

    with st.form("form_reporte_acc"):
        col1, col2 = st.columns(2)
        with col1:
            nombre = st.text_input("Nombre del Reporte*", value="" if seleccion == "-- Crear Nuevo Reporte --" else seleccion, placeholder="Ej: Reporte_Diario_RFIs_L3")
            proy_default = datos.get("proyecto", "Línea 3")
            idx_proy = list(PROYECTOS_ACC.keys()).index(proy_default) if proy_default in PROYECTOS_ACC else 3
            proyecto = st.selectbox("Proyecto ACC", list(PROYECTOS_ACC.keys()), index=idx_proy)
        with col2:
            project_id_custom = st.text_input("Project ID (solo si elegiste 'Otro')", value=datos.get("project_id", "") if idx_proy == 3 else "")
            es_diario = st.checkbox("Activar en ejecución diaria automática", value=datos.get("es_diario", True))
            separar_hojas = st.checkbox("Crear una pestaña de Excel por cada carpeta", value=datos.get("separar_hojas", True))

        urls_txt = st.text_area(
            "URLs de Carpetas ACC (una por línea. Formato opcional: NombreHoja | URL)",
            value="\n".join(datos.get("urls", [])),
            height=100,
            placeholder="RFIs | https://acc.autodesk.com/docs/files/projects/...\nCartas | https://acc.autodesk.com/docs/files/projects/..."
        )

        col_c1, col_c2 = st.columns(2)
        with col_c1:
            columnas_txt = st.text_input(
                "Columnas a exportar (separadas por coma; vacío = todas)",
                value=", ".join(datos.get("columnas", []))
            )
        with col_c2:
            condiciones_txt = st.text_area(
                "Condiciones de filtro (una por línea: ==, !=, contiene)",
                value="\n".join(datos.get("condiciones", [])),
                height=80,
                placeholder="CAMPO_CUSTOM_6011838 == ABIERTO\nCAMPO_CUSTOM_6011840 == SHP"
            )

        st.markdown("---")
        st.subheader("📧 Configuración de Correo")
        activar_mail = st.checkbox("Enviar Excel por correo al finalizar", value=cfg_mail.get("activar", True))
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            para = st.text_input("Para (correos separados por coma)", value=cfg_mail.get("para", ""))
        with col_m2:
            cc = st.text_input("CC (opcional, separados por coma)", value=cfg_mail.get("cc", ""))

        asunto = st.text_input(
            "Asunto (variables: {REPORTE}, {PROYECTO}, {FECHA}, {TOTAL})",
            value=cfg_mail.get("asunto", "⚠️ {REPORTE} - {PROYECTO} ({FECHA}) | {TOTAL} registros")
        )
        cuerpo = st.text_area(
            "Cuerpo del correo",
            value=cfg_mail.get("cuerpo", "Estimados,\n\nSe adjunta el {REPORTE} del proyecto {PROYECTO} al {FECHA}.\nTotal de registros encontrados: {TOTAL}.\n\nSaludos,\nRoberto Cano Vilchez\nControl Documentario SHP"),
            height=120
        )

        btn_guardar = st.form_submit_button("💾 Guardar Reporte en GitHub", type="primary")

        if btn_guardar:
            if not nombre.strip():
                st.warning("Debes ingresar un nombre para el reporte.")
            else:
                pid_final = project_id_custom.strip() if proyecto == "Otro (Ingresar Project ID)" else PROYECTOS_ACC[proyecto]
                config_actual[nombre.strip()] = {
                    "proyecto": proyecto,
                    "project_id": pid_final,
                    "es_diario": es_diario,
                    "separar_hojas": separar_hojas,
                    "urls": [u.strip() for u in urls_txt.splitlines() if u.strip()],
                    "columnas": [c.strip() for c in columnas_txt.split(",") if c.strip()],
                    "condiciones": [c.strip() for c in condiciones_txt.splitlines() if c.strip()],
                    "correo": {
                        "activar": activar_mail,
                        "para": para.strip(),
                        "cc": cc.strip(),
                        "asunto": asunto,
                        "cuerpo": cuerpo
                    }
                }
                res = guardar_config_github(config_actual, sha_actual)
                if res.status_code in [200, 201]:
                    st.success(f"✅ Reporte '{nombre.strip()}' guardado correctamente.")
                    st.rerun()
                else:
                    st.error(f"Error al guardar en GitHub: {res.text}")

    # Botón para eliminar reporte existente
    if seleccion != "-- Crear Nuevo Reporte --":
        if st.button(f"🗑️ Eliminar reporte '{seleccion}'"):
            del config_actual[seleccion]
            res = guardar_config_github(config_actual, sha_actual)
            if res.status_code in [200, 201]:
                st.success(f"Reporte '{seleccion}' eliminado.")
                st.rerun()
