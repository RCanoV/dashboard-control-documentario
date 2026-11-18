import streamlit as st
import pandas as pd
import requests
import base64
import json

GITHUB_REPO = "rcanov/dashboard-control-documentario"
RUTA_JSON = "config_reportes.json"

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
        "message": "Actualización de reportes jerárquicos desde Streamlit",
        "content": contenido_b64
    }
    if sha:
        payload["sha"] = sha
    return requests.put(url, headers=_get_headers(), json=payload)

def mostrar_formulario_reportes(df_catalogo: pd.DataFrame):
    """
    df_catalogo debe contener las columnas:
    ['ALCANCE', 'SUBPROYECTO', 'TIPO_OBRA', 'CONTRATO', 'COMPONENTE', 'PROJECT_ID', 'URL']
    """
    st.header("📝 Programador de Reportes Diarios por Jerarquía")

    if "GITHUB_TOKEN" not in st.secrets:
        st.error("Falta configurar `GITHUB_TOKEN` en los Secrets de Streamlit.")
        return

    config_actual, sha_actual = cargar_config_github()
    opciones_rep = ["-- Crear Nuevo Reporte --"] + list(config_actual.keys())
    seleccion_rep = st.selectbox("Selecciona un reporte guardado o crea uno nuevo:", opciones_rep)

    datos = config_actual.get(seleccion_rep, {}) if seleccion_rep != "-- Crear Nuevo Reporte --" else {}
    cfg_mail = datos.get("correo", {})

    # --- SELECTOR JERÁRQUICO DINÁMICO (fuera del form para actualización en vivo) ---
    st.subheader("1. Alcance del Reporte")
    niveles = {
        "Alcance": "ALCANCE",
        "Subproyecto": "SUBPROYECTO",
        "Tipo de Obra": "TIPO_OBRA",
        "Contrato": "CONTRATO",
        "Componente": "COMPONENTE"
    }

    col_n1, col_n2 = st.columns([1, 2])
    with col_n1:
        nivel_default = datos.get("nivel", "Contrato")
        idx_nivel = list(niveles.keys()).index(nivel_default) if nivel_default in niveles else 3
        nivel_elegido = st.selectbox("Programar reporte por:", list(niveles.keys()), index=idx_nivel)
        col_filtro = niveles[nivel_elegido]

    with col_n2:
        valores_disponibles = sorted(df_catalogo[col_filtro].dropna().unique().tolist())
        valores_previos = [v for v in datos.get("valores_nivel", []) if v in valores_disponibles]
        valores_elegidos = st.multiselect(f"Selecciona {nivel_elegido}(s):", valores_disponibles, default=valores_previos)

    # Filtro secundario opcional por Componente (si no se eligió Componente como nivel principal)
    df_filtrado = df_catalogo[df_catalogo[col_filtro].isin(valores_elegidos)] if valores_elegidos else pd.DataFrame()
    
    componentes_elegidos = []
    if nivel_elegido != "Componente" and not df_filtrado.empty:
        comps_disp = sorted(df_filtrado["COMPONENTE"].dropna().unique().tolist())
        comps_prev = [c for c in datos.get("componentes_filtro", []) if c in comps_disp]
        componentes_elegidos = st.multiselect(
            "Filtrar por Componente(s) específico(s) (vacío = todos):",
            comps_disp,
            default=comps_prev
        )
        if componentes_elegidos:
            df_filtrado = df_filtrado[df_filtrado["COMPONENTE"].isin(componentes_elegidos)]

    if not df_filtrado.empty:
        st.info(f"📂 Este reporte escaneará **{len(df_filtrado)} carpeta(s)** de **{df_filtrado['CONTRATO'].nunique()} contrato(s)**.")

    # --- FORMULARIO DE CONDICIONES, FORMATO Y CORREO ---
    with st.form("form_reporte_jerarquico"):
        st.subheader("2. Reglas, Formato y Envío")
        c1, c2 = st.columns(2)
        with c1:
            nombre = st.text_input(
                "Nombre del Reporte*",
                value="" if seleccion_rep == "-- Crear Nuevo Reporte --" else seleccion_rep,
                placeholder="Ej: Diario_Obras_Civiles_RFIs"
            )
            es_diario = st.checkbox("Activar en ejecución diaria automática", value=datos.get("es_diario", True))
        with c2:
            separar_hojas = st.checkbox("Separar en pestañas de Excel por Contrato/Componente", value=datos.get("separar_hojas", True))

        col_c1, col_c2 = st.columns(2)
        with col_c1:
            columnas_txt = st.text_input(
                "Columnas a exportar (separadas por coma; vacío = todas)",
                value=", ".join(datos.get("columnas", []))
            )
        with col_c2:
            condiciones_txt = st.text_area(
                "Condiciones de filtro (1 por línea: ==, !=, contiene)",
                value="\n".join(datos.get("condiciones", [])),
                height=80,
                placeholder="CAMPO_CUSTOM_6011838 == ABIERTO\nCAMPO_CUSTOM_6011840 == SHP"
            )

        st.markdown("---")
        activar_mail = st.checkbox("Enviar Excel por correo al finalizar", value=cfg_mail.get("activar", True))
        m1, m2 = st.columns(2)
        with m1:
            para = st.text_input("Para (separados por coma)", value=cfg_mail.get("para", ""))
        with m2:
            cc = st.text_input("CC (opcional)", value=cfg_mail.get("cc", ""))

        asunto = st.text_input(
            "Asunto (variables: {REPORTE}, {NIVEL}, {SELECCION}, {FECHA}, {TOTAL})",
            value=cfg_mail.get("asunto", "⚠️ {REPORTE} | {NIVEL}: {SELECCION} ({FECHA}) - {TOTAL} registros")
        )
        cuerpo = st.text_area(
            "Cuerpo del correo",
            value=cfg_mail.get("cuerpo", "Estimados,\n\nSe adjunta el {REPORTE} ({NIVEL}: {SELECCION}) actualizado al {FECHA}.\nTotal de registros filtrados: {TOTAL}.\n\nSaludos,\nRoberto Cano Vilchez\nControl Documentario SHP"),
            height=110
        )

        if st.form_submit_button("💾 Guardar Reporte en GitHub", type="primary"):
            if not nombre.strip() or df_filtrado.empty:
                st.warning("Ingresa un nombre y selecciona al menos un elemento de la jerarquía.")
            else:
                # Resolvemos las carpetas exactas para que GitHub Actions las ejecute directo
                objetivos = []
                for _, row in df_filtrado.iterrows():
                    objetivos.append({
                        "alcance": str(row["ALCANCE"]),
                        "subproyecto": str(row["SUBPROYECTO"]),
                        "tipo_obra": str(row["TIPO_OBRA"]),
                        "contrato": str(row["CONTRATO"]),
                        "componente": str(row["COMPONENTE"]),
                        "project_id": str(row["PROJECT_ID"]),
                        "url": str(row["URL"])
                    })

                config_actual[nombre.strip()] = {
                    "nivel": nivel_elegido,
                    "valores_nivel": valores_elegidos,
                    "componentes_filtro": componentes_elegidos,
                    "es_diario": es_diario,
                    "separar_hojas": separar_hojas,
                    "columnas": [c.strip() for c in columnas_txt.split(",") if c.strip()],
                    "condiciones": [c.strip() for c in condiciones_txt.splitlines() if c.strip()],
                    "objetivos": objetivos,
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
                    st.success(f"✅ Reporte '{nombre.strip()}' guardado con {len(objetivos)} carpeta(s).")
                    st.rerun()
                else:
                    st.error(f"Error al guardar en GitHub: {res.text}")

    if seleccion_rep != "-- Crear Nuevo Reporte --":
        if st.button(f"🗑️ Eliminar reporte '{seleccion_rep}'"):
            del config_actual[seleccion_rep]
            res = guardar_config_github(config_actual, sha_actual)
            if res.status_code in [200, 201]:
                st.success(f"Reporte '{seleccion_rep}' eliminado.")
                st.rerun()
