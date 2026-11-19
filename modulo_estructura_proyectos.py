import streamlit as st
import pandas as pd
import requests
import base64
import json

GITHUB_REPO = "rcanov/dashboard-control-documentario"
RUTA_ESTRUCTURA = "estructura_proyectos.json"

def _get_headers():
    return {
        "Authorization": f"token {st.secrets['GITHUB_TOKEN']}",
        "Accept": "application/vnd.github.v3+json"
    }

def cargar_estructura_github(estructura_respaldo: dict):
    """Carga la estructura desde GitHub; si aún no existe o está vacía, usa la de respaldo."""
    if "GITHUB_TOKEN" not in st.secrets:
        return estructura_respaldo, None
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{RUTA_ESTRUCTURA}"
    r = requests.get(url, headers=_get_headers())
    if r.status_code == 200:
        data = r.json()
        contenido = base64.b64decode(data["content"]).decode("utf-8").strip()
        if contenido:
            try:
                parsed = json.loads(contenido)
                if parsed:
                    return parsed, data["sha"]
            except json.JSONDecodeError:
                pass
        return estructura_respaldo, data["sha"]
    return estructura_respaldo, None

def guardar_estructura_github(estructura_dict: dict, sha=None):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{RUTA_ESTRUCTURA}"
    contenido_b64 = base64.b64encode(
        json.dumps(estructura_dict, indent=4, ensure_ascii=False).encode("utf-8")
    ).decode("utf-8")
    payload = {
        "message": "Actualización completa de Estructura de Proyectos desde Streamlit",
        "content": contenido_b64
    }
    if sha:
        payload["sha"] = sha
    return requests.put(url, headers=_get_headers(), json=payload)

def mostrar_editor_estructura(estructura_actual: dict, sha_actual=None):
    st.header("🏗️ Administrar Estructura de Proyectos y Contratos")

    if "GITHUB_TOKEN" not in st.secrets:
        st.error("Falta configurar `GITHUB_TOKEN` en los Secrets de Streamlit.")
        return

    # ====================================================================
    # 1. UBICACIÓN EN LA JERARQUÍA (Fuera del form para actualización en vivo)
    # ====================================================================
    st.subheader("1. Selecciona o crea la ubicación en la jerarquía")
    c1, c2 = st.columns(2)

    with c1:
        alcances = list(estructura_actual.keys()) + ["➕ Nuevo Tipo de Alcance"]
        alc_sel = st.selectbox("1. Tipo de Alcance:", alcances)
        alcance_final = st.text_input("Nombre del nuevo Alcance:") if alc_sel == "➕ Nuevo Tipo de Alcance" else alc_sel

        subs_dict = estructura_actual.get(alcance_final, {}) if alcance_final in estructura_actual else {}
        subproyectos = [k for k in subs_dict.keys() if k != "config"] + ["➕ Nuevo Subproyecto"]
        sub_sel = st.selectbox("2. Subproyecto:", subproyectos)
        sub_final = st.text_input("Nombre del nuevo Subproyecto:") if sub_sel == "➕ Nuevo Subproyecto" else sub_sel

    with c2:
        sub_data = subs_dict.get(sub_final, {}) if sub_final in subs_dict else {}
        tipos_obra = [k for k in sub_data.keys() if k != "config"] + ["➕ Nuevo Tipo de Obra"]
        obra_sel = st.selectbox("3. Tipo de Obra:", tipos_obra)
        obra_final = st.text_input("Nombre del nuevo Tipo de Obra:") if obra_sel == "➕ Nuevo Tipo de Obra" else obra_sel

        contratos_dict = sub_data.get(obra_final, {}) if obra_final in sub_data else {}
        contratos = list(contratos_dict.keys()) + ["➕ Nuevo Contrato / Componente"]
        cont_sel = st.selectbox("4. Contrato / Componente:", contratos)
        cont_final = st.text_input("Nombre del nuevo Contrato / Componente:") if cont_sel == "➕ Nuevo Contrato / Componente" else cont_sel

    # Precargar datos existentes si se está editando
    info_cont = contratos_dict.get(cont_final, {}) if cont_final in contratos_dict else {}
    cfg_sub = sub_data.get("config", {
        "campos_personalizados": {
            "6011837": "REV",
            "6011840": "PENDIENTE POR",
            "6011839": "FECHA",
            "6011838": "ESTADO"
        },
        "id_campo_fecha": "6011839",
        "id_campo_pendiente": "6011840"
    })

    st.markdown("---")

    # ====================================================================
    # 2. CONFIGURACIÓN DE CAMPOS PERSONALIZADOS DEL SUBPROYECTO
    # ====================================================================
    st.subheader("2. Campos Personalizados del Subproyecto (Autodesk ACC)")
    st.caption("Edita los IDs y los nombres de columna, o agrega nuevas filas al final de la tabla.")

    campos_dict = cfg_sub.get("campos_personalizados", {})
    df_campos_init = pd.DataFrame([
        {"ID_CAMPO": str(k), "NOMBRE_COLUMNA": str(v)} for k, v in campos_dict.items()
    ]) if campos_dict else pd.DataFrame(columns=["ID_CAMPO", "NOMBRE_COLUMNA"])

    df_campos_edit = st.data_editor(
        df_campos_init,
        num_rows="dynamic",
        use_container_width=True,
        key=f"editor_campos_{sub_final}"
    )

    col_id1, col_id2 = st.columns(2)
    with col_id1:
        id_campo_fecha = st.text_input(
            "ID del campo usado para FECHA (cálculo de días):",
            value=str(cfg_sub.get("id_campo_fecha", "6011839"))
        )
    with col_id2:
        id_campo_pendiente = st.text_input(
            "ID del campo usado para PENDIENTE POR (estados y colores):",
            value=str(cfg_sub.get("id_campo_pendiente", "6011840"))
        )

    st.markdown("---")

    # ====================================================================
    # 3. DATOS DEL CONTRATO, CARPETAS Y URLS EXCLUIDAS
    # ====================================================================
    st.subheader("3. Datos Generales del Contrato")

    f1, f2 = st.columns(2)
    with f1:
        nombre_proy = st.text_input("Nombre del Proyecto:", value=info_cont.get("nombre_proyecto", ""))
        tipos_c = ["Con Supervisión", "Contrato Directo", "Tripartito"]
        idx_tc = tipos_c.index(info_cont.get("tipo_contrato", "Con Supervisión")) if info_cont.get("tipo_contrato") in tipos_c else 0
        tipo_contrato = st.selectbox("Tipo de Contrato:", tipos_c, index=idx_tc)
        supervisor = st.text_input("Nombre del Supervisor:", value=info_cont.get("nombre_supervisor", ""))
    with f2:
        coordinador = st.text_input("Coordinador SHP:", value=info_cont.get("coordinador_shp", ""))
        admin_cont = st.text_input("Administrador de Contratos:", value=info_cont.get("administrador_contratos", ""))

    # Campos adicionales solo si es Tripartito
    empresa_1, empresa_2 = "", ""
    if tipo_contrato == "Tripartito":
        t1, t2 = st.columns(2)
        with t1:
            empresa_1 = st.text_input("Nombre Empresa 1 (Tripartito):", value=info_cont.get("nombre_empresa_1", ""))
        with t2:
            empresa_2 = st.text_input("Nombre Empresa 2 (Tripartito):", value=info_cont.get("nombre_empresa_2", ""))

    st.markdown("---")
    st.subheader("4. Links de Carpetas a Buscar (`carpetas_config`)")
    st.caption("Agrega o edita cada carpeta del contrato con su URL de Autodesk ACC y la longitud máxima de subcarpeta.")

    carpetas_existentes = info_cont.get("carpetas_config", [])
    df_carpetas_init = pd.DataFrame(carpetas_existentes) if carpetas_existentes else pd.DataFrame(
        [{"nombre": "Construccion", "url": "", "longitud_maxima": 10}]
    )
    for col_req, default_val in [("nombre", ""), ("url", ""), ("longitud_maxima", 10)]:
        if col_req not in df_carpetas_init.columns:
            df_carpetas_init[col_req] = default_val
    df_carpetas_init = df_carpetas_init[["nombre", "url", "longitud_maxima"]]

    df_carpetas_edit = st.data_editor(
        df_carpetas_init,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "nombre": st.column_config.TextColumn("Nombre Carpeta / Pestaña", required=True),
            "url": st.column_config.TextColumn("URL de Autodesk ACC", width="large", required=True),
            "longitud_maxima": st.column_config.NumberColumn("Longitud Máx.", min_value=1, max_value=100, step=1, default=10)
        },
        key=f"editor_carpetas_{cont_final}"
    )

    st.subheader("5. URLs Excluidas")
    excluidas_txt = st.text_area(
        "Pega las URLs de carpetas a excluir del escaneo (una por línea):",
        value="\n".join(info_cont.get("urls_excluidas", [])),
        height=90
    )

    st.markdown("---")

    # ====================================================================
    # 4. BOTONES DE GUARDADO Y ELIMINACIÓN
    # ====================================================================
    b1, b2 = st.columns([2, 1])
    with b1:
        if st.button("💾 Guardar Configuración en GitHub", type="primary", use_container_width=True):
            if not all([alcance_final.strip(), sub_final.strip(), obra_final.strip(), cont_final.strip()]):
                st.warning("Completa los 4 niveles de la jerarquía (Alcance, Subproyecto, Tipo de Obra y Contrato).")
            else:
                # Reconstruir diccionario de campos personalizados
                nuevos_campos = {}
                for _, row in df_campos_edit.iterrows():
                    k_id = str(row.get("ID_CAMPO", "")).strip()
                    v_nom = str(row.get("NOMBRE_COLUMNA", "")).strip()
                    if k_id and v_nom and k_id.lower() != "nan":
                        nuevos_campos[k_id] = v_nom

                # Reconstruir lista de carpetas
                nuevas_carpetas = []
                for _, row in df_carpetas_edit.iterrows():
                    c_nom = str(row.get("nombre", "")).strip()
                    c_url = str(row.get("url", "")).strip()
                    c_len = row.get("longitud_maxima", 10)
                    if c_nom and c_url and c_nom.lower() != "nan" and c_url.lower() != "nan":
                        nuevas_carpetas.append({
                            "nombre": c_nom,
                            "url": c_url,
                            "longitud_maxima": int(c_len) if pd.notna(c_len) else 10
                        })

                # Armar objeto del contrato
                datos_contrato = {
                    "nombre_proyecto": nombre_proy.strip(),
                    "tipo_contrato": tipo_contrato,
                    "nombre_supervisor": supervisor.strip(),
                    "coordinador_shp": coordinador.strip(),
                    "administrador_contratos": admin_cont.strip(),
                    "urls_excluidas": [u.strip() for u in excluidas_txt.splitlines() if u.strip()],
                    "carpetas_config": nuevas_carpetas
                }
                if tipo_contrato == "Tripartito":
                    datos_contrato["nombre_empresa_1"] = empresa_1.strip()
                    datos_contrato["nombre_empresa_2"] = empresa_2.strip()

                # Guardar en la jerarquía
                estructura_actual.setdefault(alcance_final.strip(), {})
                estructura_actual[alcance_final.strip()].setdefault(sub_final.strip(), {})
                estructura_actual[alcance_final.strip()][sub_final.strip()]["config"] = {
                    "campos_personalizados": nuevos_campos,
                    "id_campo_fecha": id_campo_fecha.strip(),
                    "id_campo_pendiente": id_campo_pendiente.strip()
                }
                estructura_actual[alcance_final.strip()][sub_final.strip()].setdefault(obra_final.strip(), {})
                estructura_actual[alcance_final.strip()][sub_final.strip()][obra_final.strip()][cont_final.strip()] = datos_contrato

                res = guardar_estructura_github(estructura_actual, sha_actual)
                if res.status_code in [200, 201]:
                    st.success(f"✅ Contrato '{cont_final.strip()}' guardado en GitHub correctamente.")
                    st.rerun()
                else:
                    st.error(f"Error al guardar en GitHub: {res.text}")

    with b2:
        if cont_sel != "➕ Nuevo Contrato / Componente" and cont_final in contratos_dict:
            if st.button(f"🗑️ Eliminar Contrato", use_container_width=True):
                del estructura_actual[alcance_final][sub_final][obra_final][cont_final]
                res = guardar_estructura_github(estructura_actual, sha_actual)
                if res.status_code in [200, 201]:
                    st.success(f"Contrato '{cont_final}' eliminado.")
                    st.rerun()
                else:
                    st.error(f"Error al eliminar: {res.text}")
