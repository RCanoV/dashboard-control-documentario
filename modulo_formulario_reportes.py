import streamlit as st
import pandas as pd
import requests
import base64
import json
import re

GITHUB_REPO = "rcanov/dashboard-control-documentario"
RUTA_JSON = "config_reportes.json"
ARCHIVO_WORKFLOW = "reportes_diarios.yml"

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
HORARIOS_PERU = [f"{h:02d}:00" for h in range(5, 23)]
COLUMNAS_BASE = [
    "ALCANCE", "SUBPROYECTO", "TIPO_OBRA", "CONTRATO", "CARPETA",
    "CODIGO", "DESCRIPCION", "DIAS_TRANSCURRIDOS"
]

MODOS_ENVIO = [
    "☁️ Nube (GitHub + Gmail)",
    "💻 Local (Mi PC + Outlook SHP)"
]

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _validar_correos(texto_correos: str):
    """Separa por comas o punto y coma y devuelve (lista_validos, lista_invalidos)."""
    crudos = [x.strip() for x in re.split(r"[;,]", texto_correos) if x.strip()]
    validos = [c for c in crudos if EMAIL_REGEX.match(c)]
    invalidos = [c for c in crudos if not EMAIL_REGEX.match(c)]
    return validos, invalidos


def _get_headers():
    return {
        "Authorization": f"token {st.secrets['GITHUB_TOKEN']}",
        "Accept": "application/vnd.github.v3+json"
    }


def cargar_config_github():
    if "GITHUB_TOKEN" not in st.secrets:
        return {}, None
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{RUTA_JSON}"
    r = requests.get(url, headers=_get_headers())
    if r.status_code == 200:
        data = r.json()
        contenido = base64.b64decode(data["content"]).decode("utf-8").strip()
        if not contenido:
            return {}, data["sha"]
        try:
            return json.loads(contenido), data["sha"]
        except json.JSONDecodeError:
            return {}, data["sha"]
    return {}, None


def guardar_config_github(config_dict, sha=None):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{RUTA_JSON}"
    config_ordenado = dict(
        sorted(config_dict.items(), key=lambda item: item[0].strip().lower())
    )
    contenido_b64 = base64.b64encode(
        json.dumps(config_ordenado, indent=4, ensure_ascii=False).encode("utf-8")
    ).decode("utf-8")
    payload = {
        "message": "Actualización de reportes programados desde Streamlit",
        "content": contenido_b64
    }
    if sha:
        payload["sha"] = sha
    return requests.put(url, headers=_get_headers(), json=payload)


def disparar_envio_github_actions(reporte_objetivo="", rama="main"):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/actions/workflows/{ARCHIVO_WORKFLOW}/dispatches"
    payload = {
        "ref": rama,
        "inputs": {
            "reporte_objetivo": reporte_objetivo
        }
    }
    return requests.post(url, headers=_get_headers(), json=payload)


def _extraer_catalogo_completo(estructura_maestra):
    """Extrae contratos, carpetas configuradas, ruta local y campos personalizados."""
    registros = []
    for alcance, subs in estructura_maestra.items():
        if not isinstance(subs, dict):
            continue
        for sub, obras in subs.items():
            if sub == "config" or not isinstance(obras, dict):
                continue
            cfg_sub = obras.get("config", {})
            campos_pers = list(cfg_sub.get("campos_personalizados", {}).values())
            for obra, contratos in obras.items():
                if obra == "config" or not isinstance(contratos, dict):
                    continue
                for contrato, info in contratos.items():
                    if not isinstance(info, dict):
                        continue
                    carpetas = info.get("carpetas_config", [])
                    ruta_local = info.get("ruta_guardado_local", "")
                    for carp in carpetas:
                        registros.append({
                            "ALCANCE": alcance,
                            "SUBPROYECTO": sub,
                            "TIPO_OBRA": obra,
                            "CONTRATO": contrato,
                            "CARPETA": carp.get("nombre", ""),
                            "URL": carp.get("url", ""),
                            "LONGITUD_MAX": carp.get("longitud_maxima", 10),
                            "CAMPOS_PERS": campos_pers,
                            "CONFIG_SUB": cfg_sub,
                            "URLS_EXCLUIDAS": info.get("urls_excluidas", []),
                            "RUTA_LOCAL": ruta_local
                        })
    return pd.DataFrame(registros)


def mostrar_formulario_reportes(estructura_maestra):
    st.header("📝 Programar Reportes Automáticos")

    if "GITHUB_TOKEN" not in st.secrets:
        st.error("Falta configurar `GITHUB_TOKEN` en los Secrets de Streamlit.")
        return

    config_actual, sha_actual = cargar_config_github()
    df_cat = _extraer_catalogo_completo(estructura_maestra)

    if df_cat.empty:
        st.warning("No hay carpetas configuradas en la Estructura de Proyectos.")
        return

    # 0. Resumen general de todos los reportes programados
    if config_actual:
        with st.expander("📋 Resumen General de Reportes Programados", expanded=True):
            filas_resumen = []
            for nom_r, cfg_r in sorted(config_actual.items(), key=lambda item: item[0].strip().lower()):
                conds = cfg_r.get("condiciones_filtro", [])
                filtro_txt = " | ".join(
                    f"{c.get('campo', '')} {c.get('operador', '')} {c.get('valor', '')}"
                    for c in conds
                ) or "Sin filtro"

                dias_r = cfg_r.get("dias_semana", [])
                dias_txt = (
                    "Lun - Vie"
                    if dias_r == ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
                    else ", ".join(dias_r)
                )

                filas_resumen.append({
                    "Reporte": nom_r,
                    "Estado": "🟢 Activo" if cfg_r.get("activo", True) else "🔴 Inactivo",
                    "Horario": ", ".join(cfg_r.get("horarios", [])),
                    "Días": dias_txt,
                    "Filtro": filtro_txt,
                    "Para": ", ".join(cfg_r.get("para", [])),
                    "CC": ", ".join(cfg_r.get("cc", [])),
                })

            df_res_rep = pd.DataFrame(filas_resumen)
            horarios_disp = ["Todos"] + sorted(
                {h for cfg_r in config_actual.values() for h in cfg_r.get("horarios", [])}
            )
            filtro_hora = st.radio(
                "Filtrar por horario de envío:",
                horarios_disp,
                horizontal=True
            )

            if filtro_hora != "Todos":
                df_res_rep = df_res_rep[df_res_rep["Horario"].str.contains(filtro_hora, na=False)]

            st.caption(f"Mostrando **{len(df_res_rep)}** de **{len(filas_resumen)}** reportes configurados.")
            st.dataframe(df_res_rep, use_container_width=True, hide_index=True)

    # 1. Elegir reporte existente o crear uno nuevo
    opciones_rep = ["➕ Nuevo Reporte"] + sorted(
        config_actual.keys(), key=lambda x: x.strip().lower()
    )
    rep_sel = st.selectbox("Selecciona un reporte para editar o crea uno nuevo:", opciones_rep)
    datos_prev = config_actual.get(rep_sel, {}) if rep_sel != "➕ Nuevo Reporte" else {}

    col_n1, col_n2 = st.columns([2, 1.3])
    with col_n1:
        nombre_reporte = st.text_input(
            "Nombre identificador del Reporte:",
            value="" if rep_sel == "➕ Nuevo Reporte" else rep_sel,
            placeholder="Ej: Reporte Diario Tercera Linea - Ingenieria"
        )
    with col_n2:
        modo_prev = datos_prev.get("modo_envio", MODOS_ENVIO[0])
        idx_modo = MODOS_ENVIO.index(modo_prev) if modo_prev in MODOS_ENVIO else 0
        modo_envio = st.selectbox("Modo de Envío:", MODOS_ENVIO, index=idx_modo)

    activo = st.checkbox("Reporte Activo", value=datos_prev.get("activo", True))

    if modo_envio == MODOS_ENVIO[1]:
        st.info("💻 Este reporte se ejecutará desde tu PC (`ejecutar_reportes_local.py`), se guardará en la carpeta local del contrato y se enviará por tu Outlook corporativo.")

    st.markdown("---")

    # 2. Selección de Contratos y Carpetas dinámicas
    st.subheader("1. Alcance y Carpetas a Escanear")
    c1, c2 = st.columns(2)
    with c1:
        alc_disp = sorted(df_cat["ALCANCE"].unique().tolist())
        alc_sel = st.multiselect("Alcance:", alc_disp, default=datos_prev.get("filtro_alcance", alc_disp))
        df_f1 = df_cat[df_cat["ALCANCE"].isin(alc_sel)] if alc_sel else df_cat

        sub_disp = sorted(df_f1["SUBPROYECTO"].unique().tolist())
        sub_prev = [s for s in datos_prev.get("filtro_subproyecto", sub_disp) if s in sub_disp]
        sub_sel = st.multiselect("Subproyecto:", sub_disp, default=sub_prev or sub_disp)
        df_f2 = df_f1[df_f1["SUBPROYECTO"].isin(sub_sel)] if sub_sel else df_f1

    with c2:
        cont_disp = sorted(df_f2["CONTRATO"].unique().tolist())
        cont_prev = [c for c in datos_prev.get("filtro_contrato", []) if c in cont_disp]
        cont_sel = st.multiselect("Contrato / Componente:", cont_disp, default=cont_prev)
        df_f3 = df_f2[df_f2["CONTRATO"].isin(cont_sel)] if cont_sel else df_f2

    df_f3 = df_f3.copy()
    df_f3["ETIQUETA_CARPETA"] = df_f3["CONTRATO"] + "  ➤  " + df_f3["CARPETA"]
    carpetas_disp = df_f3["ETIQUETA_CARPETA"].tolist()
    carpetas_prev = [
        f"{c['contrato']}  ➤  {c['carpeta']}"
        for c in datos_prev.get("carpetas_seleccionadas", [])
        if f"{c['contrato']}  ➤  {c['carpeta']}" in carpetas_disp
    ]
    carpetas_sel = st.multiselect(
        "Carpetas configuradas a incluir en este reporte:",
        carpetas_disp,
        default=carpetas_prev if carpetas_prev else carpetas_disp
    )

    df_final_carp = df_f3[df_f3["ETIQUETA_CARPETA"].isin(carpetas_sel)]

    campos_dinamicos = set(COLUMNAS_BASE)
    for lista_c in df_final_carp["CAMPOS_PERS"]:
        campos_dinamicos.update(lista_c)
    lista_campos_filtro = sorted(list(campos_dinamicos))

    st.markdown("---")

    # 3. Condiciones de Filtro dinámicas
    st.subheader("2. Condiciones de Filtro (Campos Disponibles)")
    st.caption(
        "Usa **'está en (lista)'** para varios valores separados por coma (ej: `SHP, SUPERVISION`). "
        "Usa **Conector** (`Y (AND)` / `O (OR)`) para combinar filas distintas."
    )

    filtros_prev = datos_prev.get("condiciones_filtro", [])
    df_filtros_init = pd.DataFrame(filtros_prev) if filtros_prev else pd.DataFrame(
        columns=["logica", "campo", "operador", "valor"]
    )
    if "logica" not in df_filtros_init.columns:
        df_filtros_init.insert(0, "logica", "Y (AND)")
    for col in ["campo", "operador", "valor"]:
        if col not in df_filtros_init.columns:
            df_filtros_init[col] = ""
    df_filtros_init["logica"] = df_filtros_init["logica"].replace("", "Y (AND)").fillna("Y (AND)")

    df_filtros_edit = st.data_editor(
        df_filtros_init[["logica", "campo", "operador", "valor"]],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "logica": st.column_config.SelectboxColumn(
                "Conector",
                options=["Y (AND)", "O (OR)"],
                default="Y (AND)",
                required=True
            ),
            "campo": st.column_config.SelectboxColumn("Campo a Filtrar", options=lista_campos_filtro, required=True),
            "operador": st.column_config.SelectboxColumn(
                "Condición",
                options=[
                    "igual a", "diferente de", "está en (lista)", "no está en (lista)",
                    "contiene", "no contiene", "mayor que", "menor que"
                ],
                required=True
            ),
            "valor": st.column_config.TextColumn("Valor(es)", required=True)
        },
        key=f"filtros_{rep_sel}"
    )

    st.markdown("---")

    # 4. Programación de Días y Horarios
    st.subheader("3. Programación de Envío (Hora Perú)")
    h1, h2 = st.columns(2)
    with h1:
        dias_sel = st.multiselect(
            "Días de la semana:",
            DIAS_SEMANA,
            default=datos_prev.get("dias_semana", ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"])
        )
    with h2:
        horarios_sel = st.multiselect(
            "Horarios de envío:",
            HORARIOS_PERU,
            default=datos_prev.get("horarios", ["06:00", "14:00"])
        )

    st.markdown("---")

    # 5. Correo y Columnas del Excel
    st.subheader("4. Destinatarios y Formato de Correo")
    e1, e2 = st.columns(2)
    with e1:
        para = st.text_input("Para (separar con comas):", value=", ".join(datos_prev.get("para", ["rcano@shp.pe"])))
        cc = st.text_input("CC (opcional, separar con comas):", value=", ".join(datos_prev.get("cc", [])))
    with e2:
        asunto = st.text_input(
            "Asunto del correo:",
            value=datos_prev.get("asunto", "Reporte Control Documentario - {nombre_reporte} ({fecha})")
        )
        cols_excel = st.multiselect(
            "Columnas a incluir en el Excel (vacío = todas):",
            lista_campos_filtro,
            default=[c for c in datos_prev.get("columnas_excel", []) if c in lista_campos_filtro]
        )

    cuerpo = st.text_area(
        "Cuerpo del correo (variables: {nombre_reporte}, {fecha}, {total_registros}):",
        value=datos_prev.get("cuerpo", "Estimados,\n\nSe adjunta el reporte {nombre_reporte} actualizado al {fecha} ({total_registros} registros).\n\nSaludos."),
        height=100
    )

    # 6. Guardar, Enviar Ahora o Eliminar
    b1, b2, b3 = st.columns([2, 1.5, 1])
    with b1:
        if st.button("💾 Guardar Reporte en GitHub", type="primary", use_container_width=True):
            para_validos, para_invalidos = _validar_correos(para)
            cc_validos, cc_invalidos = _validar_correos(cc)

            if not nombre_reporte.strip():
                st.warning("Ingresa un nombre para el reporte.")
            elif df_final_carp.empty:
                st.warning("Selecciona al menos una carpeta para escanear.")
            elif not dias_sel or not horarios_sel:
                st.warning("Selecciona al menos un día de la semana y un horario.")
            elif not para_validos:
                st.error("❌ Debes ingresar al menos un correo válido en el campo 'Para'.")
            elif para_invalidos or cc_invalidos:
                errores = para_invalidos + cc_invalidos
                st.error(f"❌ Corrige estos correos inválidos antes de guardar: {', '.join(errores)}")
            else:
                carpetas_guardar = [
                    {
                        "alcance": r["ALCANCE"],
                        "subproyecto": r["SUBPROYECTO"],
                        "tipo_obra": r["TIPO_OBRA"],
                        "contrato": r["CONTRATO"],
                        "carpeta": r["CARPETA"],
                        "url": r["URL"],
                        "longitud_maxima": int(r["LONGITUD_MAX"]),
                        "config_sub": r["CONFIG_SUB"],
                        "urls_excluidas": r["URLS_EXCLUIDAS"],
                        "ruta_guardado_local": r["RUTA_LOCAL"]
                    }
                    for _, r in df_final_carp.iterrows()
                ]

                condiciones_guardar = []
                for _, f in df_filtros_edit.iterrows():
                    logica_raw = f.get("logica")
                    logica = "Y (AND)" if pd.isna(logica_raw) or not str(logica_raw).strip() else str(logica_raw).strip()
                    campo = str(f.get("campo", "")).strip()
                    op = str(f.get("operador", "")).strip()
                    val = str(f.get("valor", "")).strip()
                    if campo and op and val and campo.lower() != "nan":
                        condiciones_guardar.append({
                            "logica": logica,
                            "campo": campo,
                            "operador": op,
                            "valor": val
                        })

                config_actual[nombre_reporte.strip()] = {
                    "activo": activo,
                    "modo_envio": modo_envio,
                    "filtro_alcance": alc_sel,
                    "filtro_subproyecto": sub_sel,
                    "filtro_contrato": cont_sel,
                    "carpetas_seleccionadas": carpetas_guardar,
                    "condiciones_filtro": condiciones_guardar,
                    "dias_semana": dias_sel,
                    "horarios": horarios_sel,
                    "para": para_validos,
                    "cc": cc_validos,
                    "asunto": asunto.strip(),
                    "cuerpo": cuerpo,
                    "columnas_excel": cols_excel
                }

                res = guardar_config_github(config_actual, sha_actual)
                if res.status_code in [200, 201]:
                    st.success(f"✅ Reporte '{nombre_reporte.strip()}' guardado en GitHub.")
                    st.rerun()
                else:
                    st.error(f"Error al guardar en GitHub: {res.text}")

    with b2:
        es_nube = (modo_envio == MODOS_ENVIO[0])
        if st.button("📤 Enviar ahora", use_container_width=True, disabled=not es_nube):
            nombre_limpio = nombre_reporte.strip()
            if not nombre_limpio or nombre_limpio not in config_actual:
                st.warning("Primero haz clic en '💾 Guardar Reporte en GitHub' para poder enviarlo.")
            else:
                res_envio = disparar_envio_github_actions(reporte_objetivo=nombre_limpio)
                if res_envio.status_code == 204:
                    st.success(f"✅ Envío inmediato disparado para '{nombre_limpio}' vía Gmail. Llegará en 1 a 2 minutos.")
                else:
                    st.error(f"No se pudo disparar el envío ({res_envio.status_code}): {res_envio.text}")

    with b3:
        if rep_sel != "➕ Nuevo Reporte" and rep_sel in config_actual:
            if st.button("🗑️ Eliminar", use_container_width=True):
                del config_actual[rep_sel]
                res = guardar_config_github(config_actual, sha_actual)
                if res.status_code in [200, 201]:
                    st.success(f"Reporte '{rep_sel}' eliminado.")
                    st.rerun()
                else:
                    st.error(f"Error al eliminar: {res.text}")
