import os
import re
import json
import time
import urllib.parse
from io import BytesIO
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import requests
import pandas as pd
import numpy as np
import plotly.express as px

import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, PatternFill, Alignment

from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.platypus import Image as RLImage

from modulo_estructura_proyectos import (
    ESTRUCTURA_PROYECTOS,
    _SHA_INICIAL,
    obtener_subproyectos,
    obtener_tipos_obra,
    obtener_contratos,
    obtener_info_contrato,
    cargar_estructura_github,
    mostrar_editor_estructura,
)
from modulo_formulario_reportes import mostrar_formulario_reportes

try:
    from modulo_bd import bd_disponible, guardar_documentos_bd, cargar_documentos_bd
except Exception:
    bd_disponible = lambda: False
    guardar_documentos_bd = lambda *args, **kwargs: 0
    cargar_documentos_bd = lambda *args, **kwargs: pd.DataFrame()

st.set_page_config(
    page_title="Control Documentario SHP - Autodesk ACC",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

RUTA_LOGO = "logoshp.jpeg"
ORDEN_COLUMNAS_DASHBOARD = [
    "CODIGO",
    "REV",
    "DESCRIPCION",
    "FECHA",
    "ESTADO",
    "PENDIENTE POR",
    "DIAS",
]


def obtener_hora_peru() -> datetime:
    return datetime.now(timezone(timedelta(hours=-5)))


def obtener_periodo_actual() -> str:
    ahora_pe = obtener_hora_peru()
    fecha_corte = (ahora_pe - timedelta(hours=4)).strftime("%Y-%m-%d")
    return f"{fecha_corte}_04AM"


def _leer_secreto_o_env(clave: str, defecto: str = "") -> str:
    try:
        if clave in st.secrets:
            val = str(st.secrets[clave]).strip()
            if val:
                return val
    except Exception:
        pass
    return str(os.environ.get(clave, defecto)).strip()


CLIENT_ID = _leer_secreto_o_env("CLIENT_ID")
CLIENT_SECRET = _leer_secreto_o_env("CLIENT_SECRET")
ADMIN_USER = _leer_secreto_o_env("ADMIN_USER", "rcano")
ADMIN_PASSWORD = _leer_secreto_o_env("ADMIN_PASSWORD", "shp2026")


def listar_carpetas_de_estructura(estructura: dict) -> list[dict]:
    items = []
    if not isinstance(estructura, dict):
        return items

    for alcance, subs in estructura.items():
        if not isinstance(subs, dict):
            continue
        for sub_nom, info_sub in subs.items():
            if sub_nom == "config" or not isinstance(info_sub, dict):
                continue
            cfg_sub = info_sub.get("config", {}) if isinstance(info_sub.get("config"), dict) else {}
            obras_dict = (
                info_sub.get("tipos_obra")
                if isinstance(info_sub.get("tipos_obra"), dict)
                else {k: v for k, v in info_sub.items() if k != "config" and isinstance(v, dict)}
            )
            for tipo_obra, contratos_dict in obras_dict.items():
                if not isinstance(contratos_dict, dict):
                    continue
                for cont_nom, info_cont in contratos_dict.items():
                    if not isinstance(info_cont, dict):
                        continue
                    carpetas_raw = info_cont.get("carpetas") or info_cont.get("carpetas_config") or []
                    for c in carpetas_raw:
                        if not isinstance(c, dict):
                            continue
                        carp_nom = str(c.get("carpeta") or c.get("nombre") or "Carpeta").strip()
                        url_c = str(c.get("url", "")).strip()
                        len_max = int(c.get("longitud_maxima", 10) or 10)
                        items.append({
                            "etiqueta": f"{sub_nom} | {cont_nom} | {carp_nom}",
                            "alcance": alcance,
                            "subproyecto": sub_nom,
                            "tipo_obra": tipo_obra,
                            "contrato": cont_nom,
                            "carpeta": carp_nom,
                            "url": url_c,
                            "longitud_maxima": len_max,
                            "urls_excluidas": info_cont.get("urls_excluidas", []),
                            "config_sub": {
                                "campos_personalizados": info_cont.get("campos_personalizados") or cfg_sub.get("campos_personalizados", {}),
                                "id_campo_fecha": info_cont.get("id_campo_fecha") or cfg_sub.get("id_campo_fecha", "6011839"),
                                "id_campo_pendiente": info_cont.get("id_campo_pendiente") or cfg_sub.get("id_campo_pendiente", "6011840"),
                            },
                        })
    return items


@st.cache_data(show_spinner=False, ttl=2800, max_entries=4)
def obtener_token_acc(client_id: str, client_secret: str) -> str:
    if not client_id or not client_secret:
        return ""
    url = "https://developer.api.autodesk.com/authentication/v2/token"
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json",
    }
    resp = requests.post(
        url,
        headers=headers,
        data={"grant_type": "client_credentials", "scope": "data:read"},
        auth=(client_id, client_secret),
        timeout=25,
    )
    if resp.status_code == 200:
        return str(resp.json().get("access_token", "")).strip()
    return ""


def extraer_project_y_folder_id(url_carpeta: str):
    if not url_carpeta:
        return "", ""
    match_proj = re.search(r"/projects/([^/?#]+)", url_carpeta)
    project_id = match_proj.group(1) if match_proj else ""
    if project_id and not project_id.startswith("b."):
        project_id = f"b.{project_id}"

    if "folderUrn=" in url_carpeta:
        fid = url_carpeta.split("folderUrn=")[1].split("&")[0]
    elif "folders/" in url_carpeta:
        fid = url_carpeta.split("folders/")[1].split("/")[0].split("?")[0]
    else:
        fid = url_carpeta

    fid = urllib.parse.unquote(fid).replace("%3A", ":")
    return project_id, fid


def deconstruir_diccionario(datos, prefijo: str, resultado: dict):
    if isinstance(datos, dict):
        for llave, valor in datos.items():
            nueva_llave = (
                f"{prefijo}_{llave.strip().upper()}" if prefijo else llave.strip().upper()
            )
            if any(x in nueva_llave for x in ["LINKS", "SCHEMA", "HREF", "DERIVATIVES"]):
                continue
            deconstruir_diccionario(valor, nueva_llave, resultado)
    elif isinstance(datos, list):
        for i, elemento in enumerate(datos):
            if isinstance(elemento, dict):
                name = (
                    elemento.get("name")
                    or elemento.get("displayName")
                    or elemento.get("title")
                )
                val = elemento.get("value") or elemento.get("displayValue")
                if name and val is not None:
                    resultado[f"CAMPO_CUSTOM_{str(name).strip().upper()}"] = val
                else:
                    deconstruir_diccionario(elemento, f"{prefijo}_{i}", resultado)
            else:
                resultado[prefijo] = datos
    else:
        resultado[prefijo] = datos


def obtener_tip_version_metadata(token: str, project_id: str, item_id: str) -> dict:
    if not token or not project_id or not item_id:
        return {}
    clean_item_id = item_id.split("?version=")[0]
    url = (
        f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/"
        f"items/{urllib.parse.quote(clean_item_id, safe='')}/tip"
    )
    for region in ("US", "EMEA"):
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "x-ads-region": region,
        }
        for _ in range(3):
            r = requests.get(url, headers=headers, timeout=25)
            if r.status_code == 429:
                time.sleep(2)
                continue
            if r.status_code == 200:
                return r.json().get("data", {}) or {}
            break
    return {}


def procesar_archivo_pdf_acc(
    token: str,
    item: dict,
    project_id: str,
    ruta_actual: str,
    nombre_archivo: str,
    campos_personalizados: dict,
) -> dict:
    item_id = str(item.get("id", "")).strip()
    url_webview = (
        item.get("links", {}).get("webView", {}).get("href", "")
        if isinstance(item.get("links"), dict)
        else ""
    )

    registro = {
        "NOMBRE_DEL_ARCHIVO": nombre_archivo,
        "RUTA_EN_AUTODESK": ruta_actual,
        "ITEM_ID": item_id,
        "URL_DOC": str(url_webview or "").strip(),
    }
    deconstruir_diccionario(item, "ITEM", registro)

    tip_version = obtener_tip_version_metadata(token, project_id, item_id)
    if tip_version:
        if not registro["URL_DOC"] and isinstance(tip_version.get("links"), dict):
            registro["URL_DOC"] = str(
                tip_version.get("links", {}).get("webView", {}).get("href", "")
            ).strip()
        deconstruir_diccionario(tip_version, "VERSION", registro)

    vid = (
        item.get("relationships", {})
        .get("tip", {})
        .get("data", {})
        .get("id", "")
    )
    if vid:
        vid_encoded = urllib.parse.quote(vid, safe="")
        url_attr = (
            f"https://developer.api.autodesk.com/bim360/docs/v1/projects/"
            f"{project_id.replace('b.', '')}/versions/{vid_encoded}/custom-attributes"
        )
        for intento in range(4):
            res_attr = requests.get(
                url_attr,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                },
                timeout=25,
            )
            if res_attr.status_code == 429:
                time.sleep(2 * (intento + 1))
                continue
            if res_attr.status_code == 200:
                campos_objetivo = set(str(k) for k in campos_personalizados.keys())
                for attr in res_attr.json():
                    attr_id = str(attr.get("id"))
                    if attr_id in campos_objetivo:
                        registro[f"CAMPO_CUSTOM_{attr_id}"] = str(
                            attr.get("value", "")
                        ).strip()
            break

    return registro


def escanear_recursivo_acc(
    token: str,
    project_id: str,
    folder_id: str,
    lista_datos: list,
    longitud_maxima: int,
    ids_excluidos: set,
    campos_personalizados: dict,
    ruta_actual: str = "Raíz",
):
    if not token or not project_id or not folder_id or folder_id in ids_excluidos:
        return

    encoded_fid = urllib.parse.quote(folder_id, safe="")
    url_api = (
        f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/"
        f"folders/{encoded_fid}/contents"
    )

    while url_api:
        res = None
        for region in ("US", "EMEA"):
            for _ in range(3):
                headers = {"Authorization": f"Bearer {token}", "x-ads-region": region}
                r = requests.get(url_api, headers=headers, timeout=30)
                if r.status_code == 429:
                    time.sleep(4)
                    continue
                res = r
                break
            if res is not None and res.status_code == 200:
                break

        if res is None or res.status_code != 200:
            break

        data = res.json()
        items_pdf = []
        for item in data.get("data", []):
            tipo = item.get("type")
            name = str(item.get("attributes", {}).get("displayName", "Sin Nombre")).strip()
            if tipo == "folders":
                sub_id = str(item.get("id", "")).replace("%3A", ":")
                if sub_id in ids_excluidos:
                    continue
                if len(name) <= longitud_maxima:
                    escanear_recursivo_acc(
                        token,
                        project_id,
                        sub_id,
                        lista_datos,
                        longitud_maxima,
                        ids_excluidos,
                        campos_personalizados,
                        f"{ruta_actual} > {name}",
                    )
            elif tipo == "items" and ".pdf" in name.lower():
                items_pdf.append((item, name))

        if items_pdf:
            with ThreadPoolExecutor(max_workers=3) as executor:
                futures = [
                    executor.submit(
                        procesar_archivo_pdf_acc,
                        token,
                        it,
                        project_id,
                        ruta_actual,
                        nom,
                        campos_personalizados,
                    )
                    for it, nom in items_pdf
                ]
                for fut in as_completed(futures):
                    try:
                        res_item = fut.result()
                        if res_item:
                            lista_datos.append(res_item)
                    except Exception:
                        pass

        enlaces = data.get("links", {})
        url_api = enlaces["next"].get("href") if "next" in enlaces else None


def _limpiar_dias_entero(df: pd.DataFrame, col_pendiente: str) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = df.copy()
    if "DIAS" in df.columns:
        df["DIAS"] = pd.to_numeric(df["DIAS"], errors="coerce").fillna(0).round(0).astype(int)
        if col_pendiente and col_pendiente in df.columns:
            df.loc[
                df[col_pendiente].astype(str).str.strip().str.upper() == "CERRADO",
                "DIAS",
            ] = 0
    return df


def _normalizar_df_carpeta(
    datos_carpeta: list,
    campos_personalizados: dict,
    id_campo_fecha: str,
    id_campo_pendiente: str,
    col_pendiente: str,
    orden_columnas_deseado: list,
    alcance: str,
    subproyecto: str,
    tipo_obra: str,
    contrato: str,
    carpeta: str,
):
    if not datos_carpeta:
        cols_vacias = [c for c in orden_columnas_deseado if c] or list(ORDEN_COLUMNAS_DASHBOARD)
        vacio = pd.DataFrame(columns=cols_vacias)
        return vacio, vacio

    fecha_hoy_ts = pd.Timestamp(obtener_hora_peru().date())
    df_completo = pd.DataFrame(datos_carpeta)

    if "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION" not in df_completo.columns:
        df_completo["ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION"] = ""
    if "URL_DOC" not in df_completo.columns:
        df_completo["URL_DOC"] = ""

    for campo_id in (campos_personalizados or {}).keys():
        col_c = f"CAMPO_CUSTOM_{campo_id}"
        if col_c not in df_completo.columns:
            df_completo[col_c] = pd.NaT if str(campo_id) == str(id_campo_fecha) else ""

    col_fecha_raw = f"CAMPO_CUSTOM_{id_campo_fecha}"
    if col_fecha_raw in df_completo.columns:
        fechas_conv = pd.to_datetime(df_completo[col_fecha_raw], errors="coerce")
        if fechas_conv.dt.tz is not None:
            fechas_conv = fechas_conv.dt.tz_localize(None)
        df_completo["DIAS"] = (fecha_hoy_ts - fechas_conv).dt.days
        df_completo[col_fecha_raw] = fechas_conv.dt.strftime("%d/%m/%Y").fillna("")
    else:
        df_completo["DIAS"] = 0

    renombres = {
        "NOMBRE_DEL_ARCHIVO": "CODIGO",
        "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION": "DESCRIPCION",
    }
    for cid, nombre_col in (campos_personalizados or {}).items():
        renombres[f"CAMPO_CUSTOM_{cid}"] = nombre_col

    df_todo = df_completo.rename(columns=renombres)
    df_todo = _limpiar_dias_entero(df_todo, col_pendiente)

    df_todo["ALCANCE"] = alcance
    df_todo["SUBPROYECTO"] = subproyecto
    df_todo["TIPO_OBRA"] = tipo_obra
    df_todo["CONTRATO"] = contrato
    df_todo["CARPETA"] = carpeta

    cols_vista = [c for c in (orden_columnas_deseado or ORDEN_COLUMNAS_DASHBOARD) if c in df_todo.columns]
    if not cols_vista:
        cols_vista = [c for c in ORDEN_COLUMNAS_DASHBOARD if c in df_todo.columns]

    df_vista = df_todo[cols_vista].copy()
    return df_vista, df_todo


def _procesar_carpeta_config(
    config_carpeta: dict,
    ids_excluidos: set,
    campos_personalizados: dict,
    id_campo_fecha: str,
    id_campo_pendiente: str,
    col_pendiente: str,
    orden_columnas_deseado: list,
    periodo_corte: str = "",
    forzar_api: bool = False,
):
    nombre_pestana = str(
        config_carpeta.get("nombre_pestana")
        or config_carpeta.get("carpeta")
        or config_carpeta.get("nombre")
        or "Carpeta"
    ).strip()
    alcance = str(config_carpeta.get("_alcance", "")).strip()
    subproyecto = str(config_carpeta.get("_subproyecto", "")).strip()
    tipo_obra = str(config_carpeta.get("_tipo_obra", "")).strip()
    contrato = str(config_carpeta.get("_contrato", "")).strip()
    corte = periodo_corte or obtener_periodo_actual()
    diagnostico = ""

    if not forzar_api and bd_disponible() and contrato:
        try:
            df_bd = cargar_documentos_bd(
                periodo_corte=corte,
                contrato=contrato,
                carpeta=nombre_pestana,
                incluir_extra=True,
            )
            if df_bd is not None and not df_bd.empty:
                df_todo_bd = _limpiar_dias_entero(df_bd, col_pendiente)
                cols_vista = [
                    c for c in (orden_columnas_deseado or ORDEN_COLUMNAS_DASHBOARD) if c in df_todo_bd.columns
                ]
                if not cols_vista:
                    cols_vista = [c for c in ORDEN_COLUMNAS_DASHBOARD if c in df_todo_bd.columns]
                df_vista_bd = df_todo_bd[cols_vista].copy()
                return nombre_pestana, df_vista_bd, df_todo_bd, f"🟢 Cargado desde Supabase PostgreSQL ({len(df_todo_bd)} docs)"
        except Exception as e_bd:
            diagnostico = f"⚠️ Error leyendo BD: {e_bd}"

    url_carp = str(config_carpeta.get("url", "")).strip()
    project_id, fid = extraer_project_y_folder_id(url_carp)
    token = obtener_token_acc(CLIENT_ID, CLIENT_SECRET)

    if not token:
        diagnostico = "⚠️ Falta configurar CLIENT_ID o CLIENT_SECRET válidos para consultar Autodesk ACC."
    elif not project_id or not fid:
        diagnostico = f"⚠️ URL de carpeta inválida o vacía en '{nombre_pestana}'."

    datos_carpeta = []
    if token and project_id and fid:
        escanear_recursivo_acc(
            token=token,
            project_id=project_id,
            folder_id=fid,
            lista_datos=datos_carpeta,
            longitud_maxima=int(config_carpeta.get("longitud_maxima", 10)),
            ids_excluidos=ids_excluidos,
            campos_personalizados=campos_personalizados or {},
            ruta_actual=nombre_pestana,
        )
        if not datos_carpeta:
            diagnostico = (
                f"ℹ️ No se encontraron archivos PDF en '{nombre_pestana}' (revisa URL, "
                f"longitud máxima de subcarpeta o URLs excluidas)."
            )
        else:
            diagnostico = f"☁️ Consultado en vivo desde Autodesk ACC ({len(datos_carpeta)} PDFs)"

    df_vista, df_todo = _normalizar_df_carpeta(
        datos_carpeta=datos_carpeta,
        campos_personalizados=campos_personalizados or {},
        id_campo_fecha=id_campo_fecha,
        id_campo_pendiente=id_campo_pendiente,
        col_pendiente=col_pendiente,
        orden_columnas_deseado=orden_columnas_deseado,
        alcance=alcance,
        subproyecto=subproyecto,
        tipo_obra=tipo_obra,
        contrato=contrato,
        carpeta=nombre_pestana,
    )

    if bd_disponible() and not df_todo.empty:
        try:
            guardar_documentos_bd(
                df_todo=df_todo,
                periodo_corte=corte,
                alcance=alcance,
                subproyecto=subproyecto,
                tipo_obra=tipo_obra,
                contrato=contrato,
                carpeta=nombre_pestana,
            )
        except Exception:
            pass

    return nombre_pestana, df_vista, df_todo, diagnostico


def _congelar_estructura(obj):
    if isinstance(obj, dict):
        return tuple(sorted((k, _congelar_estructura(v)) for k, v in obj.items()))
    if isinstance(obj, list):
        return tuple(_congelar_estructura(x) for x in obj)
    return obj


def _descongelar_estructura(obj):
    if isinstance(obj, tuple):
        if all(isinstance(x, tuple) and len(x) == 2 and isinstance(x[0], str) for x in obj):
            return {k: _descongelar_estructura(v) for k, v in obj}
        return [_descongelar_estructura(x) for x in obj]
    return obj


@st.cache_data(show_spinner=False, ttl=14400, max_entries=60)
def extraer_datos_contrato_completo(
    contrato_key: str,
    info_contrato_tuple: tuple,
    periodo_corte: str,
    version_cache: int = 0,
):
    info_contrato = _descongelar_estructura(info_contrato_tuple)
    if not isinstance(info_contrato, dict):
        info_contrato = {}

    alcance_ctx = str(info_contrato.get("_alcance", "")).strip()
    subproyecto_ctx = str(info_contrato.get("_subproyecto", "")).strip()
    tipo_obra_ctx = str(info_contrato.get("_tipo_obra", "")).strip()
    contrato_ctx = str(info_contrato.get("_contrato") or contrato_key).strip()

    urls_excl = info_contrato.get("urls_excluidas", [])
    if not isinstance(urls_excl, list):
        urls_excl = []
    ids_excluidos = {
        extraer_project_y_folder_id(u)[1]
        for u in urls_excl
        if u
    }

    campos_personalizados = info_contrato.get("campos_personalizados") or {}
    if not isinstance(campos_personalizados, dict):
        campos_personalizados = {}

    id_campo_fecha = str(info_contrato.get("id_campo_fecha", ""))
    id_campo_pendiente = str(info_contrato.get("id_campo_pendiente", ""))
    col_pendiente = campos_personalizados.get(id_campo_pendiente, "PENDIENTE POR")

    orden_cols = info_contrato.get("columnas_visibles")
    if not isinstance(orden_cols, list) or not orden_cols:
        orden_cols = ORDEN_COLUMNAS_DASHBOARD

    carpetas_cfg = info_contrato.get("carpetas") or info_contrato.get("carpetas_config") or []
    if not isinstance(carpetas_cfg, list):
        carpetas_cfg = []

    resultados_vista = {}
    resultados_todo = {}
    diagnosticos = {}

    for cfg_c in carpetas_cfg:
        if not isinstance(cfg_c, dict):
            continue
        cfg_c_enriquecida = dict(cfg_c)
        if "carpeta" not in cfg_c_enriquecida and "nombre" in cfg_c_enriquecida:
            cfg_c_enriquecida["carpeta"] = cfg_c_enriquecida["nombre"]
        cfg_c_enriquecida["_alcance"] = alcance_ctx
        cfg_c_enriquecida["_subproyecto"] = subproyecto_ctx
        cfg_c_enriquecida["_tipo_obra"] = tipo_obra_ctx
        cfg_c_enriquecida["_contrato"] = contrato_ctx

        pest, df_p, df_t, diag = _procesar_carpeta_config(
            config_carpeta=cfg_c_enriquecida,
            ids_excluidos=ids_excluidos,
            campos_personalizados=campos_personalizados,
            id_campo_fecha=id_campo_fecha,
            id_campo_pendiente=id_campo_pendiente,
            col_pendiente=col_pendiente,
            orden_columnas_deseado=orden_cols,
            periodo_corte=periodo_corte,
            forzar_api=(version_cache > 0),
        )
        resultados_vista[pest] = df_p
        resultados_todo[pest] = df_t
        diagnosticos[pest] = diag

    return {
        "datos": resultados_vista,
        "datos_completos": resultados_todo,
        "col_pendiente": col_pendiente,
        "diagnosticos": diagnosticos,
    }


def ejecutar_actualizacion_contrato_actual(info_contrato_ctx: dict, periodo_corte: str):
    carpetas = info_contrato_ctx.get("carpetas") or info_contrato_ctx.get("carpetas_config") or []
    if not carpetas:
        st.warning("Este contrato no tiene carpetas configuradas.")
        return

    total = len(carpetas)
    contrato_nom = info_contrato_ctx.get("_contrato", "")
    barra = st.progress(0, text=f"Sincronizando {contrato_nom}...")
    estado_txt = st.empty()

    campos_pers = info_contrato_ctx.get("campos_personalizados") or {}
    id_fecha = str(info_contrato_ctx.get("id_campo_fecha", "6011839"))
    id_pend = str(info_contrato_ctx.get("id_campo_pendiente", "6011840"))
    col_pen = campos_pers.get(id_pend, "PENDIENTE POR")
    ids_excl = {
        extraer_project_y_folder_id(u)[1]
        for u in info_contrato_ctx.get("urls_excluidas", [])
        if u
    }

    for idx, c in enumerate(carpetas, start=1):
        nom_c = str(c.get("carpeta") or c.get("nombre") or "Carpeta").strip()
        estado_txt.info(f"🔄 ({idx}/{total}) Escaneando **{contrato_nom}** → `{nom_c}`...")
        cfg_c = dict(c)
        cfg_c["carpeta"] = nom_c
        cfg_c["nombre_pestana"] = nom_c
        cfg_c["_alcance"] = info_contrato_ctx.get("_alcance", "")
        cfg_c["_subproyecto"] = info_contrato_ctx.get("_subproyecto", "")
        cfg_c["_tipo_obra"] = info_contrato_ctx.get("_tipo_obra", "")
        cfg_c["_contrato"] = contrato_nom

        _, _, df_t, _ = _procesar_carpeta_config(
            config_carpeta=cfg_c,
            ids_excluidos=ids_excl,
            campos_personalizados=campos_pers,
            id_campo_fecha=id_fecha,
            id_campo_pendiente=id_pend,
            col_pendiente=col_pen,
            orden_columnas_deseado=ORDEN_COLUMNAS_DASHBOARD,
            periodo_corte=periodo_corte,
            forzar_api=True,
        )
        barra.progress(idx / total, text=f"{idx}/{total} carpetas actualizadas ({len(df_t)} docs en {nom_c})")

    st.cache_data.clear()
    st.session_state.ultima_actualizacion = obtener_hora_peru().strftime("%d/%m/%Y %H:%M")
    estado_txt.success(f"✅ Contrato **{contrato_nom}** actualizado en Supabase.")
    time.sleep(1.0)
    st.rerun()


def construir_resumen_carpeta(df_full: pd.DataFrame, col_pendiente: str) -> dict:
    total = len(df_full) if df_full is not None else 0
    if total == 0:
        return {
            "Cerrados": 0,
            "Pendientes": 0,
            "Total": 0,
            "Avance (%)": 0.0,
            "Mayor Retraso": "Sin datos",
        }

    cerrados = 0
    mayor_retraso = "Sin pendientes"
    if col_pendiente and col_pendiente in df_full.columns:
        est_u = df_full[col_pendiente].astype(str).str.strip().str.upper()
        cerrados = int((est_u == "CERRADO").sum())

        df_abiertos = df_full[est_u != "CERRADO"].copy()
        if not df_abiertos.empty and "DIAS" in df_abiertos.columns:
            df_abiertos["DIAS_NUM"] = pd.to_numeric(df_abiertos["DIAS"], errors="coerce").fillna(0)
            if df_abiertos["DIAS_NUM"].max() > 0:
                idx_max = df_abiertos["DIAS_NUM"].idxmax()
                cod = str(df_abiertos.loc[idx_max, "CODIGO"]) if "CODIGO" in df_abiertos.columns else "Doc"
                d_max = int(df_abiertos.loc[idx_max, "DIAS_NUM"])
                resp = str(df_abiertos.loc[idx_max, col_pendiente]).strip()
                mayor_retraso = f"{cod} ({d_max} d) [{resp}]"
            else:
                mayor_retraso = "Sin retrasos"

    pendientes = total - cerrados
    avance = round((cerrados / total * 100), 1) if total > 0 else 0.0
    return {
        "Cerrados": cerrados,
        "Pendientes": pendientes,
        "Total": total,
        "Avance (%)": avance,
        "Mayor Retraso": mayor_retraso,
    }


def construir_tabla_resumen_por_estado(df_full: pd.DataFrame, incluir_total: bool = True) -> pd.DataFrame:
    if df_full is None or df_full.empty or "ESTADO" not in df_full.columns:
        return pd.DataFrame(columns=["Estado", "Cantidad", "Porcentaje (%)"])
    s = df_full["ESTADO"].fillna("").astype(str).str.strip()
    s = s.replace({"": "Sin Estado", "nan": "Sin Estado", "None": "Sin Estado"})
    vc = s.value_counts()
    total = int(len(df_full))
    filas = [
        {
            "Estado": k,
            "Cantidad": int(v),
            "Porcentaje (%)": round(int(v) / max(1, total) * 100, 1),
        }
        for k, v in vc.items()
    ]
    if incluir_total and filas:
        filas.append({
            "Estado": "TOTAL",
            "Cantidad": total,
            "Porcentaje (%)": 100.0,
        })
    return pd.DataFrame(filas)


def obtener_colores_estado(estado):
    estado = str(estado).strip().upper()
    if estado == "CERRADO":
        return "#d4edda", "#155724"
    elif "CONTRATISTA" in estado or "EMPRESA" in estado:
        return "#fff3cd", "#856404"
    elif estado == "SUPERVISION":
        return "#f8d7da", "#721c24"
    else:
        return "#cce5ff", "#004085"


def color_barra_responsable(valor: str) -> str:
    v = str(valor).strip().upper()
    if v == "CERRADO":
        return "#28a745"
    if "CONTRATISTA" in v or "EMPRESA" in v:
        return "#ffc107"
    if "SUPERVISION" in v:
        return "#dc3545"
    if "SHP" in v:
        return "#007bff"
    return "#6c757d"


def renderizar_graficos_carpeta(df_t: pd.DataFrame, col_pendiente: str, nombre_carpeta: str):
    if df_t is None or df_t.empty:
        return

    g1, g2 = st.columns(2)
    with g1:
        if col_pendiente in df_t.columns:
            serie_p = df_t[col_pendiente].fillna("").astype(str).str.strip()
            serie_p = serie_p.replace({"": "Sin Asignar", "nan": "Sin Asignar", "None": "Sin Asignar"})
            df_p = serie_p.value_counts().reset_index()
            df_p.columns = ["Responsable", "Cantidad"]
            mapa_colores = {r: color_barra_responsable(r) for r in df_p["Responsable"]}

            fig_p = px.bar(
                df_p,
                x="Responsable",
                y="Cantidad",
                color="Responsable",
                color_discrete_map=mapa_colores,
                text="Cantidad",
                title=f"📊 Distribución por Responsable — {nombre_carpeta}",
            )
            fig_p.update_traces(textposition="outside")
            fig_p.update_layout(
                showlegend=False,
                height=340,
                margin=dict(l=10, r=10, t=45, b=10),
                xaxis_title="",
                yaxis_title="Documentos",
            )
            st.plotly_chart(fig_p, width="stretch", key=f"graf_pen_{nombre_carpeta}")

    with g2:
        df_est = construir_tabla_resumen_por_estado(df_t, incluir_total=False)
        if not df_est.empty:
            fig_e = px.bar(
                df_est,
                x="Estado",
                y="Cantidad",
                color="Estado",
                text="Cantidad",
                title=f"📊 Distribución por Estado — {nombre_carpeta}",
            )
            fig_e.update_traces(textposition="outside")
            fig_e.update_layout(
                showlegend=False,
                height=340,
                margin=dict(l=10, r=10, t=45, b=10),
                xaxis_title="",
                yaxis_title="Documentos",
            )
            st.plotly_chart(fig_e, width="stretch", key=f"graf_est_{nombre_carpeta}")


def preparar_df_con_links(df_vista: pd.DataFrame, df_full: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    df_out = df_vista.copy()
    col_cfg = {}

    if "DIAS" in df_out.columns:
        df_out["DIAS"] = pd.to_numeric(df_out["DIAS"], errors="coerce").fillna(0).round(0).astype(int)
        col_cfg["DIAS"] = st.column_config.NumberColumn("DIAS", format="%d")

    if "CODIGO" in df_out.columns and df_full is not None and "URL_DOC" in df_full.columns:
        urls_serie = df_full["URL_DOC"].reindex(df_out.index).fillna("").astype(str).str.strip()
        tiene_links = urls_serie.str.startswith("http").any()
        if tiene_links:
            codigos_link = []
            for cod, url in zip(df_out["CODIGO"].astype(str), urls_serie):
                if url.startswith("http"):
                    codigos_link.append(f"{url}#doc={cod}")
                else:
                    codigos_link.append(cod)
            df_out["CODIGO"] = codigos_link
            col_cfg["CODIGO"] = st.column_config.LinkColumn(
                "CODIGO",
                display_text=r"#doc=(.+)$",
                help="Haz clic en el código para abrir el documento en Autodesk ACC",
            )

    return df_out, col_cfg


def estilizar_tabla_streamlit(df: pd.DataFrame, col_pendiente: str = None):
    if df is None or df.empty:
        return df
    if not col_pendiente or col_pendiente not in df.columns:
        return df.style

    def _aplicar_estilo_fila(row):
        bg, fg = obtener_colores_estado(row[col_pendiente])
        return [f"background-color: {bg}; color: {fg}; font-weight: 500;" for _ in row]

    return df.style.apply(_aplicar_estilo_fila, axis=1)


def exportar_df_a_excel(
    df: pd.DataFrame,
    titulo: str,
    metadata: dict,
    ruta_logo: str,
    col_pendiente: str = None,
) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte"

    if os.path.exists(ruta_logo):
        img = ExcelImage(ruta_logo)
        img.height = 40
        img.width = 120
        ws.add_image(img, "G1")

    ws["A1"] = titulo
    ws["A1"].font = Font(size=14, bold=True, color="1A365D")
    ws["A3"], ws["B3"] = "Fecha de Exportación:", metadata.get("fecha", "")
    ws["A4"], ws["B4"] = "Alcance:", metadata.get("alcance", "")
    ws["A5"], ws["B5"] = "Subproyecto:", metadata.get("subproyecto", "")
    ws["A6"], ws["B6"] = "Tipo de Obra:", metadata.get("tipo_obra", "")
    ws["A7"], ws["B7"] = "Contrato:", metadata.get("contrato", "")
    for r_i in range(3, 8):
        ws[f"A{r_i}"].font = Font(bold=True)

    fila_ini = 9
    if df is not None and not df.empty:
        for col_num, col_name in enumerate(df.columns, 1):
            c = ws.cell(row=fila_ini, column=col_num, value=str(col_name))
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
            c.alignment = Alignment(horizontal="center")
            col_u = str(col_name).upper()
            ancho = 42 if col_u == "DESCRIPCION" else (28 if col_u == "CODIGO" else 18)
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_num)].width = ancho

        idx_pen = (
            df.columns.tolist().index(col_pendiente)
            if (col_pendiente and col_pendiente in df.columns)
            else None
        )
        for r_idx, row_data in enumerate(df.values, fila_ini + 1):
            bg_hex, fg_hex = "FFFFFF", "000000"
            if idx_pen is not None:
                bg, fg = obtener_colores_estado(row_data[idx_pen])
                bg_hex, fg_hex = bg.replace("#", ""), fg.replace("#", "")
            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value="" if pd.isna(val) else val)
                cell.fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
                cell.font = Font(color=fg_hex)

        fila_pie = fila_ini + len(df) + 2
    else:
        fila_pie = fila_ini + 2

    pie = ws.cell(row=fila_pie, column=1, value="Desarrollado por Roberto Cano - Shougang Hierro Perú")
    pie.font = Font(italic=True, color="4A5568")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def exportar_df_a_pdf(
    df: pd.DataFrame,
    titulo: str,
    metadata: dict,
    ruta_logo: str,
    col_pendiente: str = None,
) -> bytes:
    def _pie_pagina(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(colors.HexColor("#4A5568"))
        canvas.drawString(doc.leftMargin, 15, "Desarrollado por Roberto Cano - Shougang Hierro Perú")
        canvas.restoreState()

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=25,
        rightMargin=25,
        topMargin=25,
        bottomMargin=35,
    )
    ancho_util = 742
    elements = []
    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle(name="CellPdf", parent=styles["Normal"], fontSize=7, leading=9)
    header_style = ParagraphStyle(
        name="HdrPdf",
        parent=styles["Normal"],
        fontSize=7,
        leading=9,
        textColor=colors.whitesmoke,
        fontName="Helvetica-Bold",
    )

    logo = (
        RLImage(ruta_logo, width=100, height=30, kind="proportional")
        if os.path.exists(ruta_logo)
        else Paragraph("", styles["Normal"])
    )
    meta_txt = (
        f"<b><font size=12>{titulo}</font></b><br/><br/>"
        f"<b>Fecha de Exportación:</b> {metadata.get('fecha', '')}<br/>"
        f"<b>Alcance:</b> {metadata.get('alcance', '')}<br/>"
        f"<b>Subproyecto:</b> {metadata.get('subproyecto', '')}<br/>"
        f"<b>Tipo de Obra:</b> {metadata.get('tipo_obra', '')}<br/>"
        f"<b>Contrato:</b> {metadata.get('contrato', '')}"
    )
    cab = Table([[logo, Paragraph(meta_txt, styles["Normal"])]], colWidths=[120, 620])
    cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    elements.append(cab)
    elements.append(Spacer(1, 12))

    if df is not None and not df.empty:
        enc = [Paragraph(str(c), header_style) for c in df.columns]
        mat = [enc]
        estilos = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
        ]
        for r_idx, (_, row) in enumerate(df.iterrows(), start=1):
            if col_pendiente and col_pendiente in df.columns:
                bg_c, fg_c = obtener_colores_estado(row[col_pendiente])
            else:
                bg_c = "#ffffff" if r_idx % 2 != 0 else "#f9fafb"
                fg_c = "#000000"
            estilos.append(("BACKGROUND", (0, r_idx), (-1, r_idx), colors.HexColor(bg_c)))
            fila_p = []
            for col_n in df.columns:
                val_s = (
                    ("" if pd.isna(row[col_n]) else str(row[col_n]))
                    .replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                )
                if col_pendiente and col_pendiente in df.columns:
                    val_s = f'<font color="{fg_c}">{val_s}</font>'
                fila_p.append(Paragraph(val_s, cell_style))
            mat.append(fila_p)

        anchos = [ancho_util / max(1, len(df.columns))] * len(df.columns)
        t = Table(mat, repeatRows=1, colWidths=anchos)
        t.setStyle(TableStyle(estilos))
        elements.append(t)
    else:
        elements.append(Paragraph("Sin registros para mostrar.", styles["Normal"]))

    doc.build(elements, onFirstPage=_pie_pagina, onLaterPages=_pie_pagina)
    buf.seek(0)
    return buf.getvalue()


def _render_modo_administrador():
    st.subheader("🔐 Modo Administrador — Estructura y Reportes Diarios")

    if "admin_autenticado" not in st.session_state:
        st.session_state["admin_autenticado"] = False

    if not st.session_state["admin_autenticado"]:
        st.info("Ingresa tu usuario y contraseña de administrador para gestionar la estructura de proyectos y los reportes automáticos.")
        with st.form("form_login_admin"):
            usr_ingresado = st.text_input("Usuario:", placeholder="Ej: rcano").strip()
            pwd_ingresada = st.text_input("Contraseña:", type="password")
            btn_entrar = st.form_submit_button("🔓 Ingresar", type="primary")
            if btn_entrar:
                if usr_ingresado.lower() == ADMIN_USER.lower() and pwd_ingresada == ADMIN_PASSWORD:
                    st.session_state["admin_autenticado"] = True
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos.")
        return

    col_adm_info, col_adm_logout = st.columns([4, 1])
    with col_adm_info:
        st.success(f"✅ Sesión de administrador activa ({ADMIN_USER}).")
    with col_adm_logout:
        if st.button("🔒 Cerrar Sesión", width="stretch"):
            st.session_state["admin_autenticado"] = False
            st.rerun()

    estructura_fresca, sha_fresco = cargar_estructura_github(ESTRUCTURA_PROYECTOS)

    tab_rep, tab_est = st.tabs([
        "✉️ Programar y Enviar Reportes Diarios",
        "🏗️ Estructura de Proyectos y Contratos",
    ])

    with tab_rep:
        mostrar_formulario_reportes(estructura_fresca)

    with tab_est:
        mostrar_editor_estructura(estructura_fresca, sha_fresco)


def _render_vista_carpeta(
    nombre_carpeta: str,
    df_c: pd.DataFrame,
    df_t: pd.DataFrame,
    col_pendiente: str,
    diag_carp: str,
    contrato_sel: str,
    metadata_base: dict,
):
    if diag_carp:
        st.caption(f"Fuente: {diag_carp}")

    res_k = construir_resumen_carpeta(df_t, col_pendiente)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📄 Total Documentos", res_k["Total"])
    k2.metric("✅ Cerrados", res_k["Cerrados"])
    k3.metric("⏳ Pendientes", res_k["Pendientes"])
    k4.metric("📈 Avance (%)", f"{res_k['Avance (%)']}%")

    if res_k["Mayor Retraso"] not in ("Sin datos", "Sin pendientes", "Sin retrasos"):
        st.warning(f"⏰ **Mayor retraso abierto:** {res_k['Mayor Retraso']}")

    renderizar_graficos_carpeta(df_t, col_pendiente, nombre_carpeta)

    with st.expander("📌 Tabla Resumen por Estado del Documento", expanded=True):
        st.dataframe(
            construir_tabla_resumen_por_estado(df_t, incluir_total=True),
            width="stretch",
            hide_index=True,
        )

    st.markdown("#### 🔍 Filtros y Detalle Documentario")
    col_f1, col_f2, col_f3 = st.columns(3)

    estados_disp = ["Todos"]
    if "ESTADO" in df_c.columns:
        serie_est_norm = (
            df_c["ESTADO"]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace({"": "Sin Estado", "nan": "Sin Estado", "None": "Sin Estado"})
        )
        estados_disp += sorted(serie_est_norm.unique().tolist())
    f_estado = col_f1.selectbox("Filtrar por Estado", estados_disp, key=f"est_{nombre_carpeta}")

    pend_disp = ["Todos"]
    if col_pendiente in df_c.columns:
        serie_pen_norm = (
            df_c[col_pendiente]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace({"": "Sin Asignar", "nan": "Sin Asignar", "None": "Sin Asignar"})
        )
        pend_disp += sorted(serie_pen_norm.unique().tolist())
    f_pend = col_f2.selectbox("Filtrar por Pendiente Por", pend_disp, key=f"pen_{nombre_carpeta}")
    f_busq = col_f3.text_input("Buscar código o descripción", key=f"busq_{nombre_carpeta}").strip()

    df_filtrado = df_c.copy()
    if f_estado != "Todos" and "ESTADO" in df_filtrado.columns:
        est_f = (
            df_filtrado["ESTADO"]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace({"": "Sin Estado", "nan": "Sin Estado", "None": "Sin Estado"})
        )
        df_filtrado = df_filtrado[est_f == f_estado]
    if f_pend != "Todos" and col_pendiente in df_filtrado.columns:
        pen_f = (
            df_filtrado[col_pendiente]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace({"": "Sin Asignar", "nan": "Sin Asignar", "None": "Sin Asignar"})
        )
        df_filtrado = df_filtrado[pen_f == f_pend]
    if f_busq:
        mask_b = pd.Series(False, index=df_filtrado.index)
        for col_b in ["CODIGO", "DESCRIPCION"]:
            if col_b in df_filtrado.columns:
                mask_b = mask_b | df_filtrado[col_b].astype(str).str.contains(f_busq, case=False, na=False)
        df_filtrado = df_filtrado[mask_b]

    st.caption(f"Mostrando **{len(df_filtrado)}** de **{len(df_c)}** documentos en `{nombre_carpeta}`")
    if df_filtrado.empty:
        st.warning("No hay registros para mostrar en esta carpeta con el filtro o corte actual.")
    else:
        df_display, col_cfg = preparar_df_con_links(df_filtrado, df_t)
        st.dataframe(
            estilizar_tabla_streamlit(df_display, col_pendiente),
            width="stretch",
            hide_index=True,
            column_config=col_cfg,
        )

    c_b1, c_b2 = st.columns(2)
    excel_bytes = exportar_df_a_excel(
        df_filtrado,
        f"{contrato_sel} - {nombre_carpeta}",
        metadata_base,
        RUTA_LOGO,
        col_pendiente,
    )
    c_b1.download_button(
        label="📥 Descargar Excel",
        data=excel_bytes,
        file_name=f"{contrato_sel}_{nombre_carpeta}.xlsx".replace(" ", "_"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_xlsx_{nombre_carpeta}",
        width="stretch",
    )

    pdf_bytes = exportar_df_a_pdf(
        df_filtrado,
        f"{contrato_sel} - {nombre_carpeta}",
        metadata_base,
        RUTA_LOGO,
        col_pendiente,
    )
    c_b2.download_button(
        label="📄 Descargar PDF",
        data=pdf_bytes,
        file_name=f"{contrato_sel}_{nombre_carpeta}.pdf".replace(" ", "_"),
        mime="application/pdf",
        key=f"dl_pdf_{nombre_carpeta}",
        width="stretch",
    )


# --- Estado inicial de sesión ---
if "version_cache" not in st.session_state:
    st.session_state.version_cache = 0
if "ultima_actualizacion" not in st.session_state:
    st.session_state.ultima_actualizacion = obtener_hora_peru().strftime("%d/%m/%Y %H:%M")
if "ejecutar_sync_acc" not in st.session_state:
    st.session_state.ejecutar_sync_acc = False

periodo_actual = obtener_periodo_actual()

# --- Barra lateral de navegación ---
with st.sidebar:
    if os.path.exists(RUTA_LOGO):
        st.image(RUTA_LOGO, width=180)
    st.title("📊 Control Documentario")
    st.caption(f"🗓️ Corte activo: `{periodo_actual}`")
    st.caption(f"🕒 Última sincronización: {st.session_state.ultima_actualizacion}")

    estado_bd = "🟢 Supabase PostgreSQL activo" if bd_disponible() else "⚪ Caché en memoria (sin BD)"
    st.info(estado_bd)

    if st.button("🔄 Actualizar Datos desde ACC", width="stretch"):
        st.session_state.ejecutar_sync_acc = True
        st.rerun()

    st.divider()
    alcances = list(ESTRUCTURA_PROYECTOS.keys()) if isinstance(ESTRUCTURA_PROYECTOS, dict) else []
    alcance_sel = st.selectbox("📌 Alcance", alcances, index=0 if alcances else None)

    subproyectos = obtener_subproyectos(alcance_sel) if alcance_sel else []
    sub_sel = st.selectbox("🏗️ Subproyecto", subproyectos, index=0 if subproyectos else None)

    info_sub = ESTRUCTURA_PROYECTOS.get(alcance_sel, {}).get(sub_sel, {}) if (alcance_sel and sub_sel) else {}
    tipos_obra = obtener_tipos_obra(alcance_sel, sub_sel) if (alcance_sel and sub_sel) else []
    tipo_sel = st.selectbox("🔧 Tipo de Obra", tipos_obra, index=0 if tipos_obra else None)

    contratos = obtener_contratos(alcance_sel, sub_sel, tipo_sel) if (alcance_sel and sub_sel and tipo_sel) else []
    contrato_sel = st.selectbox("📄 Contrato", contratos, index=0 if contratos else None)

    st.divider()
    modo_vista = st.radio(
        "🧭 Vista del Dashboard",
        [
            "📁 Detalle de Contrato",
            "🗂️ Consolidado de Subproyecto",
            "🌐 Reporte Global",
            "🛠️ Constructor Dinámico",
            "🔐 Modo Administrador",
        ],
    )

# --- Cabecera principal ---
st.title("📊 Tablero de Control Documentario - Shougang Hierro Perú")

if modo_vista == "🔐 Modo Administrador":
    _render_modo_administrador()
    st.stop()

if not alcance_sel or not sub_sel or not tipo_sel or not contrato_sel:
    st.warning(
        "Selecciona un Alcance, Subproyecto, Tipo de Obra y Contrato en la barra lateral "
        "(o entra a **🔐 Modo Administrador** para configurar la estructura)."
    )
    st.stop()

info_contrato_raw = obtener_info_contrato(alcance_sel, sub_sel, tipo_sel, contrato_sel) or {}
info_contrato_ctx = dict(info_contrato_raw)
info_contrato_ctx["_alcance"] = alcance_sel
info_contrato_ctx["_subproyecto"] = sub_sel
info_contrato_ctx["_tipo_obra"] = tipo_sel
info_contrato_ctx["_contrato"] = contrato_sel

if st.session_state.ejecutar_sync_acc:
    st.session_state.ejecutar_sync_acc = False
    ejecutar_actualizacion_contrato_actual(info_contrato_ctx, periodo_actual)

with st.spinner(f"Cargando datos de {contrato_sel}..."):
    paquete_contrato = extraer_datos_contrato_completo(
        contrato_key=contrato_sel,
        info_contrato_tuple=_congelar_estructura(info_contrato_ctx),
        periodo_corte=periodo_actual,
        version_cache=st.session_state.version_cache,
    )

datos_vista = paquete_contrato.get("datos", {})
datos_completos = paquete_contrato.get("datos_completos", {})
col_pendiente = paquete_contrato.get("col_pendiente", "PENDIENTE POR")
diagnosticos = paquete_contrato.get("diagnosticos", {})

fecha_export_str = obtener_hora_peru().strftime("%d/%m/%Y %H:%M")
metadata_base = {
    "fecha": fecha_export_str,
    "alcance": alcance_sel,
    "subproyecto": sub_sel,
    "tipo_obra": tipo_sel,
    "contrato": contrato_sel,
}

# --- 1. Detalle de Contrato ---
if modo_vista == "📁 Detalle de Contrato":
    st.subheader(f"📁 Detalle de Contrato: {contrato_sel}")

    with st.expander("ℹ️ Información del Contrato", expanded=False):
        ic1, ic2, ic3 = st.columns(3)
        ic1.markdown(f"**Proyecto:** {info_contrato_raw.get('nombre_proyecto') or 'N/A'}")
        ic1.markdown(f"**Tipo de Contrato:** {info_contrato_raw.get('tipo_contrato') or 'N/A'}")
        ic2.markdown(f"**Supervisor:** {info_contrato_raw.get('nombre_supervisor') or 'N/A'}")
        ic2.markdown(f"**Coordinador SHP:** {info_contrato_raw.get('coordinador_shp') or 'N/A'}")
        ic3.markdown(f"**Admin. Contratos:** {info_contrato_raw.get('administrador_contratos') or 'N/A'}")

    if not datos_vista:
        st.info("No hay carpetas configuradas para este contrato en `estructura_proyectos.json`.")
    else:
        nombres_carpetas = list(datos_vista.keys())
        etiquetas_tabs = ["📂 Todas las Carpetas"] + [f"📂 {k}" for k in nombres_carpetas]
        pestanas = st.tabs(etiquetas_tabs)

        # Pestaña 0: Todas las Carpetas
        with pestanas[0]:
            dfs_v_all = []
            dfs_t_all = []
            for nom_c in nombres_carpetas:
                df_vi = datos_vista[nom_c].copy()
                df_ti = datos_completos.get(nom_c, df_vi).copy()
                if not df_vi.empty:
                    if "CARPETA" not in df_vi.columns:
                        df_vi.insert(0, "CARPETA", nom_c)
                    dfs_v_all.append(df_vi)
                if not df_ti.empty:
                    dfs_t_all.append(df_ti)

            df_v_todas = (
                pd.concat(dfs_v_all, ignore_index=True)
                if dfs_v_all
                else pd.DataFrame(columns=["CARPETA"] + ORDEN_COLUMNAS_DASHBOARD)
            )
            df_t_todas = pd.concat(dfs_t_all, ignore_index=True) if dfs_t_all else df_v_todas.copy()
            _render_vista_carpeta(
                nombre_carpeta="Todas las Carpetas",
                df_c=df_v_todas,
                df_t=df_t_todas,
                col_pendiente=col_pendiente,
                diag_carp=f"Consolidado de {len(nombres_carpetas)} carpetas ({len(df_t_todas)} documentos)",
                contrato_sel=contrato_sel,
                metadata_base=metadata_base,
            )

        # Pestañas individuales por carpeta
        for tab, nombre_carpeta in zip(pestanas[1:], nombres_carpetas):
            with tab:
                _render_vista_carpeta(
                    nombre_carpeta=nombre_carpeta,
                    df_c=datos_vista[nombre_carpeta],
                    df_t=datos_completos.get(nombre_carpeta, datos_vista[nombre_carpeta]),
                    col_pendiente=col_pendiente,
                    diag_carp=diagnosticos.get(nombre_carpeta, ""),
                    contrato_sel=contrato_sel,
                    metadata_base=metadata_base,
                )

# --- 2. Consolidado de Subproyecto ---
elif modo_vista == "🗂️ Consolidado de Subproyecto":
    st.subheader(f"🗂️ Consolidado de Subproyecto: {sub_sel}")
    filas_cons = []
    obras_sub = (
        info_sub.get("tipos_obra")
        if isinstance(info_sub.get("tipos_obra"), dict)
        else {k: v for k, v in info_sub.items() if k != "config" and isinstance(v, dict)}
    )
    for t_obra, c_dict in obras_sub.items():
        if not isinstance(c_dict, dict):
            continue
        for c_nom, c_info in c_dict.items():
            if not isinstance(c_info, dict):
                continue
            c_ctx = dict(c_info)
            c_ctx["_alcance"] = alcance_sel
            c_ctx["_subproyecto"] = sub_sel
            c_ctx["_tipo_obra"] = t_obra
            c_ctx["_contrato"] = c_nom
            res_c = extraer_datos_contrato_completo(
                contrato_key=c_nom,
                info_contrato_tuple=_congelar_estructura(c_ctx),
                periodo_corte=periodo_actual,
                version_cache=st.session_state.version_cache,
            )
            cp = res_c.get("col_pendiente", "PENDIENTE POR")
            for carp_nom, df_full in res_c.get("datos_completos", {}).items():
                rk = construir_resumen_carpeta(df_full, cp)
                filas_cons.append({
                    "Tipo de Obra": t_obra,
                    "Contrato": c_nom,
                    "Carpeta": carp_nom,
                    **rk,
                })

    df_cons = pd.DataFrame(filas_cons)
    if not df_cons.empty:
        tot_s = int(df_cons["Total"].sum())
        cer_s = int(df_cons["Cerrados"].sum())
        pen_s = int(df_cons["Pendientes"].sum())
        av_s = round(cer_s / tot_s * 100, 1) if tot_s > 0 else 0.0

        ks1, ks2, ks3, ks4 = st.columns(4)
        ks1.metric("📄 Total Subproyecto", tot_s)
        ks2.metric("✅ Cerrados", cer_s)
        ks3.metric("⏳ Pendientes", pen_s)
        ks4.metric("📈 Avance Global", f"{av_s}%")

        df_graf_s = df_cons.groupby("Contrato", as_index=False)[["Cerrados", "Pendientes"]].sum()
        fig_s = px.bar(
            df_graf_s,
            x="Contrato",
            y=["Cerrados", "Pendientes"],
            barmode="stack",
            color_discrete_map={"Cerrados": "#28a745", "Pendientes": "#ffc107"},
            title=f"📊 Avance por Contrato — {sub_sel}",
        )
        fig_s.update_layout(height=360, xaxis_title="", yaxis_title="Documentos")
        st.plotly_chart(fig_s, width="stretch")

        st.dataframe(df_cons, width="stretch", hide_index=True)
        st.download_button(
            label="📥 Descargar Consolidado (Excel)",
            data=exportar_df_a_excel(df_cons, f"Consolidado {sub_sel}", metadata_base, RUTA_LOGO),
            file_name=f"Consolidado_{sub_sel}.xlsx".replace(" ", "_"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_cons_sub_xlsx",
        )
    else:
        st.info("No hay información consolidada disponible para este subproyecto.")

# --- 3. Reporte Global ---
elif modo_vista == "🌐 Reporte Global":
    st.subheader("🌐 Reporte Global de Contratos")
    if st.button("🚀 Cargar Todos los Contratos Configurados", type="primary", key="btn_cargar_global"):
        filas_glob = []
        with st.spinner("Cargando todos los contratos configurados..."):
            for item_c in listar_carpetas_de_estructura(ESTRUCTURA_PROYECTOS):
                c_ctx = {
                    "_alcance": item_c["alcance"],
                    "_subproyecto": item_c["subproyecto"],
                    "_tipo_obra": item_c["tipo_obra"],
                    "_contrato": item_c["contrato"],
                    "carpetas": [{
                        "carpeta": item_c["carpeta"],
                        "nombre_pestana": item_c["carpeta"],
                        "url": item_c["url"],
                        "longitud_maxima": item_c["longitud_maxima"],
                    }],
                    "urls_excluidas": item_c["urls_excluidas"],
                    "campos_personalizados": item_c["config_sub"].get("campos_personalizados", {}),
                    "id_campo_fecha": item_c["config_sub"].get("id_campo_fecha", "6011839"),
                    "id_campo_pendiente": item_c["config_sub"].get("id_campo_pendiente", "6011840"),
                }
                paq = extraer_datos_contrato_completo(
                    contrato_key=f"{item_c['contrato']}::{item_c['carpeta']}",
                    info_contrato_tuple=_congelar_estructura(c_ctx),
                    periodo_corte=periodo_actual,
                    version_cache=st.session_state.version_cache,
                )
                cp = paq.get("col_pendiente", "PENDIENTE POR")
                df_full = paq.get("datos_completos", {}).get(item_c["carpeta"], pd.DataFrame())
                rk = construir_resumen_carpeta(df_full, cp)
                filas_glob.append({
                    "Subproyecto": item_c["subproyecto"],
                    "Contrato": item_c["contrato"],
                    "Carpeta": item_c["carpeta"],
                    **rk,
                })
        df_glob = pd.DataFrame(filas_glob)
        if not df_glob.empty:
            tot_g = int(df_glob["Total"].sum())
            cer_g = int(df_glob["Cerrados"].sum())
            pen_g = int(df_glob["Pendientes"].sum())
            av_g = round(cer_g / tot_g * 100, 1) if tot_g > 0 else 0.0

            kg1, kg2, kg3, kg4 = st.columns(4)
            kg1.metric("📄 Total Global", tot_g)
            kg2.metric("✅ Cerrados", cer_g)
            kg3.metric("⏳ Pendientes", pen_g)
            kg4.metric("📈 Avance Global", f"{av_g}%")

            df_graf_g = df_glob.groupby("Subproyecto", as_index=False)[["Cerrados", "Pendientes"]].sum()
            fig_g = px.bar(
                df_graf_g,
                x="Subproyecto",
                y=["Cerrados", "Pendientes"],
                barmode="stack",
                color_discrete_map={"Cerrados": "#28a745", "Pendientes": "#ffc107"},
                title="📊 Estado Documentario Global por Subproyecto",
            )
            fig_g.update_layout(height=360, xaxis_title="", yaxis_title="Documentos")
            st.plotly_chart(fig_g, width="stretch")

            st.dataframe(df_glob, width="stretch", hide_index=True)
            st.download_button(
                label="📥 Descargar Reporte Global (Excel)",
                data=exportar_df_a_excel(df_glob, "Reporte Global", metadata_base, RUTA_LOGO),
                file_name="Reporte_Global_ACC.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_global_xlsx",
            )
    else:
        st.info("Haz clic en **🚀 Cargar Todos los Contratos Configurados** para generar el resumen global.")

# --- 4. Constructor Dinámico de Reporte Personalizado ---
elif modo_vista == "🛠️ Constructor Dinámico":
    st.subheader(f"🛠️ Constructor Dinámico de Reporte Personalizado - {contrato_sel}")
    if datos_completos:
        carpetas_disp = list(datos_completos.keys())
        carpetas_sel = st.multiselect(
            "Selecciona carpetas",
            carpetas_disp,
            default=carpetas_disp[:1] if carpetas_disp else [],
            key="pers_carpetas",
        )
        if carpetas_sel:
            dfs_unidos = []
            for cp_nom in carpetas_sel:
                df_tmp = datos_completos[cp_nom].copy()
                if "CARPETA" not in df_tmp.columns:
                    df_tmp["CARPETA"] = cp_nom
                dfs_unidos.append(df_tmp)
            df_base_pers = pd.concat(dfs_unidos, ignore_index=True)

            cols_disp = list(df_base_pers.columns)
            default_cols = [c for c in ORDEN_COLUMNAS_DASHBOARD if c in cols_disp] or cols_disp[:6]
            campos_sel = st.multiselect(
                "Selecciona campos para el reporte",
                cols_disp,
                default=default_cols,
                key="pers_campos",
            )

            if campos_sel:
                df_pers_final = df_base_pers[campos_sel].copy()
                df_pers_disp, col_cfg_p = preparar_df_con_links(df_pers_final, df_base_pers)
                st.dataframe(
                    estilizar_tabla_streamlit(df_pers_disp, col_pendiente),
                    width="stretch",
                    hide_index=True,
                    column_config=col_cfg_p,
                )

                c_pbtn1, c_pbtn2 = st.columns(2)
                with c_pbtn1:
                    excel_pers = exportar_df_a_excel(
                        df_pers_final,
                        f"Reporte Dinámico - {contrato_sel}",
                        metadata_base,
                        RUTA_LOGO,
                        col_pendiente if col_pendiente in df_pers_final.columns else None,
                    )
                    st.download_button(
                        label="📥 Descargar Reporte Dinámico (Excel)",
                        data=excel_pers,
                        file_name=f"Reporte_Dinamico_{contrato_sel.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_descarga_pers_xlsx",
                        width="stretch",
                    )

                with c_pbtn2:
                    pdf_pers = exportar_df_a_pdf(
                        df_pers_final,
                        f"Reporte Dinámico - {contrato_sel}",
                        metadata_base,
                        RUTA_LOGO,
                        col_pendiente if col_pendiente in df_pers_final.columns else None,
                    )
                    st.download_button(
                        label="📄 Descargar Reporte Dinámico (PDF)",
                        data=pdf_pers,
                        file_name=f"Reporte_Dinamico_{contrato_sel.replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        key="btn_descarga_pers_pdf",
                        width="stretch",
                    )
            else:
                st.warning("Selecciona al menos un campo para generar la vista previa y exportar.")
        else:
            st.info("Selecciona al menos una carpeta para explorar los campos disponibles.")
    else:
        st.info("No hay datos extendidos disponibles para este contrato.")
