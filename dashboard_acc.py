import os
import re
import time
import urllib.parse
from io import BytesIO
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import requests
import pandas as pd
import numpy as np

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
    obtener_subproyectos,
    obtener_contratos,
    obtener_info_contrato,
)

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


def _leer_credencial_acc(clave: str) -> str:
    try:
        if clave in st.secrets:
            val = str(st.secrets[clave]).strip()
            if val:
                return val
    except Exception:
        pass
    return str(os.environ.get(clave, "")).strip()


CLIENT_ID = _leer_credencial_acc("CLIENT_ID")
CLIENT_SECRET = _leer_credencial_acc("CLIENT_SECRET")


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
    registro = {
        "NOMBRE_DEL_ARCHIVO": nombre_archivo,
        "RUTA_EN_AUTODESK": ruta_actual,
        "ITEM_ID": item_id,
    }
    deconstruir_diccionario(item, "ITEM", registro)

    tip_version = obtener_tip_version_metadata(token, project_id, item_id)
    if tip_version:
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
        for _ in range(3):
            res_attr = requests.get(
                url_attr,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                },
                timeout=25,
            )
            if res_attr.status_code == 429:
                time.sleep(2)
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
            with ThreadPoolExecutor(max_workers=5) as executor:
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
        df_completo["DIAS"] = df_completo["DIAS"].fillna("")
    else:
        df_completo["DIAS"] = ""

    renombres = {
        "NOMBRE_DEL_ARCHIVO": "CODIGO",
        "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION": "DESCRIPCION",
    }
    for cid, nombre_col in (campos_personalizados or {}).items():
        renombres[f"CAMPO_CUSTOM_{cid}"] = nombre_col

    df_todo = df_completo.rename(columns=renombres)

    if col_pendiente in df_todo.columns:
        df_todo.loc[
            df_todo[col_pendiente].astype(str).str.strip().str.upper() == "CERRADO",
            "DIAS",
        ] = 0

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
    nombre_pestana = str(config_carpeta.get("nombre_pestana") or config_carpeta.get("carpeta") or "Carpeta").strip()
    alcance = str(config_carpeta.get("_alcance", "")).strip()
    subproyecto = str(config_carpeta.get("_subproyecto", "")).strip()
    tipo_obra = str(config_carpeta.get("_tipo_obra", "")).strip()
    contrato = str(config_carpeta.get("_contrato", "")).strip()
    corte = periodo_corte or obtener_periodo_actual()

    if not forzar_api and bd_disponible() and contrato:
        try:
            df_bd = cargar_documentos_bd(
                periodo_corte=corte,
                contrato=contrato,
                carpeta=nombre_pestana,
                incluir_extra=True,
            )
            if df_bd is not None and not df_bd.empty:
                df_todo_bd = df_bd.rename(columns={"CODIGO": "CODIGO"})
                if "DIAS" in df_todo_bd.columns and col_pendiente in df_todo_bd.columns:
                    df_todo_bd.loc[
                        df_todo_bd[col_pendiente].astype(str).str.strip().str.upper() == "CERRADO",
                        "DIAS",
                    ] = 0
                cols_vista = [
                    c for c in (orden_columnas_deseado or ORDEN_COLUMNAS_DASHBOARD) if c in df_todo_bd.columns
                ]
                if not cols_vista:
                    cols_vista = [c for c in ORDEN_COLUMNAS_DASHBOARD if c in df_todo_bd.columns]
                df_vista_bd = df_todo_bd[cols_vista].copy()
                return nombre_pestana, df_vista_bd, df_todo_bd
        except Exception:
            pass

    token = obtener_token_acc(CLIENT_ID, CLIENT_SECRET)
    project_id, fid = extraer_project_y_folder_id(config_carpeta.get("url", ""))
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

    return nombre_pestana, df_vista, df_todo


@st.cache_data(show_spinner=False, ttl=14400, max_entries=40)
def extraer_datos_contrato_completo(
    contrato_key: str,
    info_contrato_tuple: tuple,
    periodo_corte: str,
    version_cache: int = 0,
):
    info_contrato = dict(info_contrato_tuple)
    alcance_ctx = str(info_contrato.get("_alcance", "")).strip()
    subproyecto_ctx = str(info_contrato.get("_subproyecto", "")).strip()
    tipo_obra_ctx = str(info_contrato.get("_tipo_obra", "")).strip()
    contrato_ctx = str(info_contrato.get("_contrato") or contrato_key).strip()

    ids_excluidos = {
        extraer_project_y_folder_id(u)[1]
        for u in info_contrato.get("urls_excluidas", [])
        if u
    }
    campos_personalizados = info_contrato.get("campos_personalizados", {}) or {}
    id_campo_fecha = str(info_contrato.get("id_campo_fecha", ""))
    id_campo_pendiente = str(info_contrato.get("id_campo_pendiente", ""))
    col_pendiente = campos_personalizados.get(id_campo_pendiente, "PENDIENTE POR")
    orden_cols = info_contrato.get("columnas_visibles") or ORDEN_COLUMNAS_DASHBOARD

    carpetas_cfg = info_contrato.get("carpetas", [])
    resultados_vista = {}
    resultados_todo = {}

    for cfg_c in carpetas_cfg:
        cfg_c_enriquecida = dict(cfg_c)
        cfg_c_enriquecida["_alcance"] = alcance_ctx
        cfg_c_enriquecida["_subproyecto"] = subproyecto_ctx
        cfg_c_enriquecida["_tipo_obra"] = tipo_obra_ctx
        cfg_c_enriquecida["_contrato"] = contrato_ctx

        pest, df_p, df_t = _procesar_carpeta_config(
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

    return {
        "datos": resultados_vista,
        "datos_completos": resultados_todo,
        "col_pendiente": col_pendiente,
    }


def _congelar_estructura(obj):
    if isinstance(obj, dict):
        return tuple(sorted((k, _congelar_estructura(v)) for k, v in obj.items()))
    if isinstance(obj, list):
        return tuple(_congelar_estructura(x) for x in obj)
    return obj


def precargar_todos_los_contratos(periodo_corte: str, version_cache: int = 0) -> dict:
    cache_global = {}
    for alcance, subs in ESTRUCTURA_PROYECTOS.items():
        for sub_nom, info_sub in subs.items():
            for tipo_obra, contratos_dict in info_sub.get("tipos_obra", {}).items():
                for cont_nom, info_cont in contratos_dict.items():
                    info_cont_ctx = dict(info_cont)
                    info_cont_ctx["_alcance"] = alcance
                    info_cont_ctx["_subproyecto"] = sub_nom
                    info_cont_ctx["_tipo_obra"] = tipo_obra
                    info_cont_ctx["_contrato"] = cont_nom
                    cache_global[cont_nom] = extraer_datos_contrato_completo(
                        contrato_key=cont_nom,
                        info_contrato_tuple=_congelar_estructura(info_cont_ctx),
                        periodo_corte=periodo_corte,
                        version_cache=version_cache,
                    )
    return cache_global


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
                cell = ws.cell(row=r_idx, column=c_idx, value="" if pd.isna(val) else str(val))
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


# --- Estado inicial de sesión ---
if "version_cache" not in st.session_state:
    st.session_state.version_cache = 0
if "ultima_actualizacion" not in st.session_state:
    st.session_state.ultima_actualizacion = obtener_hora_peru().strftime("%d/%m/%Y %H:%M")

periodo_actual = obtener_periodo_actual()

# --- Barra lateral de navegación ---
with st.sidebar:
    if os.path.exists(RUTA_LOGO):
        st.image(RUTA_LOGO, width=180)
    st.title("Control Documentario")
    st.caption(f"Corte activo: `{periodo_actual}`")
    st.caption(f"Última sincronización: {st.session_state.ultima_actualizacion}")

    estado_bd = "🟢 Supabase PostgreSQL activo" if bd_disponible() else "⚪ Caché en memoria (sin BD)"
    st.info(estado_bd)

    if st.button("🔄 Actualizar Datos desde ACC", use_container_width=True):
        st.session_state.version_cache += 1
        st.session_state.ultima_actualizacion = obtener_hora_peru().strftime("%d/%m/%Y %H:%M")
        st.rerun()

    st.divider()
    alcances = list(ESTRUCTURA_PROYECTOS.keys())
    alcance_sel = st.selectbox("Alcance", alcances, index=0 if alcances else None)

    subproyectos = obtener_subproyectos(alcance_sel) if alcance_sel else []
    sub_sel = st.selectbox("Subproyecto", subproyectos, index=0 if subproyectos else None)

    info_sub = ESTRUCTURA_PROYECTOS.get(alcance_sel, {}).get(sub_sel, {})
    tipos_obra = list(info_sub.get("tipos_obra", {}).keys())
    tipo_sel = st.selectbox("Tipo de Obra", tipos_obra, index=0 if tipos_obra else None)

    contratos = obtener_contratos(alcance_sel, sub_sel, tipo_sel) if (alcance_sel and sub_sel and tipo_sel) else []
    contrato_sel = st.selectbox("Contrato", contratos, index=0 if contratos else None)

    st.divider()
    modo_vista = st.radio(
        "Vista del Dashboard",
        [
            "Detalle de Contrato",
            "Consolidado de Subproyecto",
            "Reporte Global",
            "Constructor Dinámico",
        ],
    )

# --- Cabecera principal ---
st.title("📊 Tablero de Control Documentario - Shougang Hierro Perú")

if not alcance_sel or not sub_sel or not tipo_sel or not contrato_sel:
    st.warning("Selecciona un Alcance, Subproyecto, Tipo de Obra y Contrato en la barra lateral.")
    st.stop()

info_contrato_raw = obtener_info_contrato(alcance_sel, sub_sel, tipo_sel, contrato_sel) or {}
info_contrato_ctx = dict(info_contrato_raw)
info_contrato_ctx["_alcance"] = alcance_sel
info_contrato_ctx["_subproyecto"] = sub_sel
info_contrato_ctx["_tipo_obra"] = tipo_sel
info_contrato_ctx["_contrato"] = contrato_sel

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

fecha_export_str = obtener_hora_peru().strftime("%d/%m/%Y %H:%M")
metadata_base = {
    "fecha": fecha_export_str,
    "alcance": alcance_sel,
    "subproyecto": sub_sel,
    "tipo_obra": tipo_sel,
    "contrato": contrato_sel,
}

# --- 1. Detalle de Contrato ---
if modo_vista == "Detalle de Contrato":
    st.subheader(f"📁 Detalle de Contrato: {contrato_sel}")
    if not datos_vista:
        st.info("No hay carpetas configuradas para este contrato.")
    else:
        pestanas = st.tabs(list(datos_vista.keys()))
        for tab, nombre_carpeta in zip(pestanas, datos_vista.keys()):
            with tab:
                df_c = datos_vista[nombre_carpeta]
                df_t = datos_completos.get(nombre_carpeta, df_c)

                col_f1, col_f2, col_f3 = st.columns(3)
                estados_disp = ["Todos"]
                if "ESTADO" in df_c.columns:
                    estados_disp += sorted(
                        [x for x in df_c["ESTADO"].dropna().astype(str).str.strip().unique() if x]
                    )
                f_estado = col_f1.selectbox("Filtrar por Estado", estados_disp, key=f"est_{nombre_carpeta}")

                pend_disp = ["Todos"]
                if col_pendiente in df_c.columns:
                    pend_disp += sorted(
                        [x for x in df_c[col_pendiente].dropna().astype(str).str.strip().unique() if x]
                    )
                f_pend = col_f2.selectbox("Filtrar por Pendiente Por", pend_disp, key=f"pen_{nombre_carpeta}")
                f_busq = col_f3.text_input("Buscar código o descripción", key=f"busq_{nombre_carpeta}").strip()

                df_filtrado = df_c.copy()
                if f_estado != "Todos" and "ESTADO" in df_filtrado.columns:
                    df_filtrado = df_filtrado[df_filtrado["ESTADO"].astype(str).str.strip() == f_estado]
                if f_pend != "Todos" and col_pendiente in df_filtrado.columns:
                    df_filtrado = df_filtrado[df_filtrado[col_pendiente].astype(str).str.strip() == f_pend]
                if f_busq:
                    mask_b = pd.Series(False, index=df_filtrado.index)
                    for col_b in ["CODIGO", "DESCRIPCION"]:
                        if col_b in df_filtrado.columns:
                            mask_b = mask_b | df_filtrado[col_b].astype(str).str.contains(f_busq, case=False, na=False)
                    df_filtrado = df_filtrado[mask_b]

                st.caption(f"Mostrando {len(df_filtrado)} de {len(df_c)} documentos en {nombre_carpeta}")
                st.dataframe(estilizar_tabla_streamlit(df_filtrado, col_pendiente), use_container_width=True, hide_index=True)

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
                    use_container_width=True,
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
                    use_container_width=True,
                )

# --- 2. Consolidado de Subproyecto ---
elif modo_vista == "Consolidado de Subproyecto":
    st.subheader(f"🗂️ Consolidado de Subproyecto: {sub_sel}")
    filas_cons = []
    for t_obra, c_dict in info_sub.get("tipos_obra", {}).items():
        for c_nom, c_info in c_dict.items():
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
                total_d = len(df_full)
                cerrados_d = 0
                pendientes_d = total_d
                if cp in df_full.columns and total_d > 0:
                    est_u = df_full[cp].astype(str).str.strip().str.upper()
                    cerrados_d = int((est_u == "CERRADO").sum())
                    pendientes_d = total_d - cerrados_d
                avance_d = (cerrados_d / total_d * 100) if total_d > 0 else 0.0
                filas_cons.append({
                    "Tipo de Obra": t_obra,
                    "Contrato": c_nom,
                    "Carpeta": carp_nom,
                    "Cerrados": cerrados_d,
                    "Pendientes": pendientes_d,
                    "Total": total_d,
                    "Avance (%)": round(avance_d, 1),
                })

    df_cons = pd.DataFrame(filas_cons)
    if not df_cons.empty:
        st.dataframe(df_cons, use_container_width=True, hide_index=True)
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
elif modo_vista == "Reporte Global":
    st.subheader("🌐 Reporte Global de Contratos")
    if st.button("Cargar Todos los Contratos Configurados", key="btn_cargar_global"):
        with st.spinner("Precargando todos los contratos..."):
            cache_all = precargar_todos_los_contratos(
                periodo_corte=periodo_actual,
                version_cache=st.session_state.version_cache,
            )
        filas_glob = []
        for c_nom, paq in cache_all.items():
            cp = paq.get("col_pendiente", "PENDIENTE POR")
            for carp_nom, df_full in paq.get("datos_completos", {}).items():
                tot = len(df_full)
                cer = 0
                if cp in df_full.columns and tot > 0:
                    cer = int((df_full[cp].astype(str).str.strip().str.upper() == "CERRADO").sum())
                filas_glob.append({
                    "Contrato": c_nom,
                    "Carpeta": carp_nom,
                    "Cerrados": cer,
                    "Pendientes": tot - cer,
                    "Total": tot,
                    "Avance (%)": round((cer / tot * 100) if tot > 0 else 0.0, 1),
                })
        df_glob = pd.DataFrame(filas_glob)
        if not df_glob.empty:
            st.dataframe(df_glob, use_container_width=True, hide_index=True)
            st.download_button(
                label="📥 Descargar Reporte Global (Excel)",
                data=exportar_df_a_excel(df_glob, "Reporte Global", metadata_base, RUTA_LOGO),
                file_name="Reporte_Global_ACC.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_global_xlsx",
            )
    else:
        st.info("Haz clic en **Cargar Todos los Contratos Configurados** para generar el resumen global.")

# --- 4. Constructor Dinámico de Reporte Personalizado ---
elif modo_vista == "Constructor Dinámico":
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
                st.dataframe(
                    estilizar_tabla_streamlit(df_pers_final, col_pendiente),
                    use_container_width=True,
                    hide_index=True,
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
                        use_container_width=True,
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
                        use_container_width=True,
                    )
            else:
                st.warning("Selecciona al menos un campo para generar la vista previa y exportar.")
        else:
            st.info("Selecciona al menos una carpeta para explorar los campos disponibles.")
    else:
        st.info("No hay datos extendidos disponibles para este contrato.")
