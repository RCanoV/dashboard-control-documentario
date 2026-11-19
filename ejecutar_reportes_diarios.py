import os
import re
import json
import time
import urllib.parse
import smtplib
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed
from email.message import EmailMessage
import requests
import pandas as pd

# Librerías para Excel (OpenPyXL)
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, PatternFill

# Librerías para PDF (ReportLab)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.platypus import Image as RLImage

CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
SMTP_SERVER = os.environ.get("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASS = os.environ.get("SMTP_PASS", "")

RUTA_LOGO = "logoshp.jpeg"
DIAS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
ORDEN_COLUMNAS_DASHBOARD = ["CODIGO", "REV", "DESCRIPCION", "FECHA", "ESTADO", "PENDIENTE POR", "DIAS"]

token_global = None
tiempo_inicio_token = 0


# ====================================================================
# --- 1. HORARIOS Y FILTROS DINÁMICOS ---
# ====================================================================
def corresponde_ejecutar_ahora(reporte: dict) -> bool:
    """Verifica si toca enviar el reporte según la hora de Perú (UTC-5)."""
    if not reporte.get("activo", True):
        return False
    if os.getenv("GITHUB_EVENT_NAME") == "workflow_dispatch":
        return True

    ahora_peru = datetime.now(timezone(timedelta(hours=-5)))
    dia_actual = DIAS_ES[ahora_peru.weekday()]
    hora_actual = f"{ahora_peru.hour:02d}:00"

    dias_conf = reporte.get("dias_semana", DIAS_ES)
    horas_conf = reporte.get("horarios", [])
    return (dia_actual in dias_conf) and (hora_actual in horas_conf)


def aplicar_condiciones_filtro(df: pd.DataFrame, condiciones: list) -> pd.DataFrame:
    """Aplica dinámicamente las reglas configuradas en el formulario."""
    if df.empty or not condiciones:
        return df

    df_res = df.copy()
    for cond in condiciones:
        campo = cond.get("campo")
        op = cond.get("operador")
        val = str(cond.get("valor", "")).strip()

        # Compatibilidad si el filtro usa DIAS_TRANSCURRIDOS pero la columna es DIAS
        if campo == "DIAS_TRANSCURRIDOS" and "DIAS_TRANSCURRIDOS" not in df_res.columns and "DIAS" in df_res.columns:
            campo = "DIAS"

        if campo not in df_res.columns:
            continue

        serie_str = df_res[campo].astype(str).str.strip()
        if op == "igual a":
            df_res = df_res[serie_str.str.lower() == val.lower()]
        elif op == "diferente de":
            df_res = df_res[serie_str.str.lower() != val.lower()]
        elif op == "contiene":
            df_res = df_res[serie_str.str.lower().str.contains(val.lower(), na=False)]
        elif op == "no contiene":
            df_res = df_res[~serie_str.str.lower().str.contains(val.lower(), na=False)]
        elif op in ["mayor que", "menor que"]:
            serie_num = pd.to_numeric(df_res[campo], errors="coerce")
            try:
                val_num = float(val)
                df_res = df_res[serie_num > val_num] if op == "mayor que" else df_res[serie_num < val_num]
            except ValueError:
                pass
    return df_res


# ====================================================================
# --- 2. UTILIDADES Y ESTILOS IDÉNTICOS AL DASHBOARD ---
# ====================================================================
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


def extraer_project_y_folder_id(url_carpeta: str):
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


def limpiar_nombre_hoja(nombre: str, usados: set) -> str:
    limpio = re.sub(r'[\\/*?:\[\]]', '_', nombre).strip()[:31] or "Documentos"
    base = limpio
    contador = 2
    while limpio in usados:
        sufijo = f"_{contador}"
        limpio = f"{base[:31 - len(sufijo)]}{sufijo}"
        contador += 1
    usados.add(limpio)
    return limpio


def exportar_hojas_a_excel(hojas_info: list, ruta_excel: str, ruta_logo: str):
    """Genera el Excel multi-hoja con el mismo formato, logo y colores del dashboard."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for info in hojas_info:
        ws = wb.create_sheet(title=info["hoja"])
        df = info["df"]
        meta = info["metadatos"]
        col_pendiente = info["col_pendiente"]

        if os.path.exists(ruta_logo):
            img = ExcelImage(ruta_logo)
            img.height = 40
            img.width = 120
            ws.add_image(img, "H1")

        ws["A1"] = info["titulo"]
        ws["A1"].font = Font(size=14, bold=True)

        ws["A3"] = "Fecha de Exportación:"
        ws["B3"] = meta["fecha"]
        ws["A4"] = "Alcance:"
        ws["B4"] = meta["alcance"]
        ws["A5"] = "Subproyecto:"
        ws["B5"] = meta["subproyecto"]
        ws["A6"] = "Tipo de Obra:"
        ws["B6"] = meta["tipo_obra"]
        ws["A7"] = "Contrato:"
        ws["B7"] = meta["contrato"]

        for row in range(3, 8):
            ws[f"A{row}"].font = Font(bold=True)

        fila_inicio = 9
        for col_num, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=fila_inicio, column=col_num, value=col_name)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
            col_u = str(col_name).upper()
            ancho = 42 if col_u == "DESCRIPCION" else (28 if col_u == "CODIGO" else 18)
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_num)].width = ancho

        col_idx_pendiente = df.columns.tolist().index(col_pendiente) if (col_pendiente and col_pendiente in df.columns) else None

        for r_idx, row_data in enumerate(df.values, fila_inicio + 1):
            bg_hex, font_hex = "FFFFFF", "000000"
            if col_idx_pendiente is not None:
                bg, text = obtener_colores_estado(row_data[col_idx_pendiente])
                bg_hex = bg.replace("#", "")
                font_hex = text.replace("#", "")

            for c_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value="" if pd.isna(value) else str(value))
                if col_idx_pendiente is not None:
                    cell.fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
                    cell.font = Font(color=font_hex)

        fila_pie = fila_inicio + len(df) + 2
        pie_cell = ws.cell(row=fila_pie, column=1, value="Desarrollado por Roberto Cano - Shougang Hierro Perú")
        pie_cell.font = Font(italic=True, color="4A5568")

    wb.save(ruta_excel)


def exportar_hojas_a_pdf(hojas_info: list, ruta_pdf: str, ruta_logo: str):
    """Genera el PDF con el mismo formato, logo, colores y pie de página del dashboard."""
    def agregar_pie_pagina(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(colors.HexColor("#4A5568"))
        canvas.drawString(doc.leftMargin, 15, "Desarrollado por Roberto Cano - Shougang Hierro Perú")
        canvas.restoreState()

    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=landscape(letter),
        leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=35
    )
    ancho_util = 742
    elements = []
    styles = getSampleStyleSheet()

    cell_style = ParagraphStyle(name="CellText", parent=styles["Normal"], fontSize=7, leading=9)
    header_style = ParagraphStyle(
        name="HeaderCellText", parent=styles["Normal"],
        fontSize=7, leading=9, textColor=colors.whitesmoke, fontName="Helvetica-Bold"
    )

    for idx, info in enumerate(hojas_info):
        if idx > 0:
            elements.append(PageBreak())

        df = info["df"]
        meta = info["metadatos"]
        col_pendiente = info["col_pendiente"]

        logo = RLImage(ruta_logo, width=100, height=30, kind="proportional") if os.path.exists(ruta_logo) else Paragraph("", styles["Normal"])

        meta_text = (
            f"<b><font size=12>{info['titulo']}</font></b><br/><br/>"
            f"<b>Fecha de Emisión:</b> {meta['fecha']}<br/>"
            f"<b>Alcance:</b> {meta['alcance']}<br/>"
            f"<b>Subproyecto:</b> {meta['subproyecto']}<br/>"
            f"<b>Tipo de Obra:</b> {meta['tipo_obra']}<br/>"
            f"<b>Contrato:</b> {meta['contrato']}"
        )

        tabla_cabecera = Table([[logo, Paragraph(meta_text, styles["Normal"])]], colWidths=[120, 600])
        tabla_cabecera.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ]))
        elements.append(tabla_cabecera)
        elements.append(Spacer(1, 15))

        encabezados = [Paragraph(str(col), header_style) for col in df.columns]
        data_matrix = [encabezados]
        estilos_tabla = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
        ]

        for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
            if col_pendiente and col_pendiente in df.columns:
                bg_color, text_color = obtener_colores_estado(row[col_pendiente])
            else:
                bg_color = "#ffffff" if row_idx % 2 != 0 else "#f9fafb"
                text_color = "#000000"

            estilos_tabla.append(("BACKGROUND", (0, row_idx), (-1, row_idx), colors.HexColor(bg_color)))

            fila = []
            for col_name in df.columns:
                val = row[col_name]
                text_str = "" if pd.isna(val) else str(val)
                text_str = text_str.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                if col_pendiente and col_pendiente in df.columns:
                    text_str = f'<font color="{text_color}">{text_str}</font>'
                fila.append(Paragraph(text_str, cell_style))
            data_matrix.append(fila)

        anchos_columnas = []
        for col in df.columns:
            col_u = str(col).upper()
            if col_u == "CODIGO": anchos_columnas.append(0.18 * ancho_util)
            elif col_u == "DESCRIPCION": anchos_columnas.append(0.35 * ancho_util)
            elif col_u in ["REV", "DIAS"]: anchos_columnas.append(0.06 * ancho_util)
            elif col_u == "FECHA": anchos_columnas.append(0.09 * ancho_util)
            elif col_u == "ESTADO": anchos_columnas.append(0.11 * ancho_util)
            elif col_u == "PENDIENTE POR": anchos_columnas.append(0.15 * ancho_util)
            else: anchos_columnas.append(0.10 * ancho_util)

        factor_norm = ancho_util / sum(anchos_columnas)
        anchos_columnas = [w * factor_norm for w in anchos_columnas]

        tabla = Table(data_matrix, repeatRows=1, colWidths=anchos_columnas)
        tabla.setStyle(TableStyle(estilos_tabla))
        elements.append(tabla)

    doc.build(elements, onFirstPage=agregar_pie_pagina, onLaterPages=agregar_pie_pagina)


# ====================================================================
# --- 3. API AUTODESK ACC (EXACTA AL DASHBOARD) ---
# ====================================================================
def obtener_token(scope="data:read"):
    url = "https://developer.api.autodesk.com/authentication/v2/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"}
    data = {"grant_type": "client_credentials", "scope": scope}
    r = requests.post(url, headers=headers, data=data, auth=(CLIENT_ID, CLIENT_SECRET))
    return r.json().get("access_token") if r.status_code == 200 else None


def asegurar_token():
    global token_global, tiempo_inicio_token
    if not token_global or (time.time() - tiempo_inicio_token > 3000):
        token_global = obtener_token("data:read")
        tiempo_inicio_token = time.time()


def deconstruir_diccionario(datos, prefijo, resultado):
    if isinstance(datos, dict):
        for llave, valor in datos.items():
            nueva_llave = f"{prefijo}_{llave.strip().upper()}" if prefijo else llave.strip().upper()
            if any(x in nueva_llave for x in ["LINKS", "SCHEMA", "HREF", "DERIVATIVES"]):
                continue
            deconstruir_diccionario(valor, nueva_llave, resultado)
    elif isinstance(datos, list):
        for i, elemento in enumerate(datos):
            if isinstance(elemento, dict):
                name = elemento.get("name") or elemento.get("displayName") or elemento.get("title")
                val = elemento.get("value") or elemento.get("displayValue")
                if name and val is not None:
                    resultado[f"CAMPO_CUSTOM_{str(name).strip().upper()}"] = val
                else:
                    deconstruir_diccionario(elemento, f"{prefijo}_{i}", resultado)
            else:
                resultado[prefijo] = datos
    else:
        resultado[prefijo] = datos


def obtener_tip_version_metadata(project_id, item_id):
    asegurar_token()
    clean_item_id = item_id.split("?version=")[0]
    url = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/items/{urllib.parse.quote(clean_item_id, safe='')}/tip"
    headers = {"Authorization": f"Bearer {token_global}", "Content-Type": "application/json", "x-ads-region": "US"}

    while True:
        headers["x-ads-region"] = "US"
        r = requests.get(url, headers=headers)
        if r.status_code == 429:
            time.sleep(2)
            continue
        elif r.status_code != 200:
            headers["x-ads-region"] = "EMEA"
            r = requests.get(url, headers=headers)
            if r.status_code == 429:
                time.sleep(2)
                continue
        break
    return r.json().get("data", {}) if r.status_code == 200 else {}


def procesar_archivo_pdf(item, project_id, ruta_actual, name, campos_personalizados):
    asegurar_token()
    item_id = item.get("id", "")
    registro_plano = {
        "NOMBRE_DEL_ARCHIVO": name,
        "RUTA_EN_AUTODESK": ruta_actual,
        "ITEM_ID": item_id
    }
    deconstruir_diccionario(item, "ITEM", registro_plano)

    tip_version = obtener_tip_version_metadata(project_id, item_id)
    if tip_version:
        deconstruir_diccionario(tip_version, "VERSION", registro_plano)

    vid = item.get("relationships", {}).get("tip", {}).get("data", {}).get("id", "")
    if vid:
        vid_encoded = urllib.parse.quote(vid, safe="")
        url_attr = f"https://developer.api.autodesk.com/bim360/docs/v1/projects/{project_id.replace('b.', '')}/versions/{vid_encoded}/custom-attributes"
        while True:
            asegurar_token()
            res_attr = requests.get(url_attr, headers={"Authorization": f"Bearer {token_global}", "Content-Type": "application/json"})
            if res_attr.status_code == 429:
                time.sleep(2)
                continue
            break
        if res_attr.status_code == 200:
            campos_objetivo = set(campos_personalizados.keys())
            for attr in res_attr.json():
                attr_id = str(attr.get("id"))
                if attr_id in campos_objetivo:
                    registro_plano[f"CAMPO_CUSTOM_{attr_id}"] = str(attr.get("value", "")).strip()

    return registro_plano


def escanear_recursivo(project_id, folder_id, lista_datos, longitud_maxima, ids_excluidos, campos_personalizados, ruta_actual="Raíz"):
    if folder_id in ids_excluidos:
        return

    encoded_fid = urllib.parse.quote(folder_id, safe="")
    url_api = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/folders/{encoded_fid}/contents"

    while url_api:
        asegurar_token()
        headers = {"Authorization": f"Bearer {token_global}", "x-ads-region": "US"}
        res = requests.get(url_api, headers=headers)
        if res.status_code == 429:
            time.sleep(5)
            continue
        elif res.status_code != 200:
            headers["x-ads-region"] = "EMEA"
            res = requests.get(url_api, headers=headers)
            if res.status_code != 200:
                break

        data = res.json()
        items_pdf = []

        for item in data.get("data", []):
            tipo = item.get("type")
            name = str(item.get("attributes", {}).get("displayName", "Sin Nombre")).strip()

            if tipo == "folders":
                sub_id = item["id"].replace("%3A", ":")
                if sub_id in ids_excluidos:
                    continue
                if len(name) <= longitud_maxima:
                    escanear_recursivo(project_id, sub_id, lista_datos, longitud_maxima, ids_excluidos, campos_personalizados, f"{ruta_actual} > {name}")
            elif tipo == "items" and ".pdf" in name.lower():
                items_pdf.append((item, name))

        if items_pdf:
            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = [
                    executor.submit(procesar_archivo_pdf, i[0], project_id, ruta_actual, i[1], campos_personalizados)
                    for i in items_pdf
                ]
                for future in as_completed(futures):
                    res_item = future.result()
                    if res_item:
                        lista_datos.append(res_item)

        enlaces = data.get("links", {})
        url_api = enlaces["next"].get("href") if "next" in enlaces else None


# ====================================================================
# --- 4. ENVÍO DE CORREO (CON EXCEL + PDF ADJUNTOS) ---
# ====================================================================
def enviar_correo(cfg: dict, ruta_excel: str, ruta_pdf: str, variables: dict):
    para_lista = cfg.get("para", [])
    cc_lista = cfg.get("cc", [])

    para_str = ", ".join(para_lista) if isinstance(para_lista, list) else str(para_lista)
    cc_str = ", ".join(cc_lista) if isinstance(cc_lista, list) else str(cc_lista)

    if not para_str or not SMTP_USER:
        print("Correo omitido: falta destinatario o SMTP_USER.")
        return

    asunto_tpl = cfg.get("asunto", "Reporte Control Documentario - {nombre_reporte} ({fecha})")
    cuerpo_tpl = cfg.get("cuerpo", "Adjunto el reporte {nombre_reporte} al {fecha} ({total_registros} registros).")

    msg = EmailMessage()
    msg["Subject"] = asunto_tpl.format(**variables)
    msg["From"] = SMTP_USER
    msg["To"] = para_str
    if cc_str:
        msg["Cc"] = cc_str
    msg.set_content(cuerpo_tpl.format(**variables))

    # Adjuntar Excel
    with open(ruta_excel, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(ruta_excel)
        )

    # Adjuntar PDF
    with open(ruta_pdf, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="pdf",
            filename=os.path.basename(ruta_pdf)
        )

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)
    print(f"✅ Correo enviado (Excel + PDF) a: {para_str}")


# ====================================================================
# --- 5. EJECUCIÓN PRINCIPAL ---
# ====================================================================
def main():
    global token_global, tiempo_inicio_token
    if not os.path.exists("config_reportes.json"):
        print("No existe config_reportes.json")
        return

    with open("config_reportes.json", "r", encoding="utf-8") as f:
        contenido = f.read().strip()
        plantillas = json.loads(contenido) if contenido else {}

    if not plantillas:
        print("No hay reportes configurados en config_reportes.json.")
        return

    token_global = obtener_token("data:read")
    tiempo_inicio_token = time.time()
    if not token_global:
        raise RuntimeError("No se pudo obtener el token de Autodesk ACC.")

    hora_peru = datetime.now(timezone(timedelta(hours=-5)))
    fecha_str = hora_peru.strftime("%d/%m/%Y")
    fecha_hora_str = hora_peru.strftime("%d/%m/%Y %H:%M")
    fecha_archivo = hora_peru.strftime("%Y%m%d")
    fecha_hoy_ts = pd.Timestamp(hora_peru.date())

    reporte_objetivo = os.getenv("REPORTE_OBJETIVO", "").strip()

    for nombre_rep, cfg in plantillas.items():
        if reporte_objetivo and nombre_rep != reporte_objetivo:
            continue
        if not reporte_objetivo and not corresponde_ejecutar_ahora(cfg):
            print(f"⏭️ Omitido '{nombre_rep}' (inactivo o fuera de su día/horario configurado).")
            continue

        print(f"🚀 Procesando reporte: {nombre_rep}...")
        hojas_info = []
        hojas_usadas = set()
        total_registros = 0

        for obj in cfg.get("carpetas_seleccionadas", []):
            url_carpeta = obj.get("url", "")
            project_id, fid = extraer_project_y_folder_id(url_carpeta)
            if not project_id or not fid:
                continue

            excluidas_ids = {extraer_project_y_folder_id(u)[1] for u in obj.get("urls_excluidas", []) if u}
            config_sub = obj.get("config_sub", {})
            campos_pers = config_sub.get("campos_personalizados", {})
            id_fecha = str(config_sub.get("id_campo_fecha", ""))
            id_pendiente = str(config_sub.get("id_campo_pendiente", ""))
            col_pendiente = campos_pers.get(id_pendiente, "PENDIENTE POR")
            longitud_max = int(obj.get("longitud_maxima", 10))

            datos_carpeta = []
            escanear_recursivo(project_id, fid, datos_carpeta, longitud_max, excluidas_ids, campos_pers, obj.get("carpeta", "Raíz"))
            if not datos_carpeta:
                continue

            df_completo = pd.DataFrame(datos_carpeta)

            # Garantizar columnas de campos personalizados y descripción
            if "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION" not in df_completo.columns:
                df_completo["ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION"] = ""

            for campo_id in campos_pers.keys():
                col_c = f"CAMPO_CUSTOM_{campo_id}"
                if col_c not in df_completo.columns:
                    df_completo[col_c] = pd.NaT if campo_id == id_fecha else ""

            # Formatear FECHA (dd/mm/yyyy) y calcular DIAS exactamente igual al dashboard
            col_fecha_raw = f"CAMPO_CUSTOM_{id_fecha}"
            if col_fecha_raw in df_completo.columns:
                fechas_conv = pd.to_datetime(df_completo[col_fecha_raw], errors="coerce")
                if fechas_conv.dt.tz is not None:
                    fechas_conv = fechas_conv.dt.tz_localize(None)
                df_completo["DIAS"] = (fecha_hoy_ts - fechas_conv).dt.days
                df_completo[col_fecha_raw] = fechas_conv.dt.strftime("%d/%m/%Y").fillna("")
                df_completo["DIAS"] = df_completo["DIAS"].fillna("")
            else:
                df_completo["DIAS"] = ""

            # Renombrar columnas igual que el dashboard
            renombres = {
                "NOMBRE_DEL_ARCHIVO": "CODIGO",
                "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION": "DESCRIPCION",
            }
            for cid, nombre_col in campos_pers.items():
                renombres[f"CAMPO_CUSTOM_{cid}"] = nombre_col

            df = df_completo.rename(columns=renombres)

            if col_pendiente in df.columns:
                df.loc[df[col_pendiente].astype(str).str.strip().str.upper() == "CERRADO", "DIAS"] = 0

            # Agregar columnas de contexto temporalmente para permitir filtrar por ellas
            df["ALCANCE"] = obj.get("alcance", "")
            df["SUBPROYECTO"] = obj.get("subproyecto", "")
            df["TIPO_OBRA"] = obj.get("tipo_obra", "")
            df["CONTRATO"] = obj.get("contrato", "")
            df["CARPETA"] = obj.get("carpeta", "")

            # Aplicar condiciones de filtro del formulario
            df = aplicar_condiciones_filtro(df, cfg.get("condiciones_filtro", []))
            if df.empty:
                continue

            # Elegir columnas finales (si el usuario no eligió específicas, usa el orden exacto del dashboard)
            cols_usuario = cfg.get("columnas_excel", [])
            if cols_usuario:
                cols_mapeadas = ["DIAS" if c == "DIAS_TRANSCURRIDOS" else c for c in cols_usuario]
                cols_finales = [c for c in cols_mapeadas if c in df.columns]
            else:
                cols_finales = [c for c in ORDEN_COLUMNAS_DASHBOARD if c in df.columns]

            df_export = df[cols_finales].copy()
            total_registros += len(df_export)

            etiqueta_hoja = f"{obj.get('contrato', '')}_{obj.get('carpeta', '')}"
            hojas_info.append({
                "hoja": limpiar_nombre_hoja(etiqueta_hoja, hojas_usadas),
                "titulo": f"Detalle de Documentos: {obj.get('carpeta', '')}",
                "df": df_export,
                "col_pendiente": col_pendiente,
                "metadatos": {
                    "fecha": fecha_hora_str,
                    "alcance": obj.get("alcance", ""),
                    "subproyecto": obj.get("subproyecto", ""),
                    "tipo_obra": obj.get("tipo_obra", ""),
                    "contrato": obj.get("contrato", ""),
                }
            })

        if not hojas_info:
            print(f"ℹ️ Sin resultados que cumplan los filtros para '{nombre_rep}'.")
            continue

        base_archivo = re.sub(r'[^\w\-]', '_', nombre_rep)
        ruta_excel = f"{base_archivo}_{fecha_archivo}.xlsx"
        ruta_pdf = f"{base_archivo}_{fecha_archivo}.pdf"

        exportar_hojas_a_excel(hojas_info, ruta_excel, RUTA_LOGO)
        exportar_hojas_a_pdf(hojas_info, ruta_pdf, RUTA_LOGO)

        variables = {
            "nombre_reporte": nombre_rep,
            "fecha": fecha_str,
            "total_registros": total_registros,
            "REPORTE": nombre_rep,
            "FECHA": fecha_str,
            "TOTAL": total_registros
        }
        enviar_correo(cfg, ruta_excel, ruta_pdf, variables)


if __name__ == "__main__":
    main()
