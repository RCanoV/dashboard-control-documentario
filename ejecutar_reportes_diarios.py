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
import numpy as np

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# Librerías para Excel (OpenPyXL)
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, PatternFill, Alignment

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
# --- 2. RESUMEN EJECUTIVO Y GRÁFICO DASHBOARD ---
# ====================================================================
def construir_resumen_ejecutivo(hojas_info: list):
    """Construye la tabla resumen por contrato/carpeta y los KPIs globales."""
    roles_detectados = []
    for info in hojas_info:
        df_full = info["df_full"]
        col_pen = info["col_pendiente"]
        if col_pen in df_full.columns:
            vals = df_full[col_pen].astype(str).str.strip().str.upper().unique()
            for v in vals:
                if v and v not in ["CERRADO", "NAN", "NAT", "NONE", "NULL", ""] and v not in roles_detectados:
                    roles_detectados.append(v)

    filas_resumen = []
    total_global = 0
    cerrados_global = 0

    for info in hojas_info:
        df_full = info["df_full"]
        col_pen = info["col_pendiente"]
        meta = info["metadatos"]

        fila = {
            "Contrato": meta["contrato"],
            "Carpeta": meta["carpeta"],
            "Cerrados": 0
        }

        total_carp = len(df_full)
        cerrados_carp = 0

        if col_pen in df_full.columns:
            estados = df_full[col_pen].astype(str).str.strip().str.upper()
            cerrados_carp = int((estados == "CERRADO").sum())
            fila["Cerrados"] = cerrados_carp

            for rol in roles_detectados:
                fila[rol.title()] = int((estados == rol).sum())

            df_abiertos = df_full[estados != "CERRADO"].copy()
            if not df_abiertos.empty and "DIAS" in df_abiertos.columns:
                df_abiertos["DIAS_NUM"] = pd.to_numeric(df_abiertos["DIAS"], errors="coerce").fillna(0)
                if df_abiertos["DIAS_NUM"].max() > 0:
                    idx_max = df_abiertos["DIAS_NUM"].idxmax()
                    cod = str(df_abiertos.loc[idx_max, "CODIGO"]) if "CODIGO" in df_abiertos.columns else "Doc"
                    dias_max = int(df_abiertos.loc[idx_max, "DIAS_NUM"])
                    resp = str(df_abiertos.loc[idx_max, col_pen]).strip()
                    fila["Mayor Retraso"] = f"{cod} ({dias_max} d) [{resp}]"
                else:
                    fila["Mayor Retraso"] = "Sin retrasos"
            else:
                fila["Mayor Retraso"] = "Sin pendientes"
        else:
            for rol in roles_detectados:
                fila[rol.title()] = 0
            fila["Mayor Retraso"] = "N/A"

        fila["Total"] = total_carp
        avance = (cerrados_carp / total_carp * 100) if total_carp > 0 else 0.0
        fila["Avance (%)"] = f"{avance:.1f}%"

        total_global += total_carp
        cerrados_global += cerrados_carp
        filas_resumen.append(fila)

    roles_cols = [r.title() for r in roles_detectados]
    orden_res = ["Contrato", "Carpeta", "Cerrados"] + roles_cols + ["Total", "Avance (%)", "Mayor Retraso"]
    df_resumen = pd.DataFrame(filas_resumen)[orden_res]

    pendientes_global = total_global - cerrados_global
    avance_global = (cerrados_global / total_global * 100) if total_global > 0 else 0.0

    kpis = {
        "total": total_global,
        "cerrados": cerrados_global,
        "pendientes": pendientes_global,
        "avance": f"{avance_global:.1f}%"
    }
    return df_resumen, roles_cols, kpis


def generar_grafico_resumen(df_resumen: pd.DataFrame, roles_cols: list, ruta_img: str):
    """Genera el gráfico de barras apiladas por Carpeta y Responsable."""
    if df_resumen.empty:
        return None

    etiquetas = [
        f"{row['Contrato'][:14]} - {row['Carpeta']}" if len(df_resumen["Contrato"].unique()) > 1 else str(row["Carpeta"])
        for _, row in df_resumen.iterrows()
    ]
    columnas_graf = ["Cerrados"] + roles_cols
    colores_paleta = ["#28a745", "#ffc107", "#dc3545", "#007bff", "#6f42c1", "#17a2b8", "#fd7e14"]

    fig, ax = plt.subplots(figsize=(9, 3.2))
    x = np.arange(len(etiquetas))
    acumulado = np.zeros(len(etiquetas))

    for i, col in enumerate(columnas_graf):
        valores = pd.to_numeric(df_resumen[col], errors="coerce").fillna(0).values
        ax.bar(x, valores, bottom=acumulado, label=col, color=colores_paleta[i % len(colores_paleta)], width=0.55)
        acumulado += valores

    ax.set_title("Estado Documentario por Carpeta y Responsable", fontsize=11, fontweight="bold", color="#1A365D")
    ax.set_xticks(x)
    ax.set_xticklabels(etiquetas, rotation=18, ha="right", fontsize=8)
    ax.set_ylabel("Cantidad de Documentos", fontsize=8)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.legend(fontsize=8, bbox_to_anchor=(1.01, 1), loc="upper left")
    plt.tight_layout()
    plt.savefig(ruta_img, format="png", dpi=180, bbox_inches="tight")
    plt.close(fig)
    return ruta_img


# ====================================================================
# --- 3. ESTILOS Y EXPORTACIÓN A EXCEL Y PDF ---
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


def exportar_hojas_a_excel(
    hojas_info: list,
    ruta_excel: str,
    ruta_logo: str,
    nombre_rep: str,
    fecha_hora_str: str,
    df_resumen: pd.DataFrame,
    kpis: dict,
    ruta_grafico: str
):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # --- HOJA 1: RESUMEN EJECUTIVO / DASHBOARD ---
    ws_res = wb.create_sheet(title="Resumen")
    if os.path.exists(ruta_logo):
        img_logo = ExcelImage(ruta_logo)
        img_logo.height = 40
        img_logo.width = 120
        ws_res.add_image(img_logo, "G1")

    ws_res["A1"] = f"Resumen Ejecutivo - {nombre_rep}"
    ws_res["A1"].font = Font(size=14, bold=True, color="1A365D")
    ws_res["A2"] = f"Fecha de Emisión: {fecha_hora_str}"
    ws_res["A2"].font = Font(size=10, italic=True, color="4A5568")

    # Tarjetas KPI (Filas 4 y 5)
    tarjetas = [
        ("A", "TOTAL DOCUMENTOS", kpis["total"], "E2E8F0", "1A365D"),
        ("C", "CERRADOS", kpis["cerrados"], "D4EDDA", "155724"),
        ("E", "PENDIENTES", kpis["pendientes"], "FFF3CD", "856404"),
        ("G", "AVANCE GLOBAL", kpis["avance"], "CCE5FF", "004085"),
    ]
    for col_l, titulo_kpi, valor_kpi, bg_hex, fg_hex in tarjetas:
        c_tit = ws_res[f"{col_l}4"]
        c_val = ws_res[f"{col_l}5"]
        c_tit.value = titulo_kpi
        c_val.value = valor_kpi
        c_tit.font = Font(size=9, bold=True, color=fg_hex)
        c_val.font = Font(size=14, bold=True, color=fg_hex)
        c_tit.fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
        c_val.fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
        c_tit.alignment = Alignment(horizontal="center")
        c_val.alignment = Alignment(horizontal="center")

    # Tabla Resumen (desde Fila 7)
    fila_ini_res = 7
    for col_num, col_name in enumerate(df_resumen.columns, 1):
        cell = ws_res.cell(row=fila_ini_res, column=col_num, value=col_name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
        col_u = str(col_name).upper()
        ancho = 38 if col_u == "MAYOR RETRASO" else (24 if col_u in ["CONTRATO", "CARPETA"] else 14)
        ws_res.column_dimensions[openpyxl.utils.get_column_letter(col_num)].width = ancho

    for r_idx, row_data in enumerate(df_resumen.values, fila_ini_res + 1):
        bg_row = "FFFFFF" if r_idx % 2 == 0 else "F7FAFC"
        for c_idx, value in enumerate(row_data, 1):
            cell = ws_res.cell(row=r_idx, column=c_idx, value="" if pd.isna(value) else value)
            cell.fill = PatternFill(start_color=bg_row, end_color=bg_row, fill_type="solid")

    # Incrustar Gráfico debajo de la tabla
    fila_graf = fila_ini_res + len(df_resumen) + 2
    if ruta_grafico and os.path.exists(ruta_grafico):
        img_graf = ExcelImage(ruta_grafico)
        img_graf.width = 680
        img_graf.height = 250
        ws_res.add_image(img_graf, f"A{fila_graf}")
        fila_pie_res = fila_graf + 15
    else:
        fila_pie_res = fila_graf + 1

    pie_res = ws_res.cell(row=fila_pie_res, column=1, value="Desarrollado por Roberto Cano - Shougang Hierro Perú")
    pie_res.font = Font(italic=True, color="4A5568")

    # --- HOJAS DE DETALLE ---
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


def exportar_hojas_a_pdf(
    hojas_info: list,
    ruta_pdf: str,
    ruta_logo: str,
    nombre_rep: str,
    fecha_hora_str: str,
    df_resumen: pd.DataFrame,
    kpis: dict,
    ruta_grafico: str
):
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
    kpi_style = ParagraphStyle(name="KPIText", parent=styles["Normal"], fontSize=9, leading=12, alignment=1)

    # --- PÁGINA 1: RESUMEN EJECUTIVO / DASHBOARD ---
    logo_p1 = RLImage(ruta_logo, width=100, height=30, kind="proportional") if os.path.exists(ruta_logo) else Paragraph("", styles["Normal"])
    meta_p1 = (
        f"<b><font size=13>Resumen Ejecutivo: {nombre_rep}</font></b><br/><br/>"
        f"<b>Fecha de Emisión:</b> {fecha_hora_str}<br/>"
        f"<b>Institución:</b> Shougang Hierro Perú - Control Documentario"
    )
    tabla_cab_p1 = Table([[logo_p1, Paragraph(meta_p1, styles["Normal"])]], colWidths=[120, 620])
    tabla_cab_p1.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    elements.append(tabla_cab_p1)
    elements.append(Spacer(1, 10))

    # Tarjetas KPI
    kpi_data = [[
        Paragraph(f"<b>TOTAL DOCUMENTOS</b><br/><font size=13>{kpis['total']}</font>", kpi_style),
        Paragraph(f"<b>CERRADOS</b><br/><font size=13 color='#155724'>{kpis['cerrados']}</font>", kpi_style),
        Paragraph(f"<b>PENDIENTES</b><br/><font size=13 color='#856404'>{kpis['pendientes']}</font>", kpi_style),
        Paragraph(f"<b>AVANCE GLOBAL</b><br/><font size=13 color='#004085'>{kpis['avance']}</font>", kpi_style),
    ]]
    tabla_kpi = Table(kpi_data, colWidths=[ancho_util / 4] * 4)
    tabla_kpi.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#E2E8F0")),
        ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#D4EDDA")),
        ("BACKGROUND", (2, 0), (2, 0), colors.HexColor("#FFF3CD")),
        ("BACKGROUND", (3, 0), (3, 0), colors.HexColor("#CCE5FF")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.white),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(tabla_kpi)
    elements.append(Spacer(1, 12))

    # Tabla Resumen
    enc_res = [Paragraph(str(c), header_style) for c in df_resumen.columns]
    mat_res = [enc_res]
    for _, row in df_resumen.iterrows():
        mat_res.append([
            Paragraph(str("" if pd.isna(v) else v).replace("&", "&amp;").replace("<", "&lt;"), cell_style)
            for v in row
        ])

    anchos_res = []
    for c in df_resumen.columns:
        cu = str(c).upper()
        if cu == "CONTRATO": anchos_res.append(0.18 * ancho_util)
        elif cu == "CARPETA": anchos_res.append(0.14 * ancho_util)
        elif cu == "MAYOR RETRASO": anchos_res.append(0.28 * ancho_util)
        elif cu in ["TOTAL", "AVANCE (%)", "CERRADOS"]: anchos_res.append(0.08 * ancho_util)
        else: anchos_res.append(0.08 * ancho_util)
    norm_r = ancho_util / sum(anchos_res)
    anchos_res = [w * norm_r for w in anchos_res]

    tabla_res = Table(mat_res, repeatRows=1, colWidths=anchos_res)
    tabla_res.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F9FAFB")]),
    ]))
    elements.append(tabla_res)
    elements.append(Spacer(1, 14))

    # Gráfico Resumen
    if ruta_grafico and os.path.exists(ruta_grafico):
        elements.append(RLImage(ruta_grafico, width=540, height=195, kind="proportional"))

    # --- PÁGINAS DE DETALLE ---
    for info in hojas_info:
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
# --- 4. API AUTODESK ACC ---
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
# --- 5. ENVÍO DE CORREO (EXCEL + PDF) ---
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

    with open(ruta_excel, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(ruta_excel)
        )

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
    print(f"✅ Correo enviado (Excel + PDF con Resumen) a: {para_str}")


# ====================================================================
# --- 6. EJECUCIÓN PRINCIPAL ---
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

            if "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION" not in df_completo.columns:
                df_completo["ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION"] = ""

            for campo_id in campos_pers.keys():
                col_c = f"CAMPO_CUSTOM_{campo_id}"
                if col_c not in df_completo.columns:
                    df_completo[col_c] = pd.NaT if campo_id == id_fecha else ""

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

            renombres = {
                "NOMBRE_DEL_ARCHIVO": "CODIGO",
                "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION": "DESCRIPCION",
            }
            for cid, nombre_col in campos_pers.items():
                renombres[f"CAMPO_CUSTOM_{cid}"] = nombre_col

            df = df_completo.rename(columns=renombres)

            if col_pendiente in df.columns:
                df.loc[df[col_pendiente].astype(str).str.strip().str.upper() == "CERRADO", "DIAS"] = 0

            df["ALCANCE"] = obj.get("alcance", "")
            df["SUBPROYECTO"] = obj.get("subproyecto", "")
            df["TIPO_OBRA"] = obj.get("tipo_obra", "")
            df["CONTRATO"] = obj.get("contrato", "")
            df["CARPETA"] = obj.get("carpeta", "")

            df = aplicar_condiciones_filtro(df, cfg.get("condiciones_filtro", []))
            if df.empty:
                continue

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
                "df_full": df,
                "col_pendiente": col_pendiente,
                "metadatos": {
                    "fecha": fecha_hora_str,
                    "alcance": obj.get("alcance", ""),
                    "subproyecto": obj.get("subproyecto", ""),
                    "tipo_obra": obj.get("tipo_obra", ""),
                    "contrato": obj.get("contrato", ""),
                    "carpeta": obj.get("carpeta", ""),
                }
            })

        if not hojas_info:
            print(f"ℹ️ Sin resultados que cumplan los filtros para '{nombre_rep}'.")
            continue

        # Construir Resumen Ejecutivo y Gráfico
        df_resumen, roles_cols, kpis = construir_resumen_ejecutivo(hojas_info)
        base_archivo = re.sub(r'[^\w\-]', '_', nombre_rep)
        ruta_grafico = f"grafico_{base_archivo}.png"
        generar_grafico_resumen(df_resumen, roles_cols, ruta_grafico)

        ruta_excel = f"{base_archivo}_{fecha_archivo}.xlsx"
        ruta_pdf = f"{base_archivo}_{fecha_archivo}.pdf"

        exportar_hojas_a_excel(hojas_info, ruta_excel, RUTA_LOGO, nombre_rep, fecha_hora_str, df_resumen, kpis, ruta_grafico)
        exportar_hojas_a_pdf(hojas_info, ruta_pdf, RUTA_LOGO, nombre_rep, fecha_hora_str, df_resumen, kpis, ruta_grafico)

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
