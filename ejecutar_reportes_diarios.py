import os
import re
import json
import time
import urllib.parse
import smtplib
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
import requests
import pandas as pd

CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
SMTP_SERVER = os.environ.get("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASS = os.environ.get("SMTP_PASS", "")

DIAS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

token_global = None
tiempo_inicio_token = 0


def corresponde_ejecutar_ahora(reporte: dict) -> bool:
    """Verifica si toca enviar el reporte según la hora de Perú (UTC-5)."""
    if not reporte.get("activo", True):
        return False

    # Si se disparó manualmente desde GitHub Actions ("Run workflow"), se ejecuta siempre
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


def extraer_project_y_folder_id(url_carpeta: str):
    """Extrae project_id (con prefijo b.) y folder_id desde una URL de Autodesk ACC."""
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
    """Limpia caracteres inválidos de Excel y garantiza nombres únicos de máx. 31 caracteres."""
    limpio = re.sub(r'[\\/*?:\[\]]', '_', nombre).strip()[:31] or "Reporte"
    base = limpio
    contador = 2
    while limpio in usados:
        sufijo = f"_{contador}"
        limpio = f"{base[:31 - len(sufijo)]}{sufijo}"
        contador += 1
    usados.add(limpio)
    return limpio


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


def obtener_tip_version_metadata(project_id, item_id):
    asegurar_token()
    clean_item_id = item_id.split("?version=")[0]
    encoded_item_id = urllib.parse.quote(clean_item_id, safe="")
    url = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/items/{encoded_item_id}/tip"
    headers = {"Authorization": f"Bearer {token_global}", "Content-Type": "application/json", "x-ads-region": "US"}
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        headers["x-ads-region"] = "EMEA"
        r = requests.get(url, headers=headers)
    return r.json().get("data", {}) if r.status_code == 200 else {}


def obtener_atributos_personalizados_lote(project_id: str, urns_version: list) -> dict:
    """Consulta los campos personalizados de ACC Docs en lotes de 50 versiones."""
    if not urns_version:
        return {}

    pid_sin_b = project_id[2:] if project_id.startswith("b.") else project_id
    url = f"https://developer.api.autodesk.com/bim360/docs/v1/projects/{pid_sin_b}/versions:batch-get"
    mapa_attrs = {}

    for i in range(0, len(urns_version), 50):
        asegurar_token()
        lote = urns_version[i:i + 50]
        headers = {"Authorization": f"Bearer {token_global}", "Content-Type": "application/json"}
        try:
            r = requests.post(url, headers=headers, json={"urns": lote}, timeout=30)
            if r.status_code == 200:
                for res in r.json().get("results", []):
                    urn = res.get("urn", "")
                    attrs = {}
                    for ca in res.get("customAttributes", []):
                        cid = str(ca.get("id", ""))
                        if cid:
                            attrs[cid] = ca.get("value", "")
                    mapa_attrs[urn] = attrs
        except Exception as e:
            print(f"Aviso al leer atributos personalizados: {e}")

    return mapa_attrs


def escanear_carpeta(project_id, folder_id, acumulador, excluidas_ids, ruta_actual="Raíz"):
    if folder_id in excluidas_ids:
        return

    encoded_folder_id = urllib.parse.quote(folder_id, safe="")
    url = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/folders/{encoded_folder_id}/contents"
    subcarpetas = []

    while url:
        asegurar_token()
        headers = {"Authorization": f"Bearer {token_global}", "Content-Type": "application/json", "x-ads-region": "US"}
        r = requests.get(url, headers=headers)
        if r.status_code == 429:
            time.sleep(3)
            continue
        if r.status_code != 200:
            headers["x-ads-region"] = "EMEA"
            r = requests.get(url, headers=headers)
        if r.status_code != 200:
            break

        data = r.json()
        for item in data.get("data", []):
            tipo = item.get("type")
            attrs = item.get("attributes", {})
            nombre = attrs.get("displayName", "")
            item_id = item.get("id", "")

            if tipo == "folders":
                if item_id not in excluidas_ids:
                    subcarpetas.append({"id": item_id, "nombre": nombre})
            elif tipo == "items":
                tip = obtener_tip_version_metadata(project_id, item_id)
                tip_id = tip.get("id", "")
                tip_attrs = tip.get("attributes", {})
                codigo_base = os.path.splitext(nombre)[0]
                reg = {
                    "ITEM_ID": item_id,
                    "VERSION_URN": tip_id,
                    "RUTA_EN_AUTODESK": ruta_actual,
                    "CODIGO": codigo_base,
                    "DESCRIPCION": tip_attrs.get("description", "") or nombre,
                    "NOMBRE_DEL_ARCHIVO": nombre,
                }
                acumulador.append(reg)

        enlaces = data.get("links", {})
        url = enlaces["next"].get("href") if "next" in enlaces else None

    for carpeta in subcarpetas:
        escanear_carpeta(
            project_id,
            carpeta["id"],
            acumulador,
            excluidas_ids,
            f"{ruta_actual} > {carpeta['nombre']}"
        )


def enriquecer_con_campos_personalizados(acumulador: list, project_id: str, config_sub: dict, hoy_peru: datetime):
    """Agrega las columnas de campos personalizados configurados y calcula DIAS_TRANSCURRIDOS."""
    campos_pers = config_sub.get("campos_personalizados", {})
    id_fecha = str(config_sub.get("id_campo_fecha", ""))
    id_pendiente = str(config_sub.get("id_campo_pendiente", ""))

    col_fecha = campos_pers.get(id_fecha, "FECHA")
    col_pendiente = campos_pers.get(id_pendiente, "PENDIENTE POR")

    urns = [r["VERSION_URN"] for r in acumulador if r.get("VERSION_URN")]
    mapa_attrs = obtener_atributos_personalizados_lote(project_id, urns)

    for reg in acumulador:
        attrs_item = mapa_attrs.get(reg.get("VERSION_URN", ""), {})
        for cid, nombre_col in campos_pers.items():
            reg[nombre_col] = attrs_item.get(str(cid), "")

        # Calcular DIAS_TRANSCURRIDOS
        val_pend = str(reg.get(col_pendiente, "")).strip().upper()
        val_fecha = reg.get(col_fecha, "")
        dias = None
        if val_pend == "CERRADO":
            dias = 0
        elif val_fecha:
            fecha_dt = pd.to_datetime(val_fecha, errors="coerce")
            if pd.notna(fecha_dt):
                dias = max(0, (hoy_peru.date() - fecha_dt.date()).days)
        reg["DIAS_TRANSCURRIDOS"] = dias


def enviar_correo(cfg: dict, ruta_excel: str, variables: dict):
    para_lista = cfg.get("para", [])
    cc_lista = cfg.get("cc", [])

    # Compatibilidad con formato anterior por si queda algún reporte viejo
    if isinstance(para_lista, list):
        para_str = ", ".join(para_lista)
    else:
        para_str = str(para_lista)

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

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)
    print(f"✅ Correo enviado a: {para_str}")


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
    fecha_archivo = hora_peru.strftime("%Y%m%d")

    for nombre_rep, cfg in plantillas.items():
        if not corresponde_ejecutar_ahora(cfg):
            print(f"⏭️ Omitido '{nombre_rep}' (inactivo o fuera de su día/horario configurado).")
            continue

        print(f"🚀 Procesando reporte: {nombre_rep}...")
        hojas_excel = {}
        hojas_usadas = set()
        total_registros = 0

        carpetas_objetivo = cfg.get("carpetas_seleccionadas", [])
        for obj in carpetas_objetivo:
            url_carpeta = obj.get("url", "")
            project_id, fid = extraer_project_y_folder_id(url_carpeta)
            if not project_id or not fid:
                print(f"⚠️ URL inválida en {obj.get('contrato')} - {obj.get('carpeta')}")
                continue

            excluidas_ids = {
                extraer_project_y_folder_id(u)[1]
                for u in obj.get("urls_excluidas", [])
                if u
            }

            etiqueta = f"{obj.get('contrato', '')}_{obj.get('carpeta', '')}"
            nombre_hoja = limpiar_nombre_hoja(etiqueta, hojas_usadas)

            acumulador = []
            escanear_carpeta(project_id, fid, acumulador, excluidas_ids, ruta_actual=obj.get("carpeta", "Raíz"))
            if not acumulador:
                continue

            enriquecer_con_campos_personalizados(
                acumulador,
                project_id,
                obj.get("config_sub", {}),
                hora_peru
            )

            df = pd.DataFrame(acumulador)
            df.insert(0, "CARPETA", obj.get("carpeta", ""))
            df.insert(0, "CONTRATO", obj.get("contrato", ""))
            df.insert(0, "TIPO_OBRA", obj.get("tipo_obra", ""))
            df.insert(0, "SUBPROYECTO", obj.get("subproyecto", ""))
            df.insert(0, "ALCANCE", obj.get("alcance", ""))

            # Aplicar condiciones de filtro del formulario
            df = aplicar_condiciones_filtro(df, cfg.get("condiciones_filtro", []))

            if df.empty:
                continue

            # Filtrar columnas si el usuario eligió un subconjunto
            cols = cfg.get("columnas_excel", [])
            if cols:
                cols_validas = [c for c in cols if c in df.columns]
                if cols_validas:
                    df = df[cols_validas]
            else:
                df = df.drop(columns=["ITEM_ID", "VERSION_URN"], errors="ignore")

            total_registros += len(df)
            hojas_excel[nombre_hoja] = df

        if not hojas_excel:
            print(f"ℹ️ Sin resultados que cumplan los filtros para '{nombre_rep}'.")
            continue

        nombre_archivo = re.sub(r'[^\w\-]', '_', nombre_rep)
        ruta_excel = f"{nombre_archivo}_{fecha_archivo}.xlsx"
        with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
            for hoja, df_h in hojas_excel.items():
                df_h.to_excel(writer, sheet_name=hoja, index=False)

        variables = {
            "nombre_reporte": nombre_rep,
            "fecha": fecha_str,
            "total_registros": total_registros,
            "REPORTE": nombre_rep,
            "FECHA": fecha_str,
            "TOTAL": total_registros
        }
        enviar_correo(cfg, ruta_excel, variables)


if __name__ == "__main__":
    main()
