import os
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

token_global = None
tiempo_inicio_token = 0

def obtener_token(scope="data:read"):
    url = "https://developer.api.autodesk.com/authentication/v2/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"}
    data = {"grant_type": "client_credentials", "scope": scope}
    r = requests.post(url, headers=headers, data=data, auth=(CLIENT_ID, CLIENT_SECRET))
    return r.json().get("access_token") if r.status_code == 200 else None

def obtener_tip_version_metadata(token, project_id, item_id):
    clean_item_id = item_id.split("?version=")[0]
    encoded_item_id = urllib.parse.quote(clean_item_id, safe="")
    url = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/items/{encoded_item_id}/tip"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json", "x-ads-region": "US"}
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        headers["x-ads-region"] = "EMEA"
        r = requests.get(url, headers=headers)
    return r.json().get("data", {}) if r.status_code == 200 else {}

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

def escanear_carpeta(project_id, folder_id, acumulador, ruta_actual="Raíz"):
    global token_global, tiempo_inicio_token
    encoded_folder_id = urllib.parse.quote(folder_id, safe="")
    url = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/folders/{encoded_folder_id}/contents"
    subcarpetas = []

    while url:
        if time.time() - tiempo_inicio_token > 3000:
            token_global = obtener_token("data:read")
            tiempo_inicio_token = time.time()

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
            if tipo == "folders":
                subcarpetas.append({"id": item.get("id"), "nombre": nombre})
            elif tipo == "items":
                tip = obtener_tip_version_metadata(token_global, project_id, item.get("id", ""))
                reg = {
                    "ITEM_ID": item.get("id", ""),
                    "RUTA_EN_AUTODESK": ruta_actual,
                    "NOMBRE_DEL_ARCHIVO": nombre
                }
                deconstruir_diccionario(item, "ITEM", reg)
                if tip:
                    deconstruir_diccionario(tip, "VERSION", reg)
                acumulador.append(reg)

        enlaces = data.get("links", {})
        url = enlaces["next"].get("href") if "next" in enlaces else None

    for carpeta in subcarpetas:
        escanear_carpeta(project_id, carpeta["id"], acumulador, f"{ruta_actual} > {carpeta['nombre']}")

def aplicar_condiciones(df, condiciones):
    if df.empty or not condiciones:
        return df
    df_f = df.copy()
    for regla in condiciones:
        for op in ["==", "!=", " contiene "]:
            if op in regla:
                col, val = [p.strip() for p in regla.split(op, 1)]
                if col not in df_f.columns:
                    continue
                serie = df_f[col].astype(str).str.strip().str.upper()
                val_up = val.upper()
                if op == "==":
                    df_f = df_f[serie == val_up]
                elif op == "!=":
                    df_f = df_f[serie != val_up]
                elif op == " contiene ":
                    df_f = df_f[serie.str.contains(val_up, na=False)]
                break
    return df_f

def enviar_correo(cfg_correo, ruta_excel, variables):
    if not cfg_correo.get("activar") or not cfg_correo.get("para") or not SMTP_USER:
        return
    msg = EmailMessage()
    msg["Subject"] = cfg_correo["asunto"].format(**variables)
    msg["From"] = SMTP_USER
    msg["To"] = cfg_correo["para"]
    if cfg_correo.get("cc"):
        msg["Cc"] = cfg_correo["cc"]
    msg.set_content(cfg_correo["cuerpo"].format(**variables))

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
    print(f"Correo enviado a: {cfg_correo['para']}")

def main():
    global token_global, tiempo_inicio_token
    if not os.path.exists("config_reportes.json"):
        print("No existe config_reportes.json")
        return

    with open("config_reportes.json", "r", encoding="utf-8") as f:
        plantillas = json.load(f)

    token_global = obtener_token("data:read")
    tiempo_inicio_token = time.time()
    if not token_global:
        raise RuntimeError("No se pudo obtener el token de Autodesk ACC.")

    hora_peru = datetime.now(timezone(timedelta(hours=-5)))
    fecha_str = hora_peru.strftime("%d/%m/%Y")
    fecha_archivo = hora_peru.strftime("%Y%m%d")

    for nombre_rep, cfg in plantillas.items():
        if not cfg.get("es_diario"):
            continue

        print(f"Procesando reporte: {nombre_rep}...")
        hojas_excel = {}
        consolidado = []
        total_registros = 0

        for obj in cfg.get("objetivos", []):
            project_id = obj["project_id"]
            url_carpeta = obj["url"]
            nombre_hoja = f"{obj['contrato']}_{obj['componente']}"[:31]

            fid = (
                url_carpeta.split("folderUrn=")[1].split("&")[0].replace("%3A", ":")
                if "folderUrn=" in url_carpeta
                else (url_carpeta.split("folders/")[1].split("/")[0].replace("%3A", ":") if "folders/" in url_carpeta else url_carpeta)
            )
            fid = urllib.parse.unquote(fid)

            acumulador = []
            escanear_carpeta(project_id, fid, acumulador, ruta_actual=nombre_hoja)
            if not acumulador:
                continue

            df = pd.DataFrame(acumulador)
            # Inyectamos la jerarquía en el reporte
            df.insert(0, "COMPONENTE", obj["componente"])
            df.insert(0, "CONTRATO", obj["contrato"])
            df.insert(0, "TIPO_OBRA", obj["tipo_obra"])
            df.insert(0, "SUBPROYECTO", obj["subproyecto"])
            df.insert(0, "ALCANCE", obj["alcance"])

            df = aplicar_condiciones(df, cfg.get("condiciones", []))

            cols = cfg.get("columnas", [])
            if cols:
                cols_validas = [c for c in cols if c in df.columns]
                if cols_validas:
                    df = df[cols_validas]

            total_registros += len(df)
            if cfg.get("separar_hojas", True):
                hojas_excel[nombre_hoja] = df
            else:
                consolidado.append(df)

        if not hojas_excel and not consolidado:
            print(f"Sin resultados para {nombre_rep}.")
            continue

        ruta_excel = f"{nombre_rep.replace(' ', '_')}_{fecha_archivo}.xlsx"
        with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
            if cfg.get("separar_hojas", True):
                for hoja, df_h in hojas_excel.items():
                    df_h.to_excel(writer, sheet_name=hoja, index=False)
            else:
                pd.concat(consolidado, ignore_index=True).to_excel(writer, sheet_name="Consolidado", index=False)

        variables = {
            "REPORTE": nombre_rep,
            "NIVEL": cfg.get("nivel", ""),
            "SELECCION": ", ".join(cfg.get("valores_nivel", [])),
            "FECHA": fecha_str,
            "TOTAL": total_registros
        }
        enviar_correo(cfg.get("correo", {}), ruta_excel, variables)

if __name__ == "__main__":
    main()
