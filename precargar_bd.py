import os
import re
import json
import time
import urllib.parse
from datetime import datetime, timedelta, timezone

import requests
import pandas as pd

from modulo_bd import bd_disponible, guardar_documentos_bd

CLIENT_ID = os.environ.get("CLIENT_ID", "").strip()
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "").strip()
RUTA_ESTRUCTURA = "estructura_proyectos.json"

token_global = ""
tiempo_inicio_token = 0.0


def obtener_hora_peru() -> datetime:
    return datetime.now(timezone(timedelta(hours=-5)))


def obtener_periodo_actual() -> str:
    ahora_pe = obtener_hora_peru()
    fecha_corte = (ahora_pe - timedelta(hours=4)).strftime("%Y-%m-%d")
    return f"{fecha_corte}_04AM"


def asegurar_token() -> str:
    global token_global, tiempo_inicio_token
    if token_global and (time.time() - tiempo_inicio_token) < 2700:
        return token_global

    url = "https://developer.api.autodesk.com/authentication/v2/token"
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json",
    }
    resp = requests.post(
        url,
        headers=headers,
        data={"grant_type": "client_credentials", "scope": "data:read"},
        auth=(CLIENT_ID, CLIENT_SECRET),
        timeout=30,
    )
    if resp.status_code != 200:
        raise RuntimeError(f"Error obteniendo token ACC ({resp.status_code}): {resp.text}")

    token_global = str(resp.json().get("access_token", "")).strip()
    tiempo_inicio_token = time.time()
    return token_global


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


def obtener_atributos_en_lote_acc(
    project_id: str,
    version_ids: list[str],
    campos_personalizados: dict,
) -> dict:
    mapa_attrs = {}
    if not project_id or not version_ids or not campos_personalizados:
        return mapa_attrs

    proj_sin_b = project_id.replace("b.", "")
    url_batch = (
        f"https://developer.api.autodesk.com/bim360/docs/v1/projects/"
        f"{proj_sin_b}/versions:batch-get"
    )
    campos_objetivo = set(str(k) for k in campos_personalizados.keys())

    for i in range(0, len(version_ids), 50):
        lote = version_ids[i : i + 50]
        for intento in range(5):
            token = asegurar_token()
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            }
            r = requests.post(url_batch, headers=headers, json={"urns": lote}, timeout=45)
            if r.status_code == 429:
                espera = int(r.headers.get("Retry-After", 3 * (intento + 1)))
                time.sleep(espera)
                continue
            if r.status_code == 200:
                for res_v in r.json().get("results", []):
                    vid = str(res_v.get("urn", "")).strip()
                    vals = {}
                    for attr in res_v.get("customAttributes", []) or []:
                        aid = str(attr.get("id", ""))
                        if aid in campos_objetivo:
                            val = attr.get("value")
                            vals[f"CAMPO_CUSTOM_{aid}"] = "" if val is None else str(val).strip()
                    if vid:
                        mapa_attrs[vid] = vals
                break
            if r.status_code in (500, 502, 503, 504):
                time.sleep(2 * (intento + 1))
                continue
            break

    return mapa_attrs


def escanear_recursivo_acc(
    project_id: str,
    folder_id: str,
    lista_datos: list,
    longitud_maxima: int,
    ids_excluidos: set,
    campos_personalizados: dict,
    ruta_actual: str = "Raíz",
):
    if not project_id or not folder_id or folder_id in ids_excluidos:
        return

    encoded_fid = urllib.parse.quote(folder_id, safe="")
    url_api = (
        f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/"
        f"folders/{encoded_fid}/contents"
    )

    while url_api:
        res = None
        for region in ("US", "EMEA"):
            for intento in range(4):
                token = asegurar_token()
                headers = {"Authorization": f"Bearer {token}", "x-ads-region": region}
                r = requests.get(url_api, headers=headers, timeout=30)
                if r.status_code == 429:
                    time.sleep(3 * (intento + 1))
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
            vids_lote = []
            registros_lote = []
            for it, nom in items_pdf:
                item_id = str(it.get("id", "")).strip()
                url_webview = (
                    it.get("links", {}).get("webView", {}).get("href", "")
                    if isinstance(it.get("links"), dict)
                    else ""
                )
                vid = (
                    it.get("relationships", {})
                    .get("tip", {})
                    .get("data", {})
                    .get("id", "")
                )
                reg = {
                    "NOMBRE_DEL_ARCHIVO": nom,
                    "RUTA_EN_AUTODESK": ruta_actual,
                    "ITEM_ID": item_id,
                    "URL_DOC": str(url_webview or "").strip(),
                    "_VID": vid,
                }
                deconstruir_diccionario(it, "ITEM", reg)
                registros_lote.append(reg)
                if vid:
                    vids_lote.append(vid)

            mapa_lote = obtener_atributos_en_lote_acc(
                project_id=project_id,
                version_ids=vids_lote,
                campos_personalizados=campos_personalizados,
            )
            for reg in registros_lote:
                vid = reg.pop("_VID", "")
                if vid in mapa_lote:
                    reg.update(mapa_lote[vid])
                lista_datos.append(reg)

        enlaces = data.get("links", {})
        url_api = enlaces["next"].get("href") if "next" in enlaces else None


def listar_carpetas(estructura: dict) -> list[dict]:
    items = []
    for alcance, subs in (estructura or {}).items():
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
                        items.append({
                            "alcance": alcance,
                            "subproyecto": sub_nom,
                            "tipo_obra": tipo_obra,
                            "contrato": cont_nom,
                            "carpeta": carp_nom,
                            "url": str(c.get("url", "")).strip(),
                            "longitud_maxima": int(c.get("longitud_maxima", 10) or 10),
                            "urls_excluidas": info_cont.get("urls_excluidas", []),
                            "campos_personalizados": info_cont.get("campos_personalizados") or cfg_sub.get("campos_personalizados", {}),
                            "id_campo_fecha": str(info_cont.get("id_campo_fecha") or cfg_sub.get("id_campo_fecha", "6011839")),
                            "id_campo_pendiente": str(info_cont.get("id_campo_pendiente") or cfg_sub.get("id_campo_pendiente", "6011840")),
                        })
    return items


def main():
    if not bd_disponible():
        raise RuntimeError("DATABASE_URL no está configurada.")
    if not os.path.exists(RUTA_ESTRUCTURA):
        raise FileNotFoundError(f"No existe {RUTA_ESTRUCTURA} en el repositorio.")

    with open(RUTA_ESTRUCTURA, "r", encoding="utf-8") as f:
        estructura = json.load(f)

    carpetas = listar_carpetas(estructura)
    periodo_corte = obtener_periodo_actual()
    fecha_hoy_ts = pd.Timestamp(obtener_hora_peru().date())
    total_carpetas = len(carpetas)
    total_docs_guardados = 0

    print(f"🚀 Iniciando precarga en Supabase | Corte: {periodo_corte} | Carpetas: {total_carpetas}")
    asegurar_token()

    for idx, item in enumerate(carpetas, start=1):
        project_id, fid = extraer_project_y_folder_id(item["url"])
        if not project_id or not fid:
            print(f"⚠️ [{idx}/{total_carpetas}] URL inválida: {item['contrato']} -> {item['carpeta']}")
            continue

        ids_excl = {extraer_project_y_folder_id(u)[1] for u in item["urls_excluidas"] if u}
        campos_pers = item["campos_personalizados"] or {}
        id_fecha = item["id_campo_fecha"]
        id_pend = item["id_campo_pendiente"]
        col_pend = campos_pers.get(id_pend, "PENDIENTE POR")

        datos_carpeta = []
        escanear_recursivo_acc(
            project_id=project_id,
            folder_id=fid,
            lista_datos=datos_carpeta,
            longitud_maxima=item["longitud_maxima"],
            ids_excluidos=ids_excl,
            campos_personalizados=campos_pers,
            ruta_actual=item["carpeta"],
        )

        if not datos_carpeta:
            print(f"ℹ️ [{idx}/{total_carpetas}] Sin PDFs: {item['contrato']} -> {item['carpeta']}")
            continue

        df = pd.DataFrame(datos_carpeta)
        if "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION" not in df.columns:
            df["ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION"] = ""
        if "URL_DOC" not in df.columns:
            df["URL_DOC"] = ""

        for cid in campos_pers.keys():
            col_c = f"CAMPO_CUSTOM_{cid}"
            if col_c not in df.columns:
                df[col_c] = pd.NaT if str(cid) == str(id_fecha) else ""

        col_fecha_raw = f"CAMPO_CUSTOM_{id_fecha}"
        if col_fecha_raw in df.columns:
            fechas_conv = pd.to_datetime(df[col_fecha_raw], errors="coerce")
            if fechas_conv.dt.tz is not None:
                fechas_conv = fechas_conv.dt.tz_localize(None)
            df["DIAS"] = (fecha_hoy_ts - fechas_conv).dt.days
            df[col_fecha_raw] = fechas_conv.dt.strftime("%d/%m/%Y").fillna("")
        else:
            df["DIAS"] = 0

        renombres = {
            "NOMBRE_DEL_ARCHIVO": "CODIGO",
            "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION": "DESCRIPCION",
        }
        for cid, nom_col in campos_pers.items():
            renombres[f"CAMPO_CUSTOM_{cid}"] = nom_col
        df = df.rename(columns=renombres)

        df["DIAS"] = pd.to_numeric(df["DIAS"], errors="coerce").fillna(0).round(0).astype(int)
        if col_pend in df.columns:
            df.loc[df[col_pend].astype(str).str.strip().str.upper() == "CERRADO", "DIAS"] = 0

        guardados = guardar_documentos_bd(
            df_todo=df,
            periodo_corte=periodo_corte,
            alcance=item["alcance"],
            subproyecto=item["subproyecto"],
            tipo_obra=item["tipo_obra"],
            contrato=item["contrato"],
            carpeta=item["carpeta"],
        )
        total_docs_guardados += guardados
        print(f"✅ [{idx}/{total_carpetas}] {item['contrato']} -> {item['carpeta']}: {guardados} docs guardados")

    print(f"🎉 Precarga finalizada. Total documentos guardados en Supabase: {total_docs_guardados}")


if __name__ == "__main__":
    main()
