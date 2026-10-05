import os
import time
import requests
import pandas as pd
import urllib.parse
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import streamlit as st
import plotly.express as px

# ====================================================================
# --- 1. CONFIGURACIÓN GENERAL DEL PROYECTO ---
# ====================================================================
NOMBRE_PROYECTO = "994440 Paquete OEM 01 - C601 CHEC BISA"

CAMPOS_PERSONALIZADOS = {
    "6011837": "REV",
    "6011840": "PENDIENTE POR",
    "6011839": "FECHA",
    "6011838": "ESTADO"
}
ID_CAMPO_FECHA = "6011839"
ID_CAMPO_PENDIENTE = "6011840"

COL_CODIGO = "CODIGO"
COL_DESCRIPCION = "DESCRIPCION"
ORDEN_COLUMNAS_DESEADO = ["CODIGO", "REV", "DESCRIPCION", "FECHA", "ESTADO", "PENDIENTE POR", "DIAS"]

# ====================================================================
# --- 2. CREDENCIALES SEGURAS Y RUTAS ---
# ====================================================================
# Streamlit leerá estas credenciales desde su bóveda segura (Secrets) en la nube
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]

PROJECT_ID = "3fe40740-4483-4ed9-895b-9cb579cd7b1d"

URLS_EXCLUIDAS = [
    "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Bji4Uop7SSGwBSj2kKdmLg&viewModel=detail&moduleId=folders",
    "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.ZUirAImDTluEAH7D1z2xCA&viewModel=detail&moduleId=folders",
]

CARPETAS_CONFIG = [
    {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.6ayy3aotR1iZErdo8ZQZpA&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
    {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.cWeBjsVeQJOe9T2dgiDLRQ&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
    {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.iLK0CFTcRE-f5vPgXsuWGQ&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
    {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Sgn_HpAZSfK84UcH7Od1Pw&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
    {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.J8zAgoEdRFOlCo4FwXmVHg&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
]

# ====================================================================
# --- 3. LÓGICA DE API AUTODESK ---
# ====================================================================
def obtener_ids_excluidos(urls):
    ids = set()
    for url in urls:
        if "folderUrn=" in url:
            ids.add(url.split("folderUrn=")[1].split("&")[0].replace("%3A", ":"))
        elif "folders/" in url:
            ids.add(url.split("folders/")[1].split("/")[0].replace("%3A", ":"))
    return ids

def obtener_token():
    url = "https://developer.api.autodesk.com/authentication/v2/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"}
    data = {"grant_type": "client_credentials", "scope": "data:read"}
    response = requests.post(url, headers=headers, data=data, auth=(CLIENT_ID, CLIENT_SECRET))
    return response.json().get("access_token") if response.status_code == 200 else None

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

def obtener_tip_version_metadata(token, project_id, item_id):
    clean_item_id = item_id.split("?version=")[0]
    url = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/items/{clean_item_id}/tip"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json", "x-ads-region": "US"}
    
    while True:
        headers["x-ads-region"] = "US" 
        response = requests.get(url, headers=headers)
        if response.status_code == 429:
            time.sleep(2)
            continue
        elif response.status_code == 401:
            nuevo_token = obtener_token()
            if nuevo_token:
                headers["Authorization"] = f"Bearer {nuevo_token}"
                continue
        elif response.status_code != 200:
            headers["x-ads-region"] = "EMEA"
            response = requests.get(url, headers=headers)
            if response.status_code == 429:
                time.sleep(2)
                continue
        break
    return response.json().get("data", {}) if response.status_code == 200 else {}

def procesar_archivo_pdf(item, project_id, headers, ruta_actual, name):
    item_id = item.get("id", "")
    token_val = headers["Authorization"].split(" ")[1]
    
    registro_plano = {
        "NOMBRE_DEL_ARCHIVO": name,
        "RUTA_EN_AUTODESK": ruta_actual,
        "ITEM_ID": item_id
    }
    deconstruir_diccionario(item, "ITEM", registro_plano)
    
    tip_version = obtener_tip_version_metadata(token=token_val, project_id=project_id, item_id=item_id)
    if tip_version:
        deconstruir_diccionario(tip_version, "VERSION", registro_plano)

    vid = item.get("relationships", {}).get("tip", {}).get("data", {}).get("id", "")
    if vid:
        vid_encoded = urllib.parse.quote(vid, safe='')
        url_attr = f"https://developer.api.autodesk.com/bim360/docs/v1/projects/{project_id.replace('b.', '')}/versions/{vid_encoded}/custom-attributes"
        
        while True:
            res_attr = requests.get(url_attr, headers={"Authorization": headers["Authorization"], "Content-Type": "application/json"})
            if res_attr.status_code == 429:
                time.sleep(2)
                continue
            elif res_attr.status_code == 401:
                nuevo_token_attr = obtener_token()
                if nuevo_token_attr:
                    headers["Authorization"] = f"Bearer {nuevo_token_attr}"
                    continue
            break
            
        if res_attr.status_code == 200:
            campos_objetivo = list(CAMPOS_PERSONALIZADOS.keys())
            for attr in res_attr.json():
                attr_id = str(attr.get("id"))
                if attr_id in campos_objetivo:
                    registro_plano[f"CAMPO_CUSTOM_{attr_id}"] = str(attr.get("value", "")).strip()
                    
    return registro_plano

def escanear_recursivo(project_id, folder_id, headers, lista_datos, longitud_maxima, ids_excluidos, ruta_actual="Raíz"):
    if folder_id.replace("%3A", ":") in ids_excluidos:
        return

    url_api = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/folders/{folder_id}/contents"
    while url_api:
        headers["x-ads-region"] = "US"
        res = requests.get(url_api, headers=headers)
        
        if res.status_code == 401:
            nuevo_token = obtener_token()
            if nuevo_token:
                headers["Authorization"] = f"Bearer {nuevo_token}"
                continue
            else:
                break
        elif res.status_code == 429:
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
            attributes = item.get("attributes", {})
            name = str(attributes.get("displayName", "Sin Nombre")).strip()

            if tipo == "folders":
                sub_urn_raw = item["id"].replace("%3A", ":")
                if sub_urn_raw in ids_excluidos:
                    continue
                if len(name) <= longitud_maxima:
                    sub_urn = urllib.parse.quote(item["id"], safe='')
                    escanear_recursivo(project_id, sub_urn, headers, lista_datos, longitud_maxima, ids_excluidos, f"{ruta_actual} > {name}")
            
            elif tipo == "items":
                if ".pdf" in name.lower():
                    items_pdf.append((item, name))

        if items_pdf:
            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = [executor.submit(procesar_archivo_pdf, i[0], project_id, headers, ruta_actual, i[1]) for i in items_pdf]
                for future in as_completed(futures):
                    res_item = future.result()
                    if res_item:
                        lista_datos.append(res_item)

        enlaces = data.get("links", {})
        url_api = enlaces["next"].get("href") if "next" in enlaces else None

# ====================================================================
# --- 4. EXTRACCIÓN CENTRALIZADA PARA EL DASHBOARD ---
# ====================================================================
@st.cache_data(ttl=3600)
def extraer_datos_autodesk():
    token = obtener_token()
    if not token:
        return {}
    
    headers = {"Authorization": f"Bearer {token}", "x-ads-region": "US"}
    ids_excluidos = obtener_ids_excluidos(URLS_EXCLUIDAS)
    
    dataframes_por_pestana = {}
    col_fecha = CAMPOS_PERSONALIZADOS.get(ID_CAMPO_FECHA)
    col_pendiente = CAMPOS_PERSONALIZADOS.get(ID_CAMPO_PENDIENTE)

    for config_carpeta in CARPETAS_CONFIG:
        url_carpeta = config_carpeta.get("url")
        nombre_pestana = config_carpeta.get("nombre", "Sin_Nombre")
        longitud_maxima = config_carpeta.get("longitud_maxima", 10)
        
        # Mensaje para monitorear el progreso en consola
        print(f"⏳ Descargando datos de la carpeta: {nombre_pestana}...")
        
        if not url_carpeta:
            continue
            
        try:
            project_id_raw = url_carpeta.split("projects/")[1].split("?")[0].split("/")[0]
            project_id = project_id_raw if project_id_raw.startswith("b.") else "b." + project_id_raw
            
            if "folderUrn=" in url_carpeta:
                folder_id = url_carpeta.split("folderUrn=")[1].split("&")[0].replace("%3A", ":")
            elif "folders/" in url_carpeta:
                folder_id = url_carpeta.split("folders/")[1].split("/")[0].replace("%3A", ":")
            else:
                continue
            
            datos_filtrados = []
            escanear_recursivo(project_id, folder_id, headers, datos_filtrados, longitud_maxima, ids_excluidos)
            
            if not datos_filtrados:
                continue
                
            df_completo = pd.DataFrame(datos_filtrados)
            
            for campo_id in CAMPOS_PERSONALIZADOS.keys():
                col_name = f"CAMPO_CUSTOM_{campo_id}"
                if col_name not in df_completo.columns:
                    df_completo[col_name] = pd.NaT if campo_id == ID_CAMPO_FECHA else ""

            fechas_convertidas = pd.to_datetime(df_completo[f"CAMPO_CUSTOM_{ID_CAMPO_FECHA}"], errors='coerce')
            if fechas_convertidas.dt.tz is not None:
                fechas_convertidas = fechas_convertidas.dt.tz_localize(None)
            
            fecha_hoy = pd.Timestamp.now().normalize()
            df_completo["DIAS"] = (fecha_hoy - fechas_convertidas).dt.days
            df_completo[f"CAMPO_CUSTOM_{ID_CAMPO_FECHA}"] = fechas_convertidas.dt.strftime('%d/%m/%Y').fillna("")
            df_completo["DIAS"] = df_completo["DIAS"].fillna("")

            columnas_finales = ["NOMBRE_DEL_ARCHIVO", "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION"] + [f"CAMPO_CUSTOM_{c}" for c in CAMPOS_PERSONALIZADOS.keys()] + ["DIAS"]
            df = df_completo[columnas_finales].copy()
            
            diccionario_renombres = {
                "NOMBRE_DEL_ARCHIVO": COL_CODIGO, 
                "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION": COL_DESCRIPCION,
            }
            for campo_id, nombre_columna in CAMPOS_PERSONALIZADOS.items():
                diccionario_renombres[f"CAMPO_CUSTOM_{campo_id}"] = nombre_columna
                
            df = df.rename(columns=diccionario_renombres)
            columnas_existentes = [col for col in ORDEN_COLUMNAS_DESEADO if col in df.columns]
            df = df[columnas_existentes]

            if col_pendiente in df.columns:
                df.loc[df[col_pendiente].astype(str).str.strip().str.upper() == "CERRADO", "DIAS"] = 0
            
            dataframes_por_pestana[nombre_pestana] = df

        except Exception as e:
            continue

    return dataframes_por_pestana


# ====================================================================
# --- 5. INTERFAZ WEB STREAMLIT ---
# ====================================================================
st.set_page_config(page_title="Control Documentario SHP", layout="wide", page_icon="📊")

st.title(f"📊 Dashboard de Control Documentario")
st.markdown(f"**Proyecto:** {NOMBRE_PROYECTO}")

with st.spinner('Conectando a BIM 360 y procesando metadatos...'):
    datos = extraer_datos_autodesk()

if not datos:
    st.error("No se encontraron registros o hubo un problema al conectar con la API.")
else:
    # --- MENÚ LATERAL ---
    st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1a/Shougang_Group_logo.svg/1200px-Shougang_Group_logo.svg.png", width=150)
    st.sidebar.title("Filtros")
    
    opciones_carpetas = ["Resumen Global"] + list(datos.keys())
    seleccion = st.sidebar.selectbox("Seleccionar Vista:", opciones_carpetas)
    
    # --- MÉTRICAS GLOBALES ---
    col_pendiente = CAMPOS_PERSONALIZADOS.get(ID_CAMPO_PENDIENTE, "PENDIENTE POR")
    total_cerrados = 0
    total_abiertos = 0
    
    for df in datos.values():
        if col_pendiente in df.columns:
            estados = df[col_pendiente].astype(str).str.strip().str.upper()
            total_cerrados += (estados == "CERRADO").sum()
            total_abiertos += (estados != "CERRADO").sum()

    st.markdown("---")
    m1, m2, m3 = st.columns(3)
    m1.metric("Total de Documentos", total_cerrados + total_abiertos)
    m2.metric("Documentos Cerrados", total_cerrados)
    m3.metric("Documentos Pendientes", total_abiertos, delta="Requieren atención", delta_color="inverse")
    st.markdown("---")

    # --- VISTA: RESUMEN GLOBAL ---
    if seleccion == "Resumen Global":
        st.subheader("Estado Documentario General")
        
        resumen_data = []
        for nombre, df in datos.items():
            if col_pendiente in df.columns:
                estados = df[col_pendiente].astype(str).str.strip().str.upper()
                cerrados = (estados == 'CERRADO').sum()
                contratista = (estados == 'CONTRATISTA').sum()
                supervision = (estados == 'SUPERVISION').sum()
                shp = (estados == 'SHP').sum()
                
                resumen_data.append({
                    "Carpeta": nombre,
                    "Cerrado": cerrados,
                    "Contratista": contratista,
                    "Supervision": supervision,
                    "SHP": shp
                })
        
        df_resumen = pd.DataFrame(resumen_data)
        
        c1, c2 = st.columns([2, 1])
        with c1:
            fig_bar = px.bar(df_resumen, x="Carpeta", y=["Contratista", "Supervision", "SHP", "Cerrado"],
                             title="Distribución de Pendientes por Carpeta",
                             labels={"value": "Cantidad de Documentos", "variable": "Estado"},
                             barmode="stack",
                             color_discrete_map={"Cerrado": "#28a745", "Contratista": "#dc3545", "Supervision": "#ffc107", "SHP": "#17a2b8"})
            st.plotly_chart(fig_bar, width='stretch')
            
        with c2:
            totales = df_resumen[["Cerrado", "Contratista", "Supervision", "SHP"]].sum().reset_index()
            totales.columns = ["Estado", "Total"]
            fig_pie = px.pie(totales, values="Total", names="Estado", title="Estados del Proyecto (%)",
                             hole=0.4, color="Estado",
                             color_discrete_map={"Cerrado": "#28a745", "Contratista": "#dc3545", "Supervision": "#ffc107", "SHP": "#17a2b8"})
            st.plotly_chart(fig_pie, width='stretch')

    # --- VISTA: DETALLE POR CARPETA ---
    else:
        st.subheader(f"Directorio: {seleccion}")
        df_actual = datos[seleccion].copy()
        
        if col_pendiente in df_actual.columns:
            estados_disponibles = df_actual[col_pendiente].astype(str).unique().tolist()
            filtro_estado = st.multiselect("Filtrar por Responsable:", estados_disponibles, default=estados_disponibles)
            
            df_filtrado = df_actual[df_actual[col_pendiente].isin(filtro_estado)]
            
            st.dataframe(df_filtrado, width='stretch', height=400)
            
            if "DIAS" in df_filtrado.columns:
                df_filtrado["DIAS_NUM"] = pd.to_numeric(df_filtrado["DIAS"], errors="coerce").fillna(0)
                retrasos_df = df_filtrado[df_filtrado["DIAS_NUM"] > 0].sort_values(by="DIAS_NUM", ascending=False)
                
                if not retrasos_df.empty:
                    st.warning("⚠️ Top Documentos con mayor cantidad de días en revisión:")
                    st.dataframe(retrasos_df[[COL_CODIGO, COL_DESCRIPCION, col_pendiente, "DIAS"]].head(5), width='stretch')
        else:
            st.dataframe(df_actual, width='stretch', height=400)