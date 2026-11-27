import os
import json
import math
from datetime import datetime
import pytz
import pandas as pd
import psycopg2
from psycopg2.extras import execute_values, RealDictCursor

try:
    import streamlit as st
except Exception:
    st = None


SQL_CREAR_TABLA = """
CREATE TABLE IF NOT EXISTS documentos_acc (
    id BIGSERIAL PRIMARY KEY,
    periodo_corte VARCHAR(32) NOT NULL,
    alcance VARCHAR(120) NOT NULL DEFAULT '',
    subproyecto VARCHAR(180) NOT NULL DEFAULT '',
    tipo_obra VARCHAR(120) NOT NULL DEFAULT '',
    contrato VARCHAR(200) NOT NULL DEFAULT '',
    carpeta VARCHAR(120) NOT NULL DEFAULT '',
    item_id VARCHAR(255) NOT NULL DEFAULT '',
    codigo VARCHAR(300) NOT NULL DEFAULT '',
    rev VARCHAR(60) NOT NULL DEFAULT '',
    descripcion TEXT NOT NULL DEFAULT '',
    fecha VARCHAR(40) NOT NULL DEFAULT '',
    estado VARCHAR(120) NOT NULL DEFAULT '',
    pendiente_por VARCHAR(120) NOT NULL DEFAULT '',
    dias INTEGER NOT NULL DEFAULT 0,
    ruta_autodesk TEXT NOT NULL DEFAULT '',
    url_doc TEXT NOT NULL DEFAULT '',
    datos_extra JSONB NOT NULL DEFAULT '{}'::jsonb,
    actualizado_en TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT uq_doc_corte UNIQUE (periodo_corte, contrato, carpeta, item_id)
);

CREATE INDEX IF NOT EXISTS idx_docs_corte_contrato
    ON documentos_acc (periodo_corte, contrato, carpeta);

CREATE INDEX IF NOT EXISTS idx_docs_estado_pendiente
    ON documentos_acc (estado, pendiente_por);
"""


def _obtener_secreto(clave: str, defecto: str = "") -> str:
    if st is not None:
        try:
            if clave in st.secrets:
                return str(st.secrets[clave]).strip()
        except Exception:
            pass
    return str(os.environ.get(clave, defecto)).strip()


def obtener_dsn() -> str:
    url = _obtener_secreto("DATABASE_URL", "")
    if url:
        return url

    host = _obtener_secreto("DB_HOST", "")
    port = _obtener_secreto("DB_PORT", "5432")
    dbname = _obtener_secreto("DB_NAME", "postgres")
    user = _obtener_secreto("DB_USER", "")
    password = _obtener_secreto("DB_PASSWORD", "")

    if host and user and password:
        return f"postgresql://{user}:{password}@{host}:{port}/{dbname}?sslmode=require"
    return ""


def bd_disponible() -> bool:
    return bool(obtener_dsn())


def obtener_conexion():
    dsn = obtener_dsn()
    if not dsn:
        raise RuntimeError("No está configurada DATABASE_URL ni las variables DB_HOST / DB_USER / DB_PASSWORD.")
    return psycopg2.connect(dsn, connect_timeout=15)


def inicializar_bd() -> bool:
    if not bd_disponible():
        return False
    with obtener_conexion() as conn:
        with conn.cursor() as cur:
            cur.execute(SQL_CREAR_TABLA)
        conn.commit()
    return True


def _limpiar_valor(val):
    if val is None:
        return ""
    if isinstance(val, float) and math.isnan(val):
        return ""
    if pd.isna(val):
        return ""
    return val


def guardar_documentos_bd(
    df_todo: pd.DataFrame,
    periodo_corte: str,
    alcance: str,
    subproyecto: str,
    tipo_obra: str,
    contrato: str,
    carpeta: str
) -> int:
    if df_todo is None or df_todo.empty or not bd_disponible():
        return 0

    inicializar_bd()
    columnas_fijas = {
        "CARPETA", "CODIGO", "REV", "DESCRIPCION", "FECHA",
        "ESTADO", "PENDIENTE POR", "DIAS", "RUTA_EN_AUTODESK",
        "URL_DOC", "ITEM_ID"
    }

    filas_insert = []
    ahora_utc = datetime.now(pytz.UTC)

    for _, row in df_todo.iterrows():
        item_id = str(_limpiar_valor(row.get("ITEM_ID", ""))).strip()
        codigo = str(_limpiar_valor(row.get("CODIGO", ""))).strip()
        if not item_id and not codigo:
            continue
        if not item_id:
            item_id = f"cod::{codigo}::rev::{_limpiar_valor(row.get('REV', ''))}"

        dias_raw = _limpiar_valor(row.get("DIAS", 0))
        try:
            dias_int = int(float(dias_raw)) if str(dias_raw).strip() != "" else 0
        except Exception:
            dias_int = 0

        extra_dict = {}
        for col_name, val in row.items():
            if col_name not in columnas_fijas:
                val_limpio = _limpiar_valor(val)
                if val_limpio != "":
                    extra_dict[str(col_name)] = (
                        val_limpio.isoformat()
                        if hasattr(val_limpio, "isoformat")
                        else str(val_limpio)
                    )

        filas_insert.append((
            str(periodo_corte),
            str(alcance or ""),
            str(subproyecto or ""),
            str(tipo_obra or ""),
            str(contrato or ""),
            str(carpeta or ""),
            item_id,
            codigo,
            str(_limpiar_valor(row.get("REV", ""))).strip(),
            str(_limpiar_valor(row.get("DESCRIPCION", ""))).strip(),
            str(_limpiar_valor(row.get("FECHA", ""))).strip(),
            str(_limpiar_valor(row.get("ESTADO", ""))).strip(),
            str(_limpiar_valor(row.get("PENDIENTE POR", ""))).strip(),
            dias_int,
            str(_limpiar_valor(row.get("RUTA_EN_AUTODESK", ""))).strip(),
            str(_limpiar_valor(row.get("URL_DOC", ""))).strip(),
            json.dumps(extra_dict, ensure_ascii=False),
            ahora_utc
        ))

    if not filas_insert:
        return 0

    sql_upsert = """
    INSERT INTO documentos_acc (
        periodo_corte, alcance, subproyecto, tipo_obra, contrato, carpeta,
        item_id, codigo, rev, descripcion, fecha, estado, pendiente_por,
        dias, ruta_autodesk, url_doc, datos_extra, actualizado_en
    ) VALUES %s
    ON CONFLICT (periodo_corte, contrato, carpeta, item_id)
    DO UPDATE SET
        alcance = EXCLUDED.alcance,
        subproyecto = EXCLUDED.subproyecto,
        tipo_obra = EXCLUDED.tipo_obra,
        codigo = EXCLUDED.codigo,
        rev = EXCLUDED.rev,
        descripcion = EXCLUDED.descripcion,
        fecha = EXCLUDED.fecha,
        estado = EXCLUDED.estado,
        pendiente_por = EXCLUDED.pendiente_por,
        dias = EXCLUDED.dias,
        ruta_autodesk = EXCLUDED.ruta_autodesk,
        url_doc = EXCLUDED.url_doc,
        datos_extra = EXCLUDED.datos_extra,
        actualizado_en = EXCLUDED.actualizado_en
    """

    with obtener_conexion() as conn:
        with conn.cursor() as cur:
            execute_values(cur, sql_upsert, filas_insert, page_size=500)
        conn.commit()

    return len(filas_insert)


def cargar_documentos_bd(
    periodo_corte: str = None,
    contrato: str = None,
    carpeta: str = None,
    incluir_extra: bool = True
) -> pd.DataFrame:
    if not bd_disponible():
        return pd.DataFrame()

    inicializar_bd()
    condiciones = []
    params = []

    if periodo_corte:
        condiciones.append("periodo_corte = %s")
        params.append(periodo_corte)
    if contrato:
        condiciones.append("contrato = %s")
        params.append(contrato)
    if carpeta:
        condiciones.append("carpeta = %s")
        params.append(carpeta)

    where_sql = f"WHERE {' AND '.join(condiciones)}" if condiciones else ""
    sql_select = f"""
    SELECT
        periodo_corte, alcance, subproyecto, tipo_obra, contrato, carpeta,
        item_id, codigo, rev, descripcion, fecha, estado, pendiente_por,
        dias, ruta_autodesk, url_doc, datos_extra, actualizado_en
    FROM documentos_acc
    {where_sql}
    ORDER BY contrato, carpeta, codigo
    """

    with obtener_conexion() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(sql_select, params)
            rows = cur.fetchall()

    if not rows:
        return pd.DataFrame()

    registros = []
    for r in rows:
        base = {
            "PERIODO_CORTE": r["periodo_corte"],
            "ALCANCE": r["alcance"],
            "SUBPROYECTO": r["subproyecto"],
            "TIPO_OBRA": r["tipo_obra"],
            "CONTRATO": r["contrato"],
            "CARPETA": r["carpeta"],
            "ITEM_ID": r["item_id"],
            "CODIGO": r["codigo"],
            "REV": r["rev"],
            "DESCRIPCION": r["descripcion"],
            "FECHA": r["fecha"],
            "ESTADO": r["estado"],
            "PENDIENTE POR": r["pendiente_por"],
            "DIAS": r["dias"],
            "RUTA_EN_AUTODESK": r["ruta_autodesk"],
            "URL_DOC": r["url_doc"],
        }
        if incluir_extra and r.get("datos_extra"):
            extra = r["datos_extra"] if isinstance(r["datos_extra"], dict) else json.loads(r["datos_extra"])
            for k, v in extra.items():
                if k not in base:
                    base[k] = v
        registros.append(base)

    return pd.DataFrame(registros)
