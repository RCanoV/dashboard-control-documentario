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
