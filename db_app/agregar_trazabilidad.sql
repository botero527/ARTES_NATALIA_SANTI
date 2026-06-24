-- ================================================================
-- 1. Columnas modificado_por / modificado_en en todas las tablas
-- ================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('mallas.vitrojet')   AND name='modificado_por') ALTER TABLE mallas.vitrojet   ADD modificado_por NVARCHAR(200) NULL, modificado_en DATETIME2 NULL;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('mallas.grandes')    AND name='modificado_por') ALTER TABLE mallas.grandes    ADD modificado_por NVARCHAR(200) NULL, modificado_en DATETIME2 NULL;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('mallas.pequenas')   AND name='modificado_por') ALTER TABLE mallas.pequenas   ADD modificado_por NVARCHAR(200) NULL, modificado_en DATETIME2 NULL;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('mallas.pasta_plata') AND name='modificado_por') ALTER TABLE mallas.pasta_plata ADD modificado_por NVARCHAR(200) NULL, modificado_en DATETIME2 NULL;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('mallas.vinilos')    AND name='modificado_por') ALTER TABLE mallas.vinilos    ADD modificado_por NVARCHAR(200) NULL, modificado_en DATETIME2 NULL;
GO

-- ================================================================
-- 2. Tabla de trazabilidad
-- ================================================================
IF NOT EXISTS (
    SELECT 1 FROM sys.tables t JOIN sys.schemas s ON t.schema_id=s.schema_id
    WHERE s.name='MALLAS' AND t.name='TRAZABILIDAD'
)
CREATE TABLE MALLAS.TRAZABILIDAD (
    id             INT IDENTITY(1,1) PRIMARY KEY,
    tabla          NVARCHAR(100)  NOT NULL,
    pk_campo       NVARCHAR(100)  NOT NULL,
    pk_valor       NVARCHAR(200)  NOT NULL,
    campo          NVARCHAR(100)  NOT NULL,
    valor_anterior NVARCHAR(MAX),
    valor_nuevo    NVARCHAR(MAX),
    usuario        NVARCHAR(200),
    fecha          DATETIME2      DEFAULT SYSDATETIME()
);
GO

-- Índice para consultas rápidas por tabla/pk
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_TRAZ_TABLA_PK' AND object_id=OBJECT_ID('MALLAS.TRAZABILIDAD'))
    CREATE INDEX IX_TRAZ_TABLA_PK ON MALLAS.TRAZABILIDAD (tabla, pk_valor, fecha DESC);
GO
