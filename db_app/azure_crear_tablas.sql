-- ============================================================
-- AGP Ingenieria — Esquema MALLAS
-- Servidor: agpcolombia.database.windows.net
-- Base de datos: AGP_Ingenieria
-- Ejecutar este script UNA SOLA VEZ para crear el esquema y tablas
-- ============================================================

-- Crear esquema
IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = 'mallas')
    EXEC('CREATE SCHEMA mallas');
GO

-- ── mallas.grandes ────────────────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='grandes')
CREATE TABLE mallas.grandes (
    codigo        NVARCHAR(30)  NOT NULL PRIMARY KEY,
    cod_veh       NVARCHAR(30)  NULL,
    descripcion   NVARCHAR(200) NULL,
    pieza         NVARCHAR(100) NULL,
    tipo          NVARCHAR(20)  NULL,
    version       NVARCHAR(100) NULL,
    concatenar    NVARCHAR(300) NULL,
    cambio        NVARCHAR(50)  NULL,
    created_at    DATETIME2     DEFAULT GETDATE(),
    updated_at    DATETIME2     DEFAULT GETDATE()
);
GO

-- ── mallas.pequenas ───────────────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='pequenas')
CREATE TABLE mallas.pequenas (
    codigo        INT           NOT NULL PRIMARY KEY,
    cod_veh       NVARCHAR(30)  NULL,
    descripcion   NVARCHAR(200) NULL,
    pieza         NVARCHAR(100) NULL,
    tipo          NVARCHAR(20)  NULL,
    version       NVARCHAR(100) NULL,
    concatenar    NVARCHAR(300) NULL,
    part_number   NVARCHAR(100) NULL,
    vitro_ref     NVARCHAR(50)  NULL,
    cambio        NVARCHAR(50)  NULL,
    created_at    DATETIME2     DEFAULT GETDATE(),
    updated_at    DATETIME2     DEFAULT GETDATE()
);
GO

-- ── mallas.vitrojet ───────────────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='vitrojet')
CREATE TABLE mallas.vitrojet (
    vitro         NVARCHAR(30)  NOT NULL PRIMARY KEY,
    codigo_malla  NVARCHAR(30)  NULL,
    tipo_malla    NCHAR(1)      DEFAULT 'G',
    cod_completo  NVARCHAR(100) NULL,
    bnerig        NVARCHAR(20)  NULL,
    vehiculo      NVARCHAR(200) NULL,
    version       NVARCHAR(100) NULL,
    ruta          NVARCHAR(500) NULL,
    cambio        NVARCHAR(50)  NULL,
    created_at    DATETIME2     DEFAULT GETDATE(),
    updated_at    DATETIME2     DEFAULT GETDATE()
);
GO

-- ── mallas.pasta_plata ────────────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='pasta_plata')
CREATE TABLE mallas.pasta_plata (
    consecutivo   NVARCHAR(30)  NOT NULL PRIMARY KEY,
    tipo          NVARCHAR(20)  NULL,
    vehiculo      NVARCHAR(200) NULL,
    cod_vehiculo  NVARCHAR(30)  NULL,
    version       NVARCHAR(100) NULL,
    pieza         NVARCHAR(100) NULL,
    ruta_archivo  NVARCHAR(500) NULL,
    caso          NVARCHAR(200) NULL,
    cambio        NVARCHAR(50)  NULL,
    created_at    DATETIME2     DEFAULT GETDATE(),
    updated_at    DATETIME2     DEFAULT GETDATE()
);
GO

-- ── mallas.glassjet_viejo ─────────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='glassjet_viejo')
CREATE TABLE mallas.glassjet_viejo (
    id                INT           IDENTITY(1,1) PRIMARY KEY,
    malla             NVARCHAR(50)  NULL,
    glassjet          NVARCHAR(50)  NULL,
    part_number       NVARCHAR(100) NULL,
    tipo              NVARCHAR(20)  NULL,
    vehiculo          NVARCHAR(200) NULL,
    homologacion_vitro NVARCHAR(50) NULL
);
GO

-- ── mallas.vinilos ────────────────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='vinilos')
CREATE TABLE mallas.vinilos (
    herramental   NVARCHAR(30)  NOT NULL PRIMARY KEY,
    vehiculo      NVARCHAR(200) NULL,
    cod_vehiculo  NVARCHAR(30)  NULL,
    version       NVARCHAR(100) NULL,
    pieza         NVARCHAR(100) NULL,
    tipo          NVARCHAR(20)  NULL,
    ruta          NVARCHAR(500) NULL,
    cambio        NVARCHAR(50)  NULL,
    created_at    DATETIME2     DEFAULT GETDATE(),
    updated_at    DATETIME2     DEFAULT GETDATE()
);
GO

-- Verificar creación
SELECT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_SCHEMA = 'mallas'
ORDER BY TABLE_NAME;
