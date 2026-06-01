-- AGP Glass DB — SQL Server Schema
-- Ejecutar en la base de datos Vitros_Mallas

USE Vitros_Mallas;
GO

-- ─── 1. MALLAS GRANDES (códigos A-XXXXX) ──────────────────────────────────────
CREATE TABLE mallas_grandes (
    codigo       NVARCHAR(20)  NOT NULL PRIMARY KEY,
    cod_veh      NVARCHAR(20),
    descripcion  NVARCHAR(200),
    pieza        NVARCHAR(10),
    tipo         NVARCHAR(10),
    version      NVARCHAR(20),
    concatenar   NVARCHAR(300),
    cambio       NVARCHAR(100),
    created_at   DATETIME2     DEFAULT GETDATE(),
    updated_at   DATETIME2     DEFAULT GETDATE()
);
GO

-- ─── 2. MALLAS PEQUEÑAS (códigos enteros) ──────────────────────────────────────
CREATE TABLE mallas_pequenas (
    codigo       INT           NOT NULL PRIMARY KEY,
    cod_veh      NVARCHAR(20),
    descripcion  NVARCHAR(200),
    pieza        NVARCHAR(10),
    tipo         NVARCHAR(10),
    version      NVARCHAR(20),
    concatenar   NVARCHAR(300),
    part_number  NVARCHAR(30),
    cambio       NVARCHAR(100),
    created_at   DATETIME2     DEFAULT GETDATE(),
    updated_at   DATETIME2     DEFAULT GETDATE()
);
GO

-- ─── 3. VITROJET ──────────────────────────────────────────────────────────────
CREATE TABLE vitrojet (
    vitro          NVARCHAR(20)  NOT NULL PRIMARY KEY,
    codigo_malla   NVARCHAR(20)  NOT NULL,
    tipo_malla     NCHAR(1)      DEFAULT 'G',  -- 'G'=grandes, 'P'=pequeñas
    cod_completo   NVARCHAR(60),
    bnerig         NVARCHAR(10),
    vehiculo       NVARCHAR(200),
    version        NVARCHAR(20),
    ruta           NVARCHAR(500),
    cambio         NVARCHAR(100),
    created_at     DATETIME2     DEFAULT GETDATE(),
    updated_at     DATETIME2     DEFAULT GETDATE()
);
GO

-- ─── 4. PASTA DE PLATA ────────────────────────────────────────────────────────
CREATE TABLE pasta_plata (
    consecutivo   NVARCHAR(20)  NOT NULL PRIMARY KEY,
    tipo          NVARCHAR(10),
    vehiculo      NVARCHAR(200),
    cod_vehiculo  NVARCHAR(20),
    version       NVARCHAR(20),
    pieza         NVARCHAR(10),
    ruta_archivo  NVARCHAR(500),
    caso          NVARCHAR(20),
    cambio        NVARCHAR(100),
    created_at    DATETIME2     DEFAULT GETDATE(),
    updated_at    DATETIME2     DEFAULT GETDATE()
);
GO

-- ─── 5. GLASSJET VIEJO (histórico, solo lectura) ──────────────────────────────
CREATE TABLE glassjet_viejo (
    id                  INT           IDENTITY(1,1) PRIMARY KEY,
    malla               NVARCHAR(20),
    glassjet            NVARCHAR(20),
    part_number         NVARCHAR(30),
    tipo                NVARCHAR(10),
    vehiculo            NVARCHAR(200),
    homologacion_vitro  NVARCHAR(20)
);
GO

-- ─── 6. VINILOS ───────────────────────────────────────────────────────────────
CREATE TABLE vinilos (
    herramental   NVARCHAR(20)  NOT NULL PRIMARY KEY,
    vehiculo      NVARCHAR(200),
    cod_vehiculo  NVARCHAR(20),
    version       NVARCHAR(20),
    pieza         NVARCHAR(10),
    tipo          NVARCHAR(10),
    ruta          NVARCHAR(500),
    cambio        NVARCHAR(100),
    created_at    DATETIME2     DEFAULT GETDATE(),
    updated_at    DATETIME2     DEFAULT GETDATE()
);
GO

-- ─── Índices ──────────────────────────────────────────────────────────────────
CREATE INDEX idx_grandes_desc     ON mallas_grandes(descripcion);
CREATE INDEX idx_grandes_cod_veh  ON mallas_grandes(cod_veh);
CREATE INDEX idx_pequenas_desc    ON mallas_pequenas(descripcion);
CREATE INDEX idx_pequenas_cod_veh ON mallas_pequenas(cod_veh);
CREATE INDEX idx_vitrojet_malla   ON vitrojet(codigo_malla);
CREATE INDEX idx_vitrojet_veh     ON vitrojet(vehiculo);
CREATE INDEX idx_vinilos_veh      ON vinilos(vehiculo);
CREATE INDEX idx_pasta_veh        ON pasta_plata(vehiculo);
GO

PRINT 'Tablas e índices creados correctamente en Vitros_Mallas';
GO
