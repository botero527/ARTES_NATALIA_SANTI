-- ================================================================
-- AGP Ingenieria — Tabla LOGOS_LASER (pestaña "Logos Laser")
-- Servidor: agpcolombia.database.windows.net
-- Base de datos: AGP_Ingenieria
-- Ejecutar este script UNA SOLA VEZ para crear la tabla
-- ================================================================
-- Notas de diseño:
--   * id: formato ML-0001, ML-0002... 4 digitos con ceros a la izquierda.
--     Cuando el numero pase de 9999 sigue con 5 digitos sin romperse
--     (mismo comportamiento que mallas.pasta_plata "S-00001" y
--     mallas.vinilos "VC-0001"). El consecutivo lo calcula la app con
--     un INSERT...SELECT...WITH (UPDLOCK, HOLDLOCK) — asi dos personas
--     insertando al mismo tiempo desde PCs distintos nunca chocan con
--     el mismo numero (Azure SQL bloquea la lectura hasta que el
--     INSERT anterior confirma). No necesita SEQUENCE aparte.
--   * dxf / sjf / validado / cargado_genesis / azure / app: casillas
--     de check (0/1) para marcar en que paso del proceso va cada logo.
--   * created_at / updated_at: automaticos, igual al resto de mallas.*
-- ================================================================

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_SCHEMA = 'mallas' AND TABLE_NAME = 'LOGOS_LASER'
)
CREATE TABLE mallas.LOGOS_LASER (
    id               NVARCHAR(10)  NOT NULL PRIMARY KEY,   -- ML-0001, ML-0002, ...
    cliente          NVARCHAR(150) NULL,
    producto         NVARCHAR(150) NULL,
    dxf              BIT           NOT NULL DEFAULT 0,
    sjf              BIT           NOT NULL DEFAULT 0,
    validado         BIT           NOT NULL DEFAULT 0,
    cargado_genesis  BIT           NOT NULL DEFAULT 0,
    azure            BIT           NOT NULL DEFAULT 0,
    app              BIT           NOT NULL DEFAULT 0,
    observaciones    NVARCHAR(500) NULL,
    responsable      NVARCHAR(100) NULL,                   -- quien creo / edito el registro
    created_at       DATETIME2     NOT NULL DEFAULT GETDATE(),
    updated_at       DATETIME2     NOT NULL DEFAULT GETDATE()
);
GO

-- Verificar creación
SELECT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_SCHEMA = 'mallas' AND TABLE_NAME = 'LOGOS_LASER';
