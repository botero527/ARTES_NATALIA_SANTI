-- Tabla para vincular DWGs con vitros y mallas encontrados
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='rutas_arte')
CREATE TABLE mallas.rutas_arte (
    id            INT           IDENTITY(1,1) PRIMARY KEY,
    ruta_dwg      NVARCHAR(600) NOT NULL,
    vehiculo      NVARCHAR(200) NULL,
    archivo       NVARCHAR(200) NULL,
    tipo_match    NVARCHAR(10)  NOT NULL,  -- 'VITRO', 'GRANDE', 'PEQUENA'
    codigo        NVARCHAR(50)  NOT NULL,  -- el codigo encontrado
    fecha_scan    DATETIME2     DEFAULT GETDATE(),
    CONSTRAINT uq_ruta_codigo UNIQUE (ruta_dwg, tipo_match, codigo)
);
GO

SELECT 'Tabla mallas.rutas_arte creada OK' AS resultado;
