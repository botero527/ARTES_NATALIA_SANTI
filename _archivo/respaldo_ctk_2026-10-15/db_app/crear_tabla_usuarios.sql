-- ================================================================
-- PASO 1: Crear tabla de usuarios
-- Ejecutar una sola vez en Azure SQL (AGP_Ingenieria)
-- ================================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.tables t
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE s.name = 'MALLAS' AND t.name = 'APP_USUARIOS'
)
CREATE TABLE MALLAS.APP_USUARIOS (
    id             INT IDENTITY(1,1) PRIMARY KEY,
    nombre         NVARCHAR(200),
    usuario        NVARCHAR(200) NOT NULL,
    contrasenia    NVARCHAR(200) NOT NULL,
    estatus        TINYINT       NOT NULL DEFAULT 1,
    es_admin       BIT           NOT NULL DEFAULT 0,
    sp_item_id     INT,
    actualizado_en DATETIME2     DEFAULT SYSDATETIME()
);
GO

-- ================================================================
-- PASO 2: Stored procedure para UPSERT (lo llama Power Automate)
-- ================================================================

CREATE OR ALTER PROCEDURE MALLAS.SP_UPSERT_USUARIO
    @sp_item_id  INT,
    @nombre      NVARCHAR(200),
    @usuario     NVARCHAR(200),
    @contrasenia NVARCHAR(200),
    @estatus     TINYINT,
    @es_admin    BIT
AS
BEGIN
    SET NOCOUNT ON;

    IF EXISTS (SELECT 1 FROM MALLAS.APP_USUARIOS WHERE sp_item_id = @sp_item_id)
        UPDATE MALLAS.APP_USUARIOS SET
            nombre         = @nombre,
            usuario        = @usuario,
            contrasenia    = @contrasenia,
            estatus        = @estatus,
            es_admin       = @es_admin,
            actualizado_en = SYSDATETIME()
        WHERE sp_item_id = @sp_item_id;
    ELSE
        INSERT INTO MALLAS.APP_USUARIOS
            (nombre, usuario, contrasenia, estatus, es_admin, sp_item_id)
        VALUES
            (@nombre, @usuario, @contrasenia, @estatus, @es_admin, @sp_item_id);
END;
GO
