-- Agregar columna rol a la tabla de usuarios
-- Ejecutar una sola vez

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('MALLAS.APP_USUARIOS') AND name = 'rol'
)
    ALTER TABLE MALLAS.APP_USUARIOS ADD rol NVARCHAR(50) NULL;
GO

-- Asignar rol admin a los que tienen es_admin=1
UPDATE MALLAS.APP_USUARIOS SET rol = 'admin' WHERE es_admin = 1 AND rol IS NULL;
GO

-- Verificar
SELECT nombre, usuario, rol, estatus, es_admin FROM MALLAS.APP_USUARIOS ORDER BY nombre;
GO
