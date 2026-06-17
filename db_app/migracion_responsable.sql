-- Ejecutar UNA SOLA VEZ en Azure SQL (AGP_Ingenieria)
-- Agrega columnas responsable y updated_at a las tablas de mallas

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='vitrojet' AND COLUMN_NAME='responsable')
    ALTER TABLE mallas.vitrojet ADD responsable NVARCHAR(100) NULL;

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='vitrojet' AND COLUMN_NAME='updated_at')
    ALTER TABLE mallas.vitrojet ADD updated_at DATETIME NULL;

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='grandes' AND COLUMN_NAME='responsable')
    ALTER TABLE mallas.grandes ADD responsable NVARCHAR(100) NULL;

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='grandes' AND COLUMN_NAME='updated_at')
    ALTER TABLE mallas.grandes ADD updated_at DATETIME NULL;

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='pequenas' AND COLUMN_NAME='responsable')
    ALTER TABLE mallas.pequenas ADD responsable NVARCHAR(100) NULL;

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_SCHEMA='mallas' AND TABLE_NAME='pequenas' AND COLUMN_NAME='updated_at')
    ALTER TABLE mallas.pequenas ADD updated_at DATETIME NULL;

SELECT 'Migración completada OK' AS resultado;
