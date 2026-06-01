-- Ampliar columnas que causaron truncamiento
USE Vitros_Mallas;
GO

-- mallas_grandes
ALTER TABLE mallas_grandes ALTER COLUMN cod_veh      NVARCHAR(50);
ALTER TABLE mallas_grandes ALTER COLUMN version      NVARCHAR(50);
ALTER TABLE mallas_grandes ALTER COLUMN pieza        NVARCHAR(30);
ALTER TABLE mallas_grandes ALTER COLUMN tipo         NVARCHAR(20);
ALTER TABLE mallas_grandes ALTER COLUMN concatenar   NVARCHAR(500);
GO

-- mallas_pequenas
ALTER TABLE mallas_pequenas ALTER COLUMN cod_veh      NVARCHAR(50);
ALTER TABLE mallas_pequenas ALTER COLUMN version      NVARCHAR(50);
ALTER TABLE mallas_pequenas ALTER COLUMN pieza        NVARCHAR(30);
ALTER TABLE mallas_pequenas ALTER COLUMN tipo         NVARCHAR(20);
ALTER TABLE mallas_pequenas ALTER COLUMN part_number  NVARCHAR(60);
ALTER TABLE mallas_pequenas ALTER COLUMN concatenar   NVARCHAR(500);
GO

-- vitrojet
ALTER TABLE vitrojet ALTER COLUMN codigo_malla  NVARCHAR(50);
ALTER TABLE vitrojet ALTER COLUMN cod_completo  NVARCHAR(200);
ALTER TABLE vitrojet ALTER COLUMN bnerig        NVARCHAR(20);
ALTER TABLE vitrojet ALTER COLUMN version       NVARCHAR(50);
ALTER TABLE vitrojet ALTER COLUMN vehiculo      NVARCHAR(300);
GO

-- pasta_plata
ALTER TABLE pasta_plata ALTER COLUMN cod_vehiculo  NVARCHAR(50);
ALTER TABLE pasta_plata ALTER COLUMN version       NVARCHAR(50);
ALTER TABLE pasta_plata ALTER COLUMN pieza         NVARCHAR(30);
ALTER TABLE pasta_plata ALTER COLUMN tipo          NVARCHAR(20);
ALTER TABLE pasta_plata ALTER COLUMN vehiculo      NVARCHAR(300);
ALTER TABLE pasta_plata ALTER COLUMN caso          NVARCHAR(100);
GO

-- glassjet_viejo
ALTER TABLE glassjet_viejo ALTER COLUMN malla      NVARCHAR(50);
ALTER TABLE glassjet_viejo ALTER COLUMN glassjet   NVARCHAR(50);
ALTER TABLE glassjet_viejo ALTER COLUMN vehiculo   NVARCHAR(300);
ALTER TABLE glassjet_viejo ALTER COLUMN part_number NVARCHAR(60);
GO

PRINT 'Columnas ampliadas OK';
GO
