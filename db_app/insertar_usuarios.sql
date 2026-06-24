-- ================================================================
-- Actualizar stored procedure para matchear por usuario (email)
-- Ejecutar primero
-- ================================================================
CREATE OR ALTER PROCEDURE MALLAS.SP_UPSERT_USUARIO
    @sp_item_id  INT,
    @nombre      NVARCHAR(200),
    @usuario     NVARCHAR(200),
    @contrasenia NVARCHAR(200),
    @estatus     TINYINT,
    @es_admin    BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS (SELECT 1 FROM MALLAS.APP_USUARIOS WHERE LOWER(usuario)=LOWER(@usuario))
        UPDATE MALLAS.APP_USUARIOS SET
            nombre         = @nombre,
            contrasenia    = @contrasenia,
            estatus        = @estatus,
            es_admin       = ISNULL(@es_admin, 0),
            sp_item_id     = @sp_item_id,
            actualizado_en = SYSDATETIME()
        WHERE LOWER(usuario)=LOWER(@usuario);
    ELSE
        INSERT INTO MALLAS.APP_USUARIOS
            (nombre, usuario, contrasenia, estatus, es_admin, sp_item_id)
        VALUES
            (@nombre, @usuario, @contrasenia, @estatus, ISNULL(@es_admin, 0), @sp_item_id);
END;
GO

-- ================================================================
-- Insertar todos los usuarios actuales
-- ================================================================
INSERT INTO MALLAS.APP_USUARIOS (nombre, usuario, contrasenia, estatus, es_admin) VALUES
('ANDRES BOTERO',                    'abotero@agpglass.com',          '1011099721',    1, 1),
('Administrador IT',                 'atcol@agpglass.com',            'AdminIng2025_It',1, 1),
('GUERRERO CABRERA FABIO DANILO',    'fguerrero@agpglass.com',        '1022438939',    1, 0),
('GUANUMEN HUERTO JOHNATAN',         'jguanumen@agpglass.com',        '1023005676',    1, 0),
('ACOSTA ALEXANDER',                 'alexander.acosta@agpglass.com', '93437119',      1, 0),
('DELGADO GERALDINE XIOMARA',        'g.delgado@agpglass.com',        '1031180571',    1, 0),
('MORALES MORENO KAREN STEFANIA',    'kmorales@agpglass.com',         '1233501014',    1, 0),
('PINZON JORGE',                     'jpinzon@agpglass.com',          '1030596420',    1, 0),
('LAURA PELAEZ',                     'lpelaez@agpglass.com',          '1000047853',    1, 0),
('MIGUEL BERNAL',                    'mbernal@agpglass.com',          '1000007660',    1, 0),
('NICOLAS ROJAS',                    'nirojas@agpglass.com',          '1030688452',    1, 0),
('STEVEN SUAREZ',                    'asuarez@agpglass.com',          '1030690990',    1, 0),
('DANIEL GRIMALDO',                  'dgrimaldo@agpglass.com',        '1000236441',    1, 0),
('NATALIA LEON',                     'nleon@agpglass.com',            '1137624222',    1, 0),
('JUAN SEBASTIAN RAMIREZ',           'jramirezf@agpglass.com',        '1031420151',    0, 0),
('PRACT1',                           'pract1@agpglass.com',           'PRACT_ING1',    1, 0),
('PRACT2',                           'pract2@agpglass.com',           'PRACT_ING2',    1, 0),
('PRACT3',                           'pract3@agpglass.com',           'PRACT_ING3',    1, 0),
('PRACT4',                           'pract4@agpglass.com',           'PRACT_ING4',    1, 0),
('SANTIAGO PINA',                    'spina@agpglass.com',            '1010236538',    1, 1),
('LAURA SOFIA CRUZ CASALLAS',        'lcruz@agpglass.com',            '1032937021',    0, 0),
('JUAN PABLO GALVIS JIMENEZ',        'jgalvis@agpglass.com',          '1032877183',    1, 0),
('SANTIAGO PIMENTEL',                'spimentel@agpglass.com',        '1000034924',    1, 0),
('PRACTICANTE INGENIERIA',           'practingenieria@agpglass.com',  '1000971646',    1, 0),
('JEFFERSON MAHECHA',                'jmahecha@agpglass.com',         '1019982163',    1, 0),
('DARWIN ALEJANDRO FORERO DIAZ',     'dforero@agpglass.com',          '1000256251',    1, 1),
('CESAR GARCIA VELOZA',              'cegarcia@agpglass.com',         '1001092159',    1, 0),
('FALLA CONTRERAS LOGAN DAVID',      'lfalla@agpglass.com',           '1022930033',    1, 0);
GO

-- Verificar
SELECT nombre, usuario, estatus, es_admin FROM MALLAS.APP_USUARIOS ORDER BY nombre;
GO
