-- ================================================================
-- AGP Ingenieria — Rastreo de actividad (mallas.APP_SESIONES / APP_EVENTOS)
-- Servidor: agpcolombia.database.windows.net
-- Base de datos: AGP_Ingenieria
-- Ejecutar este script UNA SOLA VEZ
-- ================================================================
-- Notas de diseño:
--   * Todos los timestamps (inicio, fin, ultimo_latido, ultima_actividad,
--     momento) los calcula la APP en Python con datetime.now() — hora del
--     PC, que en las oficinas de AGP Colombia ya está en hora de Bogotá.
--     A propósito NO se usa GETDATE()/SYSDATETIME() aquí: Azure SQL corre
--     en UTC, y un reporte gerencial que diga "entró a las 3pm" cuando en
--     realidad fue a las 10am no le sirve a nadie. (Sí se sigue usando
--     GETDATE() en el resto de mallas.* para created_at/updated_at — ahí
--     lo que importa es el orden relativo, no la hora exacta que ve un
--     gerente, así que no hace falta tocar esas tablas.)
--   * Una sesión = una vez que alguien abre la app y hace login. Mientras
--     esa sesión sigue abierta, la app manda un "latido" cada pocos
--     minutos (ultimo_latido) y anota la última vez que el usuario hizo
--     clic o tecleó de verdad (ultima_actividad) — así se puede diferenciar
--     "la dejó abierta sin usarla" de "estuvo trabajando de verdad".
--   * Si el proceso se cae o lo matan sin cerrar bien (crash, corte de luz),
--     `fin` queda NULL. Al siguiente arranque de la app, se cierra sola con
--     el último latido como hora de cierre (ver limpiar_sesiones_huerfanas
--     en db_app/tracking.py) — igual de conservador que el limpiador de
--     PENDIENTE huérfanos que ya existe para vitro/malla.
-- ================================================================

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_SCHEMA = 'mallas' AND TABLE_NAME = 'APP_SESIONES'
)
CREATE TABLE mallas.APP_SESIONES (
    id                INT IDENTITY(1,1) PRIMARY KEY,
    usuario           NVARCHAR(200) NOT NULL,   -- correo de login
    nombre            NVARCHAR(200) NULL,       -- nombre completo (foto del momento del login)
    rol               NVARCHAR(50)  NULL,
    equipo            NVARCHAR(100) NULL,       -- hostname del PC donde abrió
    version_app       NVARCHAR(20)  NULL,       -- versión del .exe (ej. 'V5')
    inicio            DATETIME2     NOT NULL,   -- hora Bogotá (Python) del login
    fin               DATETIME2     NULL,       -- hora Bogotá del cierre — NULL mientras sigue abierta
    ultimo_latido     DATETIME2     NOT NULL,   -- "sigo viva" — se actualiza cada ~4 min
    ultima_actividad  DATETIME2     NULL,       -- último clic/tecla real del usuario
    duracion_seg      INT           NULL,       -- (fin - inicio) en segundos, calculado al cerrar
    cierre_tipo       NVARCHAR(20)  NULL        -- 'normal' | 'huerfana' (se cerró sola en el siguiente arranque)
);
GO

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_SCHEMA = 'mallas' AND TABLE_NAME = 'APP_EVENTOS'
)
CREATE TABLE mallas.APP_EVENTOS (
    id            BIGINT IDENTITY(1,1) PRIMARY KEY,
    sesion_id     INT           NULL,           -- NULL = evento antes de tener sesión (ej. login fallido)
    usuario       NVARCHAR(200) NOT NULL,
    momento       DATETIME2     NOT NULL,       -- hora Bogotá (Python)
    pestana       NVARCHAR(50)  NULL,           -- "Crear Arte", "Scanner", etc.
    accion        NVARCHAR(80)  NOT NULL,       -- "login" | "cambio_pestana" | "crear_arte" | "buscar_bd" | ...
    detalle       NVARCHAR(500) NULL,           -- contexto libre (código, orden, término buscado, error...)
    duracion_ms   INT           NULL,           -- para acciones con duración medible (ej. cuánto tardó el pipeline AutoCAD)
    exito         BIT           NULL,           -- 1 = ok, 0 = falló, NULL = no aplica
    CONSTRAINT FK_APP_EVENTOS_SESION FOREIGN KEY (sesion_id)
        REFERENCES mallas.APP_SESIONES(id)
);
GO

-- Índices para que los reportes gerenciales no tengan que escanear todo
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_APP_EVENTOS_SESION')
    CREATE INDEX IX_APP_EVENTOS_SESION  ON mallas.APP_EVENTOS(sesion_id);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_APP_EVENTOS_MOMENTO')
    CREATE INDEX IX_APP_EVENTOS_MOMENTO ON mallas.APP_EVENTOS(momento);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_APP_SESIONES_USUARIO')
    CREATE INDEX IX_APP_SESIONES_USUARIO ON mallas.APP_SESIONES(usuario, inicio);
GO

-- ================================================================
-- Vista gerencial: actividad por usuario y día, lista para abrir en
-- Excel/Power BI directo contra la vista, sin escribir ningún JOIN.
-- ================================================================
CREATE OR ALTER VIEW mallas.VW_ACTIVIDAD_DIARIA AS
SELECT
    usuario,
    nombre,
    CAST(inicio AS DATE)                                             AS dia,
    COUNT(*)                                                          AS sesiones,
    MIN(inicio)                                                       AS primera_entrada,
    MAX(COALESCE(fin, ultimo_latido))                                 AS ultima_salida,
    ROUND(SUM(COALESCE(
        duracion_seg,
        DATEDIFF(SECOND, inicio, COALESCE(fin, ultimo_latido))
    )) / 60.0, 1)                                                     AS minutos_totales,
    SUM(CASE WHEN cierre_tipo = 'huerfana' THEN 1 ELSE 0 END)         AS sesiones_dejadas_abiertas
FROM mallas.APP_SESIONES
GROUP BY usuario, nombre, CAST(inicio AS DATE);
GO

-- Chequeo rápido: cuánto tiempo pasó cada quien en cada pestaña (usando el
-- rastro de eventos "cambio_pestana" — el siguiente evento de la MISMA
-- sesión marca cuándo se salió de esa pestaña).
--
-- SELECT e.usuario, e.pestana,
--        SUM(DATEDIFF(SECOND, e.momento,
--            LEAD(e.momento) OVER (PARTITION BY e.sesion_id ORDER BY e.momento))) / 60.0 AS minutos
-- FROM mallas.APP_EVENTOS e
-- WHERE e.accion = 'cambio_pestana'
-- GROUP BY e.usuario, e.pestana
-- ORDER BY minutos DESC;

-- Verificar creación
SELECT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_SCHEMA = 'mallas' AND TABLE_NAME IN ('APP_SESIONES', 'APP_EVENTOS');
