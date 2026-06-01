-- AGP Glass DB Schema
-- Motor: SQLite (compatible con MariaDB con ajustes menores)

PRAGMA journal_mode=WAL;
PRAGMA foreign_keys=ON;

-- ─── 1. MALLAS GRANDES (códigos A-XXXXX) ─────────────────────────────────────
CREATE TABLE IF NOT EXISTS mallas_grandes (
    codigo      TEXT    PRIMARY KEY,   -- A-7375, A-13723, etc.
    cod_veh     TEXT,
    descripcion TEXT,
    pieza       TEXT,
    tipo        TEXT,
    version     TEXT,
    concatenar  TEXT,                  -- columna real, guardada físicamente
    cambio      TEXT,
    created_at  DATETIME DEFAULT (datetime('now','localtime')),
    updated_at  DATETIME DEFAULT (datetime('now','localtime'))
);

-- ─── 2. MALLAS PEQUEÑAS (códigos enteros) ─────────────────────────────────────
CREATE TABLE IF NOT EXISTS mallas_pequenas (
    codigo      INTEGER PRIMARY KEY,
    cod_veh     TEXT,
    descripcion TEXT,
    pieza       TEXT,
    tipo        TEXT,
    version     TEXT,
    concatenar  TEXT,
    part_number TEXT,
    cambio      TEXT,
    created_at  DATETIME DEFAULT (datetime('now','localtime')),
    updated_at  DATETIME DEFAULT (datetime('now','localtime'))
);

-- ─── 3. VITROJET ──────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS vitrojet (
    vitro           TEXT    PRIMARY KEY,   -- T-8285, T-29887
    codigo_malla    TEXT    NOT NULL,      -- FK → mallas_grandes o mallas_pequenas
    tipo_malla      TEXT    DEFAULT 'G',   -- 'G' = grandes, 'P' = pequeñas
    cod_completo    TEXT,                  -- 1867 V-000 008 (veh + versión + pieza)
    bnerig          TEXT,                  -- BN / BNI
    vehiculo        TEXT,
    version         TEXT,
    ruta            TEXT,
    cambio          TEXT,
    created_at      DATETIME DEFAULT (datetime('now','localtime')),
    updated_at      DATETIME DEFAULT (datetime('now','localtime'))
);

-- ─── 4. PASTA DE PLATA ────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS pasta_plata (
    consecutivo TEXT    PRIMARY KEY,   -- S-00001
    tipo        TEXT,                  -- RED / ANT
    vehiculo    TEXT,
    cod_vehiculo TEXT,
    version     TEXT,
    pieza       TEXT,
    ruta_archivo TEXT,
    caso        TEXT,                  -- CASO 1, CASO 2
    cambio      TEXT,
    created_at  DATETIME DEFAULT (datetime('now','localtime')),
    updated_at  DATETIME DEFAULT (datetime('now','localtime'))
);

-- ─── 5. GLASSJET VIEJO (histórico, solo lectura) ──────────────────────────────
CREATE TABLE IF NOT EXISTS glassjet_viejo (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    malla               TEXT,
    glassjet            TEXT,
    part_number         TEXT,
    tipo                TEXT,
    vehiculo            TEXT,
    homologacion_vitro  TEXT
);

-- ─── 6. VINILOS ───────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS vinilos (
    herramental  TEXT    PRIMARY KEY,   -- VC-0001
    vehiculo     TEXT,
    cod_vehiculo TEXT,
    version      TEXT,
    pieza        TEXT,
    tipo         TEXT,                  -- BN / BNI (BN2 → BN al importar)
    ruta         TEXT,
    cambio       TEXT,
    created_at   DATETIME DEFAULT (datetime('now','localtime')),
    updated_at   DATETIME DEFAULT (datetime('now','localtime'))
);

-- ─── Índices para búsquedas rápidas ──────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_grandes_desc     ON mallas_grandes(descripcion);
CREATE INDEX IF NOT EXISTS idx_grandes_cod_veh  ON mallas_grandes(cod_veh);
CREATE INDEX IF NOT EXISTS idx_pequenas_desc    ON mallas_pequenas(descripcion);
CREATE INDEX IF NOT EXISTS idx_pequenas_cod_veh ON mallas_pequenas(cod_veh);
CREATE INDEX IF NOT EXISTS idx_vitrojet_malla   ON vitrojet(codigo_malla);
CREATE INDEX IF NOT EXISTS idx_vitrojet_veh     ON vitrojet(vehiculo);
CREATE INDEX IF NOT EXISTS idx_vinilos_veh      ON vinilos(vehiculo);
CREATE INDEX IF NOT EXISTS idx_pasta_veh        ON pasta_plata(vehiculo);
