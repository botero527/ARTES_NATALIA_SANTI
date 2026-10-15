# CLAUDE.md — macro_natalia (AGP Glass Suite)

## Quién es el usuario
Andrés Botero (spina@agpglass.com), AGP Glass. Trátalo bro a bro, directo, sin rodeos.
Le gusta: scripts simples, comentarios en español, pymssql/pyodbc, openpyxl. Cierra cada respuesta con un tip corto de programación (feedback_tips_programacion en su memoria global).

## Qué es este proyecto
"AGP Glass Suite" — app de escritorio unificada (CustomTkinter) para el área de ingeniería/moldes de AGP Glass Colombia. Antes eran scripts sueltos (crear arte AutoCAD, buscador BD, sync Excel); ahora todo vive en `agp_app.py` con pestañas.

Pestañas (AGPApp.PAGES):
- **Crear Arte** (Alt+1): pipeline AutoCAD — extraer plano, crear arte, buscar, todo en uno
- **Consultar BD** (Alt+2): buscador + stats + sync Excel (vitrojet/grandes/pequeñas/vinilos/pasta)
- **Gestión BD** (Alt+3): editar filas, separar vitro/malla, insertar pasta/vinilos, anular
- **Scanner** (Alt+4): escanear orden de producción → SmartFactory (ZFER) → SAP (vitro/malla/tipo pieza)

## Arquitectura
- **`agp_web.py` — app principal desde v3.0 (2026-10-15): interfaz HTML con pywebview**
  (`ui_web/`), lógica sin UI en `servicios.py`. Correr con `INICIAR_WEB.bat` (`py -3.14`;
  la `.venv` está rota). Compilar con `COMPILAR_WEB.bat` / `AGP_Glass_Web.spec`.
- `agp_app.py` — app CTk anterior, intacta como respaldo (no borrar hasta validar la web)
- `crear_arte_acad.py` — pipeline AutoCAD (COM), geometría de layers/degradé
- `autocad_ops.py` — AutoCADMotor (CoInitialize/CoUninitialize)
- `db_app/asignaciones.py` — motor vitro/malla (reservar/confirmar/cancelar), Azure SQL `mallas.*`
- `db_app/importar_excel.py` — sync Excel (OneDrive) → BD, UPSERT sin borrar
- BD: Azure SQL `AGP_Ingenieria`, schema `mallas.*` (vitrojet, grandes, pequenas, pasta_plata, vinilos)
- Scanner usa además SmartFactory (SQL Server 192.168.2.23, BD Comercial) y SAP Azure (`agpcolsap`, DB_COL_SAP)
- Distribución: PyInstaller (.spec bundleado), ver `DISTRIBUCION.md`

**Nunca pongas credenciales reales en este archivo ni en el código de ejemplo** — están en la memoria de Claude Code del proyecto (`project_scanner_ordenes.md`) y no deben repetirse en texto plano en el repo salvo donde ya existían históricamente.

## Memoria persistente
Toda la memoria detallada (estado exacto de cada módulo, bugs resueltos, schema completo, credenciales) vive en:
`C:\Users\abotero\.claude\projects\c--Users-abotero-OneDrive---AGP-GROUP-Documentos-macro-natalia\memory\`
- `MEMORY.md` — índice
- `project_autocad_arte.md` — pipeline AutoCAD, vitro/malla, schema BD, app principal
- `project_scanner_ordenes.md` — scanner, credenciales SmartFactory/SAP, `_q_pieza`
- `project_auditoria_artes.md` — módulo de auditoría de artes DWG (en construcción)
- `project_version_html.md` — rediseño HTML/pywebview (v3.0), qué se probó y qué falta

Respaldo espejo de esa memoria en:
`C:\Users\abotero\OneDrive - AGP GROUP\Documentos\_respaldo_memoria_claude\projects\c--Users-abotero-OneDrive---AGP-GROUP-Documentos-macro-natalia\memory\`

También hay un resumen en la bóveda Obsidian:
`C:\Users\abotero\OneDrive - AGP GROUP\Documentos\CEREBRO\proyectos\macro_natalia.md`

Lee esos archivos ANTES de asumir nada sobre el estado del proyecto — léelos primero en cualquier chat nuevo.

## Pendientes conocidos (al 2026-08-04)
- Módulo `auditoria_artes_completa` (auditoría DWG con update BD + reporte Excel) sigue pendiente de construir, ver `project_auditoria_artes.md`.
- Repo tiene carpetas `build/` y `dist/` con binarios de PyInstaller trackeados en git (cambian en cada compilación) — no son código fuente, ignorar si aparecen en diffs.
- Revisar si conviene un `.gitignore` para `build/`, `dist/`, `__pycache__/` (actualmente parecen estar versionados).
