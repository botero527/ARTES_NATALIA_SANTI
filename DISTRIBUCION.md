# AGP Glass Suite — Guía de Distribución

## Requisitos en el PC que compila

```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller AGP_Glass.spec
```

El ejecutable queda en `dist\AGP_Glass\AGP_Glass.exe`.

## Distribución a otros PCs

1. Copiar toda la carpeta `dist\AGP_Glass\` a una ruta accesible (ej. red o USB).
2. En cada PC destino: ejecutar `AGP_Glass.exe` directamente — no requiere Python instalado.

## Variable de entorno AGP_EXCEL (IMPORTANTE)

Si en algún PC el Excel de mallas está en una ruta diferente al PC del desarrollador,
definir la variable de entorno antes de abrir la app:

```bat
set AGP_EXCEL=\\servidor\compartido\LISTADO DE MALLAS Y GLASSJET 2025.xlsx
AGP_Glass.exe
```

O en el INICIAR.bat agregar la línea `set AGP_EXCEL=...` antes de llamar al .exe.

## Puertos de red necesarios

| Destino                                | Puerto | Uso                   |
|----------------------------------------|--------|-----------------------|
| agpcolombia.database.windows.net       | 1433   | BD principal (Azure)  |
| agpcolsap.database.windows.net         | 1433   | SAP Azure             |
| 192.168.2.23                           | 1433   | SmartFactory (red local) |

## AutoCAD

- AutoCAD debe estar **abierto** antes de usar las funciones de Crear Arte / Extraer.
- La app se conecta via COM (`win32com.client.GetActiveObject("AutoCAD.Application")`).

## Startup multi-PC

Al iniciar, la app cancela automáticamente reservas PENDIENTE con más de **30 minutos**
de antigüedad (huérfanas de crashes). Las reservas activas de otros PCs (< 30 min) NO
se tocan.

---

## Versión HTML (v3.0, desde 2026-10-15)

Misma app con interfaz HTML (pywebview) en la familia visual del Panel AGP /
PIEZAS_PLANAS / Tickets. Ventana con marco nativo de Windows (minimizar,
maximizar/restaurar, cerrar, Aero Snap). La lógica de negocio no cambió.

| Archivo | Qué es |
|---|---|
| `agp_web.py` | Punto de entrada + puente Python ↔ HTML (clase `Api`) |
| `servicios.py` | Consultas / guardado / scanner sin UI (copiado de agp_app.py) |
| `db_app/conexiones_externas.py` | Comercial, SAP y Calendario (Scanner) |
| `ui_web/` | index.html, css/style.css, js/*.js (una por pestaña) |
| `AGP_Glass_Web.spec` / `COMPILAR_WEB.bat` | Compilación |
| `INICIAR_WEB.bat` | Correr desde código (`py -3.14 agp_web.py`) |

- Compilar: `COMPILAR_WEB.bat` → `dist_web\AGP_Glass\AGP_Glass.exe` (carpeta onedir).
- Publicar: copiar la carpeta `AGP_Glass` a
  `\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS\Users\ANDRES ING\AGP_Glass V8\`
  y cambiar el `target` de "artes" en `AGP_LAUNCHER\tools.json` a esa ruta.
  El SSO (`--sso-token`) funciona igual que en la V7.
- Requiere WebView2 (ya viene en Windows 10/11 con Edge).
- Preferencias por PC (tema, idioma, menú fijo, últimas rutas):
  `%LOCALAPPDATA%\AGP_Artes\prefs.json`.
- La versión CustomTkinter (`agp_app.py`, `AGP_Glass.spec`, `INICIAR.bat`) sigue
  intacta como respaldo; copia exacta en `_archivo/respaldo_ctk_2026-10-15/`.
