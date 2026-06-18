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
