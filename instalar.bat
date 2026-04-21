@echo off
title AGP Arte Maker — Instalador
echo.
echo  ============================================
echo   AGP GROUP — Instalador Arte Maker
echo  ============================================
echo.

REM Verificar que Python este instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python no esta instalado.
    echo  Descargalo en: https://www.python.org/downloads/
    echo  Marca la opcion "Add Python to PATH" al instalar.
    pause
    exit /b 1
)

echo  Python encontrado. Instalando dependencias...
echo.
python -m pip install --upgrade pip
python -m pip install pywin32 openpyxl

echo.
echo  Ejecutando post-instalacion de pywin32...
python -m pywin32_postinstall -install 2>nul || echo  (ya instalado)

echo.
echo  ============================================
echo   Instalacion completada correctamente.
echo   Ahora puedes usar Arte Maker y Comprobar Arte.
echo  ============================================
echo.
pause
