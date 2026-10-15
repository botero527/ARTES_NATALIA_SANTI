@echo off
chcp 65001 > nul
title Compilar AGP Glass (HTML)
cd /d "%~dp0"

echo.
echo  ================================================
echo    Compilar AGP Glass Suite — versión HTML
echo  ================================================
echo.

py -3.14 -m pip install -r requirements.txt pyinstaller --quiet

rem workpath fuera de OneDrive: compilar dentro de OneDrive da "Acceso denegado"
py -3.14 -m PyInstaller AGP_Glass_Web.spec --noconfirm --workpath "%TEMP%\agp_glass_build" --distpath dist_web

echo.
echo  ================================================
echo    Listo: dist_web\AGP_Glass\AGP_Glass.exe
echo    Copiar la carpeta AGP_Glass completa a la red.
echo  ================================================
pause
