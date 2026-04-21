@echo off
title Generando ejecutable Arte Maker...
cd /d "%~dp0"

echo Instalando PyInstaller...
python -m pip install pyinstaller

echo.
echo Generando arte_maker.exe ...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "Arte Maker AGP" ^
    --add-data "autocad_ops.py;." ^
    --add-data "config.py;." ^
    --add-data "verificacion.py;." ^
    arte_maker.py

echo.
echo El ejecutable quedo en:  dist\Arte Maker AGP.exe
echo Copialo junto a:
echo   - autocad_ops.py
echo   - config.py
echo   - LAYERS Y CAJETINES 1.dwg  (o .3dm)
echo   - arte_script.py  (para Rhino, aparte)
echo.
pause
