@echo off
rem AGP Glass Suite — versión HTML (pywebview). La versión vieja sigue en INICIAR.bat
cd /d "%~dp0"
py -3.14 agp_web.py %*
if errorlevel 1 pause
