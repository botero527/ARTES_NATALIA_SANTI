@echo off
cd /d "%~dp0.."
echo Iniciando AGP Glass DB...
echo Abre tu navegador en: http://localhost:8000
.venv\Scripts\python -m uvicorn db_app.main:app --host 0.0.0.0 --port 8080 --reload
pause
