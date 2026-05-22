@echo off
cd /d "%~dp0"
call ".venv\Scripts\activate.bat"
python crear_arte_acad.py
pause
