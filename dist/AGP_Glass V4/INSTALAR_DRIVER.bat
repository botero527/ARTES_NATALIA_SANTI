@echo off
echo ============================================
echo  AGP Glass Suite - Instalador de requisitos
echo ============================================
echo.

:: Verificar si el driver ya esta instalado
reg query "HKLM\SOFTWARE\ODBC\ODBCINST.INI\ODBC Driver 17 for SQL Server" >nul 2>&1
if %errorlevel% == 0 (
    echo [OK] ODBC Driver 17 ya esta instalado.
    goto fin
)

reg query "HKLM\SOFTWARE\ODBC\ODBCINST.INI\ODBC Driver 18 for SQL Server" >nul 2>&1
if %errorlevel% == 0 (
    echo [OK] ODBC Driver 18 ya esta instalado.
    goto fin
)

echo [!] Falta el driver de base de datos. Descargando...
echo     (requiere internet, ~6 MB)
echo.
powershell -Command "Invoke-WebRequest -Uri 'https://aka.ms/odbc17' -OutFile '%TEMP%\odbc17.msi'"
if %errorlevel% neq 0 (
    echo [ERROR] No se pudo descargar. Instala manualmente:
    echo         https://aka.ms/odbc17
    pause
    exit /b 1
)

echo Instalando...
msiexec /i "%TEMP%\odbc17.msi" /quiet /norestart
echo [OK] Driver instalado correctamente.

:fin
echo.
echo Listo! Ahora puedes abrir AGP_Glass.exe
pause
