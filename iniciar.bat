@echo off
setlocal

cd /d "%~dp0"

if not exist "venv\Scripts\python.exe" (
    echo.
    echo ==========================================================
    echo ERROR: No se encontro el entorno virtual del proyecto.
    echo ==========================================================
    echo.
    echo Ejecute primero:
    echo.
    echo     python -m venv venv
    echo     venv\Scripts\activate
    echo     pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

"venv\Scripts\python.exe" main.py

if errorlevel 1 (
    echo.
    echo La aplicacion finalizo con errores.
    echo Revise la carpeta logs.
    echo.
    pause
)

endlocal