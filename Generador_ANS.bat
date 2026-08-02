@echo off
setlocal

cd /d "%~dp0"

if not exist "main.py" (
    echo.
    echo ==========================================================
    echo ERROR: No se encontro el archivo main.py.
    echo ==========================================================
    echo.
    echo Verifique que iniciar.bat se encuentre en la carpeta
    echo principal del proyecto.
    echo.
    pause
    exit /b 1
)

if not exist "venv\Scripts\python.exe" (
    echo.
    echo ==========================================================
    echo ERROR: No se encontro el entorno virtual del proyecto.
    echo ==========================================================
    echo.
    echo Ejecute primero los siguientes comandos:
    echo.
    echo     py -m venv venv
    echo     venv\Scripts\python.exe -m pip install --upgrade pip
    echo     venv\Scripts\python.exe -m pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

"venv\Scripts\python.exe" "main.py"

if errorlevel 1 (
    echo.
    echo ==========================================================
    echo La aplicacion finalizo con errores.
    echo ==========================================================
    echo.
    echo Revise la carpeta logs para conocer el detalle.
    echo.
    pause
)

endlocal