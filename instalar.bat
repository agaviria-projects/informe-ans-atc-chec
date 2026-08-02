@echo off
setlocal

cd /d "%~dp0"

echo.
echo ==========================================================
echo INSTALACION DEL GENERADOR DE INFORMES ANS
echo ==========================================================
echo.

where py >nul 2>&1

if errorlevel 1 (
    echo ERROR: Python no se encuentra instalado o no esta
    echo disponible en las variables del sistema.
    echo.
    pause
    exit /b 1
)

if not exist "requirements.txt" (
    echo ERROR: No se encontro requirements.txt.
    echo.
    pause
    exit /b 1
)

if not exist "venv\Scripts\python.exe" (
    echo Creando entorno virtual...
    py -m venv venv

    if errorlevel 1 (
        echo.
        echo ERROR: No fue posible crear el entorno virtual.
        echo.
        pause
        exit /b 1
    )
)

echo Actualizando pip...
"venv\Scripts\python.exe" -m pip install --upgrade pip

if errorlevel 1 (
    echo.
    echo ERROR: No fue posible actualizar pip.
    echo.
    pause
    exit /b 1
)

echo Instalando dependencias...
"venv\Scripts\python.exe" -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ERROR: No fue posible instalar las dependencias.
    echo Revise la conexion a internet y requirements.txt.
    echo.
    pause
    exit /b 1
)

echo.
echo ==========================================================
echo INSTALACION FINALIZADA CORRECTAMENTE
echo ==========================================================
echo.
echo Ya puede ejecutar iniciar.bat.
echo.
pause

endlocal