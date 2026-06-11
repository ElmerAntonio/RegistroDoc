@echo off
title RegistroDoc Pro - Iniciando
echo =======================================================================
echo              RegistroDoc Pro - Asistente de Arranque Rapido
echo =======================================================================
echo.

:: 1. Verificar si Python esta instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] No se detecto Python instalado en su sistema.
    echo.
    echo Para ejecutar este programa, necesita tener Python 3.10 o superior.
    echo Vamos a abrir la pagina oficial de descarga en su navegador.
    echo.
    echo Recuerde marcar la casilla "Add Python to PATH" durante la instalacion.
    echo.
    pause
    start https://www.python.org/downloads/
    exit /b
)

:: 2. Instalar dependencias si es necesario
echo Verificando e instalando componentes necesarios...
echo [INFO] Intentando instalacion offline desde la carpeta local "dependencies".
python -m pip install --no-index --find-links=dependencies -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo.
    echo [ADVERTENCIA] No se pudo completar la instalacion offline.
    echo Intentando descargar los componentes desde internet...
    python -m pip install -r requirements.txt
)

:: 3. Iniciar la aplicacion
echo.
echo Componentes verificados con exito. Iniciando RegistroDoc Pro...
echo Puede minimizar esta ventana. No la cierre mientras usa el programa.
echo.
python src/app.py
