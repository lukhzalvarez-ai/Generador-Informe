@echo off
setlocal
title Reporte Tecnico Astara
color 0C
cd /d "%~dp0"

echo ============================================
echo   REPORTE TECNICO ASTARA
echo   Preparando entorno local...
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] No se encontro Python instalado en este PC.
    echo Instala Python desde https://www.python.org/downloads/ y vuelve a intentar.
    pause
    exit /b 1
)

if not exist "venv" (
    echo Creando entorno virtual "venv"...
    python -m venv venv
)

call venv\Scripts\activate.bat

echo Instalando/actualizando dependencias necesarias...
pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet

echo.
echo ============================================
echo   Iniciando la aplicacion...
echo   Se abrira automaticamente en tu navegador
echo   en http://localhost:8501
echo   (deja esta ventana abierta mientras la usas)
echo ============================================
echo.

streamlit run app.py

echo.
echo La aplicacion se cerro. Presiona una tecla para salir.
pause >nul
