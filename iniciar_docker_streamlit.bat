@echo off
chcp 65001 >nul
echo ========================================================
echo   INICIANDO STREAMLIT EN DOCKER - DAG NEUMATICOS
echo ========================================================
echo.
cd /d "%~dp0"
docker compose up -d --build
echo.
echo ========================================================
echo   Contenedor Streamlit iniciado correctamente.
echo   Abriendo navegador en: http://localhost:8502
echo ========================================================
start http://localhost:8502
pause
