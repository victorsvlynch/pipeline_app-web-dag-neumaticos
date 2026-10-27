@echo off
chcp 65001 >nul
echo ========================================================
echo   DETENIENDO CONTENEDOR STREAMLIT (DOCKER)
echo ========================================================
echo.
cd /d "%~dp0"
docker compose down
echo.
echo ========================================================
echo   Contenedor detenido correctamente.
echo ========================================================
pause
