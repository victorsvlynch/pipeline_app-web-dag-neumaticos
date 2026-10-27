@echo off
chcp 65001 >nul
echo ========================================================
echo   MOSTRANDO LOGS DEL CONTENEDOR STREAMLIT (DOCKER)
echo ========================================================
echo.
cd /d "%~dp0"
docker compose logs -f
pause
