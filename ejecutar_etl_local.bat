@echo off
chcp 65001 >nul
echo ========================================================
echo   EJECUTANDO ETL LOCALMENTE (SIN DOCKER)
echo ========================================================
echo.
cd /d "%~dp0"
python scripts\etl_maestro.py
echo.
pause
