@echo off
chcp 65001 >nul
echo ========================================================
echo   INICIANDO APLICACION STREAMLIT - DAG NEUMATICOS
echo ========================================================
echo.
cd /d "%~dp0"
echo Abriendo servidor Streamlit en http://localhost:8502 ...
echo.
python -m streamlit run app.py --server.port 8502 --server.headless false
pause
