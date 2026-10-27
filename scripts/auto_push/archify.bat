@echo off
chcp 65001 > nul
setlocal

echo ============================================================
echo  📦 ARCHIFY: Empaquetado y Respaldos (Batch Runner)
echo ============================================================

set "MODE=%~1"
if "%MODE%"=="" set "MODE=Plantillas"

set "TAG=%~2"

echo Ejecutando Archify en modo: %MODE%...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0archify.ps1" -Mode %MODE% -Tag "%TAG%"

echo.
echo ============================================================
echo  Operacion finalizada.
echo ============================================================
pause
