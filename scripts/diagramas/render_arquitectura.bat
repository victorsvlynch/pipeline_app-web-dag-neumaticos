@echo off
chcp 65001 > nul
setlocal

echo ============================================================
echo   ARCHIFY: Renderizador de Diagramas (Batch Runner)
echo ============================================================

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0render_arquitectura.ps1" %*

if "%~1"=="" pause
