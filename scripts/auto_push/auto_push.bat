@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

echo ============================================================
echo  🚀 AUTO-PUSH: DAG Neumáticos (Batch Runner)
echo ============================================================

cd /d "%~dp0\..\.."

set "COMMIT_MSG=%~1"

if "%COMMIT_MSG%"=="" (
    set "DATE_STAMP=%DATE:~6,4%-%DATE:~3,2%-%DATE:~0,2% %TIME:~0,8%"
    echo.
    set /p "USER_INPUT=📝 Ingrese mensaje de commit [Enter para autogenerar con IA/Fecha]: "
    if "!USER_INPUT!"=="" (
        rem Ejecutar PowerShell para generacion con IA si esta disponible
        powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0auto_push.ps1"
        goto end
    ) else (
        set "COMMIT_MSG=!USER_INPUT!"
    )
)

echo.
echo 📦 Preparando cambios con git add...
git add .

echo 🔍 Verificando estado de Git...
git status -s

echo 📝 Creando commit con mensaje: "%COMMIT_MSG%"
git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo.
    echo ℹ️ No habia cambios pendientes para commitear o la operacion se cancelo.
    goto end
)

echo.
echo 🚀 Intentando subir cambios a GitHub...
git push origin main 2>nul
if errorlevel 1 (
    git push origin master 2>nul
    if errorlevel 1 (
        git push 2>nul
        if errorlevel 1 (
            echo ❌ No se pudo realizar el push directamente. Verifique la conexion o el remoto configurado.
        ) else (
            echo ✅ Push realizado exitosamente!
        )
    ) else (
        echo ✅ Push realizado exitosamente a la rama master!
    )
) else (
    echo ✅ Push realizado exitosamente a la rama main!
)

:end
echo.
echo ============================================================
echo  Proceso de Auto-Push finalizado.
echo ============================================================
pause
