@echo off
setlocal

REM ============================================================
REM PUBLICAR.BAT
REM Genera index.html y lo publica en GitHub
REM ============================================================

cd /d "%~dp0"

echo.
echo ============================================================
echo   ACTUALIZANDO FORMULARIO DE MANTENIMIENTO
echo ============================================================
echo.

echo [1/3] Generando index.html...
python generar_formulario.py

if errorlevel 1 (
    echo.
    echo ERROR: No se pudo generar index.html.
    echo Revisa el mensaje anterior.
    pause
    exit /b 1
)

if not exist "index.html" (
    echo.
    echo ERROR: No se encontro index.html despues de ejecutar Python.
    pause
    exit /b 1
)

echo.
echo [2/3] Preparando cambios para GitHub...
git add index.html

echo.
echo [3/3] Publicando en GitHub...
git commit -m "Actualiza formulario de mantenimiento"

if errorlevel 1 (
    echo.
    echo No hay cambios nuevos en index.html.
    echo GitHub ya tiene la version actual.
) else (
    git push origin main

    if errorlevel 1 (
        echo.
        echo ERROR: No se pudo subir el cambio a GitHub.
        pause
        exit /b 1
    )

    echo.
    echo ============================================
    echo   ACTUALIZACION ENVIADA A GITHUB
    echo ============================================
)

echo.
echo Pagina:
echo https://procesosflexigrip-hash.github.io/mtto_reportes/
echo.
echo GitHub Pages puede tardar unos minutos en actualizarse.
echo.

pause