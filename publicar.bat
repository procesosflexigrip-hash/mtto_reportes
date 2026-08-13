@echo off
REM ===================================================================
REM publicar.bat — Actualiza y publica el formulario de mantenimiento
REM ===================================================================
REM Qué hace, en orden:
REM   1) Corre generar_formulario.py (lee validacion.xlsx, arma el HTML)
REM   2) Copia el HTML generado a tu repositorio de GitHub como index.html
REM   3) Sube el cambio a GitHub (git add + commit + push)
REM   4) En 1-2 minutos, GitHub Pages queda actualizado solo
REM
REM CONFIGURACIÓN (hazlo una sola vez):
REM   Cambia la línea REPO_PATH de abajo por la ruta donde clonaste tu
REM   repositorio (la carpeta donde corriste "git clone ...").
REM   Ejemplo: C:\Users\Procesos\Documents\mantenimiento-flexigrip
REM ===================================================================

setlocal

set REPO_PATH=%USERPROFILE%\Desktop\mantenimiento\captura_mtto

REM No toques nada de aquí para abajo -------------------------------

cd /d "%~dp0"

echo.
echo === 1/3: Generando formulario desde validacion.xlsx ===
python generar_formulario.py
if errorlevel 1 (
    echo.
    echo ERROR: no se pudo generar el formulario. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

if not exist "%REPO_PATH%" (
    echo.
    echo ERROR: no existe la carpeta configurada en REPO_PATH:
    echo   %REPO_PATH%
    echo Edita este archivo publicar.bat y corrige esa ruta.
    pause
    exit /b 1
)

echo.
echo === 2/3: Copiando al repositorio ===
copy /Y captura_mantenimiento.html "%REPO_PATH%\index.html" >nul

echo.
echo === 3/3: Subiendo a GitHub ===
cd /d "%REPO_PATH%"
git add index.html
git commit -m "Actualiza formulario %date% %time%"
if errorlevel 1 (
    echo.
    echo (No había cambios nuevos que subir, o el commit falló — revisa arriba.)
) else (
    git push
)

echo.
echo ===================================================================
echo  Listo. Si hubo cambios, GitHub Pages se actualiza en 1-2 minutos.
echo ===================================================================
pause
