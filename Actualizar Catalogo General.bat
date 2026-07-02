@echo off
title Actualizando Catalogo General Fevel
cd /d "%~dp0"
echo.
echo ============================================
echo    CATALOGO GENERAL FEVEL - Actualizando
echo ============================================
echo.

:: [0] Convertir .jpeg a .jpg en ambas carpetas de fotos (reemplaza si ya existe .jpg)
set COMPLEMENTO=C:\Users\ferna\Dropbox\Fevel\Fotografía\Costos Ficha Técnica (complemento)
echo [0/3] Convirtiendo .jpeg a .jpg...
for %%f in ("%~dp0*.jpeg") do (
    if exist "%~dp0%%~nf.jpg" del /f /q "%~dp0%%~nf.jpg"
    ren "%%f" "%%~nf.jpg"
    echo   Renombrado: %%~nxf
)
pushd "%COMPLEMENTO%"
for %%f in (*.jpeg) do (
    if exist "%%~nf.jpg" del /f /q "%%~nf.jpg"
    ren "%%f" "%%~nf.jpg"
    echo   Renombrado: %%~nxf
)
popd
echo   Listo.
echo.

:: [1] Generar catalogo HTML
echo [1/3] Generando catalogo (descarga Mysoft + HTML)...
echo.
python "%~dp0generar_catalogo_general.py"

if %ERRORLEVEL% NEQ 0 (
  echo.
  echo ERROR: No se pudo generar el catalogo.
  echo Verifica que Python este instalado.
  echo.
  pause
  exit /b 1
)

:: [2] Publicar en GitHub
echo.
echo [2/3] Publicando en GitHub...
git --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
  echo AVISO: Git no esta instalado. Saltando publicacion en GitHub.
  goto :abrir
)

set /p GH_TOKEN=<"%~dp0github_token.txt"

:: Guardar token en Windows Credential Manager (no queda en git)
cmdkey /add:git:https://github.com /user:fernandovelezm-cloud /pass:%GH_TOKEN%

if exist ".git" rmdir /s /q ".git"
git init
git branch -M main
git remote add origin https://github.com/fernandovelezm-cloud/Catalogo-General.git
git config credential.helper manager
git config user.email "fernandovelezm@gmail.com"
git config user.name "Fernando Velez"
git add .
git reset HEAD github_token.txt 2>nul
git commit -m "Actualizar catalogo %date%"
git push -u origin main --force

:: Limpiar credencial temporal
cmdkey /delete:git:https://github.com

if %ERRORLEVEL% EQU 0 (
  echo Publicado en: https://fernandovelezm-cloud.github.io/Catalogo-Costos/
) else (
  echo AVISO: No se pudo publicar en GitHub. El catalogo local si se genero.
)

:: [3] Abrir catalogo local
:abrir
echo.
echo [3/3] Abriendo catalogo...
start "" "%~dp0Catalogo General.html"

echo.
echo ============================================
echo    Listo!
echo ============================================
echo.
pause
