@echo off
title Convertir JPEG a JPG
cd /d "%~dp0"
echo.
echo ============================================
echo   Renombrando .jpeg a .jpg
echo ============================================
echo.

set COUNT=0
set SKIPPED=0
for %%f in (*.jpeg) do (
    if exist "%%~nf.jpg" (
        echo Omitido: ya existe %%~nf.jpg, no se toca %%f
        set /a SKIPPED+=1
    ) else (
        ren "%%f" "%%~nf.jpg"
        echo Renombrado: %%f -^> %%~nf.jpg
        set /a COUNT+=1
    )
)

echo.
if %COUNT%==0 (
    echo No habia archivos .jpeg que convertir.
) else (
    echo Listo! %COUNT% archivo(s) renombrado(s).
)
if %SKIPPED% gtr 0 (
    echo %SKIPPED% archivo(s) omitido(s) porque ya existia un .jpg con ese nombre.
)
echo.
pause
