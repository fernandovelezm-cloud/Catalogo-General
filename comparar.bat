@echo off
chcp 65001 >nul
cd /d "%~dp0"

python -c "import PIL" >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencia Pillow...
    pip install Pillow
)

python comparar_y_validar.py

echo.
pause
