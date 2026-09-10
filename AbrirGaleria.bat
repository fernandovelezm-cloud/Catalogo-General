@echo off
python "C:\Users\ferna\Claude\Projects\Automatizaciones\convertir_png_a_jpg.py" "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ActualizarGaleria.ps1"
start "" "%~dp0Catalogo.html"
call "C:\Users\ferna\Claude\Projects\Automatizaciones\publicar_catalogo.bat"
