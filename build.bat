@echo off
echo Creando carpetas necesarias...
if not exist "installer" mkdir installer

echo Compilando aplicación...
python build_installer.py

echo Creando instalador...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" inno_setup_script.iss

echo Proceso completado!
pause


