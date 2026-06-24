@echo off
color 0A
echo =========================================
echo   Despliegue Automatico de ATS a GitHub
echo =========================================
echo.

:: 1. Pedir la version
set /p version="Ingresa el numero de la nueva version (ej. 1.2): "

:: 2. Actualizar el archivo version.txt localmente
echo %version% > version.txt

:: 3. Subir el version.txt a GitHub
echo Subiendo version.txt al repositorio...
git add version.txt
git commit -m "Actualizando version a %version%"
git push origin main

:: 4. Compilar el ejecutable
echo.
echo Compilando el ejecutable con PyInstaller...
py -m PyInstaller --clean --onefile --windowed mini_ats.py

:: 5. Crear el Release en GitHub y subir el .exe
echo.
echo Subiendo el nuevo ejecutable a GitHub Releases...
gh release create v%version% dist\mini_ats.exe -R "TuUsuario/mi-ats-proyecto" -t "Version %version%" -n "Actualizacion automatica del sistema"

echo.
echo =========================================
echo ¡Despliegue de la version %version% completado!
echo =========================================
pause