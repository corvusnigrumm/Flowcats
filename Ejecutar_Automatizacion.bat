@echo off
title Generador de Excel - Automatizacion Santamaria
color 0A

echo ====================================================================
echo      INICIANDO GENERADOR DE EXCEL (EL TIEMPO y PORTAFOLIO)
echo ====================================================================
echo.

:: Esto fuerza a la consola a ubicarse en la misma carpeta donde esta el .bat
cd /d "%~dp0"

:: Ejecuta el script de Python
python automatizacion_santamaria.py

echo.
echo ====================================================================
echo EL PROCESO HA FINALIZADO. REVISA LOS ARCHIVOS EXCEL GENERADOS.
echo ====================================================================
pause
