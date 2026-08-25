@echo off
chcp 65001 > nul
echo.
echo ============================================================
echo    Automatizacion Flowcards SEO - El Tiempo en Marfeel
echo ============================================================
echo.

cd /d "%~dp0"

REM Verificar si Python esta disponible
python --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado. Instala Python 3.10+
    pause
    exit /b 1
)

REM Verificar si Playwright esta instalado
python -c "import playwright" > nul 2>&1
if errorlevel 1 (
    echo [INSTALANDO] Playwright no encontrado. Instalando dependencias...
    pip install -r requirements_playwright.txt
    playwright install chromium
    echo.
)

REM Verificar sesion guardada
if not exist "marfeel_session.json" (
    echo [AVISO] No hay sesion guardada. Se abrira el navegador para login manual.
    echo.
    python flowcards_marfeel.py --login
    echo.
    echo Sesion guardada. Ahora ejecutando la automatizacion...
    echo.
)

echo Iniciando automatizacion...
echo.
python flowcards_marfeel.py %*

echo.
echo ============================================================
echo    Proceso terminado. Revisa flowcards_log.txt para detalles.
echo ============================================================
echo.
pause
